#!/usr/bin/env python3
"""Aplica um plano JSON de edições a um .docx preservando o padrão do original.

Uso:
    python edit_docx.py entrada.docx plano.json -o saida.docx [--dry-run] [--report r.json]

O plano é tudo-ou-nada: qualquer erro cancela a gravação. A saída só é gravada
se a validação estrutural não encontrar erros novos em relação ao original.
Veja o formato do plano em references/plan-format.md.
"""
from __future__ import annotations

import argparse
import copy
import datetime as dt
import difflib
import json
import os
import posixpath
import re
import struct
import sys
import tempfile
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_lib import (  # noqa: E402
    A_NS, BARRIER, MARKER_TAGS, MC_ALT, MC_CHOICE, MC_FALLBACK, PIC_NS, PPR_ORDER, R_NS,
    RPR_ORDER, RT_IMAGE, TCPR_ORDER, TRPR_TAIL_ORDER, W14_PARA_ID, W_NS, WP_NS, XML_SPACE,
    DocxError, Package, ParagraphText, Styles, cell_vmerge, etree, field_depths,
    has_section_break, in_fallback, iter_paragraphs, iter_tables, local, ordered_insert,
    paragraph_style, regressions, row_cells_by_grid, rpr_signature, setup_stdio, validate, w,
    wval,
)

NSMAP_W = {"w": W_NS}
_ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
_TOKEN = re.compile(r"\s+|\w+|[^\w\s]", re.UNICODE)
_MARK_REVISIONS = ("ins", "del", "moveFrom", "moveTo")
_OBJECT_TAGS = {w("drawing"), w("pict"), w("object"), MC_ALT}


class PlanError(DocxError):
    pass


# ---------------------------------------------------------------------------
# Construção de runs
# ---------------------------------------------------------------------------
def clean_rpr(rpr):
    """Cópia de um rPr sem marcas de revisão (serve para run e marca de parágrafo)."""
    if rpr is None:
        return None
    new = copy.deepcopy(rpr)
    for tag in _MARK_REVISIONS + ("rPrChange",):
        for el in new.findall(w(tag)):
            new.remove(el)
    return new if len(new) else None


def make_run(text, rpr=None):
    """Cria um w:r; \\t vira w:tab e \\n vira w:br."""
    if _ILLEGAL_XML.search(text) or BARRIER in text:
        raise PlanError(f"Texto contém caractere proibido em XML: {text!r}")
    run = etree.Element(w("r"), nsmap=NSMAP_W)
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    for piece in re.split(r"(\t|\n)", text.replace("\r\n", "\n").replace("\r", "\n")):
        if piece == "\t":
            etree.SubElement(run, w("tab"))
        elif piece == "\n":
            etree.SubElement(run, w("br"))
        elif piece:
            t = etree.SubElement(run, w("t"))
            t.text = piece
            if piece != piece.strip() or "  " in piece:
                t.set(XML_SPACE, "preserve")
    return run


def _set_text(t, text):
    t.text = text
    if text != text.strip() or "  " in text:
        t.set(XML_SPACE, "preserve")


def _content(run):
    return [c for c in run if c.tag != w("rPr")]


def split_run(run, child, offset):
    """Divide `run` antes de `child`, ou no meio de seu texto se offset > 0."""
    new = etree.Element(run.tag, dict(run.attrib), nsmap=run.nsmap)
    rpr = run.find(w("rPr"))
    if rpr is not None:
        new.append(copy.deepcopy(rpr))
    content = _content(run)
    idx = content.index(child)
    if offset:
        tail = etree.SubElement(new, w("t"))
        _set_text(tail, child.text[offset:])
        _set_text(child, child.text[:offset])
        idx += 1
    for moved in content[idx:]:
        new.append(moved)
    run.addnext(new)
    return new


def split_at(p, pos, depth):
    """Garante que `pos` (posição no texto visível) caia entre dois runs."""
    pt = ParagraphText(p, depth)
    for atom in pt.atoms:
        if atom.length and atom.start <= pos < atom.start + atom.length:
            offset = pos - atom.start
            before = [a for a in pt.atoms if a.run is atom.run and a.length and a.start < atom.start]
            if offset or before:
                split_run(atom.run, atom.child, offset)
            return


def isolate_barriers(p, depth):
    """Separa em runs próprios os objetos que dividem run com texto.

    Assim remover "o texto do run" nunca leva junto uma imagem ou uma nota.
    """
    runs = []
    for atom in ParagraphText(p, depth).atoms:
        if atom.run is not None and (not runs or runs[-1] is not atom.run):
            runs.append(atom.run)
    for run in runs:
        atoms = [a for a in ParagraphText(p, depth).atoms if a.run is run]
        kinds = [a.text is None for a in atoms if a.text is None or a.length]
        if not (any(kinds) and not all(kinds)):
            continue
        for prev, cur in reversed(list(zip(atoms, atoms[1:]))):
            if (prev.text is None) != (cur.text is None):
                split_run(run, cur.child, 0)


def rpr_for_paragraph_text(p, depth=0):
    """rPr dominante do parágrafo (por caracteres); cai na marca de parágrafo."""
    weights = Counter()
    samples = {}
    for atom in ParagraphText(p, depth).atoms:
        if atom.length and not atom.in_field:
            sig = rpr_signature(atom.run)
            weights[sig] += atom.length
            samples.setdefault(sig, atom.run.find(w("rPr")))
    if weights:
        return clean_rpr(samples[weights.most_common(1)[0][0]])
    ppr = p.find(w("pPr"))
    if ppr is not None:
        return clean_rpr(ppr.find(w("rPr")))
    return None


def trim_common(old, new):
    """Comprimento do prefixo e do sufixo comuns, recuados até limite de palavra."""
    limit = min(len(old), len(new))
    pre = 0
    while pre < limit and old[pre] == new[pre]:
        pre += 1
    while pre > 0 and old[pre - 1].isalnum() and (
        (pre < len(old) and old[pre].isalnum()) or (pre < len(new) and new[pre].isalnum())
    ):
        pre -= 1
    limit -= pre
    suf = 0
    while suf < limit and old[len(old) - 1 - suf] == new[len(new) - 1 - suf]:
        suf += 1
    while suf > 0 and old[len(old) - suf].isalnum() and (
        (suf < len(old) - pre and old[len(old) - suf - 1].isalnum())
        or (suf < len(new) - pre and new[len(new) - suf - 1].isalnum())
    ):
        suf -= 1
    return pre, suf


def diff_spans(old, new):
    """Trocas mínimas por palavra: [(ini, fim, texto novo)], ou None se pouco similares."""
    a, b = _TOKEN.findall(old), _TOKEN.findall(new)
    matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
    if matcher.ratio() < 0.5:
        return None
    offsets = [0]
    for token in a:
        offsets.append(offsets[-1] + len(token))
    merged = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            # Um espaço isolado entre duas trocas não vale uma revisão separada.
            if merged and i2 - i1 == 1 and a[i1].isspace() and i2 < len(a):
                merged[-1][1] = i2
                merged[-1][3] = j2
            continue
        if merged and merged[-1][1] == i1:
            merged[-1][1] = i2
            merged[-1][3] = j2
        else:
            merged.append([i1, i2, j1, j2])
    return [(offsets[i1], offsets[i2], "".join(b[j1:j2])) for i1, i2, j1, j2 in merged]


# ---------------------------------------------------------------------------
# Sessão de edição
# ---------------------------------------------------------------------------
class PartView:
    """Foto da parte antes de qualquer edição: índices sempre se referem a ela."""

    def __init__(self, name, root):
        self.name = name
        self.root = root
        self.depths = field_depths(root)
        self.paragraphs = list(iter_paragraphs(root))
        self.texts = [ParagraphText(p, self.depths.get(p, 0)).text for p in self.paragraphs]
        self.tables = list(iter_tables(root))

    def alive(self, el) -> bool:
        # getroottree() devolve a raiz do documento mesmo para nó destacado; é preciso subir.
        top = el
        for top in el.iterancestors():
            pass
        return top is self.root


class Session:
    def __init__(self, pkg: Package, plan: dict):
        self.pkg = pkg
        self.plan = plan
        self.styles = Styles(pkg)
        self.views = {name: PartView(name, pkg.xml(name)) for name in pkg.story_parts()}
        self.warnings: list[str] = []
        self.log: list[dict] = []
        self.op_label = ""
        self.last_inserted = {}
        self.author = plan.get("author", "docx-editor")
        self.date = plan.get("date") or dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        mode = plan.get("track_changes", "auto")
        if mode not in ("auto", True, False):
            raise PlanError('track_changes deve ser "auto", true ou false')
        self.tracked = pkg.tracks_revisions() if mode == "auto" else mode
        self._next_id = self._max_annotation_id() + 1
        self._next_docpr = self._max_docpr() + 1
        self.footnotes_before = self._note_refs()
        self.image_refs_before = {n: self._rel_refs(n) for n in self.views}

    # ---- utilitários --------------------------------------------------------
    def warn(self, message):
        self.warnings.append(f"{self.op_label}: {message}")

    def _max_annotation_id(self):
        best = 0
        for name in self.pkg.names():
            if re.match(r"^word/[^/]+\.xml$", name):
                for el in self.pkg.xml(name).iter():
                    if isinstance(el.tag, str):
                        value = el.get(w("id"))
                        if value and value.lstrip("-").isdigit():
                            best = max(best, int(value))
        return best

    def _max_docpr(self):
        best = 0
        for view in self.views.values():
            for el in view.root.iter("{%s}docPr" % WP_NS):
                value = el.get("id", "0")
                if value.isdigit():
                    best = max(best, int(value))
        return best

    def next_docpr(self):
        self._next_docpr += 1
        return self._next_docpr - 1

    def revision(self, tag):
        el = etree.Element(
            w(tag),
            {w("id"): str(self._next_id), w("author"): self.author, w("date"): self.date},
            nsmap=NSMAP_W,
        )
        self._next_id += 1
        return el

    def view(self, name=None) -> PartView:
        name = name or self.pkg.main_part
        if name not in self.views:
            raise PlanError(f"Parte de texto inexistente: {name}. Disponíveis: {sorted(self.views)}")
        return self.views[name]

    def touch(self, view):
        self.pkg.touch(view.name)

    # ---- resolução de alvos -------------------------------------------------
    _TARGET_KEYS = {"part", "p", "expect", "contains", "equals", "style", "occurrence",
                    "bookmark", "para_id", "table", "row", "cell", "cell_p"}

    def resolve(self, spec, what="target"):
        """Devolve (view, parágrafo) a partir de um seletor."""
        if not isinstance(spec, dict):
            raise PlanError(f"{what} deve ser um objeto; recebido: {spec!r}")
        unknown = set(spec) - self._TARGET_KEYS
        if unknown:
            raise PlanError(f"{what}: chaves desconhecidas {sorted(unknown)}")
        view = self.view(spec.get("part"))
        p = self._locate(view, spec, what)
        if "expect" in spec:
            actual = view.texts[view.paragraphs.index(p)]
            if spec["expect"] not in actual:
                raise PlanError(
                    f"{what}: parágrafo não contém o texto esperado {spec['expect']!r}. "
                    f"Texto atual: {actual[:120]!r}. Rode o inspect de novo."
                )
        if not view.alive(p):
            raise PlanError(f"{what}: o parágrafo já foi removido por uma operação anterior do plano")
        return view, p

    def _locate(self, view, spec, what):
        if "p" in spec:
            idx = spec["p"]
            if not isinstance(idx, int) or not 0 <= idx < len(view.paragraphs):
                raise PlanError(f"{what}: índice p={idx!r} fora de 0..{len(view.paragraphs) - 1} em {view.name}")
            return view.paragraphs[idx]
        if "table" in spec:
            return self._locate_cell(view, spec, what)
        if "bookmark" in spec:
            for p in view.paragraphs:
                for mark in p.iter(w("bookmarkStart")):
                    if mark.get(w("name")) == spec["bookmark"]:
                        return p
            raise PlanError(f"{what}: indicador {spec['bookmark']!r} não encontrado em {view.name}")
        if "para_id" in spec:
            for p in view.paragraphs:
                if p.get(W14_PARA_ID) == spec["para_id"]:
                    return p
            raise PlanError(f"{what}: w14:paraId {spec['para_id']!r} não encontrado em {view.name}")
        return self._locate_by_text(view, spec, what)

    def _locate_cell(self, view, spec, what):
        tbl = self.resolve_table(view, spec["table"], what)
        rows = tbl.findall(w("tr"))
        r = spec.get("row", 0)
        if not 0 <= r < len(rows):
            raise PlanError(f"{what}: tabela {spec['table']} tem {len(rows)} linhas; row={r} inválido")
        cells = rows[r].findall(w("tc"))
        c = spec.get("cell", 0)
        if not 0 <= c < len(cells):
            raise PlanError(f"{what}: linha {r} tem {len(cells)} células; cell={c} inválido")
        paragraphs = cells[c].findall(w("p"))
        i = spec.get("cell_p", 0)
        if not 0 <= i < len(paragraphs):
            raise PlanError(f"{what}: célula tem {len(paragraphs)} parágrafos; cell_p={i} inválido")
        return paragraphs[i]

    def _locate_by_text(self, view, spec, what):
        if "contains" not in spec and "equals" not in spec and "style" not in spec:
            raise PlanError(f"{what}: informe p, contains, equals, style, bookmark, para_id ou table")
        style = spec.get("style")
        sid = None
        if style is not None:
            sid = self.styles.resolve(style)
            if sid is None:
                raise PlanError(f"{what}: estilo {style!r} não existe no documento")
        hits = []
        for idx, (p, text) in enumerate(zip(view.paragraphs, view.texts)):
            if "contains" in spec and spec["contains"] not in text:
                continue
            if "equals" in spec and spec["equals"] != text:
                continue
            if sid is not None and paragraph_style(p) != sid:
                continue
            hits.append(idx)
        if not hits:
            raise PlanError(f"{what}: nenhum parágrafo corresponde a {spec}")
        if "occurrence" in spec:
            n = spec["occurrence"]
            if not 0 <= n < len(hits):
                raise PlanError(f"{what}: occurrence={n}, mas há {len(hits)} correspondências: {hits}")
            return view.paragraphs[hits[n]]
        if len(hits) > 1:
            raise PlanError(
                f"{what}: {len(hits)} parágrafos correspondem (índices {hits[:15]}). "
                "Use p + expect, ou occurrence."
            )
        return view.paragraphs[hits[0]]

    def resolve_table(self, view, idx, what="table"):
        if not isinstance(idx, int) or not 0 <= idx < len(view.tables):
            raise PlanError(f"{what}: tabela {idx!r} inexistente; {view.name} tem {len(view.tables)}")
        tbl = view.tables[idx]
        if not view.alive(tbl):
            raise PlanError(f"{what}: a tabela {idx} já foi removida por uma operação anterior")
        return tbl

    # ---- gêmeo em mc:Fallback ----------------------------------------------
    def twin(self, p):
        """Parágrafo equivalente em mc:Fallback, para caixas de texto.

        O Word mostra mc:Choice e o LibreOffice/Google Docs podem mostrar
        mc:Fallback; editar só um lado deixa os aplicativos em desacordo.
        """
        choice = next((a for a in p.iterancestors() if a.tag == MC_CHOICE), None)
        if choice is None:
            return None
        fallback = choice.getparent().find(MC_FALLBACK)
        if fallback is None:
            return None
        ours = list(choice.iter(w("p")))
        theirs = list(fallback.iter(w("p")))
        if len(ours) != len(theirs):
            self.warn("caixa de texto com mc:Fallback divergente; apenas mc:Choice foi editado")
            return None
        return theirs[ours.index(p)]

    # ---- pós-processamento --------------------------------------------------
    def _note_refs(self):
        refs = {"footnoteReference": set(), "endnoteReference": set()}
        for view in self.views.values():
            for el in view.root.iter(w("footnoteReference"), w("endnoteReference")):
                refs[local(el)].add(el.get(w("id")))
        return refs

    def _rel_refs(self, part):
        root = self.pkg.xml(part)
        used = set()
        prefix = "{%s}" % R_NS
        for el in root.iter():
            if isinstance(el.tag, str):
                used.update(v for k, v in el.attrib.items() if k.startswith(prefix))
        return used

    def finish(self):
        """Remove notas e imagens cujo único uso foi apagado pelo plano.

        Sem isso o texto "apagado" continuaria dentro do arquivo, invisível.
        Só atinge o que deixou de ser referenciado nesta edição.
        """
        after = self._note_refs()
        for kind, part, tag in (("footnoteReference", "word/footnotes.xml", "footnote"),
                                ("endnoteReference", "word/endnotes.xml", "endnote")):
            gone = self.footnotes_before[kind] - after[kind]
            if gone and self.pkg.has(part):
                root = self.pkg.xml(part)
                for el in root.findall(w(tag)):
                    if el.get(w("id")) in gone:
                        root.remove(el)
                self.pkg.touch(part)
        for part, before in self.image_refs_before.items():
            for rid in before - self._rel_refs(part):
                self._drop_image_rel(part, rid)

    def _drop_image_rel(self, part, rid):
        rels_name = self.pkg.rels_name(part)
        if not self.pkg.has(rels_name):
            return
        root = self.pkg.xml(rels_name)
        for rel in list(root):
            if rel.get("Id") == rid and rel.get("Type") == RT_IMAGE and rel.get("TargetMode") != "External":
                target = self.pkg.rel_target(part, rel)
                root.remove(rel)
                self.pkg.touch(rels_name)
                if not self._target_in_use(target):
                    self.pkg.remove_part(target)
                    self.pkg.remove_override(target)

    def _target_in_use(self, target):
        for name in self.pkg.names():
            if name.endswith(".rels"):
                base = name.replace("_rels/", "", 1)[:-5] if name != "_rels/.rels" else ""
                for rel in self.pkg.xml(name):
                    if rel.get("TargetMode") != "External" and self.pkg.rel_target(base, rel) == target:
                        return True
        return False


# ---------------------------------------------------------------------------
# Carga útil de texto
# ---------------------------------------------------------------------------
_RUN_KEYS = {"text", "bold", "italic", "underline", "rstyle", "like"}


def build_payload(session, spec, base_rpr):
    """Converte `text` ou `runs` em lista de w:r."""
    if isinstance(spec, str):
        return [make_run(spec, base_rpr)] if spec else []
    if not isinstance(spec, list):
        raise PlanError("Conteúdo deve ser string (text) ou lista de runs (runs)")
    out = []
    for item in spec:
        if isinstance(item, str):
            item = {"text": item}
        unknown = set(item) - _RUN_KEYS
        if unknown or "text" not in item:
            raise PlanError(f"Run inválido {item}: chaves aceitas são {sorted(_RUN_KEYS)}")
        rpr = copy.deepcopy(base_rpr) if base_rpr is not None else None
        if "like" in item:
            rpr = _rpr_from_sample(session, item["like"])
        rpr = _apply_run_overrides(session, rpr, item)
        if item["text"]:
            out.append(make_run(item["text"], rpr))
    return out


def _rpr_from_sample(session, like):
    """Copia a formatação de um trecho que já existe no documento."""
    if not isinstance(like, dict) or "text" not in like:
        raise PlanError('run.like deve ter {"text": trecho, ...seletor do parágrafo}')
    selector = {k: v for k, v in like.items() if k != "text"}
    view, p = session.resolve(selector or {"contains": like["text"], "occurrence": 0}, "run.like")
    for atom in ParagraphText(p, view.depths.get(p, 0)).atoms:
        if atom.length and like["text"] in atom.text:
            return clean_rpr(atom.run.find(w("rPr")))
    raise PlanError(f"run.like: nenhum run do parágrafo contém {like['text']!r} inteiro")


def _apply_run_overrides(session, rpr, item):
    def ensure():
        nonlocal rpr
        if rpr is None:
            rpr = etree.Element(w("rPr"), nsmap=NSMAP_W)
        return rpr

    if "rstyle" in item:
        sid = session.styles.resolve(item["rstyle"], "character")
        if sid is None:
            raise PlanError(f"Estilo de caractere {item['rstyle']!r} não existe no documento")
        target = ensure()
        for old in target.findall(w("rStyle")):
            target.remove(old)
        el = etree.Element(w("rStyle"), nsmap=NSMAP_W)
        el.set(w("val"), sid)
        ordered_insert(target, el, RPR_ORDER)
    for key, tag in (("bold", "b"), ("italic", "i")):
        if key in item:
            target = ensure()
            for name in (tag, tag + "Cs"):
                for old in target.findall(w(name)):
                    target.remove(old)
                el = etree.Element(w(name), nsmap=NSMAP_W)
                if not item[key]:
                    el.set(w("val"), "0")
                ordered_insert(target, el, RPR_ORDER)
    if "underline" in item:
        target = ensure()
        for old in target.findall(w("u")):
            target.remove(old)
        el = etree.Element(w("u"), nsmap=NSMAP_W)
        el.set(w("val"), "single" if item["underline"] else "none")
        ordered_insert(target, el, RPR_ORDER)
    return rpr


# ---------------------------------------------------------------------------
# Núcleo: troca de um intervalo do texto visível
# ---------------------------------------------------------------------------
_REFUSE_CROSSING = {w(x) for x in ("sdtContent", "fldSimple", "ins", "moveTo", "customXml")}
_REMOVE_WHEN_EMPTY = {w(x) for x in ("hyperlink", "smartTag", "ins", "moveTo")}


def to_deleted(run):
    for child in run:
        if child.tag == w("t"):
            child.tag = w("delText")
            if (child.text or "") != (child.text or "").strip():
                child.set(XML_SPACE, "preserve")
        elif child.tag == w("instrText"):
            child.tag = w("delInstrText")


def wrap_deleted(session, runs):
    """Envolve runs consecutivos em w:del."""
    last_del = None
    prev = None
    for run in runs:
        to_deleted(run)
        if last_del is not None and prev is not None and prev.getparent() is last_del \
                and last_del.getnext() is run:
            last_del.append(run)
        else:
            last_del = session.revision("del")
            run.addprevious(last_del)
            last_del.append(run)
        prev = run
    return last_del


def _check_content_controls(session, runs):
    for run in runs:
        for anc in run.iterancestors(w("sdt")):
            props = anc.find(w("sdtPr"))
            if props is None:
                continue
            if props.find(w("dataBinding")) is not None:
                session.warn(
                    "o trecho está em controle de conteúdo vinculado a dados (w:dataBinding); "
                    "o Word pode restaurar o valor vinculado ao abrir"
                )
            placeholder = props.find(w("showingPlcHdr"))
            if placeholder is not None:
                props.remove(placeholder)


def replace_span(session, view, p, start, end, payload_spec, rpr_mode="first", allow_field=False):
    """Troca [start, end) do texto visível por `payload_spec`, com ou sem revisão."""
    depth = view.depths.get(p, 0)
    isolate_barriers(p, depth)
    split_at(p, end, depth)
    split_at(p, start, depth)
    pt = ParagraphText(p, depth)
    atoms = pt.text_atoms_in(start, end)
    if any(a.in_field for a in atoms) and not allow_field:
        raise PlanError(
            "o trecho está em resultado de campo (sumário, referência, numeração de página...). "
            "O Word recalcula esse texto; edite a origem ou use allow_field_result: true"
        )
    pieces = []
    for atom in atoms:
        if not pieces or pieces[-1] is not atom.run:
            pieces.append(atom.run)
    parents = []
    for run in pieces:
        if not any(run.getparent() is seen for seen in parents):
            parents.append(run.getparent())
    if len(parents) > 1 and any(par.tag in _REFUSE_CROSSING for par in parents):
        raise PlanError(
            "o trecho atravessa o limite de um controle de conteúdo, campo ou revisão existente. "
            "Divida a edição em partes, uma de cada lado do limite"
        )
    if session.tracked and any(par.tag in (w("ins"), w("moveTo")) for par in parents):
        raise PlanError(
            "o trecho está dentro de revisão pendente de outra edição. Peça ao usuário para "
            "aceitar/rejeitar essa revisão, ou edite o XML manualmente (unpack/pack)"
        )
    if not session.tracked and any(par.tag == w("ins") for par in parents):
        session.warn("o trecho estava em revisão pendente (w:ins) e foi editado sem nova revisão")
    if len(parents) > 1:
        session.warn("o trecho atravessava um hiperlink; confira se o link continua correto")
    _check_content_controls(session, pieces)

    if rpr_mode == "dominant":
        base_rpr = rpr_for_paragraph_text(p, depth)
    elif pieces:
        base_rpr = clean_rpr(pieces[0].find(w("rPr")))
    else:
        base_rpr = _rpr_at_insertion(pt, start, p)
    new_runs = build_payload(session, payload_spec, base_rpr)
    anchor_after = None
    anchor_before = None
    if not pieces:
        anchor_after, anchor_before = _insertion_point(pt, start, p)
        if session.tracked:
            anchor_after, anchor_before = _outside_revision(anchor_after, anchor_before)
    elif session.tracked:
        anchor_after = wrap_deleted(session, pieces)
    else:
        anchor_before = pieces[0]

    if new_runs:
        group = new_runs
        if session.tracked:
            ins = session.revision("ins")
            ins.extend(new_runs)
            group = [ins]
        cursor = anchor_after
        for el in group:
            if cursor is not None:
                cursor.addnext(el)
                cursor = el
            elif anchor_before is not None:
                anchor_before.addprevious(el)
            else:
                ppr = p.find(w("pPr"))
                p.insert(1 if ppr is not None else 0, el)
                cursor = el
    if not session.tracked:
        for run in pieces:
            parent = run.getparent()
            parent.remove(run)
            if parent.tag in _REMOVE_WHEN_EMPTY and len(parent) == 0:
                parent.getparent().remove(parent)
            elif parent.tag == w("sdtContent") and parent.find(w("r")) is None and len(parent) == 0:
                etree.SubElement(parent, w("r"))
    session.touch(view)


def _rpr_at_insertion(pt, pos, p):
    previous = [a for a in pt.atoms if a.length and a.start < pos and not a.in_field]
    if previous:
        return clean_rpr(previous[-1].run.find(w("rPr")))
    following = [a for a in pt.atoms if a.length and a.start >= pos and not a.in_field]
    if following:
        return clean_rpr(following[0].run.find(w("rPr")))
    ppr = p.find(w("pPr"))
    return clean_rpr(ppr.find(w("rPr"))) if ppr is not None else None


def _outside_revision(after, before):
    """Sobe até sair de w:ins/w:moveTo: revisões não devem ser aninhadas."""
    def climb(el, edge):
        while el is not None and el.getparent().tag in (w("ins"), w("moveTo")):
            if getattr(el, edge)() is not None:
                raise PlanError("a inserção cairia no meio de uma revisão pendente; aceite/rejeite "
                                "essa revisão ou edite o XML manualmente (unpack/pack)")
            el = el.getparent()
        return el
    return climb(after, "getnext"), climb(before, "getprevious")


def _insertion_point(pt, pos, p):
    """(depois de, antes de) para uma inserção pura na posição `pos`."""
    previous = [a for a in pt.atoms if a.length and a.start < pos]
    if previous:
        return previous[-1].run, None
    following = [a for a in pt.atoms if a.length and a.start >= pos]
    if following:
        return None, following[0].run
    return None, None


# ---------------------------------------------------------------------------
# Operação: replace_text
# ---------------------------------------------------------------------------
def _normalize_loose(text):
    table = {" ": " ", "’": "'", "‘": "'", "“": '"', "”": '"',
             "–": "-", "—": "-", "‑": "-"}
    return "".join(table.get(c, c) for c in text)


def op_replace_text(session, op):
    find = op["find"]
    if not isinstance(find, str) or not find:
        raise PlanError("find deve ser um texto não vazio")
    if "replace" in op and "runs" in op:
        raise PlanError("use replace ou runs, não ambos")
    if "replace" not in op and "runs" not in op:
        raise PlanError("informe replace (texto, pode ser vazio para apagar) ou runs")
    allow_field = op.get("allow_field_result", False)
    scope = _replace_scope(session, op)
    matches = []
    loose_hint = None
    barrier_hint = None
    for view, p in scope:
        pt = ParagraphText(p, view.depths.get(p, 0))
        searchable, mapping = pt.searchable
        pos = searchable.find(find)
        found_here = False
        while pos != -1:
            matches.append((view, p, mapping[pos], mapping[pos + len(find) - 1] + 1))
            found_here = True
            pos = searchable.find(find, pos + len(find))
        if not found_here:
            if find in pt.text:
                barrier_hint = pt.text[:80]
            elif _normalize_loose(find) in _normalize_loose(pt.text) and loose_hint is None:
                at = _normalize_loose(pt.text).index(_normalize_loose(find))
                loose_hint = pt.text[at:at + len(find)]
    if not matches:
        if barrier_hint is not None:
            raise PlanError(
                f"{find!r} só existe atravessando um objeto (imagem, nota, campo, quebra) no "
                f"parágrafo {barrier_hint!r}. Edite as partes de cada lado do objeto."
            )
        if loose_hint is not None:
            raise PlanError(f"{find!r} não encontrado. O documento traz {loose_hint!r} "
                            "(espaço, aspas ou travessão diferentes).")
        raise PlanError(f"{find!r} não encontrado no escopo")
    expected = op.get("count", 1)
    if expected != "all" and len(matches) != expected:
        where = [view.paragraphs.index(p) for view, p, *_ in matches if p in view.paragraphs]
        raise PlanError(
            f"{find!r} tem {len(matches)} ocorrência(s) (parágrafos {where[:15]}), mas count={expected}. "
            'Restrinja com target ou use count (número exato ou "all").'
        )
    in_field = 0
    for view, p, start, end in matches:
        if any(a.in_field for a in ParagraphText(p, view.depths.get(p, 0)).text_atoms_in(start, end)):
            in_field += 1
    if in_field and not allow_field:
        raise PlanError(
            f"{in_field} ocorrência(s) de {find!r} estão em resultado de campo (ex.: sumário). O Word "
            "recalcula esse texto ao atualizar campos. Restrinja com target ou passe allow_field_result: true."
        )
    for view, p, start, end in reversed(matches):
        _apply_replace(session, view, p, start, end, find, op, allow_field)
    session.log.append({"op": "replace_text", "find": find, "occurrences": len(matches)})


def _apply_replace(session, view, p, start, end, find, op, allow_field):
    twin = session.twin(p)
    twin_ok = twin is not None and ParagraphText(twin).text == ParagraphText(p, view.depths.get(p, 0)).text
    for para in ([p, twin] if twin_ok else [p]):
        if "replace" in op:
            new = op["replace"]
            pre, suf = trim_common(find, new)
            replace_span(session, view, para, start + pre, end - suf,
                         new[pre:len(new) - suf] if suf else new[pre:], allow_field=allow_field)
        else:
            replace_span(session, view, para, start, end, op["runs"], allow_field=allow_field)


def _replace_scope(session, op):
    if "target" in op:
        view, p = session.resolve(op["target"])
        return [(view, p)]
    parts = op.get("parts", "main")
    if parts == "main":
        names = [session.pkg.main_part]
    elif parts == "all":
        names = list(session.views)
    elif isinstance(parts, list):
        names = parts
    else:
        raise PlanError('parts deve ser "main", "all" ou lista de partes')
    scope = []
    for name in names:
        view = session.view(name)
        scope.extend((view, p) for p in view.paragraphs if view.alive(p))
    return scope


# ---------------------------------------------------------------------------
# Operação: rewrite_paragraph
# ---------------------------------------------------------------------------
def op_rewrite_paragraph(session, op):
    view, p = session.resolve(op["target"])
    if ("text" in op) == ("runs" in op):
        raise PlanError("informe text ou runs")
    twin = session.twin(p)
    if twin is not None and ParagraphText(twin).text != ParagraphText(p, view.depths.get(p, 0)).text:
        twin = None
    for para in ([p, twin] if twin is not None else [p]):
        _rewrite_one(session, view, para, op)
    session.log.append({"op": "rewrite_paragraph", "paragraph": _index_of(view, p)})


def _rewrite_one(session, view, p, op):
    depth = view.depths.get(p, 0)
    pt = ParagraphText(p, depth)
    has_field = any(a.in_field for a in pt.atoms if a.length)
    if "runs" in op:
        if has_field:
            raise PlanError("o parágrafo tem resultado de campo; use replace_text ao redor dele")
        replace_span(session, view, p, 0, len(pt.text), op["runs"], rpr_mode="dominant")
        return
    new = op["text"]
    if new == pt.text:
        session.warn("texto novo igual ao atual; nada a fazer")
        return
    spans = diff_spans(pt.text, new) if pt.text else None
    if spans is None and has_field:
        raise PlanError("o parágrafo tem resultado de campo e o texto novo é muito diferente; "
                        "use replace_text ao redor do campo")
    if spans is None:
        replace_span(session, view, p, 0, len(pt.text), new, rpr_mode="dominant")
        return
    for start, end, text in reversed(spans):
        replace_span(session, view, p, start, end, text)


def _index_of(view, p):
    return view.paragraphs.index(p) if p in view.paragraphs else None


# ---------------------------------------------------------------------------
# Novos parágrafos
# ---------------------------------------------------------------------------
def clone_ppr(model):
    """pPr de um modelo, sem sectPr: clonar sectPr criaria uma seção fantasma."""
    if model is None:
        return None
    ppr = model.find(w("pPr"))
    if ppr is None:
        return None
    new = copy.deepcopy(ppr)
    for tag in ("sectPr", "pPrChange"):
        for el in new.findall(w(tag)):
            new.remove(el)
    mark = new.find(w("rPr"))
    if mark is not None:
        for tag in _MARK_REVISIONS + ("rPrChange",):
            for el in mark.findall(w(tag)):
                mark.remove(el)
        if len(mark) == 0:
            new.remove(mark)
    return new if len(new) else None


def mark_paragraph(session, p, kind):
    ppr = p.find(w("pPr"))
    if ppr is None:
        ppr = etree.Element(w("pPr"), nsmap=NSMAP_W)
        p.insert(0, ppr)
    mark = ppr.find(w("rPr"))
    if mark is None:
        mark = etree.Element(w("rPr"), nsmap=NSMAP_W)
        ordered_insert(ppr, mark, PPR_ORDER)
    mark.insert(0, session.revision(kind))


def new_paragraph(session, model, payload_spec, depth=0, extra_runs=None):
    """Cria parágrafo com pPr e formatação de texto do modelo; nada mais é copiado."""
    p = etree.Element(w("p"), nsmap=NSMAP_W)
    ppr = clone_ppr(model)
    if ppr is not None:
        p.append(ppr)
    base = rpr_for_paragraph_text(model, depth) if model is not None else None
    runs = build_payload(session, payload_spec, base)
    if extra_runs:
        runs.extend(extra_runs)
    if session.tracked:
        mark_paragraph(session, p, "ins")
        if runs:
            ins = session.revision("ins")
            ins.extend(runs)
            p.append(ins)
    else:
        p.extend(runs)
    return p


def _model_from_like(session, like, anchor_view, anchor):
    """`like` aceita seletor comum ou {"style": ...}: exemplar mais próximo desse estilo."""
    if like is None:
        return anchor_view, anchor
    if set(like) - {"part"} == {"style"}:
        sid = session.styles.resolve(like["style"])
        if sid is None:
            names = sorted(session.styles.name_of(s) or s for s in session.styles.by_id
                           if session.styles.type_of(s) == "paragraph")
            raise PlanError(f"Estilo {like['style']!r} não existe. Estilos de parágrafo: {names[:40]}")
        view = session.view(like.get("part", anchor_view.name))
        candidates = [i for i, q in enumerate(view.paragraphs)
                      if paragraph_style(q) == sid and view.texts[i].strip() and view.alive(q)]
        if candidates:
            here = _index_of(view, anchor) if view is anchor_view else None
            if here is None:
                here = len(view.paragraphs)
            before = [i for i in candidates if i <= here]
            best = before[-1] if before else candidates[0]
            return view, view.paragraphs[best]
        synthetic = etree.Element(w("p"), nsmap=NSMAP_W)
        ppr = etree.SubElement(synthetic, w("pPr"))
        etree.SubElement(ppr, w("pStyle")).set(w("val"), sid)
        session.warn(f"estilo {like['style']!r} está definido mas sem uso no documento; confira o resultado")
        return view, synthetic
    return session.resolve(like, "like")


def _place(session, anchor, position, new):
    if position == "after":
        key = (anchor, "after")
        cursor = session.last_inserted.get(key, anchor)
        cursor.addnext(new)
        session.last_inserted[key] = new
    elif position == "before":
        anchor.addprevious(new)
    else:
        raise PlanError('position deve ser "after" ou "before"')


def op_insert_paragraph(session, op):
    view, anchor = session.resolve(op["anchor"], "anchor")
    position = op.get("position", "after")
    items = op.get("paragraphs")
    if items is None:
        if ("text" in op) == ("runs" in op):
            raise PlanError("informe text, runs ou paragraphs")
        items = [{k: op[k] for k in ("text", "runs", "like") if k in op}]
    elif "text" in op or "runs" in op:
        raise PlanError("use paragraphs ou text/runs, não ambos")
    for item in items:
        if isinstance(item, str):
            item = {"text": item}
        like = item.get("like", op.get("like"))
        model_view, model = _model_from_like(session, like, view, anchor)
        if has_section_break(anchor) and position == "after":
            session.warn("âncora termina uma seção: o parágrafo novo vai para a seção SEGUINTE. "
                         'Para ficar na mesma seção use position: "before".')
        spec = item.get("runs", item.get("text"))
        if spec is None:
            raise PlanError(f"item de parágrafo sem text/runs: {item}")
        depth = model_view.depths.get(model, 0)
        new = new_paragraph(session, model, spec, depth)
        _place(session, anchor, position, new)
        twin = session.twin(anchor)
        if twin is not None:
            twin_model = session.twin(model) if model is not anchor else twin
            twin_new = new_paragraph(session, twin_model if twin_model is not None else model, spec, depth)
            _place(session, twin, position, twin_new)
        session.touch(view)
    session.log.append({"op": "insert_paragraph", "count": len(items), "anchor": _index_of(view, anchor)})


# ---------------------------------------------------------------------------
# Remoção de parágrafos
# ---------------------------------------------------------------------------
def relocate_anchors(session, view, removed, source_paragraphs):
    """Move marcadores do trecho removido para o parágrafo mais próximo.

    Indicadores podem ser alvo de referências cruzadas e comentários têm
    conteúdo de terceiros; remover texto não autoriza perder nenhum dos dois.
    """
    markers = [el for el in removed.iter() if el.tag in MARKER_TAGS and not in_fallback(el)]
    comment_runs = [el.getparent() for el in removed.iter(w("commentReference")) if el.getparent() is not None]
    if not markers and not comment_runs:
        return
    removed_set = set(removed.iter(w("p")))
    all_ps = list(iter_paragraphs(view.root))
    order = {q: i for i, q in enumerate(all_ps)}
    in_tb = next(removed.iterancestors(w("txbxContent")), None) is not None
    usable = [q for q in all_ps if q not in removed_set
              and (next(q.iterancestors(w("txbxContent")), None) is not None) == in_tb]
    if not usable:
        raise PlanError("não há parágrafo vizinho para receber indicadores/comentários do trecho removido")
    position = order[source_paragraphs[-1]]
    following = [q for q in usable if order[q] > position]
    dest = following[0] if following else usable[-1]
    at_start = bool(following)
    ppr = dest.find(w("pPr"))
    index = 1 if ppr is not None else 0
    for el in markers + comment_runs:
        el.getparent().remove(el)
        if at_start:
            dest.insert(index, el)
            index += 1
        else:
            dest.append(el)
    named = [m.get(w("name")) for m in markers if m.tag == w("bookmarkStart")
             and not (m.get(w("name")) or "").startswith("_")]
    if named:
        session.warn(f"indicadores {named} foram preservados no parágrafo vizinho")
    if comment_runs:
        session.warn(f"{len(comment_runs)} comentário(s) ficaram ancorados no parágrafo vizinho")


def _field_spans_boundary(view, p):
    begin = end = 0
    for el in p.iter(w("fldChar")):
        if in_fallback(el):
            continue
        kind = el.get(w("fldCharType"))
        begin += kind == "begin"
        end += kind == "end"
    return begin != end or view.depths.get(p, 0) > 0


def _block_siblings(p):
    parent = p.getparent()
    return [c for c in parent if c is not p and isinstance(c.tag, str)
            and c.tag not in MARKER_TAGS and local(c) not in ("tcPr", "sectPr", "trPr", "sdtPr", "sdtEndPr")]


def _delete_one_paragraph(session, view, p, force):
    if has_section_break(p):
        raise PlanError(
            "o parágrafo carrega uma quebra de seção (w:sectPr). Removê-lo funde seções e altera "
            "cabeçalhos, margens e orientação. Apague o texto com rewrite_paragraph (text: \"\") "
            "ou confirme a fusão com o usuário e edite o XML"
        )
    if _field_spans_boundary(view, p):
        raise PlanError("o parágrafo faz parte de um campo que atravessa parágrafos (ex.: sumário); "
                        "não pode ser removido isoladamente")
    has_object = any(el.tag in _OBJECT_TAGS for el in p.iter())
    if has_object and not force:
        raise PlanError("o parágrafo contém imagem/objeto; confirme a remoção com force: true")
    if session.tracked:
        _delete_tracked_paragraph(session, view, p)
        return
    prev, nxt = p.getprevious(), p.getnext()
    siblings = _block_siblings(p)
    cell_needs_it = p.getparent().tag == w("tc") and _ends_container(p) \
        and prev is not None and prev.tag == w("tbl")
    if not siblings or cell_needs_it:
        pt = ParagraphText(p, view.depths.get(p, 0))
        if pt.text:
            replace_span(session, view, p, 0, len(pt.text), "")
        session.warn("único/último bloco do contêiner: parágrafo esvaziado em vez de removido")
        return
    if prev is not None and nxt is not None and prev.tag == w("tbl") and nxt.tag == w("tbl"):
        raise PlanError("o parágrafo separa duas tabelas; removê-lo faria o Word mesclá-las")
    if prev is not None and prev.tag == w("tbl") and (nxt is None or nxt.tag == w("sectPr")):
        raise PlanError("o parágrafo vem depois da última tabela; o Word exige um parágrafo ali")
    relocate_anchors(session, view, p, [p])
    parent = p.getparent()
    parent.remove(p)
    if parent.tag == w("sdtContent") and len(parent) == 0:
        etree.SubElement(parent, w("p"))


def _ends_container(p):
    return all(sib.tag in MARKER_TAGS or not isinstance(sib.tag, str) for sib in p.itersiblings())


def op_delete_paragraph(session, op):
    view, p = session.resolve(op["target"])
    force = op.get("force", False)
    twin = session.twin(p)
    idx = _index_of(view, p)
    _delete_one_paragraph(session, view, p, force)
    if twin is not None:
        _delete_one_paragraph(session, view, twin, force)
    session.touch(view)
    session.log.append({"op": "delete_paragraph", "paragraph": idx})


# ---------------------------------------------------------------------------
# Tabelas
# ---------------------------------------------------------------------------
def _resolve_row(session, op):
    view = session.view(op.get("part"))
    tbl = session.resolve_table(view, op["table"])
    rows = tbl.findall(w("tr"))
    r = op["row"]
    if not isinstance(r, int) or not 0 <= r < len(rows):
        raise PlanError(f"tabela {op['table']} tem {len(rows)} linhas; row={r!r} inválido")
    return view, tbl, rows, r


def _strip_identity(el):
    """Remove rsid/paraId: paraId repetido é erro; rsid em clone é só ruído."""
    for node in el.iter():
        if isinstance(node.tag, str):
            for attr in list(node.attrib):
                if attr == W14_PARA_ID or attr.endswith("}textId") or re.search(r"\}rsid\w*$", attr):
                    del node.attrib[attr]


def _cell_paragraphs(content):
    """Normaliza o conteúdo de célula: null, texto, {"runs": [...]} ou {"paragraphs": [...]}."""
    if content is None:
        return [""]
    if isinstance(content, str):
        return [content]
    if isinstance(content, dict) and set(content) == {"runs"}:
        return [content["runs"]]
    if isinstance(content, dict) and set(content) == {"paragraphs"} and content["paragraphs"]:
        out = []
        for item in content["paragraphs"]:
            out.extend(_cell_paragraphs(item))
        return out
    raise PlanError(f'Conteúdo de célula inválido: {content!r}. Use null, texto, {{"runs": [...]}} '
                    'ou {"paragraphs": [...]}')


def _fill_cell(session, tc, content):
    """Esvazia a célula e a preenche, herdando o 1º parágrafo como modelo."""
    model = next(tc.iter(w("p")), None)
    model = copy.deepcopy(model) if model is not None else None
    for child in list(tc):
        if child.tag != w("tcPr"):
            tc.remove(child)
    for item in _cell_paragraphs(content):
        tc.append(new_paragraph(session, model, item))


def clone_row(session, model, below, cells, keep_header=False):
    new = copy.deepcopy(model)
    _strip_identity(new)
    trpr = new.find(w("trPr"))
    if trpr is not None:
        for tag in ("ins", "del", "trPrChange") + (() if keep_header else ("tblHeader",)):
            for el in trpr.findall(w(tag)):
                trpr.remove(el)
        if len(trpr) == 0:
            new.remove(trpr)
    below_map = row_cells_by_grid(below) if below is not None else {}
    new_cells = new.findall(w("tc"))
    if len(new_cells) != len(cells):
        raise PlanError(f"a linha modelo tem {len(new_cells)} células; recebidas {len(cells)}. "
                        "Use null para células sem conteúdo (inclusive as mescladas).")
    merged_cols = {}
    for start, tc in row_cells_by_grid(new).items():
        merged_cols[tc] = start
    for tc, content in zip(new_cells, cells):
        tcpr = tc.find(w("tcPr"))
        continuing = False
        if tcpr is not None:
            for tag in ("cellIns", "cellDel", "cellMerge", "tcPrChange"):
                for el in tcpr.findall(w(tag)):
                    tcpr.remove(el)
            vm = tcpr.find(w("vMerge"))
            if vm is not None:
                tcpr.remove(vm)
                neighbor = below_map.get(merged_cols[tc])
                if neighbor is not None and cell_vmerge(neighbor) == "continue":
                    continuing = True
                    ordered_insert(tcpr, etree.Element(w("vMerge"), nsmap=NSMAP_W), TCPR_ORDER)
        if continuing and content not in (None, ""):
            raise PlanError("célula dentro de mesclagem vertical não aceita conteúdo; use null")
        _fill_cell(session, tc, content)
    if session.tracked:
        trpr = new.find(w("trPr"))
        if trpr is None:
            trpr = etree.Element(w("trPr"), nsmap=NSMAP_W)
            index = 0
            for i, child in enumerate(new):
                if child.tag in (w("tblPrEx"),):
                    index = i + 1
            new.insert(index, trpr)
        trpr.append(session.revision("ins"))
    return new


def op_insert_table_row(session, op):
    view, tbl, rows, r = _resolve_row(session, op)
    position = op.get("position", "after")
    if position not in ("after", "before"):
        raise PlanError('position deve ser "after" ou "before"')
    like = op.get("like_row", r)
    if not isinstance(like, int) or not 0 <= like < len(rows):
        raise PlanError(f"like_row={like!r} inválido")
    model = rows[like]
    if model.find(w("sdt")) is not None or model.find(w("customXml")) is not None:
        raise PlanError("linha modelo com células dentro de controle de conteúdo; edite o XML manualmente")
    trpr = model.find(w("trPr"))
    if trpr is not None and trpr.find(w("tblHeader")) is not None:
        session.warn("linha modelo é cabeçalho de tabela; informe like_row com uma linha de dados "
                     "se a nova linha não deve ter a formatação do cabeçalho")
    cells = op["cells"]
    if not cells or not isinstance(cells, list):
        raise PlanError("cells deve ser a lista de conteúdos da nova linha")
    anchor = rows[r]
    if position == "after":
        below = session.last_inserted.get((anchor, "after"), anchor).getnext()
    else:
        below = anchor
    while below is not None and below.tag != w("tr"):
        below = below.getnext()
    _place(session, anchor, position, clone_row(session, model, below, cells))
    session.touch(view)
    session.log.append({"op": "insert_table_row", "table": op["table"], "row": r})


def op_delete_table_row(session, op):
    view, tbl, rows, r = _resolve_row(session, op)
    if len(rows) == 1:
        raise PlanError("é a única linha da tabela; use delete_table")
    row = rows[r]
    if any(el.tag in _OBJECT_TAGS for el in row.iter()) and not op.get("force"):
        raise PlanError("a linha contém imagem/objeto; confirme a remoção com force: true")
    if any(_field_spans_boundary(view, p) for p in row.iter(w("p"))):
        raise PlanError("a linha contém campo que atravessa parágrafos; não pode ser removida isoladamente")
    trpr = row.find(w("trPr"))
    if trpr is not None and trpr.find(w("tblHeader")) is not None:
        session.warn("a linha removida é cabeçalho de tabela")
    if session.tracked:
        _tracked_delete_row(session, view, row)
    else:
        _transfer_vmerge(row, rows[r + 1] if r + 1 < len(rows) else None)
        relocate_anchors(session, view, row, list(row.iter(w("p"))))
        tbl.remove(row)
    session.touch(view)
    session.log.append({"op": "delete_table_row", "table": op["table"], "row": r})


def _transfer_vmerge(row, next_row):
    """Se a linha removida inicia uma mesclagem, a próxima linha herda início e conteúdo."""
    if next_row is None:
        return
    below = row_cells_by_grid(next_row)
    for start, tc in row_cells_by_grid(row).items():
        if cell_vmerge(tc) != "restart":
            continue
        heir = below.get(start)
        if heir is None or cell_vmerge(heir) != "continue":
            continue
        heir.find(w("tcPr")).find(w("vMerge")).set(w("val"), "restart")
        for child in list(heir):
            if child.tag != w("tcPr"):
                heir.remove(child)
        for child in list(tc):
            if child.tag != w("tcPr"):
                heir.append(child)
        etree.SubElement(tc, w("p"))


def _tracked_delete_row(session, view, row):
    trpr = row.find(w("trPr"))
    if trpr is None:
        trpr = etree.Element(w("trPr"), nsmap=NSMAP_W)
        row.insert(0, trpr)
    for p in row.iter(w("p")):
        if not in_fallback(p):
            _delete_tracked_paragraph(session, view, p)
    mark = session.revision("del")
    existing = trpr.find(w("trPrChange"))
    if existing is not None:
        existing.addprevious(mark)
    else:
        trpr.append(mark)


def _delete_tracked_paragraph(session, view, p):
    pt = ParagraphText(p, view.depths.get(p, 0))
    runs = []
    for atom in pt.atoms:
        if atom.run is not None and atom.run not in runs and atom.child.tag != w("commentReference"):
            runs.append(atom.run)
    if any(r.getparent().tag in (w("ins"), w("moveTo")) for r in runs):
        raise PlanError("há revisão pendente no trecho; aceite/rejeite antes de excluir com revisão")
    if runs:
        wrap_deleted(session, runs)
    mark_paragraph(session, p, "del")


def op_delete_table(session, op):
    view = session.view(op.get("part"))
    tbl = session.resolve_table(view, op["table"])
    if any(el.tag in _OBJECT_TAGS for el in tbl.iter()) and not op.get("force"):
        raise PlanError("a tabela contém imagem/objeto; confirme a remoção com force: true")
    if any(_field_spans_boundary(view, p) for p in tbl.iter(w("p"))):
        raise PlanError("a tabela contém campo que atravessa parágrafos; não pode ser removida isoladamente")
    if session.tracked:
        for row in tbl.findall(w("tr")):
            _tracked_delete_row(session, view, row)
    else:
        prev, nxt = tbl.getprevious(), tbl.getnext()
        if prev is not None and nxt is not None and prev.tag == w("tbl") and nxt.tag == w("tbl"):
            raise PlanError("a remoção deixaria duas tabelas adjacentes, que o Word mesclaria")
        relocate_anchors(session, view, tbl, list(tbl.iter(w("p"))))
        parent = tbl.getparent()
        parent.remove(tbl)
        if parent.tag == w("body") and parent.find(w("p")) is None:
            sect = parent.find(w("sectPr"))
            empty = etree.Element(w("p"), nsmap=NSMAP_W)
            if sect is not None:
                sect.addprevious(empty)
            else:
                parent.append(empty)
    session.touch(view)
    session.log.append({"op": "delete_table", "table": op["table"]})


def op_insert_table(session, op):
    view, anchor = session.resolve(op["anchor"], "anchor")
    position = op.get("position", "after")
    model_view = session.view(op.get("like_part", view.name))
    model = session.resolve_table(model_view, op["like_table"], "like_table")
    rows_spec = op["rows"]
    if not rows_spec or not all(isinstance(r, list) for r in rows_spec):
        raise PlanError("rows deve ser uma lista de linhas (lista de listas)")
    model_rows = model.findall(w("tr"))
    if not model_rows:
        raise PlanError("tabela modelo sem linhas diretas")
    header_model = model_rows[0]
    body_model = model_rows[1] if len(model_rows) > 1 else model_rows[0]
    has_header = op.get("header", len(model_rows) > 1)
    new = etree.Element(w("tbl"), nsmap=NSMAP_W)
    tblpr = copy.deepcopy(model.find(w("tblPr")))
    for tag in ("tblPrChange", "tblCaption", "tblDescription"):
        for el in tblpr.findall(w(tag)):
            tblpr.remove(el)
    new.append(tblpr)
    new.append(copy.deepcopy(model.find(w("tblGrid"))))
    for i, cells in enumerate(rows_spec):
        is_header = has_header and i == 0
        source = header_model if is_header else body_model
        row = clone_row(session, source, None, cells, keep_header=is_header)
        new.append(row)
    _place(session, anchor, position, new)
    _guard_table_neighbors(session, new, anchor)
    session.touch(view)
    session.log.append({"op": "insert_table", "rows": len(rows_spec), "like_table": op["like_table"]})


def _guard_table_neighbors(session, tbl, anchor):
    """Tabelas coladas se fundem, e uma tabela precisa de parágrafo depois de si."""
    def spacer():
        empty = etree.Element(w("p"), nsmap=NSMAP_W)
        ppr = clone_ppr(anchor)
        if ppr is not None:
            for tag in ("numPr", "pageBreakBefore"):
                for el in ppr.findall(w(tag)):
                    ppr.remove(el)
            if len(ppr):
                empty.append(ppr)
        if session.tracked:
            mark_paragraph(session, empty, "ins")
        return empty

    nxt = tbl.getnext()
    if nxt is None or nxt.tag in (w("tbl"), w("sectPr")):
        tbl.addnext(spacer())
        session.warn("parágrafo vazio adicionado após a tabela (exigência do formato)")
    prev = tbl.getprevious()
    if prev is not None and prev.tag == w("tbl"):
        tbl.addprevious(spacer())
        session.warn("parágrafo vazio adicionado antes da tabela para evitar fusão com a anterior")


# ---------------------------------------------------------------------------
# Imagens
# ---------------------------------------------------------------------------
_IMAGE_TYPES = {"png": "image/png", "jpg": "image/jpeg", "gif": "image/gif"}
EMU_PER_CM = 360000
EMU_PER_PX = 9525
EMU_PER_TWIP = 635


def image_info(path):
    """(extensão, largura px, altura px) lendo só o cabeçalho do arquivo."""
    with open(path, "rb") as fh:
        data = fh.read()
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) > 24:
        width, height = struct.unpack(">II", data[16:24])
        return "png", width, height, data
    if data[:6] in (b"GIF87a", b"GIF89a"):
        width, height = struct.unpack("<HH", data[6:10])
        return "gif", width, height, data
    if data.startswith(b"\xff\xd8"):
        pos = 2
        while pos + 9 < len(data):
            if data[pos] != 0xFF:
                pos += 1
                continue
            marker = data[pos + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                height, width = struct.unpack(">HH", data[pos + 5:pos + 9])
                return "jpg", width, height, data
            pos += 2 + struct.unpack(">H", data[pos + 2:pos + 4])[0]
    raise PlanError(f"Imagem não suportada (use PNG, JPEG ou GIF): {path}")


def _section_text_width(view, anchor):
    """Largura útil, em EMU, da seção onde a âncora está."""
    if local(view.root) != "document":
        return None
    sect = None
    for el in view.root.iter(w("sectPr")):
        pos_el = el.getparent() if el.getparent().tag != w("body") else None
        if pos_el is None:
            sect = el
            break
        para = pos_el.getparent()
        if para.tag == w("p") and _is_at_or_after(view, para, anchor):
            sect = el
            break
    if sect is None:
        return None
    size, margins = sect.find(w("pgSz")), sect.find(w("pgMar"))
    if size is None or margins is None:
        return None
    try:
        twips = int(size.get(w("w"))) - int(margins.get(w("left"))) - int(margins.get(w("right")))
    except (TypeError, ValueError):
        return None
    return twips * EMU_PER_TWIP


def _is_at_or_after(view, para, anchor):
    ordered = list(iter_paragraphs(view.root))
    try:
        return ordered.index(para) >= ordered.index(anchor)
    except ValueError:
        return False


def _drawing(session, rid, cx, cy, name, alt):
    docpr = session.next_docpr()
    xml = f"""<w:drawing xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}" xmlns:r="{R_NS}">
<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:docPr id="{docpr}" name=""/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
<a:graphic><a:graphicData uri="{PIC_NS}"><pic:pic><pic:nvPicPr><pic:cNvPr id="{docpr}" name=""/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"""
    el = etree.fromstring(re.sub(r">\s+<", "><", xml))
    el.find("{%s}inline/{%s}docPr" % (WP_NS, WP_NS)).set("name", name)
    el.find("{%s}inline/{%s}docPr" % (WP_NS, WP_NS)).set("descr", alt)
    next(el.iter("{%s}cNvPr" % PIC_NS)).set("name", name)
    return el


def op_insert_image(session, op):
    view, anchor = session.resolve(op["anchor"], "anchor")
    path = op["image"]
    if not os.path.isfile(path):
        raise PlanError(f"Imagem não encontrada: {path}")
    if "alt" not in op:
        session.warn("imagem sem texto alternativo (alt); recomendável para acessibilidade")
    ext, px_w, px_h, data = image_info(path)
    if not px_w or not px_h:
        raise PlanError(f"Imagem com dimensões inválidas: {path}")
    ratio = px_h / px_w
    if "width_cm" in op:
        cx = int(op["width_cm"] * EMU_PER_CM)
        cy = int(op["height_cm"] * EMU_PER_CM) if "height_cm" in op else int(cx * ratio)
    elif "height_cm" in op:
        cy = int(op["height_cm"] * EMU_PER_CM)
        cx = int(cy / ratio)
    else:
        cx, cy = px_w * EMU_PER_PX, px_h * EMU_PER_PX
    limit = _section_text_width(view, anchor)
    if limit and cx > limit:
        cy = int(cy * limit / cx)
        cx = limit
        session.warn("imagem reduzida à largura útil da seção")
    n = 1
    while session.pkg.has(f"word/media/image{n}.{ext}"):
        n += 1
    media = f"word/media/image{n}.{ext}"
    session.pkg.add_part(media, data)
    session.pkg.ensure_default_content_type(ext, _IMAGE_TYPES[ext])
    rel_target = posixpath.relpath(media, posixpath.dirname(view.name))
    rid = session.pkg.add_relationship(view.name, RT_IMAGE, rel_target)
    drawing = _drawing(session, rid, cx, cy, f"Imagem {n}", op.get("alt", ""))
    run = etree.Element(w("r"), nsmap=NSMAP_W)
    run.append(drawing)
    model_view, model = _model_from_like(session, op.get("like"), view, anchor)
    new = new_paragraph(session, model, "", model_view.depths.get(model, 0), extra_runs=[run])
    spacing = new.find("%s/%s" % (w("pPr"), w("spacing")))
    if spacing is not None and spacing.get(w("lineRule")) == "exact":
        for attr in ("line", "lineRule"):
            spacing.attrib.pop(w(attr), None)
        session.warn("entrelinha exata removida do parágrafo da imagem, pois cortaria a figura")
    _place(session, anchor, op.get("position", "after"), new)
    session.touch(view)
    session.log.append({"op": "insert_image", "media": media, "cx": cx, "cy": cy})


# ---------------------------------------------------------------------------
# Execução
# ---------------------------------------------------------------------------
OPS = {
    "replace_text": (op_replace_text, {"find"}, {"replace", "runs", "target", "count", "parts", "allow_field_result"}),
    "rewrite_paragraph": (op_rewrite_paragraph, {"target"}, {"text", "runs"}),
    "insert_paragraph": (op_insert_paragraph, {"anchor"}, {"position", "text", "runs", "like", "paragraphs"}),
    "delete_paragraph": (op_delete_paragraph, {"target"}, {"force"}),
    "insert_table_row": (op_insert_table_row, {"table", "row", "cells"}, {"part", "position", "like_row"}),
    "delete_table_row": (op_delete_table_row, {"table", "row"}, {"part", "force"}),
    "insert_table": (op_insert_table, {"anchor", "like_table", "rows"}, {"position", "like_part", "header"}),
    "delete_table": (op_delete_table, {"table"}, {"part", "force"}),
    "insert_image": (op_insert_image, {"anchor", "image"}, {"position", "like", "width_cm", "height_cm", "alt"}),
}
PLAN_KEYS = {"operations", "track_changes", "author", "date"}


def apply_plan(pkg: Package, plan: dict) -> Session:
    if not isinstance(plan, dict) or "operations" not in plan:
        raise PlanError('O plano deve ser um objeto com a chave "operations"')
    unknown = set(plan) - PLAN_KEYS
    if unknown:
        raise PlanError(f"Chaves desconhecidas no plano: {sorted(unknown)}")
    session = Session(pkg, plan)
    if not plan["operations"]:
        raise PlanError("Plano sem operações")
    for number, op in enumerate(plan["operations"], 1):
        name = op.get("op") if isinstance(op, dict) else None
        session.op_label = f"operação #{number} ({name})"
        if name not in OPS:
            raise PlanError(f"{session.op_label}: operação desconhecida. Disponíveis: {sorted(OPS)}")
        handler, required, optional = OPS[name]
        keys = set(op) - {"op"}
        missing, extra = required - keys, keys - required - optional
        if missing or extra:
            raise PlanError(f"{session.op_label}: faltam {sorted(missing)}; chaves inválidas {sorted(extra)}")
        try:
            handler(session, op)
        except PlanError as exc:
            raise PlanError(f"{session.op_label}: {exc}") from exc
    session.op_label = "pós-processamento"
    session.finish()
    return session


def main(argv=None):
    setup_stdio()
    parser = argparse.ArgumentParser(description="Aplica um plano JSON de edições a um .docx.")
    parser.add_argument("input", help="documento de entrada (nunca é alterado)")
    parser.add_argument("plan", help="plano JSON (ou - para stdin)")
    parser.add_argument("-o", "--output", help="documento de saída")
    parser.add_argument("--dry-run", action="store_true", help="aplica e valida sem gravar a saída")
    parser.add_argument("--report", help="grava relatório JSON neste caminho")
    args = parser.parse_args(argv)
    if not args.output and not args.dry_run:
        parser.error("informe -o/--output ou --dry-run")
    report = {"ok": False, "applied": [], "warnings": [], "new_findings": []}
    code = 0
    try:
        if args.output and os.path.abspath(args.output) == os.path.abspath(args.input):
            raise DocxError("A saída não pode ser o próprio arquivo de entrada: preserve o original.")
        if args.plan == "-":
            plan = json.load(sys.stdin)
        else:
            with open(args.plan, encoding="utf-8-sig") as fh:
                plan = json.load(fh)
        pkg = Package(args.input)
        session = apply_plan(pkg, plan)
        report["applied"] = session.log
        report["warnings"] = session.warnings
        report["track_changes"] = bool(session.tracked)
        target_dir = os.path.dirname(os.path.abspath(args.output)) if args.output else None
        fd, tmp = tempfile.mkstemp(suffix=".docx", dir=target_dir)
        os.close(fd)
        try:
            pkg.write_to(tmp)
            new = regressions(validate(tmp), validate(args.input))
            report["new_findings"] = [str(f) for f in new]
            errors = [f for f in new if f.severity == "error"]
            if errors:
                raise DocxError(
                    "Saída descartada: a edição introduziria erros estruturais:\n  "
                    + "\n  ".join(str(f) for f in errors)
                )
            if args.output:
                os.replace(tmp, args.output)
                report["output"] = os.path.abspath(args.output)
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)
        report["ok"] = True
    except json.JSONDecodeError as exc:
        report["error"] = f"Plano não é JSON válido: {exc}"
        code = 2
    except DocxError as exc:
        report["error"] = str(exc)
        code = 2
    if args.report:
        with open(args.report, "w", encoding="utf-8") as fh:
            json.dump(report, fh, ensure_ascii=False, indent=2)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return code


if __name__ == "__main__":
    sys.exit(main())
