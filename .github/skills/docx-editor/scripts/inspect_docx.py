#!/usr/bin/env python3
"""Mostra a estrutura e as convenções de um .docx, sem alterá-lo.

Uso:
    python inspect_docx.py doc.docx                      # índice de parágrafos do corpo
    python inspect_docx.py doc.docx --profile            # padrão visual e de redação
    python inspect_docx.py doc.docx --part word/header1.xml
    python inspect_docx.py doc.docx --range 10:40 --runs # formatação run a run
    python inspect_docx.py doc.docx --grep "cláusula"
    python inspect_docx.py doc.docx --xml 12             # XML bruto do parágrafo 12
    python inspect_docx.py doc.docx --tables
    python inspect_docx.py doc.docx --json
"""
from __future__ import annotations

import argparse
import json
import os
import re
import statistics
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_lib import (  # noqa: E402
    MC_ALT, W14_PARA_ID, DocxError, Package, ParagraphText, Styles, cell_span, cell_vmerge, etree,
    field_depths, has_section_break, human_rpr, in_fallback, iter_paragraphs, iter_tables, local,
    paragraph_numbering, paragraph_style, setup_stdio, w, wval,
)


def _position_in_table(p):
    """(célula, linha, tabela) mais internas do parágrafo."""
    tc = next(p.iterancestors(w("tc")), None)
    if tc is None:
        return None
    tr = tc.getparent()
    while tr is not None and tr.tag != w("tr"):
        tr = tr.getparent()
    tbl = tr.getparent() if tr is not None else None
    while tbl is not None and tbl.tag != w("tbl"):
        tbl = tbl.getparent()
    return tc, tr, tbl


def paragraph_flags(p, depth):
    flags = []
    tags = Counter(el.tag for el in p.iter() if isinstance(el.tag, str))
    if has_section_break(p):
        flags.append("SEÇÃO")
    if depth > 0 or tags[w("fldChar")] or tags[w("fldSimple")]:
        flags.append("campo")
    if tags[w("drawing")] or tags[w("pict")] or tags[w("object")] or tags[MC_ALT]:
        flags.append("objeto")
    if tags[w("txbxContent")]:
        flags.append("caixa-texto")
    if tags[w("footnoteReference")] or tags[w("endnoteReference")]:
        flags.append("nota")
    if tags[w("commentReference")] or tags[w("commentRangeStart")]:
        flags.append("comentário")
    if tags[w("bookmarkStart")]:
        flags.append("indicador")
    if tags[w("hyperlink")]:
        flags.append("link")
    if tags[w("ins")] or tags[w("del")] or tags[w("moveFrom")] or tags[w("moveTo")]:
        flags.append("revisão")
    if any(el.tag == w("sdt") for el in p.iterancestors()) or tags[w("sdt")]:
        flags.append("controle")
    if any(el.tag == w("txbxContent") for el in p.iterancestors()):
        flags.append("em-caixa-texto")
    return flags


def paragraph_rows(pkg, part, styles):
    root = pkg.xml(part)
    depths = field_depths(root)
    tables = {tbl: i for i, tbl in enumerate(iter_tables(root))}
    rows = []
    for idx, p in enumerate(iter_paragraphs(root)):
        depth = depths.get(p, 0)
        pt = ParagraphText(p, depth)
        cell = None
        position = _position_in_table(p)
        if position:
            tc, tr, tbl = position
            if tbl is not None:
                cell = {"table": tables.get(tbl),
                        "row": tbl.findall(w("tr")).index(tr) if tr in tbl.findall(w("tr")) else None,
                        "cell": tr.findall(w("tc")).index(tc) if tc in tr.findall(w("tc")) else None}
        sid = paragraph_style(p)
        rows.append({
            "p": idx,
            "style": sid,
            "style_name": styles.name_of(sid) if sid else None,
            "numbering": paragraph_numbering(p) or (styles.numbering_of(sid) if sid else None),
            "cell": cell,
            "flags": paragraph_flags(p, depth),
            "para_id": p.get(W14_PARA_ID),
            "text": pt.text,
        })
    return rows


def run_rows(p, depth):
    out = []
    seen = []
    for atom in ParagraphText(p, depth).atoms:
        if atom.run is None:
            out.append({"objeto": local(atom.child)})
            continue
        if atom.text is None:
            out.append({"objeto": local(atom.child)})
            continue
        if not atom.length:
            continue
        if seen and seen[-1] is atom.run:
            out[-1]["text"] += atom.text
        else:
            seen.append(atom.run)
            out.append({"format": human_rpr(atom.run.find(w("rPr"))), "text": atom.text})
            if atom.in_field:
                out[-1]["em_campo"] = True
    return out


# ---------------------------------------------------------------------------
# Perfil
# ---------------------------------------------------------------------------
def profile(pkg):
    styles = Styles(pkg)
    root = pkg.xml(pkg.main_part)
    rows = paragraph_rows(pkg, pkg.main_part, styles)
    para_by_index = list(iter_paragraphs(root))
    depths = field_depths(root)
    return {
        "pacote": _package_summary(pkg),
        "secoes": _sections(pkg, root),
        "estilos_em_uso": _style_usage(rows, para_by_index, depths, styles),
        "listas": _lists(pkg, rows, styles),
        "tabelas": _tables(root, styles),
        "objetos": _inventory(pkg),
        "redacao": _writing(rows, para_by_index, styles, root),
        "revisoes": {
            "controle_de_alteracoes_ativo": pkg.tracks_revisions(),
            "pendentes": sum(1 for el in root.iter(w("ins"), w("del")) if el.getparent().tag not in (w("rPr"), w("trPr"))),
        },
    }


def _package_summary(pkg):
    return {
        "documento_principal": pkg.main_part,
        "partes_de_texto": pkg.story_parts(),
        "partes": len(pkg.names()),
        "macros": pkg.has("word/vbaProject.bin"),
    }


def _sections(pkg, root):
    out = []
    for sect in root.iter(w("sectPr")):
        if in_fallback(sect) or local(sect.getparent()) == "sectPrChange":
            continue
        size, margins = sect.find(w("pgSz")), sect.find(w("pgMar"))
        info = {
            "tipo_quebra": wval(sect.find(w("type"))),
            "pagina_cm": None,
            "orientacao": None,
            "margens_cm": None,
            "colunas": (sect.find(w("cols")).get(w("num")) if sect.find(w("cols")) is not None else None),
            "primeira_pagina_diferente": sect.find(w("titlePg")) is not None,
        }
        if size is not None and size.get(w("w")):
            info["pagina_cm"] = [round(int(size.get(w("w"))) / 567, 2), round(int(size.get(w("h"))) / 567, 2)]
            info["orientacao"] = size.get(w("orient"), "portrait")
        if margins is not None:
            info["margens_cm"] = {k: round(int(margins.get(w(k), "0")) / 567, 2)
                                  for k in ("top", "bottom", "left", "right")}
        rels = {r.get("Id"): pkg.rel_target(pkg.main_part, r) for r in pkg.rels(pkg.main_part)}
        rid = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        info["cabecalhos"] = {r.get(w("type")): rels.get(r.get(rid)) for r in sect.findall(w("headerReference"))}
        info["rodapes"] = {r.get(w("type")): rels.get(r.get(rid)) for r in sect.findall(w("footerReference"))}
        out.append(info)
    return out


def _style_usage(rows, paragraphs, depths, styles):
    usage = defaultdict(lambda: {"paragrafos": 0, "caracteres": 0, "rpr": Counter(), "ex": None})
    for row, p in zip(rows, paragraphs):
        entry = usage[row["style"] or styles.defaults.get("paragraph") or "(padrão)"]
        entry["paragrafos"] += 1
        entry["caracteres"] += len(row["text"])
        if entry["ex"] is None and row["text"].strip():
            entry["ex"] = {"p": row["p"], "text": row["text"][:70]}
        for atom in ParagraphText(p, depths.get(p, 0)).atoms:
            if atom.length:
                entry["rpr"][human_rpr(atom.run.find(w("rPr")))] += atom.length
    out = []
    for sid, entry in sorted(usage.items(), key=lambda kv: -kv[1]["caracteres"]):
        common = entry["rpr"].most_common(1)[0][0] if entry["rpr"] else ""
        out.append({
            "style_id": sid,
            "nome": styles.name_of(sid),
            "paragrafos": entry["paragrafos"],
            "caracteres": entry["caracteres"],
            "formato_efetivo": styles.effective_run_props(sid),
            "formatacao_direta_habitual": common,
            "exemplo": entry["ex"],
        })
    return out


def _lists(pkg, rows, styles):
    levels = {}
    if pkg.has("word/numbering.xml"):
        num_root = pkg.xml("word/numbering.xml")
        abstracts = {a.get(w("abstractNumId")): a for a in num_root.findall(w("abstractNum"))}
        for num in num_root.findall(w("num")):
            abstract = abstracts.get(wval(num.find(w("abstractNumId"))))
            if abstract is None:
                continue
            for lvl in abstract.findall(w("lvl")):
                levels[(num.get(w("numId")), lvl.get(w("ilvl")))] = (
                    wval(lvl.find(w("numFmt"))), wval(lvl.find(w("lvlText")))
                )
    use = Counter()
    for row in rows:
        if row["numbering"]:
            use[(row["numbering"][0], row["numbering"][1], row["style"])] += 1
    out = []
    for (num_id, ilvl, sid), count in use.most_common():
        fmt = levels.get((num_id, ilvl), (None, None))
        out.append({"numId": num_id, "nivel": ilvl, "estilo": sid, "itens": count,
                    "formato": fmt[0], "texto_nivel": fmt[1]})
    return out


def _tables(root, styles):
    out = []
    for idx, tbl in enumerate(iter_tables(root)):
        tblpr = tbl.find(w("tblPr"))
        rows = tbl.findall(w("tr"))
        first = rows[0] if rows else None
        header = first is not None and first.find(w("trPr")) is not None \
            and first.find(w("trPr")).find(w("tblHeader")) is not None
        merges = sum(1 for tc in tbl.iter(w("tc")) if cell_vmerge(tc) or cell_span(tc) > 1)
        sid = wval(tblpr.find(w("tblStyle"))) if tblpr is not None else None
        out.append({
            "table": idx,
            "estilo": sid,
            "linhas": len(rows),
            "colunas": len(tbl.find(w("tblGrid")).findall(w("gridCol"))) if tbl.find(w("tblGrid")) is not None else None,
            "cabecalho_repetido": header,
            "celulas_mescladas": merges,
            "primeira_linha": [" ".join(visible_cell(tc).split())[:40] for tc in first.findall(w("tc"))] if first is not None else [],
        })
    return out


def visible_cell(tc):
    return " ".join(ParagraphText(p).text for p in tc.iter(w("p")) if not in_fallback(p))


def _inventory(pkg):
    counts = Counter()
    for part in pkg.story_parts():
        for el in pkg.xml(part).iter():
            if not isinstance(el.tag, str) or in_fallback(el):
                continue
            if el.tag in (w("drawing"), w("pict")):
                counts["imagens_desenhos"] += 1
            elif el.tag == w("instrText"):
                name = (el.text or "").split()
                if name:
                    counts["campo_" + name[0].upper()] += 1
            elif el.tag == w("fldSimple"):
                name = el.get(w("instr"), "").split()
                if name:
                    counts["campo_" + name[0].upper()] += 1
            elif el.tag == w("hyperlink"):
                counts["hiperlinks"] += 1
            elif el.tag == w("bookmarkStart") and not (el.get(w("name")) or "").startswith("_"):
                counts["indicadores_nomeados"] += 1
            elif el.tag == w("sdt"):
                counts["controles_de_conteudo"] += 1
            elif el.tag == w("footnoteReference"):
                counts["notas_de_rodape"] += 1
            elif el.tag == w("endnoteReference"):
                counts["notas_de_fim"] += 1
            elif el.tag == w("commentReference"):
                counts["comentarios"] += 1
            elif el.tag == w("txbxContent"):
                counts["caixas_de_texto"] += 1
    return dict(sorted(counts.items()))


_SENTENCE = re.compile(r"[^.!?…]+[.!?…]+(?:\s|$)")
_DATE_PATTERNS = {
    "dd/mm/aaaa": r"\b\d{1,2}/\d{1,2}/\d{4}\b",
    "aaaa-mm-dd": r"\b\d{4}-\d{2}-\d{2}\b",
    "dd de mês de aaaa": r"\b\d{1,2}º?\s+de\s+[a-zç]+\s+de\s+\d{4}\b",
    "Month dd, yyyy": r"\b[A-Z][a-z]+ \d{1,2}, \d{4}\b",
}


def _writing(rows, paragraphs, styles, root):
    """Sinais objetivos da voz do autor, para o texto novo não destoar."""
    body, headings, items = [], [], []
    for row, p in zip(rows, paragraphs):
        text = row["text"].strip()
        if not text or "campo" in row["flags"]:
            continue
        if row["style"] and styles.is_heading(row["style"]):
            headings.append(text)
        elif row["numbering"] or (row["style_name"] or "").casefold().startswith(("list", "lista")):
            items.append(text)
        elif row["cell"] is None:
            body.append(text)
    joined = "\n".join(body + items)
    sentences = [s for para in body for s in _SENTENCE.findall(para)]
    lengths = [len(s.split()) for s in sentences]
    languages = Counter(wval(el) for el in root.iter(w("lang")) if wval(el))
    if not languages and styles.root is not None:
        languages = Counter(wval(el) for el in styles.root.iter(w("lang")) if wval(el))
    return {
        "idioma_marcado": languages.most_common(1)[0][0] if languages else None,
        "palavras_por_frase_mediana": statistics.median(lengths) if lengths else None,
        "palavras_por_paragrafo_mediana": statistics.median(len(b.split()) for b in body) if body else None,
        "caixa_dos_titulos": _heading_case(headings),
        "titulos_terminam_com_ponto": sum(h.endswith(".") for h in headings),
        "titulos_total": len(headings),
        "fim_de_item_de_lista": dict(Counter(_terminal(i) for i in items).most_common()),
        "aspas": {"curvas": len(re.findall("[“”‘’]", joined)),
                  "retas": len(re.findall("[\"']", joined))},
        "travessoes": {"em": joined.count("—"), "en": joined.count("–"),
                       "hifen_entre_espacos": joined.count(" - ")},
        "espaco_duplo_apos_ponto": len(re.findall(r"[.!?]  +\S", joined)),
        "numeros": {"virgula_decimal": len(re.findall(r"\d,\d{1,2}\b", joined)),
                    "ponto_decimal": len(re.findall(r"\d\.\d{1,2}\b", joined))},
        "formatos_de_data": {k: n for k, n in ((k, len(re.findall(v, joined))) for k, v in _DATE_PATTERNS.items()) if n},
    }


def _heading_case(headings):
    if not headings:
        return None
    votes = Counter()
    minor = {"de", "da", "do", "das", "dos", "e", "a", "o", "as", "os", "em", "para", "com", "por",
             "of", "the", "and", "in", "on", "to", "for", "with", "at", "by"}
    for text in headings:
        words = [x for x in re.findall(r"[^\W\d_]+", text) if len(x) > 1]
        if not words:
            continue
        if all(x.isupper() for x in words):
            votes["MAIÚSCULAS"] += 1
        elif len(words) > 1 and all(x[0].isupper() or x.casefold() in minor for x in words[1:]) \
                and any(x[0].isupper() for x in words[1:]):
            votes["Cada Palavra"] += 1
        else:
            votes["Só a primeira"] += 1
    return dict(votes.most_common())


def _terminal(text):
    last = text[-1]
    return last if last in ".;:,!?" else "(nada)"


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def format_row(row, max_chars):
    bits = [f"[{row['p']}]", row["style"] or "-"]
    if row["numbering"]:
        bits.append(f"num={row['numbering'][0]}:{row['numbering'][1]}")
    if row["cell"]:
        c = row["cell"]
        bits.append(f"T{c['table']}R{c['row']}C{c['cell']}")
    if row["flags"]:
        bits.append("{" + ",".join(row["flags"]) + "}")
    text = row["text"]
    if max_chars and len(text) > max_chars:
        text = text[:max_chars] + "…"
    return " ".join(bits) + " " + json.dumps(text, ensure_ascii=False)


def main(argv=None):
    setup_stdio()
    parser = argparse.ArgumentParser(description="Inspeciona um .docx sem alterá-lo.")
    parser.add_argument("input")
    parser.add_argument("--part", help="parte a listar (padrão: documento principal)")
    parser.add_argument("--profile", action="store_true", help="padrão visual e de redação do documento")
    parser.add_argument("--tables", action="store_true", help="lista as tabelas")
    parser.add_argument("--range", help="intervalo de parágrafos, ex.: 10:40")
    parser.add_argument("--grep", help="só parágrafos que contenham este texto")
    parser.add_argument("--runs", action="store_true", help="mostra a formatação run a run")
    parser.add_argument("--xml", type=int, metavar="N", help="XML bruto do parágrafo N")
    parser.add_argument("--max-chars", type=int, default=0, help="trunca o texto listado")
    parser.add_argument("--json", action="store_true", help="saída em JSON")
    args = parser.parse_args(argv)
    try:
        pkg = Package(args.input)
        styles = Styles(pkg)
        part = args.part or pkg.main_part
        if part not in pkg.story_parts():
            raise DocxError(f"{part} não é uma parte de texto. Disponíveis: {pkg.story_parts()}")
        root = pkg.xml(part)
        if args.profile:
            print(json.dumps(profile(pkg), ensure_ascii=False, indent=1))
            return 0
        if args.tables:
            print(json.dumps(_tables(root, styles), ensure_ascii=False, indent=1))
            return 0
        paragraphs = list(iter_paragraphs(root))
        if args.xml is not None:
            if not 0 <= args.xml < len(paragraphs):
                raise DocxError(f"Parágrafo {args.xml} fora de 0..{len(paragraphs) - 1}")
            print(etree.tostring(paragraphs[args.xml], encoding="unicode"))
            return 0
        rows = paragraph_rows(pkg, part, styles)
        depths = field_depths(root)
        lo, hi = 0, len(rows)
        if args.range:
            lo_text, _, hi_text = args.range.partition(":")
            lo = int(lo_text or 0)
            hi = int(hi_text) if hi_text else len(rows)
        selected = [r for r in rows[lo:hi] if not args.grep or args.grep.casefold() in r["text"].casefold()]
        if args.runs:
            for row in selected:
                row["runs"] = run_rows(paragraphs[row["p"]], depths.get(paragraphs[row["p"]], 0))
        if args.json:
            print(json.dumps(selected, ensure_ascii=False, indent=1))
            return 0
        print(f"# {part}: {len(rows)} parágrafos")
        for row in selected:
            print(format_row(row, args.max_chars))
            for run in row.get("runs", []):
                print("     ", json.dumps(run, ensure_ascii=False))
        return 0
    except DocxError as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
