#!/usr/bin/env python3
"""Prova o que mudou (e o que não mudou) entre o original e a versão editada.

Uso:
    python compare_docx.py original.docx editado.docx [--allow PADRÃO ...] [--json]

--allow lista partes que podem mudar (padrão: word/document.xml). Aceita curingas,
por exemplo: --allow "word/media/*" --allow word/footer1.xml

VIOLAÇÃO (saída 1) é alteração fora do escopo:
  * parte não permitida alterada, adicionada ou removida;
  * formatação alterada em texto que continuou igual;
  * configuração de seção alterada.
ATENÇÃO (saída 0) pede conferência com o pedido: objetos a mais ou a menos e
formatação nova, sem precedente no original.
"""
from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_lib import (  # noqa: E402
    DocxError, Package, ParagraphText, canon, char_formats, field_depths, glob_match, has_deleted_mark,
    in_fallback, iter_paragraphs, local, ppr_signature, rpr_signature, setup_stdio, w,
)


def compare_parts(old, new, allow):
    report = {"identicas": 0, "reserializadas": [], "alteradas": [], "adicionadas": [], "removidas": []}
    violations = []
    old_names, new_names = set(old.names()), set(new.names())
    for name in sorted(old_names | new_names):
        allowed = glob_match(name, allow)
        if name not in new_names:
            report["removidas"].append(name)
            if not allowed:
                violations.append(f"parte removida fora do escopo: {name}")
        elif name not in old_names:
            report["adicionadas"].append(name)
            if not allowed:
                violations.append(f"parte adicionada fora do escopo: {name}")
        elif hashlib.sha256(old.data[name]).digest() == hashlib.sha256(new.data[name]).digest():
            report["identicas"] += 1
        elif name.endswith((".xml", ".rels")) and canon(old.xml(name)) == canon(new.xml(name)):
            report["reserializadas"].append(name)
        else:
            report["alteradas"].append(name)
            if not allowed:
                violations.append(f"parte alterada fora do escopo: {name}")
    return report, violations


def _paragraph_records(root):
    depths = field_depths(root)
    records = []
    for p in iter_paragraphs(root):
        pt = ParagraphText(p, depths.get(p, 0))
        records.append({"text": pt.text, "ppr": ppr_signature(p), "chars": char_formats(pt),
                        "object": canon_objects(pt), "deleted": has_deleted_mark(p)})
    return records


def canon_objects(pt):
    # A referência de comentário pode ser realocada de propósito ao remover um parágrafo vizinho.
    return [local(a.child) for a in pt.atoms if a.text is None and a.child.tag != w("commentReference")]


def compare_story(name, old_root, new_root, known_ppr, known_rpr):
    """Alinha parágrafos e separa edição pedida de alteração colateral."""
    old_recs, new_recs = _paragraph_records(old_root), _paragraph_records(new_root)
    matcher = difflib.SequenceMatcher(None, [r["text"] for r in old_recs], [r["text"] for r in new_recs], autojunk=False)
    result = {"parte": name, "modificados": [], "inseridos": [], "removidos": []}
    violations, attention = [], []

    def check_same_text(i, j):
        old_rec, new_rec = old_recs[i], new_recs[j]
        if old_rec["ppr"] != new_rec["ppr"]:
            violations.append(f"{name}: propriedades do parágrafo mudaram sem pedido, p{i} -> p{j}: {old_rec['text'][:60]!r}")
        if old_rec["chars"] != new_rec["chars"]:
            violations.append(f"{name}: formatação de texto inalterado mudou, p{i} -> p{j}: {old_rec['text'][:60]!r}")
        if old_rec["object"] != new_rec["object"]:
            violations.append(f"{name}: objeto (imagem, nota, campo) sumiu ou mudou em parágrafo de texto igual, "
                              f"p{i} -> p{j}: {old_rec['text'][:60]!r}")

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for offset in range(i2 - i1):
                check_same_text(i1 + offset, j1 + offset)
            continue
        pairs, unmatched_old, unmatched_new = _pair_block(old_recs, new_recs, range(i1, i2), range(j1, j2))
        for i, j in pairs:
            old_rec, new_rec = old_recs[i], new_recs[j]
            result["modificados"].append({"de": i, "para": j, "antes": old_rec["text"], "depois": new_rec["text"]})
            if old_rec["ppr"] != new_rec["ppr"]:
                violations.append(f"{name}: propriedades do parágrafo modificado mudaram, p{i} -> p{j}")
            _check_surviving_chars(name, i, j, old_rec, new_rec, violations, attention, known_rpr)
        for i in unmatched_old:
            result["removidos"].append({"p": i, "text": old_recs[i]["text"]})
        for j in unmatched_new:
            rec = new_recs[j]
            if rec["deleted"] and not rec["text"]:
                continue  # é o parágrafo removido com revisão, já listado em "removidos"
            result["inseridos"].append({"p": j, "text": rec["text"]})
            if rec["ppr"] not in known_ppr:
                attention.append(f"{name}: p{j} inserido com propriedades sem precedente no original: {rec['text'][:60]!r}")
            if set(rec["chars"]) - known_rpr:
                attention.append(f"{name}: p{j} inserido com formatação de texto sem precedente: {rec['text'][:60]!r}")
    return result, violations, attention


def _pair_block(old_recs, new_recs, old_range, new_range):
    pairs = []
    used = -1
    new_list = list(new_range)
    unmatched_old = []
    for i in old_range:
        hit = None
        for j in new_list:
            if j <= used:
                continue
            if difflib.SequenceMatcher(None, old_recs[i]["text"], new_recs[j]["text"], autojunk=False).ratio() >= 0.5:
                hit = j
                break
        if hit is None:
            unmatched_old.append(i)
        else:
            pairs.append((i, hit))
            used = hit
    taken = {j for _, j in pairs}
    return pairs, unmatched_old, [j for j in new_list if j not in taken]


def _check_surviving_chars(name, i, j, old_rec, new_rec, violations, attention, known_rpr):
    matcher = difflib.SequenceMatcher(None, old_rec["text"], new_rec["text"], autojunk=False)
    # O texto visível exclui objetos, logo os índices coincidem com char_formats.
    old_chars, new_chars = old_rec["chars"], new_rec["chars"]
    for block in matcher.get_matching_blocks():
        for k in range(block.size):
            if old_chars[block.a + k] != new_chars[block.b + k]:
                violations.append(f"{name}: formatação de trecho preservado mudou, p{i} -> p{j}: "
                                  f"{old_rec['text'][block.a:block.a + block.size][:50]!r}")
                break
    fresh = set(new_chars) - known_rpr
    if fresh:
        attention.append(f"{name}: p{j} usa formatação de texto sem precedente no original")


_INVENTORY_TAGS = {
    w("tbl"): "tabelas", w("tr"): "linhas de tabela", w("sectPr"): "seções", w("drawing"): "desenhos",
    w("pict"): "objetos VML", w("hyperlink"): "hiperlinks", w("sdt"): "controles de conteúdo",
    w("footnoteReference"): "notas de rodapé", w("endnoteReference"): "notas de fim",
    w("commentReference"): "comentários", w("bookmarkStart"): "indicadores",
    w("ins"): "revisões (inserção)", w("del"): "revisões (exclusão)",
    w("headerReference"): "referências de cabeçalho", w("footerReference"): "referências de rodapé",
}


def inventory(pkg):
    counts = Counter()
    for part in pkg.story_parts():
        for el in pkg.xml(part).iter():
            if not isinstance(el.tag, str) or in_fallback(el):
                continue
            if el.tag in _INVENTORY_TAGS:
                counts[_INVENTORY_TAGS[el.tag]] += 1
            elif el.tag == w("instrText"):
                words = (el.text or "").split()
                if words:
                    counts["campo " + words[0].upper()] += 1
    return counts


def compare_sections(old, new):
    def collect(pkg):
        return [canon(s, skip_tags={w("sectPrChange")}) for s in pkg.xml(pkg.main_part).iter(w("sectPr"))
                if not in_fallback(s)]
    return collect(old) == collect(new)


def compare_tables(old, new):
    """Propriedades de tabela/célula não aparecem na comparação de texto."""
    def tables(pkg):
        return [t for t in pkg.xml(pkg.main_part).iter(w("tbl")) if not in_fallback(t)]

    old_tables, new_tables = tables(old), tables(new)
    if len(old_tables) != len(new_tables):
        return []
    out = []
    for idx, (a, b) in enumerate(zip(old_tables, new_tables)):
        if canon(a.find(w("tblPr"))) != canon(b.find(w("tblPr"))) or canon(a.find(w("tblGrid"))) != canon(b.find(w("tblGrid"))):
            out.append(f"tabela {idx}: propriedades ou grade alteradas")
        if len(a.findall(w("tr"))) == len(b.findall(w("tr"))):
            props_a = [canon(x) for x in a.iter(w("tcPr"), w("trPr"))]
            props_b = [canon(x) for x in b.iter(w("tcPr"), w("trPr"))]
            if props_a != props_b:
                out.append(f"tabela {idx}: propriedades de linha ou célula alteradas")
    return out


def _known_formats(pkg):
    known_ppr, known_rpr = {""}, {""}
    for part in pkg.story_parts():
        root = pkg.xml(part)
        for p in iter_paragraphs(root):
            known_ppr.add(ppr_signature(p))
        for run in root.iter(w("r")):
            known_rpr.add(rpr_signature(run))
    return known_ppr, known_rpr


def main(argv=None):
    setup_stdio()
    parser = argparse.ArgumentParser(description="Compara original e versão editada de um .docx.")
    parser.add_argument("original")
    parser.add_argument("edited")
    parser.add_argument("--allow", action="append", default=[], help="parte que pode mudar (aceita curingas)")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)
    try:
        old, new = Package(args.original), Package(args.edited)
    except DocxError as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 2
    allow = args.allow or ["word/document.xml"]
    parts, violations = compare_parts(old, new, allow)
    attention = []
    known_ppr, known_rpr = _known_formats(old)
    stories = []
    for name in new.story_parts():
        if old.has(name) and (name in parts["alteradas"] or name in parts["reserializadas"]):
            result, story_violations, story_attention = compare_story(name, old.xml(name), new.xml(name), known_ppr, known_rpr)
            stories.append(result)
            violations += story_violations
            attention += story_attention
    if not compare_sections(old, new):
        violations.append("configuração de seção alterada (w:sectPr)")
    violations += compare_tables(old, new)
    before, after = inventory(old), inventory(new)
    changed_inventory = {k: [before[k], after[k]] for k in sorted(set(before) | set(after)) if before[k] != after[k]}
    for key, (a, b) in changed_inventory.items():
        attention.append(f"inventário: {key}: {a} -> {b}")
    if [i.filename for i in old.infos if new.has(i.filename)] != [i.filename for i in new.infos if old.has(i.filename)]:
        attention.append("ordem das entradas do ZIP mudou")
    attention = sorted(set(attention))
    report = {"ok": not violations, "violacoes": violations, "atencao": attention,
              "partes": parts, "historias": stories, "inventario_alterado": changed_inventory}
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=1))
    else:
        print(f"Partes: {parts['identicas']} idênticas byte a byte; alteradas: {parts['alteradas'] or '-'}; "
              f"reserializadas: {parts['reserializadas'] or '-'}; adicionadas: {parts['adicionadas'] or '-'}; "
              f"removidas: {parts['removidas'] or '-'}")
        for story in stories:
            print(f"\n== {story['parte']}: {len(story['modificados'])} modificado(s), "
                  f"{len(story['inseridos'])} inserido(s), {len(story['removidos'])} removido(s)")
            for item in story["modificados"]:
                print(f"  ~ p{item['de']}->p{item['para']}\n    - {item['antes']!r}\n    + {item['depois']!r}")
            for item in story["inseridos"]:
                print(f"  + p{item['p']}: {item['text']!r}")
            for item in story["removidos"]:
                print(f"  - p{item['p']}: {item['text']!r}")
        for item in attention:
            print(f"ATENÇÃO: {item}")
        for item in violations:
            print(f"VIOLAÇÃO: {item}")
        print("\nOK: nada mudou fora do escopo." if not violations else f"\nREPROVADO: {len(violations)} violação(ões).")
    return 0 if not violations else 1


if __name__ == "__main__":
    sys.exit(main())
