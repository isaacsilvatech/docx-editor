"""Testes da skill docx-editor. Rode com: python -m unittest discover -s tests -v"""
from __future__ import annotations

import contextlib
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))
sys.path.insert(0, HERE)

import compare_docx  # noqa: E402
import edit_docx  # noqa: E402
import inspect_docx  # noqa: E402
import make_fixture  # noqa: E402
import pack_docx  # noqa: E402
import render_docx  # noqa: E402
import unpack_docx  # noqa: E402
import validate_docx  # noqa: E402
from docx_lib import (  # noqa: E402
    DocxError, Package, ParagraphText, field_depths, iter_paragraphs, parse_xml, rpr_signature,
    regressions, validate, w,
)


def read_bytes(path):
    with open(path, "rb") as fh:
        return fh.read()


def write_bytes(path, data):
    with open(path, "wb") as fh:
        fh.write(data)


def quiet(fn, *args):
    out = io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
        code = fn(list(args))
    return code, out.getvalue()


class Base(unittest.TestCase):
    tracked_fixture = False

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="docx-editor-test-")
        self.addCleanup(shutil.rmtree, self.dir, ignore_errors=True)
        self.src = make_fixture.build(os.path.join(self.dir, "in.docx"), tracked=self.tracked_fixture)
        self.out = os.path.join(self.dir, "out.docx")

    def edit(self, operations, allow_ok=True, **plan):
        """Executa o CLI real e confere o que toda edição precisa cumprir."""
        plan_path = os.path.join(self.dir, "plan.json")
        with open(plan_path, "w", encoding="utf-8") as fh:
            json.dump({"operations": operations, **plan}, fh, ensure_ascii=False)
        code, output = quiet(edit_docx.main, self.src, plan_path, "-o", self.out)
        report = json.loads(output)
        if allow_ok:
            self.assertEqual(code, 0, report.get("error"))
            self.assertEqual(regressions(validate(self.out), validate(self.src)), [],
                             "a edição introduziu achados de validação")
        return code, report

    def error(self, operations, fragment, **plan):
        code, report = self.edit(operations, allow_ok=False, **plan)
        self.assertEqual(code, 2, report)
        self.assertIn(fragment, report["error"])
        self.assertFalse(os.path.exists(self.out), "plano recusado não pode gerar arquivo")

    def paragraphs(self, path=None, part="word/document.xml"):
        root = Package(path or self.out).xml(part)
        depths = field_depths(root)
        return [ParagraphText(p, depths.get(p, 0)).text for p in iter_paragraphs(root)]

    def para(self, index, path=None, part="word/document.xml"):
        return list(iter_paragraphs(Package(path or self.out).xml(part)))[index]

    def compare(self, *allow):
        args = [self.src, self.out, "--json"]
        for pattern in allow:
            args += ["--allow", pattern]
        code, output = quiet(compare_docx.main, *args)
        return code, json.loads(output)


class TestNoOpAndPackage(Base):
    def test_fixture_is_clean(self):
        findings = validate(self.src)
        self.assertEqual([str(f) for f in findings if f.severity != "info"], [])

    def test_untouched_parts_are_byte_identical_and_order_kept(self):
        self.edit([{"op": "replace_text", "find": "Último parágrafo.", "replace": "Parágrafo final."}])
        with zipfile.ZipFile(self.src) as a, zipfile.ZipFile(self.out) as b:
            self.assertEqual(a.namelist(), b.namelist())
            for name in a.namelist():
                if name != "word/document.xml":
                    self.assertEqual(a.read(name), b.read(name), name)

    def test_namespaces_and_ignorable_survive(self):
        self.edit([{"op": "replace_text", "find": "Último parágrafo.", "replace": "Fim."}])
        with zipfile.ZipFile(self.out) as zf:
            data = zf.read("word/document.xml")
        self.assertTrue(data.startswith(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
        self.assertIn(b'mc:Ignorable="w14 w15"', data)
        root = parse_xml(data)
        for prefix in ("w14", "w15", "wps", "v"):
            self.assertIn(prefix, root.nsmap)
        self.assertEqual(data.count(b"xmlns:w="), 1, "não pode redeclarar namespaces em elementos novos")

    def test_input_is_never_overwritten(self):
        plan = os.path.join(self.dir, "p.json")
        with open(plan, "w", encoding="utf-8") as fh:
            json.dump({"operations": [{"op": "delete_paragraph", "target": {"contains": "descartável"}}]}, fh)
        before = read_bytes(self.src)
        code, _ = quiet(edit_docx.main, self.src, plan, "-o", self.src)
        self.assertEqual(code, 2)
        self.assertEqual(before, read_bytes(self.src))

    def test_dry_run_writes_nothing(self):
        plan = os.path.join(self.dir, "p.json")
        with open(plan, "w", encoding="utf-8") as fh:
            json.dump({"operations": [{"op": "delete_paragraph", "target": {"contains": "descartável"}}]}, fh)
        code, _ = quiet(edit_docx.main, self.src, plan, "--dry-run")
        self.assertEqual(code, 0)
        self.assertEqual(sorted(os.listdir(self.dir)), ["in.docx", "p.json"])

    def test_rejects_legacy_and_junk_files(self):
        legacy = os.path.join(self.dir, "x.docx")
        with open(legacy, "wb") as fh:
            fh.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"0" * 64)
        with self.assertRaisesRegex(DocxError, "senha"):
            Package(legacy)
        with open(legacy, "wb") as fh:
            fh.write(b"isto nao e um zip")
        with self.assertRaisesRegex(DocxError, "não é um ZIP"):
            Package(legacy)

    def test_all_or_nothing(self):
        self.error([
            {"op": "delete_paragraph", "target": {"contains": "descartável"}},
            {"op": "replace_text", "find": "não existe", "replace": "x"},
        ], "operação #2")


class TestReplaceText(Base):
    def test_across_runs_keeps_formatting(self):
        self.edit([{"op": "replace_text", "find": "contrato vigente", "replace": "acordo vigente"}])
        self.assertTrue(self.paragraphs()[4].startswith("O acordo vigente vence em 2023 e pode"))
        fonts = {rpr_signature(r) for r in self.para(4).findall(w("r"))}
        original = {rpr_signature(r) for r in self.para(4, self.src).findall(w("r"))}
        self.assertEqual(fonts, original)
        self.assertEqual(self.compare()[0], 0)

    def test_replacement_inherits_bold_of_matched_text(self):
        self.edit([{"op": "replace_text", "find": "vence em 2023", "replace": "vence em 2030"}])
        runs = self.para(4).findall(w("r"))
        bold = [r for r in runs if r.find(w("rPr")) is not None and r.find(w("rPr")).find(w("b")) is not None]
        self.assertIn("2030", "".join(t.text for r in bold for t in r.findall(w("t"))))

    def test_preserves_leading_trailing_space(self):
        self.edit([{"op": "replace_text", "find": "Veja ", "replace": "Consulte  "}])
        t = self.para(9).findall(w("r"))[0].find(w("t"))
        self.assertEqual(t.get("{http://www.w3.org/XML/1998/namespace}space"), "preserve")

    def test_count_guard(self):
        self.error([{"op": "replace_text", "find": "Introdução", "replace": "X", "count": 1, "target": {"p": 99}}], "fora de")
        self.error([{"op": "replace_text", "find": "item", "replace": "ponto"}], "count=1")

    def test_count_all(self):
        self.edit([{"op": "replace_text", "find": "item", "replace": "ponto", "count": "all"}])
        self.assertIn("Primeiro ponto;", self.paragraphs())
        self.assertIn("Segundo ponto;", self.paragraphs())

    def test_field_result_refused_then_allowed(self):
        self.error([{"op": "replace_text", "find": "Resultados", "replace": "Achados", "count": "all"}], "resultado de campo")
        self.edit([{"op": "replace_text", "find": "Resultados", "replace": "Achados", "count": "all",
                    "allow_field_result": True}])
        self.assertEqual(self.paragraphs()[1], "Achados")
        self.assertEqual(self.paragraphs()[12], "Achados")

    def test_barrier_blocks_match_over_footnote(self):
        self.error([{"op": "replace_text", "find": "norma aplicável", "replace": "regra"}], "atravessando um objeto")
        self.edit([{"op": "replace_text", "find": "a norma", "replace": "a regra"}])
        root = Package(self.out).xml("word/document.xml")
        self.assertEqual(len(list(root.iter(w("footnoteReference")))), 1)
        self.assertEqual(len(list(root.iter(w("commentReference")))), 1)

    def test_image_in_paragraph_survives_text_edit(self):
        self.edit([{"op": "replace_text", "find": "legenda da figura", "replace": "descrição da figura"}])
        self.assertEqual(len(list(self.para(10).iter(w("drawing")))), 1)
        self.assertEqual(self.compare()[0], 0)

    def test_loose_match_hint(self):
        self.error([{"op": "replace_text", "find": "não existe aqui", "replace": "x"}], "não encontrado")

    def test_edit_header_and_footer_keeps_field(self):
        self.edit([
            {"op": "replace_text", "find": "Empresa Exemplo", "replace": "Empresa Nova", "parts": ["word/header1.xml"]},
            {"op": "replace_text", "find": "Página", "replace": "Pág.", "target": {"part": "word/footer1.xml", "p": 0}},
        ])
        self.assertEqual(self.paragraphs(part="word/header1.xml")[0], "Empresa Nova")
        footer = Package(self.out).xml("word/footer1.xml")
        self.assertEqual(len(list(footer.iter(w("fldChar")))), 3)
        code, report = self.compare("word/header1.xml", "word/footer1.xml")
        self.assertEqual(code, 0, report["violacoes"])

    def test_multi_run_payload_with_character_style(self):
        self.edit([{"op": "replace_text", "find": "Último parágrafo.",
                    "runs": [{"text": "Nota: ", "rstyle": "Strong"}, {"text": "fim."}]}])
        runs = self.para(-1).findall(w("r"))
        self.assertEqual(runs[0].find(w("rPr")).find(w("rStyle")).get(w("val")), "Forte")
        self.assertIsNone(runs[1].find(w("rPr")))

    def test_unknown_character_style_refused(self):
        self.error([{"op": "replace_text", "find": "Último parágrafo.",
                     "runs": [{"text": "x", "rstyle": "Inventado"}]}], "não existe")

    def test_textbox_edit_mirrors_fallback(self):
        self.edit([{"op": "replace_text", "find": "Texto na caixa", "replace": "Novo na caixa"}])
        with zipfile.ZipFile(self.out) as zf:
            data = zf.read("word/document.xml").decode("utf-8")
        self.assertEqual(data.count(">Novo<"), 2, "mc:Choice e mc:Fallback precisam mudar juntos")
        self.assertNotIn("Texto", data.split("mc:Fallback")[1])

    def test_content_control_binding_warns(self):
        _, report = self.edit([{"op": "replace_text", "find": "Valor vinculado", "replace": "Outro valor"}])
        self.assertTrue(any("dataBinding" in x for x in report["warnings"]))


class TestRewrite(Base):
    def test_word_level_diff_keeps_inline_bold(self):
        self.edit([{"op": "rewrite_paragraph", "target": {"p": 4, "expect": "contrato"},
                    "text": "O contrato vigente vence em 2030 e pode ser renovado por acordo entre as partes."}])
        bold = [r for r in self.para(4).findall(w("r"))
                if r.find(w("rPr")).find(w("b")) is not None]
        self.assertIn("vence em", "".join(t.text for r in bold for t in r.findall(w("t"))))
        self.assertEqual(self.compare()[0], 0)

    def test_whole_rewrite_uses_dominant_format_and_keeps_ppr(self):
        before_ppr = self.para(4, self.src).find(w("pPr"))
        self.edit([{"op": "rewrite_paragraph", "target": {"p": 4}, "text": "Tudo foi trocado neste parágrafo inteiro."}])
        p = self.para(4)
        self.assertEqual(ParagraphText(p).text, "Tudo foi trocado neste parágrafo inteiro.")
        self.assertEqual(p.find(w("pPr")).find(w("jc")).get(w("val")), before_ppr.find(w("jc")).get(w("val")))
        run = p.find(w("r"))
        self.assertIsNone(run.find(w("rPr")).find(w("b")), "o negrito minoritário não deve dominar")
        self.assertEqual(run.find(w("rPr")).find(w("rFonts")).get(w("ascii")), "Georgia")

    def test_keeps_footnote_and_comment_objects(self):
        self.edit([{"op": "rewrite_paragraph", "target": {"p": 5}, "text": "O prazo segue a norma federal aplicável ao caso."}])
        p = self.para(5)
        self.assertEqual(len(list(p.iter(w("footnoteReference")))), 1)
        self.assertEqual(len(list(p.iter(w("commentReference")))), 1)

    def test_section_break_paragraph_text_can_be_rewritten(self):
        self.edit([{"op": "rewrite_paragraph", "target": {"p": 11}, "text": "Encerramento da primeira seção."}])
        self.assertIsNotNone(self.para(11).find(w("pPr")).find(w("sectPr")))
        self.assertEqual(self.compare()[0], 0)

    def test_hyperlink_kept_when_outside_change(self):
        self.edit([{"op": "rewrite_paragraph", "target": {"p": 9}, "text": "Veja o portal oficial para mais detalhes."}])
        self.assertEqual(len(list(self.para(9).iter(w("hyperlink")))), 1)

    def test_stale_index_detected(self):
        self.error([{"op": "rewrite_paragraph", "target": {"p": 4, "expect": "texto que não está lá"}, "text": "x"}],
                   "texto esperado")

    def test_field_paragraph_refused_when_field_changes(self):
        self.error([{"op": "rewrite_paragraph", "target": {"contains": "Conforme o valor"}, "text": "Totalmente diferente agora mesmo sem nada igual."}],
                   "resultado de campo")


class TestInsertParagraph(Base):
    def test_list_item_clone_keeps_numbering_and_style(self):
        self.edit([{"op": "insert_paragraph", "anchor": {"p": 7}, "text": "Novo item;"}])
        p = self.para(8)
        self.assertEqual(ParagraphText(p).text, "Novo item;")
        numpr = p.find(w("pPr")).find(w("numPr"))
        self.assertEqual(numpr.find(w("numId")).get(w("val")), "1")
        self.assertEqual(p.find(w("pPr")).find(w("pStyle")).get(w("val")), "ItemLista")
        code, report = self.compare()
        self.assertEqual(code, 0, report["violacoes"])
        self.assertEqual(report["atencao"], [])

    def test_like_style_uses_nearest_exemplar(self):
        self.edit([{"op": "insert_paragraph", "anchor": {"p": 10}, "like": {"style": "Heading 1"}, "text": "Seção nova"}])
        self.assertEqual(self.para(11).find(w("pPr")).find(w("pStyle")).get(w("val")), "Titulo1")

    def test_unknown_style_refused_with_options(self):
        self.error([{"op": "insert_paragraph", "anchor": {"p": 10}, "like": {"style": "Heading 9"}, "text": "x"}],
                   "Estilos de parágrafo")

    def test_clone_of_section_paragraph_never_duplicates_section(self):
        self.edit([{"op": "insert_paragraph", "anchor": {"p": 11}, "position": "before", "text": "Antes da quebra."}])
        root = Package(self.out).xml("word/document.xml")
        self.assertEqual(len(list(root.iter(w("sectPr")))), 2)
        self.assertEqual(self.compare()[0], 0)

    def test_after_section_paragraph_warns(self):
        _, report = self.edit([{"op": "insert_paragraph", "anchor": {"p": 11}, "text": "Depois."}])
        self.assertTrue(any("SEGUINTE" in x for x in report["warnings"]))

    def test_multiple_after_same_anchor_keep_order(self):
        self.edit([
            {"op": "insert_paragraph", "anchor": {"p": 7}, "text": "A;"},
            {"op": "insert_paragraph", "anchor": {"p": 7}, "text": "B;"},
            {"op": "insert_paragraph", "anchor": {"p": 7}, "paragraphs": ["C;", "D;"]},
        ])
        self.assertEqual(self.paragraphs()[8:12], ["A;", "B;", "C;", "D;"])

    def test_indexes_refer_to_original_document(self):
        self.edit([
            {"op": "insert_paragraph", "anchor": {"p": 3}, "text": "Extra"},
            {"op": "replace_text", "target": {"p": 7, "expect": "Segundo"}, "find": "Segundo", "replace": "Dois"},
        ])
        self.assertIn("Dois item;", self.paragraphs())

    def test_clone_does_not_copy_ids_bookmarks_or_objects(self):
        self.edit([{"op": "insert_paragraph", "anchor": {"p": 3}, "text": "Outro título"}])
        new = self.para(4)
        self.assertIsNone(new.get("{http://schemas.microsoft.com/office/word/2010/wordml}paraId"))
        self.assertEqual(list(new.iter(w("bookmarkStart"))), [])
        self.edit([{"op": "insert_paragraph", "anchor": {"p": 10}, "text": "Sem imagem"}])
        self.assertEqual(list(self.para(11).iter(w("drawing"))), [])

    def test_body_paragraph_inherits_font_from_model(self):
        self.edit([{"op": "insert_paragraph", "anchor": {"p": 4}, "text": "Novo parágrafo de corpo."}])
        run = self.para(5).find(w("r"))
        self.assertEqual(run.find(w("rPr")).find(w("rFonts")).get(w("ascii")), "Georgia")
        self.assertIsNone(run.find(w("rPr")).find(w("b")))

    def test_run_like_copies_existing_format(self):
        self.edit([{"op": "insert_paragraph", "anchor": {"p": 4},
                    "runs": [{"text": "Prazo:", "like": {"p": 4, "text": "vence em 2023"}}, {"text": " curto."}]}])
        runs = self.para(5).findall(w("r"))
        self.assertIsNotNone(runs[0].find(w("rPr")).find(w("b")))
        self.assertIsNone(runs[1].find(w("rPr")).find(w("b")))


class TestDeleteParagraph(Base):
    def test_simple(self):
        self.edit([{"op": "delete_paragraph", "target": {"contains": "descartável"}}])
        self.assertNotIn("Parágrafo descartável.", self.paragraphs())
        self.assertEqual(self.compare()[0], 0)

    def test_section_break_protected(self):
        self.error([{"op": "delete_paragraph", "target": {"p": 11}}], "quebra de seção")

    def test_toc_paragraph_protected(self):
        self.error([{"op": "delete_paragraph", "target": {"p": 1}}], "atravessa parágrafos")

    def test_bookmark_survives_on_neighbor(self):
        _, report = self.edit([{"op": "delete_paragraph", "target": {"p": 3}}])
        root = Package(self.out).xml("word/document.xml")
        self.assertEqual([b.get(w("name")) for b in root.iter(w("bookmarkStart"))], ["_Toc100", "valor_alfa"])

    def test_comment_survives_on_neighbor(self):
        _, report = self.edit([{"op": "delete_paragraph", "target": {"p": 5}}])
        self.assertTrue(any("comentário" in x for x in report["warnings"]))
        root = Package(self.out).xml("word/document.xml")
        self.assertEqual(len(list(root.iter(w("commentReference")))), 1)
        self.assertEqual(len(list(root.iter(w("commentRangeStart")))), 1)

    def test_footnote_text_removed_with_its_reference(self):
        self.edit([{"op": "delete_paragraph", "target": {"p": 5}}])
        notes = Package(self.out).xml("word/footnotes.xml")
        self.assertEqual([n.get(w("id")) for n in notes.findall(w("footnote"))], ["-1", "0"])

    def test_image_requires_force_and_media_is_purged(self):
        self.error([{"op": "delete_paragraph", "target": {"p": 10}}], "force")
        self.edit([{"op": "delete_paragraph", "target": {"p": 10}, "force": True}])
        pkg = Package(self.out)
        self.assertFalse(pkg.has("word/media/image1.png"))
        self.assertFalse(any(r.get("Id") == "rId21" for r in pkg.rels("word/document.xml")))

    def test_only_paragraph_in_cell_is_emptied_not_removed(self):
        _, report = self.edit([{"op": "delete_paragraph", "target": {"table": 0, "row": 2, "cell": 1}}])
        self.assertTrue(any("esvaziado" in x for x in report["warnings"]))
        tbl = next(Package(self.out).xml("word/document.xml").iter(w("tbl")))
        cell = tbl.findall(w("tr"))[2].findall(w("tc"))[1]
        self.assertEqual(len(cell.findall(w("p"))), 1)
        self.assertEqual(ParagraphText(cell.find(w("p"))).text, "")

    def test_only_paragraph_in_footer_is_kept(self):
        self.edit([{"op": "delete_paragraph", "target": {"part": "word/header1.xml", "p": 0}}])
        self.assertEqual(self.paragraphs(part="word/header1.xml"), [""])

    def test_textbox_delete_is_mirrored(self):
        self.edit([{"op": "replace_text", "find": "Texto na caixa", "replace": ""}])
        with zipfile.ZipFile(self.out) as zf:
            self.assertNotIn("Texto na caixa", zf.read("word/document.xml").decode("utf-8"))


class TestTables(Base):
    def _table(self, path=None):
        return next(Package(path or self.out).xml("word/document.xml").iter(w("tbl")))

    def test_insert_row_clones_format_inside_merge(self):
        self.edit([{"op": "insert_table_row", "table": 0, "row": 1, "cells": [None, "Gama", "30,5"]}])
        rows = self._table().findall(w("tr"))
        self.assertEqual(len(rows), 4)
        cells = rows[2].findall(w("tc"))
        self.assertIsNotNone(cells[0].find(w("tcPr")).find(w("vMerge")), "continua a mesclagem")
        self.assertIsNone(cells[0].find(w("tcPr")).find(w("vMerge")).get(w("val")))
        self.assertEqual(cells[2].find(w("p")).find(w("pPr")).find(w("jc")).get(w("val")), "right")
        self.assertEqual(cells[2].find(w("p")).find(w("r")).find(w("rPr")).find(w("color")).get(w("val")), "1F4E79")
        self.assertEqual(list(rows[2].iter(w("bookmarkStart"))), [], "indicador do modelo não é clonado")
        self.assertIsNone(rows[2].get("{http://schemas.microsoft.com/office/word/2010/wordml}paraId"))

    def test_insert_row_after_merge_end_does_not_extend_merge(self):
        self.edit([{"op": "insert_table_row", "table": 0, "row": 2, "like_row": 1, "cells": ["Grupo B", "Delta", "40,0"]}])
        row = self._table().findall(w("tr"))[3]
        first = row.findall(w("tc"))[0]
        self.assertIsNone(first.find(w("tcPr")).find(w("vMerge")))
        self.assertEqual(ParagraphText(first.find(w("p"))).text, "Grupo B")

    def test_header_row_clone_drops_repeat_flag(self):
        _, report = self.edit([{"op": "insert_table_row", "table": 0, "row": 0, "cells": ["a", "b", "c"]}])
        new = self._table().findall(w("tr"))[1]
        self.assertTrue(any("cabeçalho" in x for x in report["warnings"]))
        self.assertIsNone(new.find(w("trPr")))

    def test_wrong_cell_count_refused(self):
        self.error([{"op": "insert_table_row", "table": 0, "row": 1, "cells": ["a", "b"]}], "3 células")

    def test_content_in_merged_cell_refused(self):
        self.error([{"op": "insert_table_row", "table": 0, "row": 1, "cells": ["x", "y", "z"]}], "mesclagem vertical")

    def test_delete_row_transfers_vertical_merge(self):
        self.edit([{"op": "delete_table_row", "table": 0, "row": 1}])
        rows = self._table().findall(w("tr"))
        self.assertEqual(len(rows), 2)
        heir = rows[1].findall(w("tc"))[0]
        self.assertEqual(heir.find(w("tcPr")).find(w("vMerge")).get(w("val")), "restart")
        self.assertEqual(ParagraphText(heir.find(w("p"))).text, "Grupo A")
        names = [b.get(w("name")) for b in Package(self.out).xml("word/document.xml").iter(w("bookmarkStart"))]
        self.assertIn("valor_alfa", names, "indicador da linha removida precisa sobreviver")

    def test_insert_table_clones_model(self):
        self.edit([{"op": "insert_table", "anchor": {"contains": "descartável"}, "like_table": 0,
                    "rows": [["A", "B", "C"], ["1", "2", "3"]]}])
        tables = list(Package(self.out).xml("word/document.xml").iter(w("tbl")))
        self.assertEqual(len(tables), 2)
        new = tables[1]
        self.assertEqual(new.find(w("tblPr")).find(w("tblStyle")).get(w("val")), "TabelaPadrao")
        rows = new.findall(w("tr"))
        self.assertIsNotNone(rows[0].find(w("trPr")).find(w("tblHeader")))
        self.assertIsNone(rows[1].findall(w("tc"))[0].find(w("tcPr")).find(w("vMerge")), "tabela nova não herda mesclagem")

    def test_delete_table_keeps_bookmark(self):
        self.edit([{"op": "delete_table", "table": 0}])
        root = Package(self.out).xml("word/document.xml")
        self.assertEqual(list(root.iter(w("tbl"))), [])
        self.assertIn("valor_alfa", [b.get(w("name")) for b in root.iter(w("bookmarkStart"))])

    def test_delete_only_row_refused(self):
        self.edit([{"op": "delete_table_row", "table": 0, "row": 2}, {"op": "delete_table_row", "table": 0, "row": 1}])
        self.src, self.out = self.out, os.path.join(self.dir, "out2.docx")
        self.error([{"op": "delete_table_row", "table": 0, "row": 0}], "única linha")


class TestImage(Base):
    def test_insert_image_registers_everything(self):
        image = os.path.join(self.dir, "novo.png")
        with open(image, "wb") as fh:
            fh.write(make_fixture.png_bytes(200, 100))
        _, report = self.edit([{"op": "insert_image", "anchor": {"p": 10}, "image": image, "width_cm": 5, "alt": "Gráfico"}])
        pkg = Package(self.out)
        self.assertTrue(pkg.has("word/media/image2.png"))
        drawing = list(self.para(11).iter(w("drawing")))[0]
        docpr = drawing[0].find("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr")
        self.assertEqual(docpr.get("descr"), "Gráfico")
        self.assertEqual(docpr.get("id"), "4", "id acima dos já usados no documento")
        ext = drawing[0].find("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent")
        self.assertEqual((int(ext.get("cx")), int(ext.get("cy"))), (1800000, 900000))
        code, cmp_report = self.compare("word/document.xml", "word/_rels/*", "word/media/*")
        self.assertEqual(code, 0, cmp_report["violacoes"])

    def test_oversized_image_capped_to_text_width(self):
        image = os.path.join(self.dir, "grande.png")
        with open(image, "wb") as fh:
            fh.write(make_fixture.png_bytes(5000, 1000))
        _, report = self.edit([{"op": "insert_image", "anchor": {"p": 10}, "image": image, "alt": "x"}])
        self.assertTrue(any("largura útil" in x for x in report["warnings"]))

    def test_unsupported_format(self):
        image = os.path.join(self.dir, "x.bmp")
        with open(image, "wb") as fh:
            fh.write(b"BM" + b"0" * 40)
        self.error([{"op": "insert_image", "anchor": {"p": 10}, "image": image}], "não suportada")


class TestTrackedChanges(Base):
    tracked_fixture = True

    def test_auto_mode_follows_document_setting(self):
        _, report = self.edit([{"op": "replace_text", "find": "Último parágrafo.", "replace": "Parágrafo final."}], author="Ana")
        self.assertTrue(report["track_changes"])
        p = self.para(-1)
        self.assertEqual(p.find(w("ins")).get(w("author")), "Ana")
        self.assertIsNotNone(p.find(w("del")))
        self.assertTrue(list(p.find(w("del")).iter(w("delText"))))
        self.assertEqual(list(p.find(w("del")).iter(w("t"))), [])

    def test_ids_do_not_collide_with_existing(self):
        self.edit([{"op": "replace_text", "find": "Último parágrafo.", "replace": "Final."}])
        ids = [el.get(w("id")) for el in Package(self.out).xml("word/document.xml").iter(w("ins"), w("del"))]
        self.assertEqual(len(ids), len(set(ids)))

    def test_word_level_granularity(self):
        self.edit([{"op": "replace_text", "find": "vence em 2023", "replace": "vence em 2030"}])
        ins = list(self.para(4).iter(w("ins")))
        self.assertEqual(["".join(t.text for t in i.iter(w("t"))) for i in ins], ["2030"])

    def test_insert_paragraph_marks_paragraph_and_run(self):
        self.edit([{"op": "insert_paragraph", "anchor": {"p": 7}, "text": "Novo;"}])
        p = self.para(8)
        self.assertIsNotNone(p.find(w("pPr")).find(w("rPr")).find(w("ins")))
        self.assertIsNotNone(p.find(w("ins")))

    def test_delete_paragraph_keeps_text_as_revision(self):
        self.edit([{"op": "delete_paragraph", "target": {"contains": "descartável"}}])
        with zipfile.ZipFile(self.out) as zf:
            self.assertIn("Parágrafo descartável.", zf.read("word/document.xml").decode("utf-8"))
        root = Package(self.out).xml("word/document.xml")
        self.assertIn("Parágrafo descartável.", ["".join(t.text for t in p.iter(w("delText"))) for p in root.iter(w("p"))])

    def test_delete_row_marks_row(self):
        self.edit([{"op": "delete_table_row", "table": 0, "row": 2}])
        tbl = next(Package(self.out).xml("word/document.xml").iter(w("tbl")))
        row = tbl.findall(w("tr"))[2]
        self.assertIsNotNone(row.find(w("trPr")).find(w("del")))

    def test_insert_row_marks_row(self):
        self.edit([{"op": "insert_table_row", "table": 0, "row": 2, "like_row": 2, "cells": [None, "Gama", "1"]}])
        tbl = next(Package(self.out).xml("word/document.xml").iter(w("tbl")))
        self.assertIsNotNone(tbl.findall(w("tr"))[3].find(w("trPr")).find(w("ins")))

    def test_pending_revision_of_other_author_refused(self):
        self.error([{"op": "replace_text", "find": "inserção", "replace": "x"}], "revisão pendente")

    def test_explicit_false_overrides_setting(self):
        _, report = self.edit([{"op": "replace_text", "find": "Último parágrafo.", "replace": "Fim."}], track_changes=False)
        self.assertFalse(report["track_changes"])
        self.assertIsNone(self.para(-1).find(w("ins")))

    def test_delete_image_paragraph_keeps_media_while_revision_pending(self):
        self.edit([{"op": "delete_paragraph", "target": {"p": 10}, "force": True}])
        self.assertTrue(Package(self.out).has("word/media/image1.png"))


class TestValidator(Base):
    def _mutate(self, part, old, new):
        path = os.path.join(self.dir, "bad.docx")
        with zipfile.ZipFile(self.src) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == part:
                    self.assertIn(old.encode(), data, old)
                    data = data.replace(old.encode(), new.encode(), 1)
                zout.writestr(item, data)
        return {f.code for f in validate(path) if f.severity == "error"}

    def test_detects_corruption_classes(self):
        doc = "word/document.xml"
        cases = [
            (doc, 'w:val="Titulo1"', 'w:val="Inexistente"', "STYLE_MISSING"),
            (doc, '<w:numId w:val="1"/>', '<w:numId w:val="9"/>', "NUM_ID_MISSING"),
            (doc, 'r:embed="rId21"', 'r:embed="rId99"', "RID_UNRESOLVED"),
            (doc, '<w:footnoteReference w:id="1"/>', '<w:footnoteReference w:id="7"/>', "NOTE_REF_MISSING"),
            (doc, '<w:commentReference w:id="0"/>', '<w:commentReference w:id="5"/>', "NOTE_REF_MISSING"),
            (doc, '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p w14:paraId="1A2B3C01"', '</w:p><w:p w14:paraId="1A2B3C01"', "FIELD_UNBALANCED"),
            (doc, 'w:pStyle w:val="ItemLista"/><w:numPr>', 'w:numPr>', None),
            (doc, '<w:jc w:val="both"/></w:pPr><w:r w:rsidR="00AA0002">', '<w:jc w:val="both"/><w:pStyle w:val="Normal"/></w:pPr><w:r w:rsidR="00AA0002">', "CHILD_ORDER"),
            (doc, '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>', '', "TABLE_NO_GRID"),
            (doc, '<w:delText>removido</w:delText>', '<w:t>removido</w:t>', "TEXT_IN_DELETION"),
            (doc, 'w14:paraId="1A2B3C01"', 'w14:paraId="FFFFFFFF"', "PARA_ID_RANGE"),
            (doc, 'mc:Ignorable="w14 w15"', 'mc:Ignorable="w14 w15 w16"', "MC_IGNORABLE_UNDECLARED"),
            ("word/_rels/document.xml.rels", 'Target="header1.xml"', 'Target="header9.xml"', "REL_DANGLING"),
            ("[Content_Types].xml", '<Default Extension="png" ContentType="image/png"/>', '', "CT_MISSING"),
            ("[Content_Types].xml", 'PartName="/word/header1.xml"', 'PartName="/word/header9.xml"', "CT_OVERRIDE_DANGLING"),
            ("word/numbering.xml", '<w:abstractNumId w:val="0"/>', '<w:abstractNumId w:val="3"/>', "NUM_ABSTRACT_MISSING"),
        ]
        for part, old, new, code in cases:
            with self.subTest(code=code, old=old[:40]):
                found = self._mutate(part, old, new)
                if code is None:
                    self.assertFalse(found - {"X"}, found)
                else:
                    self.assertIn(code, found)

    def test_detects_cell_without_paragraph_and_malformed_xml(self):
        self.assertIn("CELL_NO_PARAGRAPH", self._mutate(
            "word/document.xml", '<w:vMerge/></w:tcPr><w:p/></w:tc>', '<w:vMerge/></w:tcPr></w:tc>'))
        self.assertIn("XML_MALFORMED", self._mutate("word/styles.xml", "</w:styles>", ""))

    def test_baseline_ignores_inherited_but_catches_new(self):
        bad = os.path.join(self.dir, "herdado.docx")
        with zipfile.ZipFile(self.src) as zin, zipfile.ZipFile(bad, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = data.replace(b'w:val="Titulo1"', b'w:val="Fantasma"', 1)
                zout.writestr(item, data)
        code, _ = quiet(validate_docx.main, bad)
        self.assertEqual(code, 1)
        code, _ = quiet(validate_docx.main, bad, "--baseline", bad)
        self.assertEqual(code, 0)

    def test_update_fields_is_flagged(self):
        path = os.path.join(self.dir, "uf.docx")
        with zipfile.ZipFile(self.src) as zin, zipfile.ZipFile(path, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/settings.xml":
                    data = data.replace(b'<w:zoom', b'<w:updateFields w:val="true"/><w:zoom', 1)
                zout.writestr(item, data)
        self.assertIn("UPDATE_FIELDS_PROMPT", {f.code for f in validate(path)})


class TestCompare(Base):
    def _variant(self, part, old, new):
        with zipfile.ZipFile(self.src) as zin, zipfile.ZipFile(self.out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == part:
                    self.assertIn(old.encode(), data, old)
                    data = data.replace(old.encode(), new.encode(), 1)
                zout.writestr(item, data)

    def assert_violation(self, fragment):
        code, report = self.compare()
        self.assertEqual(code, 1, report)
        self.assertTrue(any(fragment in v for v in report["violacoes"]), report["violacoes"])

    def test_identical_is_ok(self):
        shutil.copy(self.src, self.out)
        code, report = self.compare()
        self.assertEqual(code, 0)
        self.assertEqual(report["partes"]["alteradas"], [])

    def test_catches_style_part_change(self):
        self._variant("word/styles.xml", 'w:val="32"', 'w:val="40"')
        self.assert_violation("word/styles.xml")

    def test_catches_properties_change(self):
        self._variant("docProps/core.xml", "Documento de teste", "Outro título")
        self.assert_violation("docProps/core.xml")

    def test_catches_format_drift_on_same_text(self):
        self._variant("word/document.xml", '<w:b/><w:sz w:val="22"/>', '<w:i/><w:sz w:val="22"/>')
        self.assert_violation("formatação")

    def test_catches_paragraph_property_drift(self):
        self._variant("word/document.xml", '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:commentRangeStart', '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:commentRangeStart')
        self.assert_violation("propriedades do parágrafo")

    def test_catches_section_change(self):
        self._variant("word/document.xml", 'w:w="16838" w:h="11906"', 'w:w="11906" w:h="16838"')
        self.assert_violation("seção")

    def test_catches_cell_property_change(self):
        self._variant("word/document.xml", 'w:fill="DDDDDD"', 'w:fill="FF0000"')
        self.assert_violation("célula")

    def test_catches_lost_image_in_same_text_paragraph(self):
        self._variant("word/document.xml", make_fixture.inline_image("rId21", 1, "Figura 1"), "")
        self.assert_violation("objeto")

    def test_allow_accepts_content_types_literally(self):
        self._variant("[Content_Types].xml", '<Default Extension="png"', '<Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="png"')
        self.assertEqual(self.compare("[Content_Types].xml")[0], 0)
        self.assert_violation("[Content_Types].xml")

    def test_reports_invented_format_as_attention(self):
        self.edit([{"op": "insert_paragraph", "anchor": {"p": 7},
                    "runs": [{"text": "Inventado", "italic": True, "underline": True}]}])
        code, report = self.compare()
        self.assertEqual(code, 0)
        self.assertTrue(any("sem precedente" in x for x in report["atencao"]))


class TestInspect(Base):
    def test_listing_shows_structure(self):
        code, output = quiet(inspect_docx.main, self.src)
        self.assertEqual(code, 0)
        self.assertIn('[3] Titulo1 {indicador} "Introdução"', output)
        self.assertIn("num=1:0", output)
        self.assertIn("SEÇÃO", output)
        self.assertIn("T0R1C1", output)

    def test_profile_captures_conventions(self):
        code, output = quiet(inspect_docx.main, self.src, "--profile")
        self.assertEqual(code, 0)
        data = json.loads(output)
        self.assertEqual(len(data["secoes"]), 2)
        self.assertEqual(data["secoes"][1]["orientacao"], "landscape")
        self.assertEqual(data["tabelas"][0]["estilo"], "TabelaPadrao")
        self.assertTrue(data["tabelas"][0]["cabecalho_repetido"])
        self.assertEqual(data["redacao"]["idioma_marcado"], "pt-BR")
        self.assertEqual(data["listas"][0]["formato"], "decimal")
        self.assertIn(";", data["redacao"]["fim_de_item_de_lista"])

    def test_runs_and_xml_views(self):
        code, output = quiet(inspect_docx.main, self.src, "--range", "4:5", "--runs")
        self.assertEqual(code, 0)
        self.assertIn("font=Georgia b 11pt", output)
        code, output = quiet(inspect_docx.main, self.src, "--xml", "3")
        self.assertIn("w:bookmarkStart", output)

    def test_other_parts(self):
        code, output = quiet(inspect_docx.main, self.src, "--part", "word/footer1.xml")
        self.assertEqual(code, 0)
        self.assertIn("Página", output)


class TestStyleBasedLists(Base):
    def test_numbering_inherited_from_style_is_reported(self):
        with zipfile.ZipFile(self.src) as zin:
            items = [(i, zin.read(i.filename)) for i in zin.infolist()]
        path = os.path.join(self.dir, "lista.docx")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info, data in items:
                if info.filename == "word/styles.xml":
                    data = data.replace(
                        b'<w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>',
                        b'<w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>'
                        b'<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>')
                if info.filename == "word/document.xml":
                    data = data.replace(b'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>', b'')
                zout.writestr(info, data)
        code, output = quiet(inspect_docx.main, path, "--profile")
        self.assertEqual(code, 0)
        data = json.loads(output)
        self.assertTrue(any(x["estilo"] == "ItemLista" and x["formato"] == "decimal" for x in data["listas"]), data["listas"])
        self.assertIn(";", data["redacao"]["fim_de_item_de_lista"])


class TestOtherStructures(Base):
    def _rewrite_fixture(self, old, new):
        path = os.path.join(self.dir, "var.docx")
        with zipfile.ZipFile(self.src) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    self.assertIn(old, data)
                    data = data.replace(old, new, 1)
                zout.writestr(item, data)
        self.src = path

    def test_simple_field_result_is_protected(self):
        self._rewrite_fixture(
            b'<w:r><w:t>Par\xc3\xa1grafo descart\xc3\xa1vel.</w:t></w:r>',
            b'<w:fldSimple w:instr=" DATE "><w:r><w:t>01/01/2024</w:t></w:r></w:fldSimple>')
        self.error([{"op": "replace_text", "find": "01/01/2024", "replace": "x"}], "resultado de campo")

    def test_page_break_blocks_match_and_survives(self):
        self._rewrite_fixture(
            b'<w:r><w:t>Par\xc3\xa1grafo descart\xc3\xa1vel.</w:t></w:r>',
            b'<w:r><w:t>antes</w:t><w:br w:type="page"/><w:t>depois</w:t></w:r>')
        self.error([{"op": "replace_text", "find": "antesdepois", "replace": "x"}], "atravessando um objeto")
        self.edit([{"op": "replace_text", "find": "antes", "replace": "primeiro"}])
        root = Package(self.out).xml("word/document.xml")
        self.assertEqual(len(list(root.iter(w("br")))), 1)

    def test_jpeg_registers_content_type(self):
        image = os.path.join(self.dir, "foto.jpg")
        write_bytes(image, b"\xff\xd8\xff\xc0\x00\x11\x08\x00\x64\x00\xc8" + b"\x00" * 20)
        self.edit([{"op": "insert_image", "anchor": {"p": 10}, "image": image, "width_cm": 4, "alt": "Foto"}])
        pkg = Package(self.out)
        self.assertTrue(pkg.has("word/media/image1.jpg"))
        extensions = [e.get("Extension") for e in pkg.xml("[Content_Types].xml")]
        self.assertIn("jpg", extensions)
        code, report = self.compare("word/document.xml", "word/media/*", "word/_rels/*", "[Content_Types].xml")
        self.assertEqual(code, 0, report["violacoes"])

    def test_tracked_delete_is_not_reported_as_insertion(self):
        self.edit([{"op": "delete_paragraph", "target": {"contains": "descartável"}}], track_changes=True)
        code, report = self.compare()
        self.assertEqual(code, 0, report["violacoes"])
        story = report["historias"][0]
        self.assertEqual(story["inseridos"], [])
        self.assertEqual([x["text"] for x in story["removidos"]], ["Parágrafo descartável."])

    def test_tracked_insert_has_no_false_alarm(self):
        self.edit([{"op": "insert_paragraph", "anchor": {"p": 7}, "text": "Novo;"}], track_changes=True)
        code, report = self.compare()
        self.assertEqual(code, 0, report["violacoes"])
        self.assertFalse([x for x in report["atencao"] if "sem precedente" in x], report["atencao"])


class TestRenderHelpers(Base):
    def test_missing_libreoffice_is_reported_not_hidden(self):
        missing = os.path.join(self.dir, "nao-existe.exe")
        code, output = quiet(render_docx.main, self.src, "--outdir", self.dir, "--soffice", missing)
        self.assertEqual(code, 3)
        self.assertIn("NÃO executado", output)

    def test_page_count(self):
        pdf = os.path.join(self.dir, "x.pdf")
        write_bytes(pdf, b"<< /Type /Pages /Count 2 >> << /Type /Page >> << /Type/Page >>")
        self.assertEqual(render_docx.page_count(pdf), 2)


class TestUnpackPack(Base):
    def test_roundtrip_is_byte_identical_per_part(self):
        folder = os.path.join(self.dir, "x")
        self.assertEqual(quiet(unpack_docx.main, self.src, folder)[0], 0)
        self.assertEqual(quiet(pack_docx.main, folder, self.out, "--original", self.src)[0], 0)
        with zipfile.ZipFile(self.src) as a, zipfile.ZipFile(self.out) as b:
            self.assertEqual(a.namelist(), b.namelist())
            for name in a.namelist():
                self.assertEqual(a.read(name), b.read(name), name)

    def test_pack_refuses_new_structural_errors(self):
        folder = os.path.join(self.dir, "x")
        quiet(unpack_docx.main, self.src, folder)
        doc = os.path.join(folder, "word", "document.xml")
        data = read_bytes(doc).replace(b'w:val="Titulo1"', b'w:val="Fantasma"', 1)
        write_bytes(doc, data)
        self.assertEqual(quiet(pack_docx.main, folder, self.out, "--original", self.src)[0], 1)
        self.assertFalse(os.path.exists(self.out))

    def test_pack_refuses_malformed_xml(self):
        folder = os.path.join(self.dir, "x")
        quiet(unpack_docx.main, self.src, folder)
        write_bytes(os.path.join(folder, "word", "styles.xml"), b"<a><b></a>")
        self.assertEqual(quiet(pack_docx.main, folder, self.out, "--original", self.src)[0], 2)

    def test_unpack_refuses_nonempty_destination(self):
        folder = os.path.join(self.dir, "x")
        os.makedirs(folder)
        write_bytes(os.path.join(folder, "f"), b"")
        self.assertEqual(quiet(unpack_docx.main, self.src, folder)[0], 2)


class TestPlanValidation(Base):
    def test_typos_are_rejected(self):
        self.error([{"op": "replace_txt", "find": "a"}], "operação desconhecida")
        self.error([{"op": "replace_text", "find": "a", "replac": "b"}], "chaves inválidas")
        self.error([{"op": "delete_paragraph", "target": {"paragraph": 1}}], "chaves desconhecidas")
        self.error([], "sem operações")

    def test_illegal_xml_chars_refused(self):
        self.error([{"op": "replace_text", "find": "Último parágrafo.", "replace": "a\x01b"}], "proibido")

    def test_ambiguous_target_lists_candidates(self):
        self.error([{"op": "delete_paragraph", "target": {"contains": "item"}}], "occurrence")

    def test_removed_paragraph_cannot_be_targeted_again(self):
        self.error([
            {"op": "delete_paragraph", "target": {"contains": "descartável"}},
            {"op": "delete_paragraph", "target": {"contains": "descartável"}},
        ], "já foi removido")


if __name__ == "__main__":
    unittest.main()
