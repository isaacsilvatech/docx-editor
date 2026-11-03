#!/usr/bin/env python3
"""Remonta um .docx a partir de uma pasta, seguindo a ordem do pacote original.

Uso:
    python pack_docx.py pasta/ saida.docx --original original.docx

Mantém a ordem e os metadados das entradas ZIP e não cria entradas de diretório.
Recusa a gravação se a validação encontrar erros novos em relação ao original.
"""
from __future__ import annotations

import argparse
import os
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_lib import DocxError, clone_zipinfo, parse_xml, regressions, setup_stdio, validate  # noqa: E402


def collect(folder):
    found = {}
    for dirpath, _, filenames in os.walk(folder):
        for filename in filenames:
            full = os.path.join(dirpath, filename)
            found[os.path.relpath(full, folder).replace(os.sep, "/")] = full
    return found


def build(folder, original, out_path):
    files = collect(folder)
    with zipfile.ZipFile(original) as zin:
        infos = zin.infolist()
    known = {i.filename for i in infos}
    with zipfile.ZipFile(out_path, "w", allowZip64=True) as zout:
        for info in infos:
            if info.filename in files:
                zout.writestr(clone_zipinfo(info), _read(files[info.filename]))
        for name in sorted(set(files) - known):
            zout.write(files[name], name, compress_type=zipfile.ZIP_DEFLATED)
    return sorted(known - set(files)), sorted(set(files) - known)


def _read(path):
    with open(path, "rb") as fh:
        data = fh.read()
    if path.endswith((".xml", ".rels")):
        try:
            parse_xml(data)
        except Exception as exc:
            raise DocxError(f"XML malformado em {path}: {exc}") from exc
    return data


def main(argv=None):
    setup_stdio()
    parser = argparse.ArgumentParser(description="Remonta um .docx a partir de uma pasta.")
    parser.add_argument("input_dir")
    parser.add_argument("output")
    parser.add_argument("--original", required=True, help="docx original, usado para ordem e como referência")
    args = parser.parse_args(argv)
    if os.path.abspath(args.output) == os.path.abspath(args.original):
        print("ERRO: a saída não pode ser o próprio original.", file=sys.stderr)
        return 2
    fd, tmp = tempfile.mkstemp(suffix=".docx", dir=os.path.dirname(os.path.abspath(args.output)))
    os.close(fd)
    try:
        removed, added = build(args.input_dir, args.original, tmp)
        new = regressions(validate(tmp), validate(args.original))
        for finding in new:
            print(finding)
        if any(f.severity == "error" for f in new):
            print("ERRO: saída descartada por erros estruturais novos.", file=sys.stderr)
            return 1
        os.replace(tmp, args.output)
    except DocxError as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 2
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    print(f"Gerado {args.output}. Partes removidas: {removed or '-'}; adicionadas: {added or '-'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
