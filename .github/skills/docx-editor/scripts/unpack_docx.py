#!/usr/bin/env python3
"""Extrai um .docx para edição manual do XML (rota de exceção).

Uso:
    python unpack_docx.py doc.docx pasta/

Os arquivos saem byte a byte, sem pretty-print: reformatar o XML acrescentaria
espaços que podem virar texto visível ou alterar o comportamento do documento.
Depois de editar, use pack_docx.py.
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_lib import DocxError, Package, setup_stdio  # noqa: E402


def main(argv=None):
    setup_stdio()
    parser = argparse.ArgumentParser(description="Extrai um .docx para uma pasta.")
    parser.add_argument("input")
    parser.add_argument("output_dir")
    args = parser.parse_args(argv)
    try:
        pkg = Package(args.input)
        base = os.path.abspath(args.output_dir)
        if os.path.isdir(base) and os.listdir(base):
            raise DocxError(f"A pasta de destino não está vazia: {base}")
        for name in pkg.names():
            if name.endswith("/"):
                continue
            target = os.path.abspath(os.path.join(base, *name.split("/")))
            if not target.startswith(base + os.sep):
                raise DocxError(f"Entrada ZIP aponta para fora da pasta de destino: {name}")
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with open(target, "wb") as fh:
                fh.write(pkg.data[name])
    except DocxError as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 2
    print(f"{len(pkg.names())} partes extraídas em {base}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
