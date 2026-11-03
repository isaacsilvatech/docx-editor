#!/usr/bin/env python3
"""Converte um .docx em PDF com o LibreOffice: teste real de abertura e paginação.

Uso:
    python render_docx.py doc.docx --outdir pasta/ [--soffice CAMINHO] [--compare original.docx]

Saída: 0 = convertido; 1 = o LibreOffice falhou; 3 = LibreOffice não instalado.
"""
from __future__ import annotations

import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_lib import setup_stdio  # noqa: E402

DEFAULT_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
]


def find_soffice(explicit=None):
    if explicit:
        return explicit if os.path.exists(explicit) else None
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    return next((p for p in DEFAULT_PATHS if os.path.exists(p)), None)


def page_count(pdf_path):
    with open(pdf_path, "rb") as fh:
        data = fh.read()
    return len(re.findall(rb"/Type\s*/Page(?![a-zA-Z])", data))


def convert(soffice, docx, outdir):
    os.makedirs(outdir, exist_ok=True)
    with tempfile.TemporaryDirectory() as profile:
        uri = "file:///" + profile.replace(os.sep, "/").lstrip("/")
        cmd = [soffice, f"-env:UserInstallation={uri}", "--headless", "--norestore",
               "--convert-to", "pdf", "--outdir", outdir, docx]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300, errors="replace")
    stem = os.path.splitext(os.path.basename(docx))[0]
    pdf = os.path.join(outdir, stem + ".pdf")
    if proc.returncode != 0 or not os.path.exists(pdf) or os.path.getsize(pdf) == 0:
        raise RuntimeError(f"Conversão falhou ({proc.returncode}): {proc.stderr or proc.stdout}")
    return pdf


def main(argv=None):
    setup_stdio()
    parser = argparse.ArgumentParser(description="Converte um .docx em PDF com o LibreOffice.")
    parser.add_argument("input")
    parser.add_argument("--outdir", required=True)
    parser.add_argument("--soffice", help="caminho do executável do LibreOffice")
    parser.add_argument("--compare", help="converte também o original e compara o número de páginas")
    args = parser.parse_args(argv)
    soffice = find_soffice(args.soffice)
    if not soffice:
        print("LibreOffice não encontrado; teste de renderização NÃO executado. "
              "Informe isso ao usuário em vez de presumir sucesso.", file=sys.stderr)
        return 3
    try:
        pdf = convert(soffice, os.path.abspath(args.input), os.path.abspath(args.outdir))
        pages = page_count(pdf)
        print(f"{pdf}: {pages} página(s)")
        if args.compare:
            base_dir = os.path.join(os.path.abspath(args.outdir), "original")
            original = convert(soffice, os.path.abspath(args.compare), base_dir)
            before = page_count(original)
            print(f"{original}: {before} página(s)")
            if before != pages:
                print(f"ATENÇÃO: número de páginas mudou de {before} para {pages}; confira se o pedido justifica.")
    except (RuntimeError, subprocess.TimeoutExpired, OSError) as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
