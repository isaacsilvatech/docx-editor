#!/usr/bin/env python3
"""Valida a estrutura de um .docx: os defeitos que levam ao aviso de reparo do Word.

Uso:
    python validate_docx.py saida.docx --baseline original.docx [--strict] [--json]

Com --baseline, só contam os achados NOVOS: documentos reais costumam ter
pequenas irregularidades herdadas, que não são responsabilidade da edição.

Saída: 0 = ok; 1 = erros novos (ou avisos novos, com --strict); 2 = não foi possível ler.
"""
from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_lib import regressions, setup_stdio, validate  # noqa: E402


def main(argv=None):
    setup_stdio()
    parser = argparse.ArgumentParser(description="Valida a estrutura de um .docx.")
    parser.add_argument("input")
    parser.add_argument("--baseline", help="documento original: ignora o que já existia nele")
    parser.add_argument("--strict", action="store_true", help="avisos novos também reprovam")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    findings = validate(args.input)
    if any(f.code == "PACKAGE_UNREADABLE" for f in findings):
        print(findings[0], file=sys.stderr)
        return 2
    inherited = 0
    if args.baseline:
        baseline = validate(args.baseline)
        new = regressions(findings, baseline)
        inherited = len(findings) - len(new)
        findings = new
    errors = [f for f in findings if f.severity == "error"]
    warnings = [f for f in findings if f.severity == "warning"]
    infos = [f for f in findings if f.severity == "info"]
    failed = bool(errors or (args.strict and warnings))
    if args.json:
        print(json.dumps({
            "ok": not failed,
            "findings": [dict(severity=f.severity, code=f.code, part=f.part, detail=f.detail, where=f.where)
                         for f in findings],
            "inherited": inherited,
        }, ensure_ascii=False, indent=1))
    else:
        for finding in errors + warnings + infos:
            print(finding)
        suffix = f" ({inherited} herdado(s) do original, ignorado(s))" if args.baseline else ""
        verdict = "REPROVADO" if failed else "OK"
        print(f"{verdict}: {len(errors)} erro(s), {len(warnings)} aviso(s), {len(infos)} nota(s){suffix}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
