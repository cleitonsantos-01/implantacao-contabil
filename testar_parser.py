"""
Testa o parser de PDF com os balancetes reais.
Execute: python testar_parser.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

import re
from app import extract_pdf, format_br, cls_depth, find_leaf_accounts

PDF_DIR = os.path.join(os.path.dirname(__file__), "..")

def test_pdf(path: str, label: str, min_depth: int = 4):
    print(f"\n{'='*60}")
    print(f"Testando: {label}")
    print(f"Arquivo: {os.path.basename(path)}")

    if not os.path.exists(path):
        print(f"  ARQUIVO NAO ENCONTRADO: {path}")
        return

    records = extract_pdf(open(path, "rb"))
    print(f"Total registros extraidos: {len(records)}")

    # Filtrar por profundidade
    by_depth = [r for r in records if cls_depth(r.get("classification", "")) >= min_depth]
    print(f"Contas com classif. >= {min_depth}: {len(by_depth)}")

    # Filtrar apenas folhas (sem filhas)
    leaves = find_leaf_accounts(records)
    filtered = [r for r in records if r.get("classification", "") in leaves]
    print(f"Contas analiticas (folhas): {len(filtered)}")

    # Mostrar primeiras 5
    print("\nPrimeiros 5 registros:")
    for r in filtered[:5]:
        print(
            f"  [{r['current_indicator']}] {r['code']!s:>6} | "
            f"{r['classification']!s:25} | "
            f"{r['description'][:40]:40} | "
            f"R$ {format_br(r['current_value'])}"
        )

    # Totais
    tot_d = sum(r["current_value"] for r in filtered if r["current_indicator"] == "D")
    tot_c = sum(r["current_value"] for r in filtered if r["current_indicator"] == "C")
    print(f"\nTotal Déb.: R$ {format_br(tot_d)}")
    print(f"Total Cré.: R$ {format_br(tot_c)}")
    print(f"Diferença:  R$ {format_br(abs(tot_d - tot_c))}")
    print("Balanceado: " + ("SIM OK" if abs(tot_d - tot_c) < 1 else "NAO (diferenca esperada se so folhas)"))


if __name__ == "__main__":
    test_pdf(
        os.path.join(PDF_DIR, "Balancete enviado pelo antigo contador 2026  - HAB.pdf"),
        "HAB - Contador Anterior",
        min_depth=4,
    )
    test_pdf(
        os.path.join(PDF_DIR, "Balancete enviado pelo antigo contador 2026 - NCS.pdf"),
        "NCS - Contador Anterior",
        min_depth=4,
    )
    test_pdf(
        os.path.join(PDF_DIR, "Balancete enviado pelo antigo contador 2026 - PKS.pdf"),
        "PKS - Contador Anterior",
        min_depth=4,
    )
    print("\n" + "="*60)
    print("Teste concluído.")
