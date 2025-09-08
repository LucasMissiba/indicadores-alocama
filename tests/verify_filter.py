from pathlib import Path
import sys

import pandas as pd

try:
    import app
except Exception as e:
    print(f"Falha ao importar app: {e}")
    sys.exit(1)


def main() -> None:
    project_root = Path.cwd()
    base_dir = project_root / "GRUPO SOLAR"
    groups = ["HOSPITALAR", "SOLAR", "DOMMUS"]

    files = []
    per_group_counts = {}
    for g in groups:
        dir_g = base_dir / g
        if dir_g.exists() and dir_g.is_dir():
            group_files = app.list_excel_files(dir_g, recursive=True)
            files.extend(group_files)
            per_group_counts[g] = len(group_files)
        else:
            per_group_counts[g] = 0

    print("=== Diagnóstico por grupo ===")
    print(per_group_counts)
    print(f"Arquivos totais encontrados: {len(files)}")

    if not files:
        print("Nenhum arquivo encontrado. Verifique as pastas e extensões (.xlsx/.xlsm).")
        return

    # Usa a coluna 'E' conforme solicitado
    target_column = "E"
    df_result, ignored_files, error_files = app.count_items_in_files(files, target_column, base_dir)

    print("=== Arquivos ignorados (sem coluna) ===")
    print(ignored_files)
    print("=== Arquivos com erro de leitura ===")
    print(error_files)

    print("=== Amostra do resultado (por arquivo) ===")
    print(df_result.head(10).to_string(index=False))

    if df_result.empty:
        print("Resultado vazio após contagem.")
        return

    df_totais = (
        df_result.groupby("Item", as_index=False)["Quantidade"].sum().sort_values("Quantidade", ascending=False)
    )
    print("=== Top 10 itens por total ===")
    print(df_totais.head(10).to_string(index=False))

    # Teste do filtro de produtos: pega os 2 primeiros itens e filtra
    top_items = df_totais["Item"].head(2).tolist()
    df_filtered = df_result[df_result["Item"].astype(str).str.strip().isin(top_items)]
    print("=== Itens escolhidos para filtro (2 primeiros) ===")
    print(top_items)
    print(f"Linhas antes do filtro: {len(df_result)} | Após filtro: {len(df_filtered)}")


if __name__ == "__main__":
    main()


