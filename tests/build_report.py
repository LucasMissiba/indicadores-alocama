from pathlib import Path
import sys

import pandas as pd

try:
    import app
except Exception as e:
    print(f"Falha ao importar app: {e}")
    sys.exit(1)


def build_report(base_dir: Path, out_path: Path, target_col: str = "E") -> None:
    groups = []
    for sub in base_dir.iterdir():
        if sub.is_dir():
            groups.append(sub.name)

    excel_files = []
    for g in groups:
        excel_files.extend(app.list_excel_files(base_dir / g, recursive=True))

    if not excel_files:
        print("Nenhum arquivo encontrado para relatório.")
        return

    df_result, ignored_files, error_files, _ = app.count_items_in_files(
        excel_files, target_col, base_dir, use_smart=True, only_equipment=True
    )

    if df_result.empty:
        print("Sem dados após contagem.")
        return

    # Totais por item e por arquivo
    df_totais = (
        df_result.groupby("Item", as_index=False)["Quantidade"].sum().sort_values("Quantidade", ascending=False)
    )
    item_order = df_totais["Item"].tolist()
    df_result["Item"] = pd.Categorical(df_result["Item"], categories=item_order, ordered=True)
    df_by_file_sorted = df_result.sort_values(["Item", "Arquivo", "Quantidade"], ascending=[True, True, False])

    # Monta relatório com XlsxWriter
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        df_by_file_sorted.to_excel(writer, index=False, sheet_name="por_arquivo")
        df_totais.to_excel(writer, index=False, sheet_name="totais_por_item")

        # Gráfico de barras simples
        wb = writer.book
        ws = writer.sheets["totais_por_item"]
        chart = wb.add_chart({"type": "column"})
        # Assume cabeçalho na linha 0, dados começam na linha 1
        nrows = len(df_totais)
        chart.add_series({
            "name": "Quantidade",
            "categories": ["totais_por_item", 1, 0, nrows, 0],
            "values": ["totais_por_item", 1, 1, nrows, 1],
        })
        chart.set_title({"name": "Totais por Item"})
        chart.set_x_axis({"name": "Item"})
        chart.set_y_axis({"name": "Quantidade"})
        chart.set_legend({"none": True})
        ws.insert_chart("D2", chart, {"x_scale": 1.4, "y_scale": 1.4})

    print(f"Relatório gerado em: {out_path}")


def main() -> None:
    project_root = Path.cwd()
    base_dir = project_root / "GRUPO SOLAR"
    out_path = project_root / "relatorio_grupo_solar.xlsx"
    build_report(base_dir, out_path, target_col="E")


if __name__ == "__main__":
    main()


