from pathlib import Path
import sys

import pandas as pd


# Zakladne priecinky projektu.
# Path(__file__) je cesta k tomuto skriptu.
# parents[1] znamena: chod o dva kroky vyssie, teda do hlavneho priecinka projektu.
PROJECT_DIR = Path(__file__).resolve().parents[1]
INPUT_DIR = PROJECT_DIR / "input"
OUTPUT_DIR = PROJECT_DIR / "output"
OUTPUT_FILE = OUTPUT_DIR / "hazmat_analysis_output.xlsx"


def find_excel_file():
    """Najde prvy Excel subor v zlozke input."""
    excel_files = []

    for file_path in INPUT_DIR.iterdir():
        is_excel = file_path.suffix.lower() in [".xlsx", ".xls"]
        is_temporary_excel_file = file_path.name.startswith("~$")

        if file_path.is_file() and is_excel and not is_temporary_excel_file:
            excel_files.append(file_path)

    if not excel_files:
        raise FileNotFoundError(
            "V zlozke input nie je ziaden Excel subor. "
            "Vloz tam subor s koncovkou .xlsx alebo .xls."
        )

    return sorted(excel_files)[0]


def analyze_workbook(excel_path):
    """Nacita Excel a pripravi jednoduche tabulky pre vystup."""
    sheets = pd.read_excel(excel_path, sheet_name=None)

    summary_rows = []
    column_rows = []

    for sheet_name, data in sheets.items():
        summary_rows.append(
            {
                "sheet_name": sheet_name,
                "row_count": len(data),
                "column_count": len(data.columns),
            }
        )

        for column_name in data.columns:
            non_empty_count = data[column_name].notna().sum()
            example_values = data[column_name].dropna().head(1).tolist()

            column_rows.append(
                {
                    "sheet_name": sheet_name,
                    "column_name": column_name,
                    "non_empty_count": int(non_empty_count),
                    "empty_count": int(len(data) - non_empty_count),
                    "example_value": example_values[0] if example_values else "",
                }
            )

    first_sheet_name = next(iter(sheets))
    preview = sheets[first_sheet_name].head(20).copy()

    # Tieto stlpce su zatial prazdne. Su pripravene pre buduce HAZMAT vyhodnotenie.
    preview["hazmat_status"] = ""
    preview["hazmat_reason"] = ""
    preview["hazmat_confidence"] = ""
    preview["recommended_action"] = ""

    return (
        pd.DataFrame(summary_rows),
        pd.DataFrame(column_rows),
        preview,
    )


def save_output(summary, columns, preview):
    """Ulozi vysledok do Excelu v zlozke output."""
    OUTPUT_DIR.mkdir(exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        columns.to_excel(writer, sheet_name="columns", index=False)
        preview.to_excel(writer, sheet_name="preview", index=False)


def main():
    print("Startujem analyzu Excelu...")

    try:
        excel_path = find_excel_file()
    except FileNotFoundError as error:
        print(f"Chyba: {error}")
        print("Tip: Skopiruj Excel subor do zlozky input a spusti skript znova.")
        sys.exit(1)

    print(f"Nacitavam subor: {excel_path.name}")

    summary, columns, preview = analyze_workbook(excel_path)
    save_output(summary, columns, preview)

    print("Hotovo.")
    print(f"Vystupny subor je ulozeny tu: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
