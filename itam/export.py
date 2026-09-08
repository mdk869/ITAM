from io import BytesIO

import pandas as pd


def export_to_excel(df, filename="asset_data.xlsx"):
    """Export a sanitized copy so source text cannot become Excel formulas."""
    export_df = df.copy()
    for column in export_df.select_dtypes(include=["object", "string"]).columns:
        export_df[column] = export_df[column].map(
            lambda value: "'" + value if isinstance(value, str) and value.startswith(("=", "+", "-", "@")) else value
        )
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Assets")
    output.seek(0)
    return output
