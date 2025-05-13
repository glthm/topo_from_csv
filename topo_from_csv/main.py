#!/usr/bin/env python

import polars as pl
import sys
from pathlib import Path
import xlsxwriter

pattern = 4
default_color = "#d8cfc8"

def grab_csv(path: Path|str) -> pl.DataFrame:
    """
    An interface for polars.read_csv that has the appropriate parameters for
     our use case, and selects only the columns we need.
    """
    return pl.read_csv(
                path,
                columns=[
                    "hold_colors",
                    "grade",
                    "name",
                    "description",
                    "openers",
                    "sector"
                ],
                separator="\t",
                missing_utf8_is_empty_string=True,
                raise_if_empty=True
            )


def refactor(df: pl.DataFrame) -> pl.DataFrame:
    """
    Helper function doing most of the refactoring that we can do on the raw csv.
    """

    return df.select(

            pl.when(    
                pl.col("sector").str.starts_with("Ligne ")
            ).then(
                pl.col("sector").str.slice(6)
            ).otherwise(
                pl.col("sector")
            ).alias(
                "Relais"
            ),
            
            pl.col("grade").alias("Cote"),

            pl.col("name").alias("Nom de la voie"),

            pl.col("openers").alias("Ouvreur·ses"),

            pl.col("description").alias("Description et commentaires")
        )

def write_excel(output_path: Path|str, df: pl.DataFrame) -> None:
    """
    formatting:
        A DataFrame whose columns are included on the list of columns of `data`. For each
        specified column, we get the format of each cell of the Excel containing `data`
        from the format string provided by the corresponding cell in `formatting`.
    """
    # https://docs.pola.rs/api/python/stable/reference/api/polars.DataFrame.write_excel.html
    content = refactor(df)
    with xlsxwriter.Workbook(output_path) as wb:
        ws = wb.add_worksheet()
        content.write_excel(
                workbook=wb, worksheet=ws,
                header_format={
                    "bold": True,
                    "bottom": 2,
                    },
                autofit=True,
                
                )
        for i, (val_to_format, color_str) in enumerate(zip(content["Relais"], df["hold_colors"])):
            color_list = color_str.split(",")
            cell_format_dict = {
                        "bg_color": color_list[0] or default_color,
                        "bold": True,
                        "align": "center_across"
                }
            if len(color_list) >= 2:
                # we only take the first 2 colors of colors list
                cell_format_dict |= {
                        "diag_border": 5,
                        "diag_type": 1,
                        "diag_color": color_list[1]  # this color isn't displayed
                    }
            ws.write(i+1, 0, val_to_format, wb.add_format(cell_format_dict))

def main():
    path = Path(sys.argv[1])
    write_excel(path.with_suffix(".xlsx"), grab_csv(path))
    
if __name__ == "__main__":
    try:
        main()
    except:
        print("Usage: python main.py path/to/your/csv")
        raise
