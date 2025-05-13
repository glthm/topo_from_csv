#!/usr/bin/env python

import polars as pl
from typing import Iterable
import sys
from pathlib import Path
import xlsxwriter
from os import PathLike

# Configuration that you can change to match your needs
## The default color for routes that don't have any defined color
default_color = "#d8cfc8"
## The keys for which we want to output a page where the routes are ordered by
##  these keys, given here under their original name in the Oblyk CSV file
default_ordering_keys = ("sector", "grade")
## The default matching between the field names in the raw CSV exported from
##  the Oblyk dashboard and the output file. This also defines the fields that
##  will be in the final output. Note that 'sector' has a special behaviour as
##  it will always be present in the final output, even if not pressent here
default_field_names = {
    "sector": "Relais",
    "grade": "Cote",
    "name": "Nom de la voie",
    "description": "Description et commentaires",
    "openers": "Ouvreur⋅se⋅s"
}
## How to format the header line
default_header_format = {
    "bold": True,  # bold text
    "bottom": 2,  # thick bottom border
}
## How to format the diagonals on bi-color routes
default_diag_style = {
    "diag_border": 5,
    "diag_type": 1
}

def fname[T](key: T, d:dict[T, T]|None = None) -> T:
    """
    Properly typed wrapper for `d.get(key, key)`, that will return `d[key]` if
    defined and `key` otherwise.
    If `d` is None, `default_field_names` will be used as the input dictionary.
    Notes
    -----
    This function is mostly meant to be used for `d=None`, as using it with an 
    other dictionary could result in a longer code snippet than calling `d.get`
    directly.
    """
    if d is None:
        d = default_field_names
    return d.get(key, key)


def grab_csv(path: PathLike|str|bytes) -> pl.DataFrame:
    """
    An interface for polars.read_csv that has the appropriate parameters for
     our use case, and selects only the columns we need.
    """
    cols_to_grab = {
        "hold_colors",  # needed for processing the cell colors
        "sector",       # always required for now
    }.union(
        default_field_names.keys()  # columns that will be displayed in output
    )
    return pl.read_csv(
                path,
                columns=list(cols_to_grab),
                separator="\t",
                missing_utf8_is_empty_string=True,
                raise_if_empty=True
            )


def refactor(df: pl.DataFrame) -> pl.DataFrame:
    """
    Helper function doing most of the refactoring that we can do on the raw csv.
    """

    return df.select(
            # sector field
            pl.when(    
                pl.col("sector").str.starts_with("Ligne ")
            ).then(
                pl.col("sector").str.slice(6)
            ).otherwise(
                pl.col("sector")
            ).alias(
                fname("sector")
            ),
            # other fields that do not reuire extra processing
            *[
                pl.col(x).alias(fname(x))
                for x in default_field_names
                if x != "sector"
            ]
            
        )

def write_excel(output_path: PathLike|str|bytes, input_path: PathLike|str|bytes) -> None:
    """
    Get the CSV file from `input_path` and write the formatted output to the excel file `output_path`.
    """
    # https://docs.pola.rs/api/python/stable/reference/api/polars.DataFrame.write_excel.html
    df = grab_csv(input_path)
    content = refactor(df)
    with xlsxwriter.Workbook(output_path) as wb:
        for order_key in default_ordering_keys:
            fokey = fname(order_key)
            ws = wb.add_worksheet(f"Par {fokey}")
            df = df.sort(order_key)
            content = content.sort(fokey)
            content.write_excel(
                    workbook=wb, worksheet=ws,
                    header_format=default_header_format,
                    autofit=True,
                    )
            for i, (val_to_format, color_str) in enumerate(
                    zip(content[fname("sector")], df["hold_colors"])
                    ):
                color_list = color_str.split(",")
                cell_format_dict = {
                            "bg_color": color_list[0].strip() or default_color,
                            "bold": True,
                            "align": "center"
                    }
                if len(color_list) >= 2:
                    # we only take the first 2 colors of colors list
                    cell_format_dict = (
                        cell_format_dict |
                        {"diag_color": color_list[1].strip()} |
                        default_diag_style
                    )
                ws.write(i+1, 0, val_to_format, wb.add_format(cell_format_dict))

def main():
    path = Path(sys.argv[1])
    write_excel(path.with_suffix(".xlsx"), path)
    
if __name__ == "__main__":
    try:
        main()
        print(
            "Done ! You still have to",
            " - set the font color to a lighter tone on cells where the "
            "background color is too dark",
            " - check the ordering of your sectors if you ordered your lines "
            "by sector, as they can be a mix of numbers and letters",
            " - add the cell borders to the body of the tables",
            " - center the route grades in their cell if wanted",
            " - adapt the zoom and width of your table so that it is ready "
            "to print",
            sep="\n"
            )
    except:
        print("Usage: python main.py path/to/your/csv")
        raise
