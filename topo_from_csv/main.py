#!/usr/bin/env python

import polars as pl
import sys
from pathlib import Path

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
                    "anchor"
                ],
                separator="\t",
                raise_if_empty=True
            )

def color_cols(
        df: pl.DataFrame,
        color_source_col: str,
        *color_target_cols: str
) -> pl.DataFrame: ...

def write_excel(output_path: str, df: pl.DataFrame) -> None:
    """
    https://docs.pola.rs/api/python/stable/reference/api/polars.DataFrame.write_excel.html
    """
    ...

def main():
    path = Path(sys.argv[1])
    df = pl.read_csv(
            path,
            separator="\t"
        )
    
    write_excel(path.with_suffix(".xls"), df)

if __name__ == "__main__":
  print("Only stubs are defined for now")
