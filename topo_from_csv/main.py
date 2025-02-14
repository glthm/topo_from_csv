#!/usr/bin/env python

import polars as pl


def read_csv(input_path: str) -> pl.DataFrame: ...

def remove_cols(df: pl.DataFrame, name: str, *more_names: str) -> pl.DataFrame: ...

def color_cols(
        df: pl.DataFrame,
        color_source_col: str,
        *color_target_cols: str
) -> pl.DataFrame: ...

def write_excel(output_path: str) -> None:
    """
    https://docs.pola.rs/api/python/stable/reference/api/polars.DataFrame.write_excel.html
    """
    ...

        
if __name__ == "__main__":
  print("Only stubs are defined for now")
