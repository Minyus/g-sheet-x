import json
import yaml
from pathlib import Path
import sys
import typing as t

import gspread


def _get_worksheet(
    url: str,
    sheet: str = "",
):
    gc = gspread.oauth()
    ss = gc.open_by_url(url)
    ws = ss.worksheet(sheet) if sheet else ss.get_worksheet(0)
    return ws


def get(
    url: str,
    sheet: str = "",
    format: str = "",
    separator: str = ", ",
):
    ws = _get_worksheet(url, sheet)
    if not format:
        val_2d = ws.get_all_values()
        return "\n".join([separator.join(val_1d) for val_1d in val_2d])

    dicts = ws.get_all_records()
    if format == "json":
        return json.dumps(dicts, indent=2)
    if format == "yaml":
        return yaml.dump(dicts)
    if format == "report":
        return "\n".join(
            [separator.join([f"{k} {v:+}" for (k, v) in d.items()]) for d in dicts]
        )
    if format == "polars":
        import polars as pl

        df = pl.DataFrame(dicts)
        return repr(df)

    import pandas as pd

    df = pd.DataFrame(dicts)
    if format == "pandas":
        return repr(df)
    if format == "markdown":
        return df.to_markdown(index=False)
    raise NotImplementedError(format)


def set(
    url: str,
    sheet: str = "",
    file: str = "",
    text: str = "",
    separator: str = "\t",
    cell: str = "A1",
):
    if text:
        pass
    elif file:
        text = Path(file).read_text()
    else:
        text = "\n".join(sys.stdin)
    assert text
    val_2d = [row.split(separator) for row in text.split("\n")]
    ws = _get_worksheet(url, sheet)
    ws.update(val_2d, cell)
