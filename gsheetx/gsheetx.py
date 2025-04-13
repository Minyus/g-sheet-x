from datetime import datetime
import json
import yaml
from pathlib import Path
import sys
import typing as t

import gspread


def _get_spreadsheet(
    url: str | None = None,
    spreadsheet: str | None = None,
    folder: str | None = None,
):
    gc = gspread.oauth()
    if spreadsheet:
        try:
            return gc.open(spreadsheet)
        except gspread.exceptions.SpreadsheetNotFound:
            folder_id = folder.split(r":")[-1] if folder else None
            return gc.create(spreadsheet, folder_id=folder_id)
    if url:
        return gc.open_by_url(url)
    raise ValueError(repr(locals()))


def get_spreadsheet(
    url: str | None = None,
    spreadsheet: str | None = None,
    folder: str | None = None,
):
    return repr(_get_spreadsheet(url=url, spreadsheet=spreadsheet, folder=folder))


def _get_sheets(
    url: str | None = None,
    spreadsheet: str | None = None,
):
    ss = _get_spreadsheet(url, spreadsheet=spreadsheet)
    ws_ls = [ws.title for ws in ss.worksheets()]
    return ws_ls


def get_sheets(
    url: str | None = None,
    spreadsheet: str | None = None,
):
    return repr(_get_sheets(url=url, spreadsheet=spreadsheet))


def _get_sheet(
    url: str | None = None,
    spreadsheet: str | None = None,
    folder: str | None = None,
    sheet: str = "",
):
    ss = _get_spreadsheet(url=url, spreadsheet=spreadsheet, folder=folder)
    ws_ls = [ws.title for ws in ss.worksheets()]
    if not sheet:
        return ss.get_worksheet(0)
    if sheet in ws_ls:
        return ss.worksheet(sheet)
    return ss.add_worksheet(sheet, rows=0, cols=0)


def get_sheet(
    url: str | None = None,
    spreadsheet: str | None = None,
    folder: str | None = None,
    sheet: str = "",
):
    return repr(
        _get_sheet(url=url, spreadsheet=spreadsheet, folder=folder, sheet=sheet)
    )


def get(
    url: str | None = None,
    spreadsheet: str | None = None,
    folder: str | None = None,
    sheet: str = "",
    format: str = "",
    separator: str = ", ",
):
    ws = _get_sheet(url=url, spreadsheet=spreadsheet, folder=folder, sheet=sheet)
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
    url: str | None = None,
    spreadsheet: str | None = None,
    folder: str | None = None,
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
    ws = _get_sheet(url, spreadsheet=spreadsheet, folder=folder, sheet=sheet)
    ws.update(val_2d, cell)


def apply(
    url: str | None = None,
    spreadsheet: str | None = None,
    folder: str | None = None,
    sheet: str = "",
    template_sheet: str = "",
    delete_backup: bool = False,
):
    assert template_sheet
    ss = _get_spreadsheet(url, spreadsheet=spreadsheet, folder=folder)
    if sheet:
        sheets = [sheet]
    else:
        sheets = [s.title for s in ss.worksheets() if s.title != template_sheet]

    for sheet in sheets:
        n_sheets = len(ss.worksheets())
        ws = _get_sheet(url, spreadsheet=spreadsheet, folder=folder, sheet=sheet)
        val_2d = ws.get_all_values()
        ws_index = ws.index
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_sheet = sheet + "_" + timestamp
        backup_ws = ws.duplicate(
            new_sheet_name=backup_sheet, insert_sheet_index=n_sheets
        )
        ss.del_worksheet(ws)
        template_ws = _get_sheet(
            url, spreadsheet=spreadsheet, folder=folder, sheet=template_sheet
        )
        ws = template_ws.duplicate(new_sheet_name=sheet, insert_sheet_index=ws_index)
        ws.update(val_2d)
        if delete_backup:
            ss.del_worksheet(backup_ws)
