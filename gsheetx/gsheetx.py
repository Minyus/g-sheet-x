from datetime import datetime
import json
import yaml
from pathlib import Path
import sys
import typing as t

import gspread
from gspread.utils import ValueRenderOption, ValueInputOption


def _show(ss: gspread.Spreadsheet):
    print(json.dumps(dict(url=ss.url, title=ss.title), indent=2))


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
        if url.startswith("http"):
            return gc.open_by_url(url)
        else:
            return gc.open_by_key(url)
    raise ValueError(repr(locals()))


def get_spreadsheet(
    url: str | None = None,
    spreadsheet: str | None = None,
    folder: str | None = None,
):
    return _show(_get_spreadsheet(url=url, spreadsheet=spreadsheet, folder=folder))


def _get_sheets(
    url: str | None = None,
    spreadsheet: str | None = None,
):
    ss = _get_spreadsheet(url, spreadsheet=spreadsheet)
    ws_ls = [ws.title for ws in ss.worksheets()]
    return ws_ls


def list_sheets(
    url: str | None = None,
    spreadsheet: str | None = None,
):
    return "\n".join(_get_sheets(url=url, spreadsheet=spreadsheet))


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
    render: str = "formula",
    col: str = "",
    col_sep: str = "\t",
    filter: str = "",
    filter_sep: str = "\t",
    filter_col: str = "",
    format: str = "sv",
    separator: str = "\t",
    fill_strategy: str = "",
):
    vro = getattr(ValueRenderOption, render)
    ws = _get_sheet(url=url, spreadsheet=spreadsheet, folder=folder, sheet=sheet)
    dicts = ws.get_all_records(value_render_option=vro)

    if filter or col or fill_strategy:
        filter_vals = None
        if filter:
            assert filter_sep
            assert filter_col
            filter_vals = set(filter.split(filter_sep))
        cols = []
        if col:
            assert col_sep
            cols = col.split(col_sep)
        dicts = [
            ({c: d.get(c) for c in cols} if cols else d)
            for d in dicts
            if (filter_vals is None or d[filter_col] in filter_vals)
        ]
        if fill_strategy:
            import polars as pl

            df = pl.DataFrame(dicts)
            for c in df.columns:
                df = df.with_columns(
                    pl.col(c)
                    .replace("", None)
                    .fill_null(strategy=fill_strategy)
                    .replace(None, "")
                )
            dicts = df.to_dicts()
    elif format == "sv":
        val_2d = ws.get_all_values(value_render_option=vro)
        return "\n".join(
            [separator.join([str(v) for v in val_1d]) for val_1d in val_2d]
        )

    if format == "sv":
        val_2d = [list(dicts[0].keys())] + [list(d.values()) for d in dicts]
        return "\n".join(
            [separator.join([str(v) for v in val_1d]) for val_1d in val_2d]
        )

    if format == "json":
        return json.dumps(dicts, indent=2)
    if format == "yaml":
        return yaml.dump(dicts)
    if format == "report":
        return "\n".join(
            [
                separator.join(
                    [
                        f"{k} {v:+}" if isinstance(v, float) else f"{k} {v}"
                        for (k, v) in d.items()
                    ]
                )
                for d in dicts
            ]
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


def _get_all_values(ws: gspread.Worksheet):
    vals = ws.get_all_values()
    if len(vals) == 1 and not vals[0]:
        return []
    return vals


def update(
    url: str | None = None,
    spreadsheet: str | None = None,
    folder: str | None = None,
    sheet: str = "",
    file: str = "",
    content: str | dict = "",
    separator: str = "\t",
    extra: dict | None = None,
    cell: str = "",
    parse: bool = False,
    verbose: bool = False,
):
    if content:
        text = content
    elif file:
        text = Path(file).read_text()
    else:
        text = "\n".join(sys.stdin).strip()
        try:
            text = eval(text)
        except Exception:
            pass
    assert text
    ws = _get_sheet(url, spreadsheet=spreadsheet, folder=folder, sheet=sheet)
    ll = None
    if not (cell and cell[-1] in {str(i) for i in range(10)}):
        ll = _get_all_values(ws)
        next_row = len(ll) + 1
        first_col = cell or "A"
        cell = f"{first_col}{next_row}"
    if isinstance(text, (dict, list)):
        if isinstance(text, list):
            val_dicts = text
        else:
            val_dicts = [text]
        ll = ll or _get_all_values(ws)
        if ll and ll[0]:
            keys = ll[0]
            header = []
        else:
            keys = list(val_dicts[0].keys())
            header = [keys]
        if extra:
            assert isinstance(extra, dict)
            for val_d in val_dicts:
                val_d.update(extra)
        val_2d = header + [[str(val_d.get(k, "")) for k in keys] for val_d in val_dicts]
    else:
        val_2d = [row.split(separator) for row in text.split("\n")]
    if verbose:
        for val_1d in val_2d:
            print(separator.join(val_1d))
    ws.update(val_2d, cell, raw=not parse)


def apply(
    url: str | None = None,
    spreadsheet: str | None = None,
    folder: str | None = None,
    template_sheet: str = "",
    sheet: str = "",
    sheet_sep: str = "\n",
    render: str = "formula",
    delete_backup: bool = False,
    to_copy: bool = False,
    copy_as: str | None = None,
    copy_prefix: str = "",
    copy_postfix: str = " copied",
):
    vro = getattr(ValueRenderOption, render)
    assert template_sheet
    if to_copy:
        ss = _copy_spreadsheet(
            url,
            spreadsheet=spreadsheet,
            folder=folder,
            copy_as=copy_as,
            copy_prefix=copy_prefix,
            copy_postfix=copy_postfix,
        )
        url = ss.url
        spreadsheet = None
        _show(ss)

    ss = _get_spreadsheet(url, spreadsheet=spreadsheet, folder=folder)
    if sheet:
        sheets = sheet.split(sheet_sep)
    else:
        sheets = [s.title for s in ss.worksheets() if s.title != template_sheet]

    for sheet in sheets:
        n_sheets = len(ss.worksheets())
        ws = _get_sheet(url, spreadsheet=spreadsheet, folder=folder, sheet=sheet)
        val_2d = ws.get_all_values(value_render_option=vro)
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
        ws.update(val_2d, value_input_option=ValueInputOption.user_entered)
        if delete_backup or to_copy:
            ss.del_worksheet(backup_ws)


def _copy_spreadsheet(
    url: str | None = None,
    spreadsheet: str | None = None,
    copy_as: str | None = None,
    copy_prefix: str = "",
    copy_postfix: str = " copied",
    folder: str | None = None,
    copy_permissions: bool = False,
    skip_comments: bool = False,
):
    gc = gspread.oauth()

    if spreadsheet:
        ss = gc.open(spreadsheet)
    elif url and url.startswith("http"):
        ss = gc.open_by_url(url)
    else:
        assert url
        ss = gc.open_by_key(url)
    ss_id = ss.id
    ss_title = ss.title
    title = copy_as if copy_as else (copy_prefix + ss_title + copy_postfix)
    folder_id = folder.split(r":")[-1] if folder else None
    return gc.copy(
        file_id=ss_id,
        title=title,
        copy_permissions=copy_permissions,
        folder_id=folder_id,
        copy_comments=not skip_comments,
    )


def copy_spreadsheet(
    url: str | None = None,
    spreadsheet: str | None = None,
    copy_as: str | None = None,
    copy_prefix: str = "",
    copy_postfix: str = " copied",
    folder: str | None = None,
    copy_permissions: bool = False,
    skip_comments: bool = False,
):
    return _show(
        _copy_spreadsheet(
            url=url,
            spreadsheet=spreadsheet,
            copy_as=copy_as,
            copy_prefix=copy_prefix,
            copy_postfix=copy_postfix,
            folder=folder,
            copy_permissions=copy_permissions,
            skip_comments=skip_comments,
        )
    )
