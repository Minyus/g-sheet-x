import fire

from .gsheetx import get_spreadsheet, get_sheets, get_sheet, get, update, apply


def main():
    fire.Fire(
        dict(
            get_spreadsheet=get_spreadsheet,
            get_sheets=get_sheets,
            get_sheet=get_sheet,
            get=get,
            update=update,
            apply=apply,
        )
    )


if __name__ == "__main__":
    main()
