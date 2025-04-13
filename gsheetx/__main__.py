import fire

from .gsheetx import get_spreadsheet, get_sheets, get_sheet, get, set


def main():
    fire.Fire(
        dict(
            get_spreadsheet=get_spreadsheet,
            get_sheets=get_sheets,
            get_sheet=get_sheet,
            get=get,
            set=set,
        )
    )


if __name__ == "__main__":
    main()
