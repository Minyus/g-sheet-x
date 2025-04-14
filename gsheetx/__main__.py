import fire

from .gsheetx import (
    get_spreadsheet,
    list_sheets,
    get_sheet,
    get,
    update,
    apply,
    copy_spreadsheet,
)


def main():
    fire.Fire(
        dict(
            get_spreadsheet=get_spreadsheet,
            list_sheets=list_sheets,
            get_sheet=get_sheet,
            get=get,
            update=update,
            apply=apply,
            copy_spreadsheet=copy_spreadsheet,
        )
    )


if __name__ == "__main__":
    main()
