import fire

from .gsheetx import get, set


def main():
    fire.Fire(
        dict(
            get=get,
            set=set,
        )
    )


if __name__ == "__main__":
    main()
