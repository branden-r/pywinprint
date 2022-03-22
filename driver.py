from pathlib import Path
from printer import send
from typing import Final
from win32print import EnumPrinters, GetDefaultPrinter, SetDefaultPrinter  # type: ignore


PRINTER_NAME_IDX: Final[int] = 2
PRINTERS: Final[list[str]] = [printer[PRINTER_NAME_IDX] for printer in EnumPrinters(PRINTER_NAME_IDX)]
EXPECTED_PRINTER: Final[str | None] = None
PRINTABLE_EXTS: Final[tuple[str, ...]] = (".pdf",)
MSG: Final[str] = "SENT $path TO $printer"


def get_printer_name() -> str | None:
    """
    Figure out which printer to use.
    Will prompt the user to make a selection if no valid expected printer is provided (see global constants).
    :returns: None
    """
    printer: str | None
    if EXPECTED_PRINTER and EXPECTED_PRINTER in PRINTERS:
        printer = EXPECTED_PRINTER
    else:
        print(f"EXPECTED PRINTER NOT FOUND ({EXPECTED_PRINTER})") if EXPECTED_PRINTER else None
        print("CHOOSE PRINTER OR QUIT")
        no: int
        choice: str
        for no, choice in enumerate(PRINTERS + ["QUIT"]):
            print(f"{no}. {choice}")
        try:
            printer = PRINTERS[int(input())]
        except (EOFError, ValueError, IndexError):
            printer = None
    return printer


def main() -> None:
    """
    Prints out everything in the same directory and subdirectories.
    :returns: None
    """
    printer_name: str | None = get_printer_name()
    if printer_name:
        path: Path
        paths: tuple[Path, ...] = tuple(path for path in Path(".").rglob("*") if path.suffix in PRINTABLE_EXTS)
        send(*paths, printer_name=printer_name, msg=MSG)
    else:
        print("NO PRINTER SELECTED")
    print("WORK COMPLETE")
    input("PRESS ENTER TO QUIT\n")


if __name__ == "__main__":
    main()
