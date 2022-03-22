from __future__ import annotations
from PIL import Image  # type: ignore
from PIL.ImageWin import Dib  # type: ignore
from os import startfile
from pathlib import Path
from pdf2image import convert_from_path  # type: ignore
from string import Template
from typing import Any, Callable, Final
from win32print import GetDefaultPrinter  # type: ignore
from win32ui import CreateDC  # type: ignore


class Printer:
    """
    Wrapper class for communicating with a printer on Windows.
    Should only be used as a context manager.
    """

    PAPER_WIDTH_IDX: Final[int] = 110  # magic numbers to get info from windows
    PAPER_HEIGHT_IDX: Final[int] = 111  # i hate these

    def __init__(self: Printer, name: str = "") -> None:
        """
        Stores the printer name if one was given, uses the default printer otherwise.
        Connects to the printer, then figures out the paper width, paper height, and handle.
        We use the handle later to send data to the printer.
        """
        self.name: str = name if name else GetDefaultPrinter()
        self.device: Any = CreateDC()
        self.device.CreatePrinterDC(self.name)
        self.paper_width: int = self.device.GetDeviceCaps(Printer.PAPER_WIDTH_IDX)
        self.paper_height: int = self.device.GetDeviceCaps(Printer.PAPER_HEIGHT_IDX)
        self.handle: Any = self.device.GetHandleOutput()

    def __enter__(self: Printer) -> Printer:
        """
        Returns a reference to the printer.
        Doesn't need to do anything else because we already did all the heavy lifting in __init__.
        """
        return self

    def __exit__(self: Printer, *_) -> None:
        """
        Kills the connection to the printer.
        We don't care about the other values Python gives us.
        """
        self.device.DeleteDC()

    def send(self: Printer, *paths: str | Path, msg: str = "", shell: bool = False) -> None:
        """
        Sending files to a printer prints them out as documents.
        Prints the files at the given paths.
        The argument "msg" is printed to the console after printing each document, if "msg" is given.
            This can be templated with $printer, $path, $document, and $stem.
        If shell is true, then we send files to the printer by starting the file with the print verb, otherwise we send
        files to the printer by using a device context.
            You probably don't want to use the shell -- it's not silent, and sometimes requires user input.
            The device context method, on the other hand, is silent.
        """
        print_document: Callable[[Path], None] = Printer.shell_print if shell else self.device_print
        for path in paths:
            if isinstance(path, str):
                path = Path(path)
            print_document(path)
            if msg:
                print(Template(msg).safe_substitute(printer=self.name, path=path, document=path.name, stem=path.stem))

    @staticmethod
    def shell_print(path: Path) -> None:
        """
        Prints the file at the given path using a shell command.
        Less technically involved (literally one line), but not silent.
        """
        startfile(path, "print")

    def device_print(self, path: Path) -> None:
        """
        Prints the file at the given path using a device context.
        More technically involved, but always silent.
        """
        with Document(self, path.name) as document:
            if path.suffix == ".pdf":
                document.send(*(Dib(page) for page in convert_from_path(path)))  # print each page of the pdf
            else:
                document.send(path)  # print the image on one page


class Document:
    """
    Wrapper class for printing a document on Windows.
    Should only be used as a context manager inside a Printer context.
    """

    def __init__(self: Document, printer: Printer, name: str) -> None:
        """
        Stores a reference to the printer currently in use, and stores the name of the document.
        """
        self.printer: Printer = printer
        self.name: str = name

    def __enter__(self: Document) -> Document:
        """
        Tells the printer to start printing this document with this name.
        """
        self.printer.device.StartDoc(self.name)
        return self

    def __exit__(self: Document, *_) -> None:
        """
        Tells the printer to stop printing this document.
        We don't care about the other values Python gives us.
        """
        self.printer.device.EndDoc()

    def send(self: Document, *items: Path | Dib) -> None:
        """
        Sending items to a document prints them as pages in that document.
        """
        item: Path | Dib
        for item in items:
            if isinstance(item, Path):
                item = Page.get_dib(item)
            with Page(self) as page:
                page.send(item)


class Page:
    """
    Wrapper class for printing a page on Windows.
    Should only be used as a context manager inside a Document context.
    """

    def __init__(self: Page, document: Document) -> None:
        """
        Assigns a reference to the document currently being printed.
        """
        self.document: Document = document

    def __enter__(self: Page) -> Page:
        """
        Tells the printer to start printing this page.
        """
        self.document.printer.device.StartPage()
        return self

    def __exit__(self: Page, *_) -> None:
        """
        Tells the printer to stop printing this page.
        We don't care about the other arguments Python gives us.
        """
        self.document.printer.device.EndPage()

    def send(self: Page, dib: Dib) -> None:
        """
        Sending a dib to a page prints it on that page.
        """
        dib.draw(self.document.printer.handle, Page.fit_dib(self.document.printer, dib))

    # thanks tim golden
    @staticmethod
    def get_dib(path: Path) -> Dib:
        """
        Converts the image file at the given path to a device-independent bitmap.
        Rotates the image by ninety degrees if the image is wider than it is tall.
        """
        # must be in RGB format or constructing a Dib object throws an exception
        bmp: Image = Image.open(path).convert("RGB")
        if bmp.size[0] > bmp.size[1]:
            bmp = bmp.rotate(90)
        return Dib(bmp)

    @staticmethod
    def fit_dib(printer: Printer, dib: Dib) -> tuple[int, int, int, int]:
        """
        Returns cartesian coordinates for printing a DIB on a sheet of paper as large as we possibly can based on the
        given printer.
        """
        scale: int = min([printer.paper_width / dib.size[0], printer.paper_height / dib.size[1]])
        width: int
        height: int
        width, height = [int(dim * scale) for dim in dib.size]
        x1: int = int((printer.paper_width - width) / 2)
        y1: int = int((printer.paper_height - height) / 2)
        x2: int = x1 + width
        y2: int = y1 + height
        return x1, y1, x2, y2


def send(*paths: str | Path, printer_name: str = "", msg: str = "", shell: bool = False) -> None:
    """
    Sends the given items to the default printer to be printed as documents.
    If a printer name is given, then that printer is used instead.
    If a message is provided, this message is printed after each item is sent to the printer.
        This message can be templated with $printer, $path, $document, and $stem.
    If shell is true, then the file is printed by starting the file with the "print" verb.
        This uses the default printer, so printer_name is disregarded.
    """
    with Printer(printer_name) as printer:
        printer.send(*paths, msg=msg, shell=shell)
