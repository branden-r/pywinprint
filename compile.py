from os import system
from pathlib import Path
from shutil import rmtree


def to_exe(stem: str) -> None:
    """
    Creates an .exe of the .py file with the given stem.
    The given .py file must be in the same directory as this file.
    The .exe appears in the same directory as this file.
    Deletes build, dist, and pycache folders before and after compilation.
    :returns: None
    """
    rmtree("build") if Path("build").is_dir() else None
    rmtree("dist") if Path("dist").is_dir() else None
    rmtree("__pycache__") if Path("__pycache__").is_dir() else None
    system(f"pyinstaller --onefile {stem}.py")
    Path(f"dist/{stem}.exe").replace(f"{stem}.exe")
    rmtree("build")
    Path("dist").rmdir()
    rmtree("__pycache__")
    Path(f"{stem}.spec").unlink()


def main() -> None:
    """
    Calls to_exe for driver.py.
    :returns: None
    """
    to_exe("driver")


if __name__ == "__main__":
    main()
