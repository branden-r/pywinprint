## What is this?
This is a small repo containing code that will allow you to print documents on Windows in a Pythonic way (printer.py).
Additionally, this repo contains a short program (driver.py) demonstrating how you might use printer.py.
Note that driver.py can be compiled to an executable using compile.py, and a pre-compiled executable (driver.exe) is
also available.

## Why?
I was trying to print out a large number of PDF documents at work, and it was really annoying.
So, I automated it.

## How?
This project heavily relies on libraries created by Tim Golden.
It also uses some stuff from Pillow (a fork of PIL - Python Imaging Library).
You should check out Tim Golden's Python stuff and Pillow -- both of these are incredibly useful.

## Can I use this?
Sure.
If you're doing something that requires licensing, you should probably investigate the licensing information for the
libraries I used.

## My antivirus said the .exe file is a scary trojan.
It's not.
I haven't investigated this issue all too thoroughly, but I've seen other people mention this when using pyinstaller.
You can ignore it, but if it really bothers you, you can just run your code as a Python script.
You could also try compiling your own executable.
