# beamer2ppt
Python script converting LaTeX beamer in `.pdf` file to `.ppt` or `.pptx` file, without losing formulae formatting.

## Implementation principle
pdf &rarr; img &rarr; insert into ppt &rarr; align &rarr; export

## Prereq.
- python `2.7`+ installed
- python package [`PyMuPDF`](https://pymupdf.readthedocs.io/en/latest/) installed
- python package [`python-pptx`](https://python-pptx.readthedocs.io/en/latest/) installed

**You can run the `init.bat` to automatically install the necessary dependencies.**

## Execution
1. Run terminal in the same directory of the `.pdf` file you want to convert
2. Run command: `python beamer2ppt.py [pdf directory] [ppt directory]`
