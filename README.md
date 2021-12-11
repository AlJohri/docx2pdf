# docx2pdf

[![PyPI](https://img.shields.io/pypi/v/docx2pdf)](https://pypi.org/project/docx2pdf/)

Convert `docx` to `pdf` on Windows or macOS directly using Microsoft Word (must be installed).

On Windows, this is implemented via [`win32com`](https://pypi.org/project/pywin32/) while on macOS this is implemented via [JXA](https://github.com/JXA-Cookbook/JXA-Cookbook) (Javascript for Automation, aka AppleScript in JS).

## Install

On macOS:

```
brew install aljohri/-/docx2pdf
```

Via [pipx](https://pipxproject.github.io/pipx/):

```
pipx install docx2pdf
```

Via pip:

```
pip install docx2pdf
```

## CLI

```
usage: docx2pdf [-h] [--keep-active] [--version] input [output]

Example Usage:

Convert single docx file in-place from myfile.docx to myfile.pdf:
    docx2pdf myfile.docx

Batch convert docx folder in-place. Output PDFs will go in the same folder:
    docx2pdf myfolder/

Convert single docx file with explicit output filepath:
    docx2pdf input.docx output.pdf

Convert single docx file and output to a different explicit folder:
    docx2pdf input.docx output_dir/

Batch convert docx folder. Output PDFs will go to a different explicit folder:
    docx2pdf input_dir/ output_dir/

positional arguments:
  input          input file or folder. batch converts entire folder or convert
                 single file
  output         output file or folder

optional arguments:
  -h, --help     show this help message and exit
  --keep-active  prevent closing word after conversion
  --version      display version and exit
```

## Library

```python
from docx2pdf import convert

convert("input.docx")
convert("input.docx", "output.pdf")
convert("my_docx_folder/")
```

See CLI docs above (or in `docx2pdf --help`) for all the different invocations. It is the same for the CLI and python library.

## Jupyter Notebook

If you are using this in the context of jupyter notebook, you will need `ipywidgets` for the tqdm progress bar to render properly.

```
pip install ipywidgets
jupyter nbextension enable --py widgetsnbextension
``
