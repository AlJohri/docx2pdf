import sys
import json
import subprocess
from pathlib import Path
from tqdm.auto import tqdm

try:
    # 3.8+
    from importlib.metadata import version
except ImportError:
    from importlib_metadata import version

__version__ = version(__package__)


def windows(paths, keep_active):
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    wdFormatPDF = 17

    if paths["batch"]:
        for docx_filepath in tqdm(sorted(Path(paths["input"]).glob("[!~]*.docx"))):
            pdf_filepath = Path(paths["output"]) / (str(docx_filepath.stem) + ".pdf")
            doc = word.Documents.Open(str(docx_filepath))
            doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
            doc.Close(0)
    else:
        pbar = tqdm(total=1)
        docx_filepath = Path(paths["input"]).resolve()
        pdf_filepath = Path(paths["output"]).resolve()
        doc = word.Documents.Open(str(docx_filepath))
        doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
        doc.Close(0)
        pbar.update(1)

    if not keep_active:
        word.Quit()


def macos(paths, keep_active):
    """Run the conversion on a MacOS system. Calls a JXA script, which
    runs Microsoft Word to do the actual conversion.

    See docstring for convert() for a description of the parameters.

    :raises EnvironmentError: if the JXA exits with nonzero return code
                              because Microsoft Word is not available
    :raises RuntimeError: if the JXA exits with nonzero return code for
                          any other reason
    """
    script = (Path(__file__).parent / "convert.jxa").resolve()
    cmd = [
        "/usr/bin/osascript",
        "-l",
        "JavaScript",
        str(script),
        str(paths["input"]),
        str(paths["output"]),
        str(keep_active).lower(),
    ]

    total = len(list(Path(paths["input"]).glob("*.docx"))) if paths["batch"] else 1
    pbar = tqdm(total=total)

    process = subprocess.Popen(cmd, stderr=subprocess.PIPE)
    process.wait()
    if process.returncode != 0:
        msg = process.stderr.read().decode().rstrip()
        if "Application can't be found" in msg:
            raise EnvironmentError("Microsoft Word is not available.")
        raise RuntimeError(msg)

    def stderr_results(process):
        while True:
            line = process.stderr.readline().rstrip()
            if not line:
                break
            yield line.decode("utf-8")

    for line in stderr_results(process):
        try:
            msg = json.loads(line)
        except ValueError:
            continue
        if msg["result"] == "success":
            pbar.update(1)
        elif msg["result"] == "error":
            print(msg)
            sys.exit(1)


def resolve_paths(input_path, output_path):
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve() if output_path else None
    output = {}
    if input_path.is_dir():
        output["batch"] = True
        output["input"] = str(input_path)
        if output_path:
            assert output_path.is_dir()
        else:
            output_path = str(input_path)
        output["output"] = output_path
    else:
        output["batch"] = False
        assert str(input_path).endswith((".docx", ".DOCX"))
        output["input"] = str(input_path)
        if output_path and output_path.is_dir():
            output_path = str(output_path / (str(input_path.stem) + ".pdf"))
        elif output_path:
            assert str(output_path).endswith(".pdf")
        else:
            output_path = str(input_path.parent / (str(input_path.stem) + ".pdf"))
        output["output"] = output_path
    return output


def convert(input_path, output_path=None, keep_active=False):
    """Wrapper around the conversion functions depending on whether the
    system is Windows or MacOS. The supplied paths are 'resolved' into
    a dictionary of path information before being given to macos() or
    windows().

    :param input_path: The path to the docx.
    :param output_path: The path to the pdf (by default, the same name
                        and directory as the docx, but with .pdf file
                        extension).
    :param keep_active: Whether to keep Microsoft Word running after the
                        conversion(s) are complete.
    """
    paths = resolve_paths(input_path, output_path)
    if sys.platform == "darwin":
        return macos(paths, keep_active)
    elif sys.platform == "win32":
        return windows(paths, keep_active)
    else:
        raise NotImplementedError(
            "docx2pdf is not implemented for linux as it requires Microsoft Word to be installed"
        )


def cli():

    import textwrap
    import argparse

    if "--version" in sys.argv:
        print(__version__)
        sys.exit(0)

    description = textwrap.dedent(
        """
    Example Usage:

    Convert single docx file in-place from myfile.docx to myfile.pdf:
        docx2pdf myfile.docx

    Batch convert docx folder in-place. Output PDFs will go in the same folder:
        docx2pdf myfolder/

    Convert single docx file with explicit output filepath:
        docx2pdf input.docx output.docx

    Convert single docx file and output to a different explicit folder:
        docx2pdf input.docx output_dir/

    Batch convert docx folder. Output PDFs will go to a different explicit folder:
        docx2pdf input_dir/ output_dir/
    """
    )

    formatter_class = lambda prog: argparse.RawDescriptionHelpFormatter(
        prog, max_help_position=32
    )
    parser = argparse.ArgumentParser(
        description=description, formatter_class=formatter_class
    )
    parser.add_argument(
        "input",
        help="input file or folder. batch converts entire folder or convert single file",
    )
    parser.add_argument("output", nargs="?", help="output file or folder")
    parser.add_argument(
        "--keep-active",
        action="store_true",
        default=False,
        help="prevent closing word after conversion",
    )
    parser.add_argument(
        "--version", action="store_true", default=False, help="display version and exit"
    )

    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)
    else:
        args = parser.parse_args()

    convert(args.input, args.output, args.keep_active)
