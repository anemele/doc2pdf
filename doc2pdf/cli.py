"""convert Word file (.doc, .docx) to PDF file (.pdf)"""

import argparse
import glob
import os.path
from itertools import chain
from pathlib import Path

from .core import convert


def main():
    parser = argparse.ArgumentParser(prog=__package__, description=__doc__)
    parser.add_argument(
        'file', nargs='+', type=str, help='Word file (.doc, .docx), glob supports'
    )

    args = parser.parse_args()
    args_file: list[str] = args.file

    paths = map(
        Path, filter(os.path.isfile, chain.from_iterable(map(glob.iglob, args_file)))
    )
    convert(paths)
