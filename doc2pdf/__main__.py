"""convert Word file (.doc, .docx) to PDF file (.pdf)"""

import argparse
import glob
import os.path
import sys
from itertools import chain
from pathlib import Path

from .core import convert

parser = argparse.ArgumentParser(
    prog=__package__ if len(sys.argv) == 1 else sys.argv[1], description=__doc__
)
parser.add_argument(
    'file', nargs='+', type=str, help='Word file (.doc, .docx), glob supports'
)

args = parser.parse_args(sys.argv[2:])
args_file: list[str] = args.file

paths = map(
    Path, filter(os.path.isfile, chain.from_iterable(map(glob.iglob, args_file)))
)
convert(paths)
