#!/usr/bin/env python3.8
# -*- encoding: utf-8 -*-

"""convert Word file (.doc, .docx) to PDF file (.pdf)"""

import argparse
import glob
import os.path
from itertools import chain
from pathlib import Path
from typing import List

from win32com import client

# available format code
# wdFormatDocument = 0
# wdFormatDocument97 = 0
# wdFormatDocumentDefault = 16
# wdFormatDOSText = 4
# wdFormatDOSTextLineBreaks = 5
# wdFormatEncodedText = 7
# wdFormatFilteredHTML = 10
# wdFormatFlatXML = 19
# wdFormatFlatXMLMacroEnabled = 20
# wdFormatFlatXMLTemplate = 21
# wdFormatFlatXMLTemplateMacroEnabled = 22
# wdFormatHTML = 8
# wdFormatPDF = 17
# wdFormatRTF = 6
# wdFormatTemplate = 1
# wdFormatTemplate97 = 1
# wdFormatText = 2
# wdFormatTextLineBreaks = 3
# wdFormatUnicodeText = 7
# wdFormatWebArchive = 9
# wdFormatXML = 11
# wdFormatXMLDocument = 12
# wdFormatXMLDocumentMacroEnabled = 13
# wdFormatXMLTemplate = 14
# wdFormatXMLTemplateMacroEnabled = 15


def convert(word, file: Path):
    try:
        doc_abspath = file.absolute()
        doc = word.Documents.Open(str(doc_abspath))
        pdf_abspath = doc_abspath.with_suffix('.pdf')
        doc.SaveAs(str(pdf_abspath), 17)  # ref: above code list
        print('done', file)
    except Exception as e:
        print('error', e)


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        'file', nargs='+', type=str, help='Word file (.doc, .docx), glob supports'
    )

    return parser.parse_args()


def main():
    args = parse_args()
    # print(args)
    # return
    args_file: List[str] = args.file

    files = map(
        Path, filter(os.path.isfile, chain.from_iterable(map(glob.iglob, args_file)))
    )

    word = client.DispatchEx('Word.Application')
    try:
        for file in files:
            convert(word, file)
    finally:
        word.Quit()


if __name__ == '__main__':
    main()
