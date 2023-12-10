from pathlib import Path
from typing import Iterable

from win32com import client

from .log import logger

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


def _convert(word, path: Path):
    try:
        doc_abspath = path.absolute()
        doc = word.Documents.Open(str(doc_abspath))
        pdf_abspath = doc_abspath.with_suffix('.pdf')
        doc.SaveAs(str(pdf_abspath), 17)  # ref: above code list
    except Exception as e:
        return e


def convert(paths: Iterable[Path]):
    word = client.DispatchEx('Word.Application')
    try:
        for path in paths:
            r = _convert(word, path)
            if r is None:
                logger.info(f'done: {path}')
            else:
                logger.error(r)
    finally:
        word.Quit()
