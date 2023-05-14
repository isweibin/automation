# SPDX-FileCopyrightText: © 2023 Jia Weibin <self@isweibin.com>
# SPDX-License-Identifier: Apache-2.0


import os
import win32com.client


class DocConverter:
    def __init__(self):
        self.word = win32com.client.Dispatch("Word.Application")
        self.word.Visible = 0

    @staticmethod
    def _get_path(ext):
        path = os.getcwd()
        docs = os.listdir(path)
        for doc in docs:
            if doc.endswith(ext):
                pdf = os.path.splitext(doc)[0] + ".pdf"
                if pdf in docs:
                    continue
                doc = os.path.join(path, doc)
                pdf = os.path.join(path, pdf)
                yield doc, pdf

    def convert(self):
        for doc, pdf in self._get_path(("doc", "docx")):
            newpdf = self.word.Documents.Open(doc)
            newpdf.SaveAs(pdf, FileFormat=17)
            newpdf.Close()


if __name__ == "__main__":
    converter = DocConverter()
    converter.convert()
