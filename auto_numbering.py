# SPDX-FileCopyrightText: © 2023 Jia Weibin <self@isweibin.com>
# SPDX-License-Identifier: Apache-2.0

import itertools

import win32com.client


class AutoNumbering:
    def __init__(self):
        self.acad = win32com.client.Dispatch("AutoCAD.Application")
        self.keyword = "TK"
        self.tagstring = "文件号"

    def _get_blocks(self):
        for entity in self.acad.ActiveDocument.ModelSpace:
            if entity.EntityName == "AcDbBlockReference":
                if self.keyword in entity.Name:
                    yield entity

    def _sort_blocks(self):
        return sorted(
            self._get_blocks(),
            key=lambda x: (int(x.InsertionPoint[0]), int(x.InsertionPoint[1])),
        )

    def _group_blocks(self):
        groups = itertools.groupby(
            self._sort_blocks(), lambda x: int(x.InsertionPoint[0]) // 100
        )
        for _, group in groups:
            yield group

    def update(self):
        for group in self._group_blocks():
            i = 1
            for entity in group:
                for attribute in entity.GetAttributes():
                    if attribute.TagString == self.tagstring:
                        attribute.TextString = (
                            f"{attribute.TextString[:-2]}{str(i).zfill(2)}"
                        )
                        attribute.Update()
                i += 1

        self.acad.ActiveDocument.Regen(0)


if __name__ == "__main__":
    updater = AutoNumbering()
    updater.update()
