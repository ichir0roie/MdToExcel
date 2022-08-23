import os
import json


class SheetMdInfo:
    def __init__(self, name: str, data: list[str]):
        self.sheetName: str = name
        self.data: list[str] = data


class BookMdInfo:
    def __init__(self):
        self.sheets: list[SheetMdInfo] = []

    def addSheet(self, name: str, sheetArray: list[SheetMdInfo]):
        sheet = SheetMdInfo(name, sheetArray)
        self.sheets.append(sheet)


test = False
