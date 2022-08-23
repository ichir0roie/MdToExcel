from tkinter.messagebox import NO
from excelStructure import *
import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook


import mdToArray
import os
import shutil

from openpyxl.styles import Font

import settings as stg


class ArrayToExcel:
    def __init__(self):
        self.book: BookMdInfo = None
        self.books: list[BookMdInfo] = []
        self.templateBookPath = None
        self.templateSheetName = None
        self.startRow = None
        self.startCol = None

        self.baseFont = None

        return

    def reset(self):
        self.book = None
        self.books = []

    def setBook(self, book: BookMdInfo):
        self.book = book

    def addBook(self, book: BookMdInfo):
        self.books.append(book)

    def __newWorkSheetFromTemplate(self, workbook: Workbook) -> Worksheet:
        template = None
        if stg.templateSheetName in workbook.sheetnames:
            template: Worksheet = workbook.copy_worksheet(
                workbook.get_sheet_by_name(stg.templateSheetName)
            )
            template.sheet_state = "visible"
        else:
            template = workbook.create_sheet()
        return template

    def generate(self, outputPath: str):
        wb = self.setupWorkBook(outputPath=outputPath)

        for sheet in self.book.sheets:
            self.parseSheetMdInfoToExcel(wb, sheet)

        self.finishWorkBook(wb, outputPath)

        return

    def setupWorkBook(self, outputPath: str):
        extension = outputPath[outputPath.rfind(".") + 1 :]
        if extension == "xlsx":
            raise Exception("file must be xlsm")

        # initialize as xlsm
        outputDir = outputPath[0 : outputPath.rfind("/")]
        if not os.path.exists(outputPath):
            if not os.path.exists(outputDir):
                os.makedirs(outputDir)
            shutil.copy(stg.templateBookPath, outputPath)
            # wb=load_workbook(filename=self.templateBookPath, keep_vba=True,read_only=False)
        wb = load_workbook(outputPath, keep_vba=True, read_only=False)

        # reset template sheet
        if stg.templateSheetName in wb.sheetnames:
            std: Worksheet = wb.get_sheet_by_name(stg.templateSheetName)
            # https://stackoverflow.com/questions/23157643/openpyxl-and-hidden-unhidden-excel-worksheets
            std.sheet_state = "hidden"
            # wb.remove_sheet(std)
        return wb

    def finishWorkBook(self, workbook: Workbook, outputPath: str):
        workbook.save(outputPath)
        workbook.close()

    def parseSheetMdInfoToExcel(self, wb: Workbook, sheet: SheetMdInfo):
        print(sheet.sheetName)
        print(sheet.data)

        # delete before created sheet
        if sheet.sheetName in wb.sheetnames:
            std = wb.get_sheet_by_name(sheet.sheetName)
            wb.remove_sheet(std)

        newSheet = self.__newWorkSheetFromTemplate(wb)

        newSheet.title = sheet.sheetName

        self.baseFont = Font(name=stg.font, size=stg.size)

        for r, row in enumerate(sheet.data):
            for c, column in enumerate(row):
                if column != "":
                    self.__setVal(newSheet, r + 1, c + 1, column)

    def generates(self, outputPath):
        wb = self.setupWorkBook(outputPath)

        for book in self.books:
            for sheet in book.sheets:
                self.parseSheetMdInfoToExcel(wb, sheet)

        self.finishWorkBook(wb, outputPath)

        return

    def __setVal(self, ws: openpyxl.worksheet, row, col, val):
        cell = ws.cell(row=row + stg.startRow, column=col + stg.startCol)
        cell.font = self.baseFont
        cell.value = val
        return


if __name__ == "__main__":
    mte = mdToArray.MdToArray()
    mte.read("test/test.md")

    # import pickle

    # # with open("book.pickle","wb")as f:
    # #     pickle.dump(mte.book, f)
    # with open("book.pickle", "rb") as f:
    #     book = pickle.load(f)
    ate = ArrayToExcel()
    ate.setBook(book=mte.book)
    # ate.readTemplate()
    ate.generate("test/test.xlsm")
