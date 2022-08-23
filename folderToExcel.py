import mdToArray
import arrayToExcel
import json

import glob

from excelStructure import *

# import settings as stng

import os
import shutil


import settings as stg


class FolderToExcel:
    def __init__(self):
        self.mta = mdToArray.MdToArray()
        self.ate = arrayToExcel.ArrayToExcel()

        # 追記上書きにしたい｡
        # shutil.rmtree(stng.outputExFolder)
        # os.mkdir(stng.outputExFolder)

        self.ate.__readTemplate(
            stg.templateBookPath,
            stg.templateSheetName,
            stg.startRow,
            stg.startCol,
        )

        # with open("settings.json", mode="r", encoding="utf-8")as f:
        #     text = f.read()
        #     self.settings = json.loads(text)

        self.dirs = []

        self.readFolder()

    # まずエクセル共をコピーして貼り付けておく｡
    def copyExcels(self):

        self.deleteInitOutputFolder()
        self.deleteInitCombineFolder()

        for dir in self.dirs:
            targets = [
                p
                for p in glob.glob(dir + "**.*")
                if os.path.isfile(p) and not ".md" in p
            ]

            for path in targets:
                self.save(path)

        return

    def readFolder(self):
        # 対象のフォルダ内のファイルを全部読み込んで、リストに格納
        # ↓
        # フォルダ単位で検索して､ret!

        searchPath = (stg.inputMdFolder + "/**/").replace("//", "/")
        self.dirs = [
            p for p in glob.glob(searchPath, recursive=False) if not os.path.isfile(p)
        ]
        print(self.dirs)
        self.dirs = [p for p in self.dirs if not ".git" in p]
        if not test:
            self.dirs = [p for p in self.dirs if not "test" in p]
        else:
            self.dirs = [p for p in self.dirs if "test" in p]

        for ext in stg.extracts:
            self.dirs = [p for p in self.dirs if ext not in p]

    def generate(self):
        # 対象のデータを順次変換

        targets = [p for p in self.targetPaths if ".md" in p]

        for path in targets:
            self.mta.read(path)
            self.ate.setBook(self.mta.book)
            savePath = path.replace(stg.inputMdFolder, stg.outputExcelFolder)
            extension = stg.templateBookPath[stg.templateBookPath.find(".") :]
            savePath = savePath.replace(".md", extension)
            self.ate.generateBook(savePath, stg.font, stg.size)

    def generateByDir(self):
        # 対象のデータを順次変換
        for dir in self.dirs:
            targets = [p for p in glob.glob(dir + "/**.md") if os.path.isfile(p)]

            self.ate.reset()
            for path in targets:
                self.mta.read(path)
                self.ate.addBook(self.mta.book)
            self.generates(dir)

    def generates(self, dir):
        savePath = dir.replace(stg.inputMdFolder, stg.outputExcelFolder)
        if savePath[-1] == "/":
            savePath = savePath[:-1]
        fileName = savePath.split("/")[-1]
        savePath = savePath + "/" + fileName + ".md.xlsm"
        extension = stg.templateBookPath[stg.templateBookPath.find(".") :]
        # savePath=savePath.replace(".md", extension)
        self.ate.generateBooks(savePath, stg.font, stg.size)

    def save(self, path):
        savePath = path.replace(stg.inputMdFolder, stg.outputExcelFolder)
        saveDir = saveDir[: saveDir.rfind("/")]
        if not os.path.exists(saveDir):
            os.makedirs(saveDir)
        shutil.copy(path, savePath)

    def deleteInitOutputFolder(self):
        paths = glob.glob(stg.outputExcelFolder + "*")
        for path in paths:
            print(path)
            shutil.rmtree(path)

    def deleteInitCombineFolder(self):
        files = glob.glob(stg.combineFolder + "*")
        for file in files:
            # print(file)
            os.remove(file)


if __name__ == "__main__":
    # test=True

    m = FolderToExcel()

    m.copyExcels()
    m.generateByDir()
    m.combine()
