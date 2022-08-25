from excelStructure import *

import settings as stg


class MdToArray:
    def __init__(self):
        self.book = None
        self.fileName = None
        return

    def read(self, path: str):
        self.fileName = path.replace("//", "/").split("/")[-1].replace(".md", "")
        with open(path, encoding="utf-8", mode="r") as f:
            # self.md = f.read()
            text = f.read()
        textList = self.__loadTextList(text=text)
        self.__compile(textList)
        return

    def __loadTextList(self, text: str) -> list[str]:
        return text.split("\n")

    # HACK ちょっとこれは汚い｡
    def __compile(self, textList: list[str]):
        self.book = BookMdInfo()
        sheet: list[SheetMdInfo] = []
        pcss = []
        indent = 0
        indentBefore = 0
        for line in textList:
            column = []
            if len(line) <= 0:
                sheet.append([])
                continue
            # read # .
            p = 0
            while line[p] == "#":
                p += 1
            # if p > 4:
            #     p = 4
            #     print("warning : deep indent. bad.")
            if p > 0:  # pcss
                indentBefore = indent
                indent = p - 1
                if indentBefore > indent:
                    pcss[indentBefore] = 0
                if len(pcss) <= indent:
                    pcss.append(0)
                pcss[indent] = pcss[indent] + 1
                column.extend(["" for i in range(indent)])
                column.append(str(pcss[indent]) + ".")
                column.append(line[p + 1 :])
                sheet.append(column)
                continue

            # TODO セルのフォントフラグを設定できるようにしたい。
            if line[0] == " " or line[0] == "*" or line[1] == ".":
                listIndent = 0
                listTxt = ""
                sc = 0
                while line[sc] == " ":
                    sc += 1
                # for "* "
                if line[sc] == "*":
                    listIndent = int(sc / stg.listPlusIndentSpaces)
                    listTxt = "・"
                    addTex = line[line.find("* ") + 2 :]
                # for "1. "
                elif line[sc + 1] == ".":
                    listIndent = int(sc / stg.listNumIndentSpaces)
                    listTxt = line[sc : sc + 2]
                    addTex = line[line.find(". ") + 2 :]
                column.extend(["" for i in range(indent + 1 + listIndent)])
                column.append(listTxt)
                column.append(addTex)
            else:
                # generate nomal text
                column.extend(["" for i in range(indent + 1)])
                column.append(line[p:])
            sheet.append(column)

        # change page
        self.book.addSheet(self.fileName, sheet)

    def __printBook(self):
        for sheet in self.book.sheets:  # type:Sheet
            print("sheet : " + sheet.sheetName)
            # print(i)
            for line in sheet.data:
                t = ""
                for el in line:
                    if el == "":
                        el = "  "
                    t += el + " | "
                print(t)
            print("/n/n")


if __name__ == "__main__":
    mte = MdToArray()
    mte.read("test/test.md")
    mte.__printBook()
