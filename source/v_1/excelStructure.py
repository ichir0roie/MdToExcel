class SheetMdInfo:
    def __init__(self, name: str, data: list[str]):
        self.sheetName: str = name
        self.data: list[str] = data
        # TODO self.data: list[SellInfo] = data


class SellInfo:
    def __init__(self, text: str, fontMode: int) -> None:
        self.text = text
        self.fontMode = fontMode


class BookMdInfo:
    def __init__(self):
        self.sheets: list[SheetMdInfo] = []

    def addSheet(self, name: str, sheetArray: list[SheetMdInfo]):
        sheet = SheetMdInfo(name, sheetArray)
        self.sheets.append(sheet)


test = False
