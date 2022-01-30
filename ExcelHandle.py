from tkinter import filedialog
import openpyxl


class ExcelHandle():
    def __init__(self) -> None:
        pass

    def ExcelOpen(selft, fname):
        # fnameには開く予定のファイル名が送られてくる。(str)
        dtitle = fname + 'を開く'
        typ = [('Excelファイル', '*.xlsx'), ('旧Excelファイル', '*.xls')]
        dir = './'
        fle = filedialog.askopenfilename(
            filetypes=typ, initialdir=dir, title=dtitle)

        wb = openpyxl.load_workbook(fle)

        return wb

    def getExcel(self):
        ExcelBooks = list()
        ExcelBooks.append(self.ExcelOpen('単体試験シート'))
        ExcelBooks.append(self.ExcelOpen('(仮)養生シリアル一覧シート'))
        ExcelBooks.append(self.ExcelOpen('(仮)シリアルホスト名一覧シート'))
        # print(type(ExcelBooks[0]))
        # print(type(ExcelBooks[1]))
        # print(type(ExcelBooks[2]))
        return ExcelBooks

    def getBasepoint(self, SheetNames):
        print('項番\tシート名\n')
        tmpIndex = -1
        for tmpNames in SheetNames:
            tmpIndex += 1
            print(str(tmpIndex) + '\t' + str(tmpNames))

        return int(input('項番を入力-> '))

    def SplitCellAddress(self, CellAddress):
        tmp = list(CellAddress)
        column = tmp[0]  # アルファベットの方、列
        row = int(tmp[1])  # 数字の方、行
        return column, row

    '''
    def getSerial(self, target, SerialsSheet):
        Ycolumn, Yrow = self.SplitCellAddress(input('養生番号が一覧されている最初のセル番地を入力してください。ex)A3 -> '))
        Scolumn, Srow = self.SplitCellAddress(input('シリアルが一覧されている最初のセル番地を入力してください。 ex)B3 -> '))
        while 1:
            if target == SerialsSheet[str(Ycolumn) + str(Yrow)] :
                # 養生番号の発見
    '''
