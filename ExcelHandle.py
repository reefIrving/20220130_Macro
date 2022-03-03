from tkinter import filedialog
import openpyxl
import re


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
        ExcelBooks.append(self.ExcelOpen('シート名を変更したいExcelファイル'))

        ExcelBooks.append(self.ExcelOpen('養生番号とシリアルが並んだExcelファイル'))

        ExcelBooks.append(self.ExcelOpen('シリアルとホスト名が記載されているExcelファイル'))

        # print(type(ExcelBooks[0]))
        # print(type(ExcelBooks[1]))
        # print(type(ExcelBooks[2]))
        return ExcelBooks

    def SplitCellAddress(self, CellAddress):
        column = re.split('[0-9]+', CellAddress)
        # 数字の方、行 ここができてないよ 全部くぎっちゃってる、1ケタだけにしてあるみたい
        row = re.split('[a-z][A-Z]+', CellAddress)
        return column, row

    def SearchFromBook(self, workbook, targetstr):
        detectflag = False
        duplicateflag = False
        returnlist = []

        print('ブック全体から検索します。')

        for ws in workbook.worksheets:  # ブック内の全シートを回るループ
            print(ws.title + 'を検索します。')
            for row in ws.iter_rows():  # シート内の全行を回るループ
                for cell in row:  # 行内の全セルを回るループ
                    if cell.value is None:
                        continue  # ここら辺の比較の時にStrがどうのこうの出る。
                    else:
                        # ここから重複検索をしますが、処理が大幅に遅くなる可能性があるので、場合によって重複検索しないほうに帰る必要がある
                        if str(targetstr) in str(cell.value):
                            if detectflag:
                                duplicateflag = True
                            else:
                                detectflag = True

                            returnlist.append(ws.title)
                            returnlist.append(cell.row)

        if detectflag:
            # 発見時処理
            if duplicateflag:
                # 重複時処理
                print('検索結果が複数検出されました。\n重複しています。')
                return returnlist
            else:
                return returnlist
        else:
            # 未発見時処理
            print('データが見つかりませんでした。')
            return None
