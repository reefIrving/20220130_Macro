import openpyxl
import ExcelHandle

# global変数


def main():
    # マクロ開始宣言
    print('マクロを開始します。\n ダイアログが表示されるので、3種のExcelファイルを選択してください。')
    ExHandler = ExcelHandle.ExcelHandle()
    ExcelBooks = ExHandler.getExcel()
    '''
        ExcelBooks[0] = 単体試験シート
        ExcelBooks[1] = (仮)養生シリアル一覧シート
        ExcelBooks[2] = (仮)シリアルホスト名一覧シート
    '''
    print('シート名一覧を表示するので、試験シートの該当項番を入力してください。(目次抜いた最初の試験ページ)')
    IndexNum = ExHandler.getBasepoint(ExcelBooks[0].sheetnames)
    print('(仮)養生シリアル一覧シートのシート名を表示するので、シリアルと養生番号の乗っているシート名を入力してください。')
    IndexSerials = ExHandler.getBasepoint(ExcelBooks[1].sheetnames)


# 変更対象ページ以降を持って、検索しに行く
# main呼び出し
if __name__ == '__main__':
    main()
