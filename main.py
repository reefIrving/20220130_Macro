from unittest import skip
import ExcelHandle

# global変数


def main():
    # マクロ開始宣言
    print('マクロを開始します。\n ダイアログが表示されるので、3種のExcelファイルを選択してください。')
    ExHandler = ExcelHandle.ExcelHandle()
    ExcelBooks = ExHandler.getExcel()
    '''
        ExcelBooks[0] = 変更対象のブック
        ExcelBooks[1] = 検索元のブック
        ExcelBooks[2] = 検索先になるブック
    '''
    print('養生番号とシリアルが並んだExcelファイル についてです。')
    print('シリアルが乗っているカラムを入力してください。\n注意:アルファベットでは無く、数値で入力してください\n 例) Bを選びたい場合、2')
    targetcolumns = input()
    print('養生番号が乗っているカラムを入力してください。\n 例) Aを選びたい場合、1')
    yojocolumns = input()
    print('項目名の行がある場合、何行目からデータが並んでいるか入力してください。\n 例) タイトルと項目名で、2行使っている場合、3')
    skiprow = input()
    print('\nシリアルとホスト名が記載されているExcelファイル についてです。')
    print('ホスト名が乗っているカラムを入力してください。\n 例) Fを選びたい場合は、6')
    hostcolumns = input()
    for ws in ExcelBooks[1].worksheets:
        # シート内の全行を回るループ、min_rowは開始するrow位置を指定
        for row in ws.iter_cols(min_col=int(targetcolumns), min_row=int(skiprow)):
            for val in row:
                SearchResult = ExHandler.SearchFromBook(
                    ExcelBooks[2], val.value)
                '''正常時
                    SearchResult[0] = シート名
                    SearchResult[1] = 該当セル(row)
                '''
                if len(SearchResult) != 2:  # 複数検出や発見できなかった等。
                    print('エラーとなりました。\n スキップします。')
                else:  # 正常に検索できた場合
                    print('ヒットしました。\n シート名を変更します。')

                    tmpwb1 = ExcelBooks[2]
                    tmpws1 = tmpwb1[SearchResult[0]]
                    tmpwb2 = ExcelBooks[1]
                    tmpws2 = tmpwb2.worksheets[0]
                    tmpwb3 = ExcelBooks[0]

                    tmpws3 = tmpwb3[str(tmpws2.cell(
                        row=val.row, column=int(yojocolumns)).value)]

                    # print(tmpws1.cell(row=SearchResult[1],
                    #     column=int(hostcolumns)).value)
                    tmpws3.title = str(tmpws1.cell(
                        row=SearchResult[1], column=int(hostcolumns)).value)

            '''
                tmpwb1 = ExcelBooks[2]
                tmpws1 = tmpwb1[SearchResult[0]]
                tmpws1.cell(row=SearchResult[1], column=hostcolumns) = ホスト名
                tmpwb2 = ExcelBooks[1]
                tmpws2 = tmpwb2.worksheets[0]
                tmpws2.cell(row=val.row, column=yojocolumns) = 養生番号(いま検索してる対象のやつ)
            '''
    ExcelBooks[0].save('./処理済みブック.xlsx')


if __name__ == '__main__':
    main()
