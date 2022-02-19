import openpyxl

wb = openpyxl.load_workbook('./testFile/test.xlsx')


def SearchFromBook(workbook, targetstr):
    detectflag = False
    duplicateflag = False
    returnlist = []

    print('ブック全体から検索します。')

    for ws in workbook.worksheets:  # ブック内の全シートを回るループ
        for row in ws.iter_rows():  # シート内の全行を回るループ
            for cell in row:  # 行内の全セルを回るループ
                if targetstr == cell.value:  # ここから重複検索をしますが、処理が大幅に遅くなる可能性があるので、場合によって重複検索しないほうに帰る必要がある
                    if detectflag:
                        duplicateflag = True
                    else:
                        detectflag = True

                    returnlist.append(ws.title)
                    returnlist.append(cell.coordinate)

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


target = 'さむらごうち'
resultlist = SearchFromBook(wb, target)

print(resultlist)
