<div id="top"></div>

<!-- ## 使用技術 -->
<!-- シールド一覧 -->
<!-- https://t8csp.csb.app/ -->
<p style="display: inline">
  <img src="https://img.shields.io/badge/VBA-excel%20-21a366.svg?logo=office&style=popout">
</p>


# Excel：画像を各シートに1枚ずつ張り付けていくVBA


# 概要

エビデンス・説明等で1Sheetに1画像を張りつけたブックを作成するソースです。

# 導入方法

png2book.xlsmをマクロ有効にして実行を押す

ソースは以下になります。

[png2book.txt](https://github.com/yymat/png2book/blob/main/png2book.txt)
```Sub png2book()
    Dim fd As FileDialog
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sp As Shape
    Dim sPath As String, fPath As String

    'Specify the folder where png is located.	'pngがあるフォルダを指定する
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = False Then Exit Sub

    sPath = fd.SelectedItems(1) & "\"
    fPath = Dir(sPath & "*.png", vbNormal)

    Set wb = Workbooks.Add
    
    'Add images to sheets.	'画像をシートに追加する
    While (fPath <> "")
        Set ws = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = wb.Worksheets.Count - 1
        Set sp = ws.Shapes.AddPicture(sPath & fPath, False, True, 0, ws.Cells(2, 1).Top, 0, 0)
        sp.ScaleWidth 1, True
        sp.ScaleHeight 1, True
        fPath = Dir()
    Wend

    'delete Sheet1
    Application.DisplayAlerts = False
    wb.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True

End Sub
```

# 使い方

1. 実行すると新規ブックを作成します。

2. 画像フォルダの指定するダイアログにてフォルダを指定します。

3. 画像ファイルの数だけシートが作成され、シートに画像が貼り付けられます

# ライセンス
 [MIT license](https://en.wikipedia.org/wiki/MIT_License).
