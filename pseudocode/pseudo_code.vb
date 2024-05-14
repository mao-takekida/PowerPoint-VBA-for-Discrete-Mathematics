' 与えられたテキストの２文字目以降のフォントサイズを変更する関数
Function ChangeFontSize(textRange As textRange, fontSize As Integer)
    Dim i As Integer
    For i = 2 To textRange.Characters.Count
        textRange.Characters(i, 1).Font.Size = fontSize
    Next i
End Function


' 選択中のスライドに擬似コードを挿入します。
' 擬似コードの関数名は, InpudBox に入力された文字列を使用します。
Sub add_pseudo_code()
    ' 現在のスライドを取得
    Dim cur_slide As slide
    Set cur_slide = ActiveWindow.View.slide

    ' 選択中のスライドに擬似コードを挿入
    Dim pseudo_code As shape
    Set pseudo_code = cur_slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 500, 500)

    ' 擬似コードの関数名を入力
    pseudo_code.TextFrame.TextRange.Text = InputBox("Enter function name", "Function Name")
    ' ２文字目以降のフォントサイズを小さくする関数を適用
    ' 元のフォントサイズの r 倍に変更
    Dim r As Single
    r = 0.7
    ChangeFontSize pseudo_code.TextFrame.TextRange, pseudo_code.TextFrame.TextRange.Font.Size * r

    ' 改行
    pseudo_code.TextFrame.TextRange.InsertAfter vbCrLf
End Sub
