' 選択したオブジェクトのテキストに、選択した順番に1から始まる番号を振る
Sub AssignNumbers()
    ' 一つも選択されていない場合, エラーを出力
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox ("Select at least one object")
        Exit Sub
    End If

    ' ActiveWindow 上で選択されているすべてのスライドオブジェクトを表す ShapeRange オブジェクトを取得
    Dim SelectedShapes As ShapeRange
    Set SelectedShapes = ActiveWindow.Selection.ShapeRange

    ' 選択された図形の数
    Dim n As Integer
    n = SelectedShapes.Count

    Dim i As Integer

    ' 先頭の図形に振る番号
    Dim number As Integer
    number = 1
    For i = 1 To n
        ' i 番目に選択された図形を取得
        Dim shape As Shape
        Set shape = SelectedShapes(i)

        ' 先頭の図形に既にテキストがあり、それが数字である場合, その数字を number に設定
        If i = 1 And shape.HasTextFrame And IsNumeric(shape.TextFrame.TextRange.Text) Then
            number = CInt(shape.TextFrame.TextRange.Text)
        End If

        ' テキストフレームがない場合, エラーを出力
        If shape.HasTextFrame = msoFalse Then
            MsgBox ("Select shapes with text frame")
            Exit Sub
        End If

        ' テキストを挿入
        shape.TextFrame.TextRange.Text = CStr(number + i - 1)
    Next i
End Sub
