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
    For i = 1 To n
        ' 選択された図形を取得
        Dim shape As Shape
        Set shape = SelectedShapes(i)
        ' テキストフレームがない場合, エラーを出力
        If shape.HasTextFrame = msoFalse Then
            MsgBox ("Select shapes with text frame")
            Exit Sub
        End If
        ' テキストを挿入
        shape.TextFrame.TextRange.Text = CStr(i)
    Next i
End Sub
