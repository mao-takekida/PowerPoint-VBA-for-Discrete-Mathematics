' shape が辺であれば True を返す
Function IsEdge(shape As shape) As Boolean
        IsEdge = shape.Type = msoLine
End Function

' shape が楕円であれば True を返す
Function IsOval(shape As shape) As Boolean
        IsOval = shape.AutoShapeType = msoShapeOval
End Function

' shape が十角形であれば True を返す
Function IsDecagon(shape As shape) As Boolean
        IsDecagon = shape.AutoShapeType = msoShapeDecagon
End Function

' 現在のスライドの選択したオブジェクトから 指定されたタイプのオブジェクトだけを選択した状態にする
' 引数  type_str: 選択するオブジェクトのタイプを表す文字列
Sub SelectOnly(type_str As String)
        ' 一つも選択されていない場合, エラーを出力
        If ActiveWindow.Selection.Type <> ppSelectionShapes Then
                MsgBox ("Select at least one object")
                Exit Sub
        End If

        ' ActiveWindow 上で選択されているすべてのスライドオブジェクトを表す ShapeRange オブジェクトを取得
        Dim SelectedShapes As ShapeRange
        Set SelectedShapes = ActiveWindow.Selection.ShapeRange

        ' 図形を識別するための name を保存する動的配列
        Dim names() As String
        ' 選択された図形の数
        Dim d_type_count As Integer
        d_type_count = 0
        Dim shape As shape
        For Each shape In SelectedShapes
                ' 選択された図形のタイプが指定されたタイプであるかを判定
                Dim is_target_type As Boolean
                If type_str = "edge" And IsEdge(shape) Then
                        is_target_type = True
                ElseIf type_str = "oval" And IsOval(shape) Then
                        is_target_type = True
                ElseIf type_str = "decagon" And IsDecagon(shape) Then
                        is_target_type = True
                Else
                        is_target_type = False
                End If

                If is_target_type Then
                        ' 配列に shape.Name を追加
                        ' ReDim Preserve で配列のサイズを変更
                        ReDim Preserve names(d_type_count)
                        names(d_type_count) = shape.Name
                        d_type_count = d_type_count + 1
                End If
        Next

        ' 一度すべての選択を解除
        ActiveWindow.Selection.Unselect

        ' name が names に含まれる図形を選択
        ' 現在のスライドを取得
        Dim cur_page As Slide
        Set cur_page = ActiveWindow.View.Slide
        Dim i As Integer
        For i = 1 To d_type_count
                cur_page.Shapes(names(i - 1)).Select msoFalse
        Next i
End Sub

' 選択したオブジェクトから辺のみを選択した状態にする
Sub SelectOnlyEdges()
        SelectOnly "edge"
End Sub

' 選択したオブジェクトからOvalのみを選択した状態にする
Sub SelectOnlyOvals()
        SelectOnly "oval"
End Sub

' 選択したオブジェクトから10角形のみを選択した状態にする
Sub SelectOnlyDecagons()
        SelectOnly "decagon"
End Sub
