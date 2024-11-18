' text の中心が edge の中心から edge の法線方向に k 離れるように配置
Sub CalculateTextPosition(edge As Shape, text As Shape, ByRef text_x As Single, ByRef text_y As Single, k As Single)
    ' テキストの中心と, 辺の中心が一致するように配置
    text_x = edge.Left + edge.Width / 2 - text.Width / 2
    text_y = edge.Top + edge.Height / 2 - text.Height / 2

    Dim dx As Single, dy As Single

    ' とりあえず edge の幅と高さを使う
    ' ToDo: edge のBeginX, BeginY, EndX, EndY を使う
    dx = edge.Width
    dy = edge.Height

    ' 辺の法線方向に k 離れるように配置
    Dim length As Single
    length = Sqr(dx * dx + dy * dy)
    text_x = text_x + k * dy / length
    text_y = text_y - k * dx / length
End Sub

' 選択した辺にランダムな重みを割り当てる
' 重みは1から10の範囲でランダムに選ばれる
Sub AssignWeights()
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

    Dim shape As Shape
    For Each shape In SelectedShapes
        ' 選択された図形のタイプが msoLine でない場合, エラーを出力
        If shape.Type <> msoLine Then
            MsgBox ("Select lines")
            Exit Sub
        End If
    Next

    ' 選択された辺に重みを割り当て
    Dim weight As Integer
    Dim edge As Shape
    For Each edge In SelectedShapes
        ' 1からMAX_WEIGHTの範囲でランダムな重みを選択
        Dim MAX_WEIGHT As Integer
        MAX_WEIGHT = 10
        weight = Int((MAX_WEIGHT) * Rnd + 1)
        
        ' テキストを作成
        Dim text As Shape
        Set text = ActiveWindow.View.Slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 0, 0)
        text.TextFrame.TextRange.Text = CStr(weight)
        ' テキストのフォントサイズを設定
        text.TextFrame.TextRange.Font.Size = 18

        ' テキストの位置を計算
        Dim text_x As Single, text_y As Single
        CalculateTextPosition edge, text, text_x, text_y, 10
        ' テキストを移動
        text.Left = text_x
        text.Top = text_y
    Next
End Sub
