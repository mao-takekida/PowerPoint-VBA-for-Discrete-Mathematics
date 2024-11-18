' 選択された図形の中心の座標を取得
Sub GetCenterPosition(shape As Shape, ByRef x As Single, ByRef y As Single)
    x = shape.Left + shape.Width / 2
    y = shape.Top + shape.Height / 2
End Sub

' src から dst への矢印の始点と終点の座標を計算
' (src_cx, src_cy): src の中心の座標
' (dst_cx, dst_cy): dst の中心の座標
' radius: 半径
' arrow_sx, arrow_sy: 矢印の始点の座標
' arrow_dx, arrow_dy: 矢印の終点の座標
' 両端を半径分だけ短くした線分を返す
Sub CalculateArrowPosition(src_cx As Single, src_cy As Single, dst_cx As Single, dst_cy As Single, radius As Single, ByRef arrow_sx As Single, ByRef arrow_sy As Single, ByRef arrow_dx As Single, ByRef arrow_dy As Single)
    Dim dx As Single, dy As Single
    dx = dst_cx - src_cx
    dy = dst_cy - src_cy
    Dim length As Single 'ベクトルの長さ
    length = Sqr(dx * dx + dy * dy)

    ' 半径分だけ短くした線分の端点の座標を計算
    arrow_sx = src_cx + radius * dx / length
    arrow_sy = src_cy + radius * dy / length
    arrow_dx = dst_cx - radius * dx / length
    arrow_dy = dst_cy - radius * dy / length
End Sub

' 1つ目に選択した図形から, 2つ目以降に選択した図形に向かって辺を引く
' directed = True なら有向グラフの辺を引く
Sub MakeEdge1ToN(Optional directed As Boolean)
    ' cur_pageは現在のスライド
    Dim cur_page As Slide
    Set cur_page = ActiveWindow.View.Slide

    if ActiveWindow.Selection.Type <> ppSelectionShapes Or ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox("Select 2 or more objects")
        Exit Sub
    End If

    ' 選択された図形を取得
    Dim selected_shapes As ShapeRange
    Set selected_shapes = ActiveWindow.Selection.ShapeRange
    ' 選択された図形の数
    Dim n As Integer
    n = selected_shapes.Count

    Dim shape As Shape
    For Each shape In selected_shapes
        ' msoAutoShape 以外の図形が選択された場合, エラーを出力
        If shape.Type <> msoAutoShape Then
            MsgBox("msoAutoShape オブジェクトを選択してください")
            Exit Sub
        End If
    Next 

    ' 最初に選択した図形をsrcとする
    Dim src As Shape
    Set src = selected_shapes(1)
    Dim src_cx As Single, src_cy As Single
    GetCenterPosition src, src_cx, src_cy

    ' 2番目以降に選択した図形に対して辺を引く
    Dim i As Integer
    For i = 2 To n
        ' 選択された図形を取得
        Dim dst As Shape
        Set dst = selected_shapes(i)
        ' dst の中心の座標を取得
        Dim dst_cx As Single, dst_cy As Single
        GetCenterPosition dst, dst_cx, dst_cy

        ' 半径を計算
        Dim radius As Single
        radius = dst.Width / 2

        ' src から dst への辺の始点と終点の座標を計算
        Dim arrow_sx As Single, arrow_sy As Single, arrow_dx As Single, arrow_dy As Single
        CalculateArrowPosition src_cx, src_cy, dst_cx, dst_cy, radius, arrow_sx, arrow_sy, arrow_dx, arrow_dy
        ' 辺を描画
        Dim arrow As Shape
        Set arrow = cur_page.Shapes.AddLine(arrow_sx, arrow_sy, arrow_dx, arrow_dy)

        ' 辺の設定
        ' 太さを1.0Ptに設定
        arrow.Line.Weight = 1.0

        ' 有向グラフの場合, 辺を描画
        if directed Then
                arrow.Line.EndArrowheadStyle = msoArrowheadTriangle
        End If
        ' 辺を最背面に移動
        arrow.ZOrder msoSendToBack
    Next i
End Sub

' 無向辺を引く
' 一つ目に選択した図形から, 二つ目以降に選択した図形に向かって辺を引く
Sub MakeEdgeUndirected1ToN()
    MakeEdge1ToN False
End Sub

' 有向辺を引く
' 一つ目に選択した図形から, 二つ目以降に選択した図形に向かって辺を引く
Sub MakeEdgeDirected1ToN()
    MakeEdge1ToN True
End Sub
