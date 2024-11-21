' 選択中のスライドのオブジェクトの名前, 種類, 座標を表示

Function MsoTypeToStr(ByVal shapeType As MsoShapeType) As String
    Select Case shapeType
        Case msoAutoShape: MsoTypeToStr = "オートシェイプ"
        Case msoCallout: MsoTypeToStr = "引き出し線"
        Case msoChart: MsoTypeToStr = "グラフ"
        Case msoComment: MsoTypeToStr = "コメント"
        Case msoFreeform: MsoTypeToStr = "フリーフォーム"
        Case msoGroup: MsoTypeToStr = "グループ"
        Case msoEmbeddedOLEObject: MsoTypeToStr = "埋め込みOLEオブジェクト"
        Case msoFormControl: MsoTypeToStr = "フォームコントロール"
        Case msoLine: MsoTypeToStr = "直線"
        Case msoLinkedOLEObject: MsoTypeToStr = "リンクOLEオブジェクト"
        Case msoLinkedPicture: MsoTypeToStr = "リンク画像"
        Case msoOLEControlObject: MsoTypeToStr = "OLEコントロールオブジェクト"
        Case msoPicture: MsoTypeToStr = "画像"
        Case msoPlaceholder: MsoTypeToStr = "プレースホルダ"
        Case msoTextEffect: MsoTypeToStr = "テキスト効果"
        Case msoMedia: MsoTypeToStr = "メディア"
        Case msoTextBox: MsoTypeToStr = "テキストボックス"
        Case msoScriptAnchor: MsoTypeToStr = "スクリプトアンカー"
        Case msoTable: MsoTypeToStr = "テーブル"
        Case msoCanvas: MsoTypeToStr = "キャンバス"
        Case msoDiagram: MsoTypeToStr = "ダイアグラム"
        Case msoInk: MsoTypeToStr = "墨"
        Case msoInkComment: MsoTypeToStr = "インクコメント"
        Case msoSmartArt: MsoTypeToStr = "スマートアート"
        Case msoSlicer: MsoTypeToStr = "スライサー"
        Case msoWebVideo: MsoTypeToStr = "Webビデオ"
        Case Else: MsoTypeToStr = "その他 (" & shapeType & ")"
    End Select
End Function

' 引数:
'   slide: スライド
Function DebugSelectedSlideObjects(ByRef slide As Slide) As String
    Dim shape As Shape
    Dim result As String
    For Each shape In slide.Shapes
        result = result & "名前:" & shape.Name & " " & _
                 "種類:" & MsoTypeToStr(shape.Type) & " " & _
                 "左:" & shape.Left & " " & _
                 "上:" & shape.Top & " " & _
                 "幅:" & shape.Width & " " & _
                 "高さ:" & shape.Height & vbCrLf
        result = result & vbCrLf
    Next shape
    DebugSelectedSlideObjects = result
End Function


Sub DebugSelectedSlides()
    Dim slides As SlideRange
    Dim slide As Slide

    If ActiveWindow.Selection.Type <> ppSelectionSlides Then
        MsgBox "選択したオブジェクトがスライドではありません."
        Exit Sub
    End If

    Set slides = ActiveWindow.Selection.SlideRange
    For Each slide In slides
        MsgBox "スライド: " & slide.Name & vbCrLf & DebugSelectedSlideObjects(slide)
    Next slide
End Sub
