const R = 0.6 ' 縮小率
const TARGET_SHAPE_NAME = "Slide Number" ' スライド番号を表示するテキストボックスの名前
const DIVISOR = "/" ' スライド番号の区切り文字

' スライドの総数を返す
' 非表示スライドは含まない
Function CountTotalSlides() As Integer
    CountTotalSlides = 0
    For Each sld In ActivePresentation.Slides
        If sld.SlideShowTransition.Hidden = msoFalse Then
            CountTotalSlides = CountTotalSlides + 1
        End If
    Next sld
End Function

' すべてのスライドについて、
' [現在のスライド番号]　DIVISOR　[スライドの総数] という形式でスライド番号を表示
' 非表示のスライドはスライド番号を表示しない
Sub ShowSlideNumber()
    ' 総数を取得
    Dim total_slides As Integer
    total_slides = CountTotalSlides()

    Dim sld As Slide
    Dim slide_number As Integer
    slide_number = 0
    For Each sld In ActivePresentation.Slides
        ' 非表示スライドはスキップ
        if sld.SlideShowTransition.Hidden = msoTrue Then
            Go To Continue
        End If
        ' 現在のスライド番号を計算
        slide_number = slide_number + 1
        Dim shp As Shape
        For Each shp In sld.Shapes
            ' スライド番号を表示するテキストボックスの名前であるかを判定
            If Left(shp.Name, Len(TARGET_SHAPE_NAME)) = TARGET_SHAPE_NAME Then
                ' スライド番号を表示
                shp.TextFrame.TextRange.Text = slide_number & DIVISOR & total_slides
                ' total_slides の部分のみを小さく表示する
                Dim text_range As TextRange
                ' InStr で DIVISOR の位置を取得し, その後の文字列を取得
                ' Charactes(開始位置, 長さ) で文字列を取得
                Set text_range = shp.TextFrame.TextRange.Characters(InStr(shp.TextFrame.TextRange.Text, DIVISOR) + 1, Len(shp.TextFrame.TextRange.Text))
                ' 元の大きさの R 倍にする
                text_range.Font.Size = text_range.Font.Size * R
            End If
        Next shp
Continue:
    Next sld
End Sub
