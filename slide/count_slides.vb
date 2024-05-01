const R = 0.6 ' 縮小率
const TARGET_SHAPE_NAME = "Slide Number" ' スライド番号を表示するテキストボックスの名前
const DIVISOR = "/" ' スライド番号の区切り文字

' スライドの総数を返す
Function CountTotalSlides() As Integer
    CountTotalSlides = ActivePresentation.Slides.Count
End Function

' すべてのスライドについて、
' [現在のスライド番号]　DIVISOR　[スライドの総数] という形式でスライド番号を表示
Sub ShowSlideNumber()
    ' 総数を取得
    Dim total_slides As Integer
    total_slides = CountTotalSlides()

    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        ' 現在のスライド番号を取得
        Dim slide_number As Integer
        slide_number = sld.SlideIndex
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
    Next sld
End Sub
