const BAR_HEIGHT = 1
const BAR_COLOR = 0 ' 0 は黒

' すべてのスライドにプログレスバーを表示する
Sub AddProgressBar()
    Dim i As Integer
    ' ActivePresentation.Slides.Count はスライドの総数
    Dim total_slides As Integer
    total_slides = ActivePresentation.Slides.Count
    For i = 1 To total_slides
        ' i 番目のスライドを取得
        Dim slide As slide
        Set slide = ActivePresentation.Slides(i)

        ' プログレスバーを追加
        Dim shape As shape
        ' ProgressBar という名前の長方形がすでに存在する場合は削除
        On Error Resume Next
        slide.Shapes("ProgressBar").Delete
        On Error GoTo 0

        ' スライドの幅と高さを取得
        Dim slide_height As Single
        slide_height = ActivePresentation.PageSetup.SlideHeight
        Dim slide_width As Single
        slide_width = ActivePresentation.PageSetup.SlideWidth

        ' 長方形を追加
        Set shape = slide.Shapes.AddShape(msoShapeRectangle, 0, slide_height - BAR_HEIGHT, slide_width * i / total_slides, BAR_HEIGHT)
        ' 名前を設定
        shape.Name = "ProgressBar"
        ' プログレスバーの色を黒に設定
        shape.Fill.ForeColor.RGB = BAR_COLOR
        ' プログレスバーの位置とサイズを設定
        shape.Left = 0
        shape.Top = slide_height - 10
        shape.Width = slide_width * i / total_slides
        shape.Height = BAR_HEIGHT
    Next i
End Sub
