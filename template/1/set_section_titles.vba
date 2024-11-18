' カラーテーマとアクセントカラーの定数
Const THEME_NUMBER As Integer = 1  ' 使用するテーマの番号
Const ACCENT_COLOR_INDEX As Integer = msoThemeAccent4 ' 使用するアクセントカラーのインデックス
' 薄いグレーの色 (RGB: 204,204,204)
Const GRAY_COLOR As Long = 13421772

' スライドマスターの自作レイアウト名
Const AGENDA_LAYOUT_NAME As String = "Agenda Layout" ' Agenda スライドのレイアウト名
Const CONTENT_LAYOUT_NAME As String = "Content Layout" ' Content スライドのレイアウト名


'----------------------------------------------------------------------------------------------------

' 一つ目のアジェンダのテキストボックスの位置
Const FIRST_AGENDA_TEXTBOX_Y As Integer = 100
' 最後のアジェンダのテキストボックスの位置
Const LAST_AGENDA_TEXTBOX_Y As Integer = 435
' テキストボックスと円のY位置の差
Const AGENDA_CIRCLE_TEXTBOX_DIFF As Single = 7.65
' 円の高さ、幅
Const AGENDA_CIRCLE_HEIGHT_WIDTH As Integer = 30
' テキストボックスのフォント
Const TEXTBOX_FONT_NAME  As String = "YuGothic"
' 小さい円のテキストボックスと円のY位置の差
Const AGENDA_SMALL_CIRCLE_TEXTBOX_DIFF As Single = 16.44 - 8.787
'----------------------------------------------------------------------------------------------------
' コンテンツスライドの右端の円のX位置
Const CONTENT_RIGHT_CIRCLE_X As Single = 708.9
' コンテンツスライドの右端の円のY位置
Const CONTENT_RIGHT_CIRCLE_Y As Single = 21.5
' 小さい円の間隔
Const CONTENT_SMALL_CIRCLE_INTERVAL As Single = 11
' 小さい円の高さ、幅
Const CONTENT_SMALL_CIRCLE_HEIGHT_WIDTH As Single = 8.22
'----------------------------------------------------------------------------------------------------

' Agenda スライドのレイアウトを取得
Function GetAgendaLayout() As CustomLayout
    Dim layout As CustomLayout
    ' 通常は1つのデザインがメインで使われているので, 1番目のデザインを使用する
    For Each layout In ActivePresentation.Designs(1).SlideMaster.CustomLayouts
        If layout.Name = AGENDA_LAYOUT_NAME Then
            Set GetAgendaLayout = layout
            Exit Function
        End If
    Next layout
    
    ' 見つからない場合はエラー
    MsgBox "Agenda layout not found"
    Set GetAgendaLayout = Nothing
End Function

' Content スライドのレイアウトを取得
Function GetContentLayout() As CustomLayout
    Dim layout As CustomLayout
    ' 通常は1つのデザインがメインで使われているので, 1番目のデザインを使用する
    For Each layout In ActivePresentation.Designs(1).SlideMaster.CustomLayouts
        If layout.Name = CONTENT_LAYOUT_NAME Then
            Set GetContentLayout = layout
            Exit Function
        End If
    Next layout
    ' 見つからない場合はエラー
    MsgBox "Content layout not found"
    Set GetContentLayout = Nothing
End Function

' アクセントカラーを取得する関数
Function GetAccentColor() As Long
    ' 指定したテーマとアクセントカラーを使用して色を取得
    GetAccentColor = ActivePresentation.Designs(THEME_NUMBER).SlideMaster.Theme.ThemeColorScheme.Colors(ACCENT_COLOR_INDEX).RGB
End Function

' 入力が正しいかチェック
Function CheckInputSectionTitles(inputText As String) As Boolean
    ' 空文字列の場合はエラー
    If inputText = "" Then
        MsgBox "Input is empty"
        CheckInputSectionTitles = False
        Exit Function
    End If
    ' カンマで分割して、要素数をチェック
    Dim titles As Variant
    titles = Split(inputText, ",")
    ' それぞれの要素が空文字列でないかチェック
    Dim title As Variant
    For Each title In titles
        If title = "" Then
            MsgBox "Input is invalid"
            CheckInputSectionTitles = False
            Exit Function
        End If
    Next title
    CheckInputSectionTitles = True
End Function

' 章タイトルを入力する.
Function InputSectionTitles() As String
    InputSectionTitles = InputBox("章タイトルをカンマ区切りで入力してください")
End Function

' 章タイトルのリストを取得
Function GetSectionTitles(inputText As String) As Variant
    Dim titles As Variant
    Dim i As Integer
    
    titles = Split(inputText, ",")
    ' 前後の空白を削除
    For i = 0 To UBound(titles)
        titles(i) = Trim(titles(i))
    Next i
    
    GetSectionTitles = titles
End Function

' テキストボックスのY位置を取得する関数
Function GetTextboxYPosition(section_count As Integer, current_index As Integer) As Single
    If section_count = 1 Then
        ' セクションが1つしかない場合は中央に配置
        GetTextboxYPosition = (FIRST_AGENDA_TEXTBOX_Y + LAST_AGENDA_TEXTBOX_Y) / 2
    Else
        ' セクションが複数ある場合は均等に配置
        GetTextboxYPosition = FIRST_AGENDA_TEXTBOX_Y + current_index * (LAST_AGENDA_TEXTBOX_Y - FIRST_AGENDA_TEXTBOX_Y) / (section_count - 1)
    End If
End Function

' 円のY位置を取得する関数
Function GetCircleYPosition(section_count As Integer, current_index As Integer) As Single
    GetCircleYPosition = GetTextboxYPosition(section_count, current_index) + AGENDA_CIRCLE_TEXTBOX_DIFF
End Function

' 小さい円のY位置を取得する関数
Function GetSmallCircleYPosition(section_count As Integer, current_index As Integer) As Single
    GetSmallCircleYPosition = GetCircleYPosition(section_count, current_index) + AGENDA_SMALL_CIRCLE_TEXTBOX_DIFF
End Function

' Agenda スライドを作成する.
' 指定した章のインデックスをアクセント{ACCENT_COLOR_INDEX}の色に色付けして表示する
Function CreateAgendaSlide(section_titles As Variant, section_index As Integer) As slide
    ' 章のリストをテキストボックスに表示する
    Dim i As Integer
    Dim textbox As shape
    Dim yPosition As Single
    Dim agendaCircle As shape
    Dim smallCircle As shape
    Dim verticalLine As shape

    ' スライドを作成して, 末尾に追加
    Set CreateAgendaSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)

    ' スライドの日付、フッター、スライド番号を無効化
    With CreateAgendaSlide.HeadersFooters
        .Footer.Visible = msoFalse
        .DateAndTime.Visible = msoFalse
        .SlideNumber.Visible = msoFalse
    End With

    For i = 0 To UBound(section_titles)
        ' テキストボックスを作成
        ' 縦に並べる
        yPosition = GetTextboxYPosition(UBound(section_titles) + 1, i)
        Set textbox = CreateAgendaSlide.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, _
            165, _
            yPosition, _
            530, _
            40)
        textbox.TextFrame.TextRange.text = section_titles(i)
        ' 游ゴシック
        textbox.TextFrame.TextRange.Font.Name = TEXTBOX_FONT_NAME
        textbox.TextFrame.TextRange.Font.Color.RGB = GRAY_COLOR
        textbox.TextFrame.TextRange.Font.Size = 32
        textbox.TextFrame.TextRange.Font.Bold = msoTrue

        ' 円の図形を作成
        Set agendaCircle = CreateAgendaSlide.Shapes.AddShape(msoShapeOval, _
            115, _
            GetCircleYPosition(UBound(section_titles) + 1, i), _
            AGENDA_CIRCLE_HEIGHT_WIDTH, _
            AGENDA_CIRCLE_HEIGHT_WIDTH)
        ' 円の色を白にする
        agendaCircle.Fill.ForeColor.RGB = RGB(255, 255, 255)
        ' 円の図形の枠線を非表示にする
        agendaCircle.Line.Visible = msoFalse
        ' 円の図形の名前を設定
        agendaCircle.Name = "AgendaWhiteCircle" & i

        ' 指定した章のインデックスをアクセント{ACCENT_COLOR_INDEX}の色に色付けする
        If i = section_index Then
            textbox.TextFrame.TextRange.Font.Color.RGB = GetAccentColor()
            ' 小さい円を追加
            Set smallCircle = CreateAgendaSlide.Shapes.AddShape(msoShapeOval, _
                122.7, _
                GetSmallCircleYPosition(UBound(section_titles) + 1, i), _
                AGENDA_CIRCLE_HEIGHT_WIDTH / 2, _
                AGENDA_CIRCLE_HEIGHT_WIDTH / 2)
            ' 小さい円の色をアクセント{ACCENT_COLOR_INDEX}の色にする
            smallCircle.Fill.ForeColor.RGB = GetAccentColor()
            ' 小さい円の枠線を非表示にする
            smallCircle.Line.Visible = msoFalse 
            ' 最前面に配置
            smallCircle.ZOrder msoBringToFront
        End If
    Next i

    If UBound(section_titles) > 0 Then
        ' 白い縦棒を追加
        Set verticalLine = CreateAgendaSlide.Shapes.AddShape(msoShapeRectangle, _
            126.42, _
            FIRST_AGENDA_TEXTBOX_Y + 23, _
            8.5, _
            339)
        ' 縦棒の色を白にする
        verticalLine.Fill.ForeColor.RGB = RGB(255, 255, 255)
        ' 縦棒の枠線を非表示にする
        verticalLine.Line.Visible = msoFalse
        ' 最後面に配置
        verticalLine.ZOrder msoSendToBack
        ' 名前を設定
        verticalLine.Name = "AgendaVerticalLine"
    End If

    ' レイアウトを設定
    CreateAgendaSlide.CustomLayout = GetAgendaLayout
End Function

Function CreateContentSlide(section_titles As Variant, section_index As Integer) As slide
    ' スライドを作成して, 末尾に追加
    Set CreateContentSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
    ' スライドの日付、フッターを無効化
    With CreateContentSlide.HeadersFooters
        .Footer.Visible = msoFalse
        .DateAndTime.Visible = msoFalse
    End With
    ' レイアウトを設定
    CreateContentSlide.CustomLayout = GetContentLayout

    ' セクションタイトルを表示
    CreateContentSlide.Shapes.Title.TextFrame.TextRange.Text = section_titles(section_index)

    ' セクションタイトルのフォントを游ゴシックにする
    CreateContentSlide.Shapes.Title.TextFrame.TextRange.Font.Name = TEXTBOX_FONT_NAME
    ' 色をアクセント{ACCENT_COLOR_INDEX}の色にする
    CreateContentSlide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = GetAccentColor()

    ' 小さい円を追加する.
    Dim smallCircle As shape
    For i = 0 To UBound(section_titles)
        Set smallCircle = CreateContentSlide.Shapes.AddShape(msoShapeOval, _
            CONTENT_RIGHT_CIRCLE_X - i * CONTENT_SMALL_CIRCLE_INTERVAL, _
            CONTENT_RIGHT_CIRCLE_Y, _
            CONTENT_SMALL_CIRCLE_HEIGHT_WIDTH, _
            CONTENT_SMALL_CIRCLE_HEIGHT_WIDTH)
        ' 小さい円の色をアクセント白にする
        smallCircle.Fill.ForeColor.RGB = RGB(255, 255, 255)
        ' 小さい円の枠線を灰色にする
        smallCircle.Line.ForeColor.RGB = GRAY_COLOR
        ' 名前を設定 ContentSmallCircle{UBound(section_titles) - i}
        smallCircle.Name = "ContentSmallCircle" & (UBound(section_titles) - i)
    Next i

    ' i 番目の小さい円をアクセント{ACCENT_COLOR_INDEX}の色にする
    CreateContentSlide.Shapes("ContentSmallCircle" & section_index).Fill.ForeColor.RGB = GetAccentColor()
    ' 枠線を非表示にする
    CreateContentSlide.Shapes("ContentSmallCircle" & section_index).Line.Visible = msoFalse

End Function

Sub SetSectionTitles()
    ' 入力を受け取る
    Dim inputText As String
    Dim section_titles As Variant
    Dim i As Integer
    Dim newSlide As slide

    inputText = InputSectionTitles()
    ' 入力が正しいかチェック
    If Not CheckInputSectionTitles(inputText) Then
        Exit Sub
    End If
    ' 章タイトルのリストを取得
    section_titles = GetSectionTitles(inputText)
    For i = 0 To UBound(section_titles)
        ' アジェンダスライドを作成
        Set newSlide = CreateAgendaSlide(section_titles, i)
        ' コンテンツスライドを作成
        Set newSlide = CreateContentSlide(section_titles, i)
    Next i
End Sub


