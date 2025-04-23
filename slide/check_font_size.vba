Sub CheckFontSize()
    Dim sld As Slide
    Dim shp As Shape
    Dim minSize As Single
    Dim problemFound As Boolean
    
    problemFound = False
    
    ' スライドが選択されているか確認
    On Error Resume Next
    Dim selType As PpSelectionType
    selType = ActiveWindow.Selection.Type
    On Error GoTo 0
    
    If selType <> ppSelectionSlides Then
        MsgBox "スライドが選択されていません。チェックするスライドを選択してください。", vbExclamation
        Exit Sub
    End If
    
    ' 最小フォントサイズの入力を求める
    Dim inputStr As String
    inputStr = InputBox("警告を表示する最小フォントサイズを入力してください" & vbNewLine & _
                        "（例: 18）", "フォントサイズチェック")
    
    ' キャンセルボタンが押された場合
    If inputStr = "" Then
        Exit Sub
    End If
    
    ' 数値かどうかチェック
    If Not IsNumeric(inputStr) Then
        MsgBox "整数を入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 整数かどうかチェック
    minSize = CDbl(inputStr)
    If minSize <> Int(minSize) Then
        MsgBox "整数を入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 正の値かどうかチェック
    If minSize <= 0 Then
        MsgBox "正の整数を入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 選択されているスライドをチェック
    For Each sld In ActiveWindow.Selection.SlideRange
        ' 各シェイプをチェック
        For Each shp In sld.Shapes
            ' テキストを含むオブジェクトの場合
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    ' テキストの各行をチェック
                    Dim textRange As TextRange
                    Set textRange = shp.TextFrame.TextRange
                    
                    ' フォントサイズをチェック
                    If textRange.Font.Size < minSize Then
                        problemFound = True
                        Dim displayText As String
                        If Len(textRange.Text) > 200 Then
                            displayText = Left(textRange.Text, 200) & "..."
                        Else
                            displayText = textRange.Text
                        End If
                        
                        MsgBox "警告: スライド " & sld.SlideIndex & " に" & _
                               "指定された最小サイズ(" & minSize & ")より小さい" & _
                               "フォントサイズ(" & textRange.Font.Size & ")が見つかりました。" & vbNewLine & _
                               "テキスト: " & displayText, _
                               vbExclamation
                    End If
                End If
            End If
        Next shp
    Next sld
    
    If Not problemFound Then
        MsgBox "チェックが完了しました。指定サイズより小さいフォントは見つかりませんでした。", vbInformation
    Else
        MsgBox "チェックが完了しました。", vbInformation
    End If
End Sub
