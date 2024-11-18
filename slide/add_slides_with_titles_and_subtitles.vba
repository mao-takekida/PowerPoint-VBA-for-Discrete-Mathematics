Sub AddSlidesWithTitlesAndSubtitles()
    ' タイトルとサブタイトルのリストをユーザーから取得する
    Dim titlesAndSubtitles As Variant ' タイトルとサブタイトルを格納するための配列
    Dim inputString As String ' ユーザーからの入力文字列を格納
    Dim inputArray() As String ' 入力された文字列を分割して格納する配列
    Dim i As Long ' ループ用のカウンタ変数

    ' ユーザーにタイトルとサブタイトルをまとめて入力させる（簡易化）
    ' ユーザーが一度に全てのタイトルとサブタイトルを入力できるようにして、作業の効率化を図る
    inputString = InputBox( _
        "各スライドのタイトルとサブタイトルをカンマで区切り、" & vbNewLine & _
        "スライド間を縦棒で区切って入力してください。" & vbNewLine & _
        vbNewLine & _
        "形式：" & vbNewLine & _
        "title1, subtitle1 | title2, subtitle2 ... | titleN, subtitleN" & vbNewLine & _
        vbNewLine & _
        "例：" & vbNewLine & _
        "頂点と辺, グラフの頂点と辺の定義 | 最短経路, ダイクストラ法による最短経路 | オイラー路, オイラー路とオイラー閉路の条件", _
        "タイトルとサブタイトルの入力")
    If inputString = "" Then
        ' ユーザーが入力をキャンセルした場合の処理
        ' 処理を中断し、適切なメッセージを表示してユーザーに通知する
        MsgBox "入力がキャンセルされました。", vbExclamation ' キャンセルされた場合の警告メッセージ
        Exit Sub ' サブプロシージャの実行を終了
    End If

    ' 入力文字列をパースして配列に変換
    ' 縦棒で分割して各スライドの情報を配列に格納
    inputArray = Split(inputString, "|") ' 入力文字列を縦棒で区切り配列に変換
    ReDim titlesAndSubtitles(UBound(inputArray)) ' 配列のサイズを入力の数に応じて再定義
    For i = 0 To UBound(inputArray)
        Dim pair() As String ' タイトルとサブタイトルを格納する一時的な配列
        pair = Split(inputArray(i), ",") ' タイトルとサブタイトルをカンマで区切って配列に格納
        If UBound(pair) = 1 Then
            ' タイトルとサブタイトルを配列に格納（前後の空白を除去）
            ' ユーザー入力の整形を行い、不要な空白を削除することで正確な処理を実現する
            titlesAndSubtitles(i) = Array(Trim(pair(0)), Trim(pair(1))) ' トリムで空白を削除して格納
        Else
            ' 入力形式が無効な場合のエラーメッセージ
            ' 入力が不正な場合にはメッセージを表示して再度試すよう促す
            MsgBox "無効な入力形式です。タイトルとサブタイトルをカンマで区切ってください。", vbExclamation ' エラーメッセージの表示
            Exit Sub ' サブプロシージャの実行を終了
        End If
    Next i

    ' PowerPoint のプレゼンテーションオブジェクトを取得する
    ' 既存の PowerPoint アプリケーションに接続するか、新規に起動する
    Dim pptApp As Object ' PowerPoint アプリケーションオブジェクト
    Dim pptPres As Presentation ' プレゼンテーションオブジェクト
    
    ' PowerPoint がすでに開かれているか確認し、開かれていない場合は新規で起動する
    On Error Resume Next ' エラーが発生しても処理を続行（既に開いているかを確認するため）
    Set pptApp = GetObject(, "PowerPoint.Application") ' 既存の PowerPoint アプリケーションに接続
    If pptApp Is Nothing Then
        ' PowerPoint が起動していない場合、新しいインスタンスを作成
        Set pptApp = CreateObject("PowerPoint.Application") ' PowerPoint を新規に起動
    End If
    On Error GoTo 0 ' エラーハンドリングを通常に戻す

    ' 現在開いているプレゼンテーションを取得する
    ' 既存のプレゼンテーションがない場合、エラーとして処理を終了
    If pptApp.Presentations.Count > 0 Then
        Set pptPres = pptApp.ActivePresentation ' 現在開いているプレゼンテーションを取得
    Else
        ' プレゼンテーションが開かれていない場合のエラーメッセージ
        MsgBox "プレゼンテーションが開かれていません。", vbExclamation ' プレゼンテーションがない場合の警告メッセージ
        Exit Sub ' サブプロシージャの実行を終了
    End If

    ' titlesAndSubtitles の配列に基づいてスライドを追加する
    ' 各タイトルとサブタイトルを用いてスライドを順次追加する
    For i = LBound(titlesAndSubtitles) To UBound(titlesAndSubtitles)
        ' 新規スライドを追加
        ' スライドは現在のスライド数の最後に追加される
        Dim newSlide As Slide ' 新規スライドオブジェクト
        Set newSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutTitle) ' 新規スライドを追加

        ' タイトルとサブタイトルを設定
        With newSlide
            ' スライドのタイトルを設定
            ' タイトルプレースホルダーが存在する場合にのみ設定を行う
            If .Shapes.HasTitle Then
                .Shapes.Title.TextFrame.TextRange.Text = titlesAndSubtitles(i)(0) ' タイトルの設定
            End If
            ' サブタイトル（ボディ部分）を設定
            ' ボディ部分のプレースホルダーが存在する場合にのみ設定を行う
            If .Shapes.Placeholders.Count >= 2 Then
                .Shapes.Placeholders(2).TextFrame.TextRange.Text = titlesAndSubtitles(i)(1) ' サブタイトルの設定
            End If
        End With
    Next i

    ' オブジェクトの解放
    ' メモリリークを防ぐためにオブジェクトを解放する
    Set pptPres = Nothing ' プレゼンテーションオブジェクトを解放
    Set pptApp = Nothing ' PowerPoint アプリケーションオブジェクトを解放
End Sub
