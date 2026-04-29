Option Explicit

'③【残す】Config取得 + チェック 約40行=====
Public Function GetTargetMonth() As Date

    Dim sh As Worksheet
    Dim yy As Long, mm As Long

    Set sh = ThisWorkbook.Worksheets(SH_CONFIG)

    ' Configシートのセルから年と月を取得
    yy = CLng(sh.Cells(2, 2).Value)
    mm = CLng(sh.Cells(2, 3).Value)

    ' 年が2000年〜2100年の範囲外であればエラーを発生させる
    If yy < 2000 Or yy > 2100 Then Err.Raise 1, , "年が不正です"
    
    ' 月が1月〜12月の範囲外であればエラーを発生させる
    If mm < 1 Or mm > 12 Then Err.Raise 1, , "月が不正です"

    ' 指定された年・月を日付型として返す
    GetTargetMonth = DateSerial(yy, mm, 1)

End Function

'------------------------------------------

'④【残す】月次カレンダー生成（中核） 約120〜140行 残すロジック（再構成）
Public Sub CreateCalendar()

    Dim baseDate As Date
    Dim calMap As Variant

    On Error GoTo ErrorHandler 

    ' ログ記録：処理開始
    WriteLog "CreateCalendar", "Start"

    ' 対象年月を取得
    baseDate = GetTargetMonth
    
    ' カレンダーの配列データを作成（※BuildCalendarArrayは別途定義が必要）
    calMap = BuildCalendarArray(baseDate)

    ' カレンダー情報をシートへ出力（※WriteCalendarは別途定義が必要）
    WriteCalendar calMap

    ' ログ記録：処理成功
    WriteLog "CreateCalendar", "Success"
    exit sub

ErrorHandler:
    WriteLog "CreateCalendar", "Error: " & Err.Number & " / " & Err.Description
    MsgBox "エラーが発生しました。ログを確認してください。", vbCritical
End Sub
