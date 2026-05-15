Option Explicit

Private Function BuildCalendarArray(ByVal baseDate As Date) As Variant

    Dim lastDay As Long
    Dim d As Long
    'Dim result(1 To 6, 1 To 7) As Variant   ' ← 行列が逆で Weekday の戻り値と不整合を修正 2026/05/15
    Dim result(1 To 7, 1 To 6) As Variant


    ' 月の最終日を取得
    lastDay = Day(DateSerial(Year(baseDate), Month(baseDate) + 1, 0))

    ' 1日から最終日までループ処理
    For d = 1 To lastDay
        ' Weekday関数を用いて、曜日に基づいた2次元配列のセル位置を計算
        ' result(行: 曜日[1-7], 列: 日付順) に格納していくロジック
        result(Weekday(DateSerial(Year(baseDate), Month(baseDate), d)), _
                  ((d - 1) \ 7) + 1) = d
    Next

    ' 計算結果の配列を関数の戻り値として返す
    BuildCalendarArray = result

End Function

'============================================================

'⑤【残す】書き込み処理 約40行=====
Private Sub WriteCalendar(ByVal calMap As Variant)

    Dim sh As Worksheet
    Dim r As Long, c As Long

    Set sh = ThisWorkbook.Worksheets(SH_CALEN)

    ' 既存のカレンダーデータをクリア
    sh.Cells.Clear

    ' 配列（calMap）の内容をシートへ順番に書き込む
    ' 行(r)と列(c)を入れ替えて出力するループ処理
    For r = 1 To UBound(calMap, 1)
        For c = 1 To UBound(calMap, 2)
            sh.Cells(c + 2, r).Value = calMap(r, c)
        Next
    Next

End Sub
