'①【残す】共通定義（最小限） 約40行=====
Option Explicit

Public Const SH_CONFIG As String = "Config"
Public Const SH_CALEN As String = "Calendar"
Public Const SH_LOG As String = "Log"

Public Const ROW_HEADER As Long = 1
Public Const ROW_DATA As Long = 2

'==========================================

'②【残す】ログ出力（縮小版） 約60行=====
Public Sub WriteLog(ByVal processName As String, ByVal status As String)

    Dim sh As Worksheet
    Dim nextRow As Long

    Set sh = ThisWorkbook.Worksheets(SH_LOG)

    ' ログシートが空の場合、ヘッダーを作成する
    If sh.Cells(1, 1).Value = "" Then
        sh.Range("A1:C1").Value = Array("処理名", "状態", "時刻")
    End If

    ' 最終行の次の行を取得
    nextRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row + 1

    ' ログ内容を書き込む
    sh.Cells(nextRow, 1).Value = processName
    sh.Cells(nextRow, 2).Value = status
    sh.Cells(nextRow, 3).Value = Now

End Sub

'==========================================

'③【残す】Config取得 + チェック 約40行=====