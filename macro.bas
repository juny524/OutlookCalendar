Option Explicit

Sub ExportOutlookAppointmentsToExcel()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.Namespace
    Dim olFolder As Outlook.Folder
    Dim olItems As Outlook.Items
    Dim olAppt As Outlook.AppointmentItem
    Dim ws As Worksheet
    Dim i As Integer
    Dim startDate As Date
    Dim endDate As Date
    Dim strFilter As String
    Dim olFilteredItems As Outlook.Items

    ' Outlookアプリケーションの初期化
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    ' 名前空間とカレンダーフォルダの取得
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderCalendar)

    ' 取得期間の設定（例：今日から30日後まで）
    startDate = Date
    endDate = DateAdd("d", 30, startDate)

    ' フィルターの作成
    strFilter = "[Start] >= '" & Format(startDate, "ddddd h:nn AMPM") & "' AND [End] <= '" & Format(endDate, "ddddd h:nn AMPM") & "'"

    ' フィルターを適用したアイテムの取得
    Set olItems = olFolder.Items
    olItems.Sort "[Start]"
    olItems.IncludeRecurrences = True
    Set olFilteredItems = olItems.Restrict(strFilter)

    ' 出力先のワークシートの設定
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Cells.ClearContents

    ' ヘッダーの設定
    ws.Cells(1, 1).Value = "件名"
    ws.Cells(1, 2).Value = "開始日時"
    ws.Cells(1, 3).Value = "終了日時"
    ws.Cells(1, 4).Value = "場所"
    ws.Cells(1, 5).Value = "定期的な予定"

    ' 予定の取得とExcelへの書き込み
    i = 2
    For Each olAppt In olFilteredItems
        ws.Cells(i, 1).Value = olAppt.Subject
        ws.Cells(i, 2).Value = olAppt.Start
        ws.Cells(i, 3).Value = olAppt.End
        ws.Cells(i, 4).Value = olAppt.Location
        ws.Cells(i, 5).Value = IIf(olAppt.IsRecurring, "はい", "いいえ")
        i = i + 1
    Next olAppt

    ' 後処理
    Set olAppt = Nothing
    Set olItems = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing

    MsgBox "Outlookの予定をExcelにエクスポートしました。", vbInformation
End Sub

