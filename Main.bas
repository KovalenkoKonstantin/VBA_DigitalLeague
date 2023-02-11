Attribute VB_Name = "Main"
Sub Data_insertion()
 
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim ws, this As Worksheet
 Dim pt As PivotTable
 Dim MyRange, MyCell As Range
 Dim key As String
 Dim x As Integer
 x = 24 'количество листов для вставки
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 
Application.ScreenUpdating = False
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Файл для вставки")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If
 
 'удаление предыдущих данных
 On Error GoTo ExitHandler
 For i = 1 To x
 ThisWorkbook.Sheets(i).Activate
 Range("A1:N3000").Select
 With Selection
        .Clear
 End With
 Next i
 
 ThisWorkbook.Sheets("Preferences").Activate
 Range("Q3:Q33").Select
 With Selection
        .Clear
 End With
 Range("L2").Copy
 Range("Q3:Q33").Select
 With Selection
        .PasteSpecial Paste:=xlPasteFormats
 End With

'вставка листов
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
    For i = 1 To x
    On Error Resume Next
     importWB.Sheets(i).Activate
     lLastRow = Cells(Rows.Count, "K").End(xlUp).Row
     j = lLastRow
     
     importWB.Sheets(i).Activate
     Range("A1:N" & j).Select
     Range("A1:N" & j).Copy
     ThisWorkbook.Sheets(i).Activate
     Range("A1:N" & j).Select
     With Selection
            .PasteSpecial Paste:=xlPasteAll
            .UnMerge
            .Font.Name = "Times New Roman"
            .WrapText = False
            .MergeCells = False
     End With
    Next i

'обновление сводных таблиц
ThisWorkbook.Activate
For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws
ThisWorkbook.Activate
For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws

'завершение
importWB.Close
ThisWorkbook.Sheets("Preferences").Activate

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub





