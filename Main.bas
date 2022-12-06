Attribute VB_Name = "Main"
Sub нахлабуч()

 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim importWB  As Workbook
 Dim ws As Worksheet
 Dim pt As PivotTable
 Dim this As Worksheet
 Set ThisWorkbook = ActiveWorkbook
 Dim MyRange As Range
 Dim MyCell As Range
 
 On Error GoTo ErrHandler
 
 Application.ScreenUpdating = False
 Application.DisplayAlerts = False
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="‘айл дл€ вставки")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "‘айл не выбран!"
 GoTo ExitHandler
 End If
 
 —н€ть«ащиту¬сехЋистов
 
 'удаление проверок переноса
    ThisWorkbook.Sheets("Preferences").Range("W5:X505").Select
            With Selection
                .Clear
            End With
    ThisWorkbook.Sheets("Preferences").Range("Y11:AD511").Select
            With Selection
                .Clear
            End With
    ThisWorkbook.Sheets("Preferences").Range("AE11:AE511").Select
            With Selection
                .Clear
            End With
    ThisWorkbook.Sheets("Preferences").Range("AF20:AK520").Select
            With Selection
                .Clear
            End With
 
 'восстановление формул на листе с »ЌЌ
 ThisWorkbook.Sheets("90 контр").Activate
    Range("E4").Copy

    Range("E9:E200").Select
    Set MyRange = Selection
    With Selection
            .PasteSpecial xlPasteFormulas
    End With
 
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

ThisWorkbook.Sheets("Data90").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("Data90-1").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("Data90-2").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("Preferences").Activate
 Range("AT2").Select
 With Selection
        .Clear
 End With


For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws

For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws

 importWB.Sheets(1).Activate
 Range("A1:BB300").Select
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("Data90").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(2).Activate
 Range("A1:BB300").Select
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("Data90-1").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(3).Activate
 Range("A1:BB300").Select
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("Data90-2").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

importWB.Close

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

Set this = ThisWorkbook.Sheets("Preferences")

 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"

 this.Activate
 Range("AT1").Copy
 this.Activate
 Range("AT2").Select
 Selection.PasteSpecial Paste:=xlPasteValues
 
 this.Activate
 Range("L2").Select
 «аблокировать¬сеЋисты

ExitHandler:
 Application.ScreenUpdating = True
 ThisWorkbook.Sheets("Preferences").Activate
 «аблокировать¬сеЋисты
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub


