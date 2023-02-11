Attribute VB_Name = "Support"
Sub weightloss()
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook
For i = 1 To 25
ThisWorkbook.Sheets(i).Activate
Range("A3000:BB30000").Select
Range(Selection, Selection.End(xlToRight)).Select
With Selection
    .Delete
End With
Next i
End Sub
Sub refresh()
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook
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
End Sub
Sub lastone()
Dim ThisWorkbook, importWB As Workbook
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="���� ��� �������")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "���� �� ������!"
 GoTo ExitHandler
 End If
 
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
 importWB.Sheets(1).Activate
 
 lLastRow = Cells(Rows.Count, "K").End(xlUp).Row
 ThisWorkbook.Sheets("Ranges").Activate
 Range("B15").Value = lLastRow

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
End Sub
Sub ���()

 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim importWB  As Workbook
 Dim this As Worksheet
 Set ThisWorkbook = ActiveWorkbook
 
 On Error GoTo ErrHandler
 
 Application.ScreenUpdating = False
 Application.DisplayAlerts = False
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="���� ��� �������")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "���� �� ������!"
 GoTo ExitHandler
 End If
 
 ���������������������
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

ThisWorkbook.Sheets("���").Activate
 Range("A1:BB400000").Select
 With Selection
        .Clear
 End With

 importWB.Sheets(1).Activate
 Range("A1:�400000").Select
 Range("A1:�400000").Copy
 ThisWorkbook.Sheets("���").Activate
 Range("A1:�400000").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

importWB.Close

 ThisWorkbook.Sheets("���").Activate
 Range("A1:A400000").Copy
 ThisWorkbook.Sheets("���").Activate
 Range("C1:C400000").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

Set this = ThisWorkbook.Sheets("Preferences")
 
 this.Activate
 Range("L2").Select
 ���������������������
 
' MsgBox "���������� ��� (���) ������� ��������"

ExitHandler:
 Application.ScreenUpdating = True
 ThisWorkbook.Sheets("Preferences").Activate
 ���������������������
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub
Sub ���������������������()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Protect Password:="gfhjkm"
Next ws
ActiveWorkbook.Sheets("������_��������").Unprotect Password:="gfhjkm"
ActiveWorkbook.Protect Password:="gfhjkm"
ThisWorkbook.Sheets("Preferences").Activate
End Sub

Sub ���������������������()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Unprotect Password:="gfhjkm"
Next ws
ActiveWorkbook.Unprotect Password:="gfhjkm"
ThisWorkbook.Sheets("Preferences").Activate
End Sub
