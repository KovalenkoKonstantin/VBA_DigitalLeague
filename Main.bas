Attribute VB_Name = "Main"
Sub ��������()

 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim ws, this As Worksheet
 Dim pt As PivotTable
 Dim MyRange, MyCell As Range
 Dim key As String
 Dim x As Integer
 x = 9 '���������� ������ ��� �������
 
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
 MultiSelect:=True, Title:="���� ��� �������")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If
 
 '�������� ���������� ������
 On Error GoTo ExitHandler
 For i = 1 To x
 ThisWorkbook.Sheets(i).Activate
 Range("A1:N3000").Select
 With Selection
        .Clear
 End With
 Next i
 ThisWorkbook.Sheets("Preferences").Activate
 Range("Q3:Q300").Select
 With Selection
        .Clear
 End With
    '������ �����
    ThisWorkbook.Sheets(9).Activate
    Range("A1:N12000").Select
    With Selection
           .Clear
    End With
    '������� ���
    ThisWorkbook.Sheets("���").Activate
    Range("A1:C400000").Select
    With Selection
           .Clear
    End With

'������� ������
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
 
    '������ �����
    a = 9
    For i = 6 To a
    If i = a Then
    key = "H"
    Else
    key = "I"
    End If
'    i = 5
'    key = "I"
    importWB.Sheets(i).Activate
    lLastRow = Cells(Rows.Count, key).End(xlUp).Row
    j = lLastRow
    importWB.Sheets(i).Activate
    Range("A1:L" & j).Select
    Range("A1:L" & j).Copy
    ThisWorkbook.Sheets(i).Activate
    Range("A1:L" & j).Select
    With Selection
            .PasteSpecial Paste:=xlPasteAll
            .UnMerge
            .Font.Name = "Times New Roman"
            .WrapText = False
            .MergeCells = False
    End With
'         i = 6
'         key = "I"
'        importWB.Sheets(i).Activate
'        lLastRow = Cells(Rows.Count, key).End(xlUp).Row
'        j = lLastRow
'         importWB.Sheets(i).Activate
'         Range("A1:L" & j).Select
'         Range("A1:L" & j).Copy
'         ThisWorkbook.Sheets(i).Activate
'         Range("A1:L" & j).Select
'         With Selection
'                .PasteSpecial Paste:=xlPasteAll
'                .UnMerge
'                .Font.Name = "Times New Roman"
'                .WrapText = False
'                .MergeCells = False
'         End With
'             i = 7
'             key = "I"
'            importWB.Sheets(i).Activate
'            lLastRow = Cells(Rows.Count, key).End(xlUp).Row
'            j = lLastRow
'             importWB.Sheets(i).Activate
'             Range("A1:L" & j).Select
'             Range("A1:L" & j).Copy
'             ThisWorkbook.Sheets(i).Activate
'             Range("A1:L" & j).Select
'             With Selection
'                    .PasteSpecial Paste:=xlPasteAll
'                    .UnMerge
'                    .Font.Name = "Times New Roman"
'                    .WrapText = False
'                    .MergeCells = False
'             End With
'                 i = 8
'                 key = "H"
'                importWB.Sheets(i).Activate
'                lLastRow = Cells(Rows.Count, key).End(xlUp).Row
'                j = lLastRow
'                 importWB.Sheets(i).Activate
'                 Range("A1:L" & j).Select
'                 Range("A1:L" & j).Copy
'                 ThisWorkbook.Sheets(i).Activate
'                 Range("A1:L" & j).Select
'                 With Selection
'                        .PasteSpecial Paste:=xlPasteAll
'                        .UnMerge
'                        .Font.Name = "Times New Roman"
'                        .WrapText = False
'                        .MergeCells = False
'                 End With
    Next i

'���������� ������� ������
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

'����������
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




