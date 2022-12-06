Attribute VB_Name = "Inception"
Public Sub Inception()

 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim importWB  As Workbook
 On Error GoTo ErrHandler
 Set ThisWorkbook = ActiveWorkbook
 
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False
 
 FilesToOpen = Application.GetOpenFilename(MultiSelect:=True, Title:="���� ��� �����������")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If
   
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
 
 ThisWorkbook.Activate
 ���������������������
   
 importWB.Sheets(1).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("����������").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(2).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("����������2").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(3).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("����������3").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(4).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("����������4").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(5).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("����������5").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
ThisWorkbook.Activate
���������������������
ThisWorkbook.Sheets("Parsing").Activate
    
 importWB.Close

ExitHandler:
 Application.ScreenUpdating = True
 Application.EnableEvents = True
 ActiveSheet.DisplayPageBreaks = True
 Application.DisplayStatusBar = True
 Application.DisplayAlerts = True
 ThisWorkbook.Activate
 ���������������������
 ThisWorkbook.Sheets("Parsing").Activate
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

Public Sub ������()

 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim importWB  As Workbook
 On Error GoTo ErrHandler
 Set ThisWorkbook = ActiveWorkbook
 
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False
 
 FilesToOpen = Application.GetOpenFilename(MultiSelect:=True, Title:="���� ��� �����������")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If
   
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
 
 ThisWorkbook.Activate
 ���������������������
   
 importWB.Sheets(1).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("Inception").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(2).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("Inception2").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(3).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("Inception3").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(4).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("Inception4").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(5).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("Inception5").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
ThisWorkbook.Activate
���������������������
ThisWorkbook.Sheets("Parsing").Activate
    
 importWB.Close

ExitHandler:
 Application.ScreenUpdating = True
 Application.EnableEvents = True
 ActiveSheet.DisplayPageBreaks = True
 Application.DisplayStatusBar = True
 Application.DisplayAlerts = True
 ThisWorkbook.Activate
 ���������������������
 ThisWorkbook.Sheets("Parsing").Activate
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub
