Attribute VB_Name = "Refactor"
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
 Range("A1:BB400000").Select
 Range("A1:BB400000").Copy
 ThisWorkbook.Sheets("���").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

importWB.Close

Set this = ThisWorkbook.Sheets("Preferences")
 
 this.Activate
 Range("L2").Select
 ���������������������
 
 MsgBox "���������� ��� (���) ������� ��������"

ExitHandler:
 Application.ScreenUpdating = True
 ThisWorkbook.Sheets("Preferences").Activate
 ���������������������
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub
Sub ��������_���()

 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim this As Worksheet
' Dim MyRange As Range
' Dim MyCell As Range
 
 Set ThisWorkbook = ActiveWorkbook
 
 On Error GoTo ErrHandler
 
 Application.ScreenUpdating = False
 Application.DisplayAlerts = False

 ���������������������

ThisWorkbook.Sheets("90 �����").Activate
 Range("K9:K900").Copy

 Range("E9:E900").Select
 Selection.PasteSpecial Paste:=xlPasteValues
'            Range("E9:E900").Select
'            Set MyRange = Selection
'            For Each MyCell In MyRange
'            If MyCell.Value = 0 Then
'            MyCell.Value = Empty
'            End If
'            Next MyCell
'
Set this = ThisWorkbook.Sheets("Preferences")

 this.Activate
 Range("L2").Select
 ���������������������
 
 MsgBox "��� ����������"

ExitHandler:
 Application.ScreenUpdating = True
 ThisWorkbook.Sheets("Preferences").Activate
 ���������������������
 ClearClipboard
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

