Attribute VB_Name = "Transfer"
Sub �����������������()

 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim import, this As Worksheet
 Dim MyRange, MyCell As Range
 Dim SaveName, Range1, Range2, range3, range4, range5, range6, y, Ra1, Ra2, Ra3, Ra4, Ra5, Ra6, Ra7, Ra8, Ra9, Ra10, y1, z1, y2, z2, y3, z3, y4, z4, y5, z5, y6, z6, Folder, Path, Slash As String
 Dim x As Integer '������ ��� ����������
 Dim object As Object
 Dim a As Integer '���������� ���������� ����� �������
' a = 31 '����� ������ �� ������� ����� �������� ������� �����
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
 
 Set ThisWorkbook = ActiveWorkbook
 ThisWorkbook.Sheets("Ranges").Activate
 a = Range("A1").Value '��������� ������ � Ranges
 On Error GoTo ErrHandler
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsm), *.xlsm", _
 MultiSelect:=True, Title:="���� ��� �������")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "���� �� ������!"
 GoTo ExitHandler
 End If
 
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
 
 '���������� �����
 
 '�������� ������������� ���������� ��������
 ThisWorkbook.Sheets("Preferences").Activate
 If Range("L16").Value = False Then
    GoTo Content
 End If
 
 '�������� ���������� �����
    On Error Resume Next
    For i = 2 To a

    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "B" & i
    Range2 = "A" & i
    x = Range(Range1).Value
    y = Range(Range2).Text
    importWB.Sheets(y).Activate
    
    ActiveSheet.Protect Password:="tesla", AllowInsertingRows:=True
    ActiveSheet.Unprotect Password:="tesla"
    flag = x
        counter = 1
        Do While counter <= flag
        counter = counter + 1
        lastrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
        Rows(lastrow).AutoFill Rows(lastrow).Resize(2), xlFillDefault
        On Error Resume Next
        Rows(lastrow + 1).SpecialCells(xlConstants).ClearContents
        On Error GoTo ExitHandler
        Rows(lastrow + 1).SpecialCells(xlCellTypeBlanks).Item(1).Activate
        Rows(lastrow - 1).Select
        Selection.Copy
        Rows(lastrow).Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        Loop
    ActiveSheet.Protect Password:="tesla", DrawingObjects:=True, Contents:=True, Scenarios:=True _
       , AllowFormattingCells:=False, AllowFormattingColumns:=True, _
         AllowFormattingRows:=True, AllowInsertingColumns:=False, AllowInsertingRows _
       :=False, AllowSorting:=False, AllowFiltering:=True, AllowUsingPivotTables _
       :=False
Next i
  
'����������
Content:

 '�������� ������������� ���������� ��������
 ThisWorkbook.Sheets("Preferences").Activate
 If Range("L17").Value = False Then
    GoTo 1
 End If

 '��������� ��������
 Set import = importWB.Sheets("����������")
 Set this = ThisWorkbook.Sheets("Preferences")
 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 
 '��������� ����������
import.Activate
Range("D2:E2").Select
Range("D2:E2").UnMerge

Range("D5:E5").Select
Range("D5:E5").UnMerge

Range("D6:E6").Select
Range("D6:E6").UnMerge

Range("D7:E7").Select
Range("D7:E7").UnMerge

Range("D8:E8").Select
Range("D8:E8").UnMerge

Range("D9:E9").Select
Range("D9:E9").UnMerge

Range("D10:E10").Select
Range("D10:E10").UnMerge

Range("D11:E11").Select
Range("D11:E11").UnMerge

Range("D12:E12").Select
Range("D12:E12").UnMerge

Range("D15:E15").Select
Range("D15:E15").UnMerge

this.Activate
Range("AA1").Copy
import.Activate
Range("D2").Select
Selection.PasteSpecial Paste:=xlPasteValues

this.Activate
Range("AA2:AA9").Copy
import.Activate
Range("D5:D12").Select
Selection.PasteSpecial Paste:=xlPasteValues

this.Activate
Range("AA10").Copy
import.Activate
Range("D15").Select
Selection.PasteSpecial Paste:=xlPasteValues

import.Activate
Range("D2:E2").Select
Range("D2:E2").Merge

Range("D5:E5").Select
Range("D5:E5").Merge

Range("D6:E6").Select
Range("D6:E6").Merge

Range("D7:E7").Select
Range("D7:E7").Merge

Range("D8:E8").Select
Range("D8:E8").Merge

Range("D9:E9").Select
Range("D9:E9").Merge

Range("D10:E10").Select
Range("D10:E10").Merge

Range("D11:E11").Select
Range("D11:E11").Merge

Range("D12:E12").Select
Range("D12:E12").Merge

Range("D15:E15").Select
Range("D15:E15").Merge

import.Activate
ActiveSheet.Protect Password:="tesla"

'�����
1:

i = 2
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 2
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        Ra5 = "W" & i
        '3 �������
        y3 = Range(Ra5).Text
        Ra6 = "X" & i
        z3 = Range(Ra6).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
            
        '������� 3
        this.Activate
        Range(y3).Copy
        import.Activate
        Range(z3).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"

2:
 
i = 3
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 4
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        Ra5 = "W" & i
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
4:

i = 4
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 5
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
            '�������� �����
            Range(z1).Select
            Set MyRange = Selection
            For Each MyCell In MyRange
            If MyCell.Value = 0 Then
            MyCell.Value = Empty
            End If
            Next MyCell
            '��� ������� (����������)
            import.Activate
            Range("E23:E32").Select
            With Selection
                .Clear
            End With
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
5:

i = 5
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 71
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
            '�������� �����
            Range(z1).Select
            Set MyRange = Selection
            For Each MyCell In MyRange
            If MyCell.Value = 0 Then
            MyCell.Value = Empty
            End If
            Next MyCell
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"

71:

i = 7
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 7
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"

7:

i = 6
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 8
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
8:

i = 8
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 81
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
81:

i = 9
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 9
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
9:

i = 10
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 10
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"

10:

i = 11
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 101
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"

101:

i = 12
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 19
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
19:

i = 13
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 77
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        Ra5 = "W" & i
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
77:

i = 14
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 99
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
99:

i = 15
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 4101
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
4101:

i = 16
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo �����
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
�����:

i = 17
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo Financialcontracts
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        Ra5 = "W" & i
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
5802:

i = 18
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 58021
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        Ra5 = "W" & i
        '3 �������
        y3 = Range(Ra5).Text
        Ra6 = "X" & i
        z3 = Range(Ra6).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
            
        '������� 3
        this.Activate
        Range(y3).Copy
        import.Activate
        Range(z3).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
58021:

i = 19
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 66
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        Ra5 = "W" & i
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
66:

'����� ������
i = 20
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo Financialcontracts
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        '3 �������
        Ra5 = "W" & i
        y3 = Range(Ra5).Text
        Ra6 = "X" & i
        z3 = Range(Ra6).Text
        '4 �������
        Ra7 = "AC" & i
        y4 = Range(Ra7).Text
        Ra8 = "AD" & i
        z4 = Range(Ra8).Text
        '5 �������
        Ra9 = "AI" & i
        y5 = Range(Ra9).Text
        Ra10 = "AJ" & i
        z5 = Range(Ra10).Text
        
    '������� ������ � ������
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 3
        this.Activate
        Range(y3).Copy
        import.Activate
        Range(z3).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 4
        this.Activate
        Range(y4).Copy
        import.Activate
        Range(z4).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 5
        this.Activate
        Range(y5).Copy
        import.Activate
        Range(z5).Select
        Selection.PasteSpecial Paste:=xlPasteValues

    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"

Financialcontracts:

'����� ������
i = 21
'���������� ��������
k = 2
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 68
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        
    '������� ������ � ������
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        On Error Resume Next
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
68:

'����� ������
i = 22
'���������� ��������
k = 3
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 69
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        '3 �������
        Ra5 = "W" & i
        y3 = Range(Ra5).Text
        Ra6 = "X" & i
        z3 = Range(Ra6).Text
        
    '������� ������ � ������
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        On Error Resume Next
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 3
        On Error Resume Next
        this.Activate
        Range(y3).Copy
        import.Activate
        Range(z3).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
69:

i = 23
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 73
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        Ra5 = "W" & i
        '3 �������
        Ra5 = "W" & i
        y3 = Range(Ra5).Text
        Ra6 = "X" & i
        z3 = Range(Ra6).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 3
        On Error Resume Next
        this.Activate
        Range(y3).Copy
        import.Activate
        Range(z3).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
73:

i = 24
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 80
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
80:

i = 25
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 84
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
84:

i = 26
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 96
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
96:

'����� ������
i = 27
'���������� ��������
k = 3
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 97
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        '3 �������
        Ra5 = "W" & i
        y3 = Range(Ra5).Text
        Ra6 = "X" & i
        z3 = Range(Ra6).Text
        
    '������� ������ � ������
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        On Error Resume Next
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 3
        On Error Resume Next
        this.Activate
        Range(y3).Copy
        import.Activate
        Range(z3).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
97:

'����� ������
i = 28
'���������� ��������
k = 3
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo incometax
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        '3 �������
        Ra5 = "W" & i
        y3 = Range(Ra5).Text
        Ra6 = "X" & i
        z3 = Range(Ra6).Text
        
    '������� ������ � ������
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        On Error Resume Next
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 3
        On Error Resume Next
        this.Activate
        Range(y3).Copy
        import.Activate
        Range(z3).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
incometax:

'����� ������
i = 29
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo summaryofaccounts
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        '3 �������
        Ra5 = "W" & i
        y3 = Range(Ra5).Text
        Ra6 = "X" & i
        z3 = Range(Ra6).Text
        '4 �������
        Ra7 = "AC" & i
        y4 = Range(Ra7).Text
        Ra8 = "AD" & i
        z4 = Range(Ra8).Text
        '5 �������
        Ra9 = "AI" & i
        y5 = Range(Ra9).Text
        Ra10 = "AJ" & i
        z5 = Range(Ra10).Text
        
    '������� ������ � ������
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 3
        this.Activate
        Range(y3).Copy
        import.Activate
        Range(z3).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 4
        this.Activate
        Range(y4).Copy
        import.Activate
        Range(z4).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 5
        this.Activate
        Range(y5).Copy
        import.Activate
        Range(z5).Select
        Selection.PasteSpecial Paste:=xlPasteValues

    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"

summaryofaccounts:
i = 30
ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo osv
     End If
   '����_��_������
 Set import = importWB.Sheets("����_��_������")
 Set this = ThisWorkbook.Sheets("����_��_������")

 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 '��������� ����_��_������
 '1
 this.Activate
 Range("C4").Copy
 import.Activate
 Range("E9").Select
 Selection.PasteSpecial Paste:=xlPasteValues
      '2
    this.Activate
    Range("C14").Copy
    import.Activate
    Range("E19").Select
    Selection.PasteSpecial Paste:=xlPasteValues
        '3
        this.Activate
        Range("C16:C17").Copy
        import.Activate
        Range("E21:E22").Select
        Selection.PasteSpecial Paste:=xlPasteValues
            '4
            this.Activate
            Range("C20").Copy
            import.Activate
            Range("E25").Select
            Selection.PasteSpecial Paste:=xlPasteValues
                '5
                this.Activate
                Range("C23:C25").Copy
                import.Activate
                Range("E28:E30").Select
                Selection.PasteSpecial Paste:=xlPasteValues
                    '6
                    this.Activate
                    Range("C32").Copy
                    import.Activate
                    Range("E37").Select
                    Selection.PasteSpecial Paste:=xlPasteValues
                        '7
                        this.Activate
                        Range("D22:D24").Copy
                        import.Activate
                        Range("F27:F29").Select
                        Selection.PasteSpecial Paste:=xlPasteValues
                            '8
                            this.Activate
                            Range("D26:D27").Copy
                            import.Activate
                            Range("F31:F32").Select
                            Selection.PasteSpecial Paste:=xlPasteValues
                                '9
                                this.Activate
                                Range("D34").Copy
                                import.Activate
                                Range("F39").Select
                                Selection.PasteSpecial Paste:=xlPasteValues
                                    '10
                                    this.Activate
                                    Range("E14:E15").Copy
                                    import.Activate
                                    Range("G19:G20").Select
                                    Selection.PasteSpecial Paste:=xlPasteValues
                                        '11
                                        this.Activate
                                        Range("E18:E19").Copy
                                        import.Activate
                                        Range("G23:G24").Select
                                        Selection.PasteSpecial Paste:=xlPasteValues
                                            '12
                                            this.Activate
                                            Range("E21").Copy
                                            import.Activate
                                            Range("G26").Select
                                            Selection.PasteSpecial Paste:=xlPasteValues
                                                '13
                                                this.Activate
                                                Range("G6:G11").Copy
                                                import.Activate
                                                Range("I11:I16").Select
                                                Selection.PasteSpecial Paste:=xlPasteValues
                                                    '14
                                                    this.Activate
                                                    Range("I4:J4").Copy
                                                    import.Activate
                                                    Range("K9:L9").Select
                                                    Selection.PasteSpecial Paste:=xlPasteValues
                                                        '15
                                                        this.Activate
                                                        Range("I17:J17").Copy
                                                        import.Activate
                                                        Range("K22:L22").Select
                                                        Selection.PasteSpecial Paste:=xlPasteValues
                                                            '16
                                                            this.Activate
                                                            Range("I22:I29").Copy
                                                            import.Activate
                                                            Range("K27:K34").Select
                                                            Selection.PasteSpecial Paste:=xlPasteValues
                                                                '17
                                                                this.Activate
                                                                Range("J22:J27").Copy
                                                                import.Activate
                                                                Range("L27:L32").Select
                                                                Selection.PasteSpecial Paste:=xlPasteValues
                                                                    '18
                                                                    this.Activate
                                                                    Range("K23:K24").Copy
                                                                    import.Activate
                                                                    Range("M28:M29").Select
                                                                    Selection.PasteSpecial Paste:=xlPasteValues
                                                                        '19
                                                                        this.Activate
                                                                        Range("I32:J32").Copy
                                                                        import.Activate
                                                                        Range("K37:L37").Select
                                                                        Selection.PasteSpecial Paste:=xlPasteValues
                                                                            '20
                                                                            this.Activate
                                                                            Range("L34").Copy
                                                                            import.Activate
                                                                            Range("N39").Select
                                                                            Selection.PasteSpecial Paste:=xlPasteValues
                                                                                '21
                                                                                this.Activate
                                                                                Range("M14:M15").Copy
                                                                                import.Activate
                                                                                Range("O19:O20").Select
                                                                                Selection.PasteSpecial Paste:=xlPasteValues
                                                                                    '22
                                                                                    this.Activate
                                                                                    Range("M18:M19").Copy
                                                                                    import.Activate
                                                                                    Range("O23:O24").Select
                                                                                    Selection.PasteSpecial Paste:=xlPasteValues
                                                                                        '23
                                                                                        this.Activate
                                                                                        Range("M21").Copy
                                                                                        import.Activate
                                                                                        Range("O26").Select
                                                                                        Selection.PasteSpecial Paste:=xlPasteValues
                                                                                        
                                                                                            '24
                                                                                            this.Activate
                                                                                            Range("O6:O11").Copy
                                                                                            import.Activate
                                                                                            Range("Q11:Q16").Select
                                                                                            Selection.PasteSpecial Paste:=xlPasteValues
                                                                                                '25
                                                                                                this.Activate
                                                                                                Range("Q4:Q32").Copy
                                                                                                import.Activate
                                                                                                Range("S9:S37").Select
                                                                                                Selection.PasteSpecial Paste:=xlPasteValues
                                                                                                    '26
                                                                                                    this.Activate
                                                                                                    Range("Q34").Copy
                                                                                                    import.Activate
                                                                                                    Range("S39").Select
                                                                                                    Selection.PasteSpecial Paste:=xlPasteValues

 import.Activate
 ActiveSheet.Protect Password:="tesla"
 
osv:

i = 31
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo 5804
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        
    '������� ������ � ������
        
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
    
5804:

'����� ������
i = 32
'���������� ��������
k = 3
'�������� ������������� ���������� ��������
     ThisWorkbook.Sheets("Preferences").Activate
     If Range("S" & i + 1).Value = False Then
        GoTo saver
     End If
     
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '������ ����������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    ActiveWindow.Zoom = 100
    this.Activate
    ActiveSheet.Unprotect Password:="gfhjkm"
   
   '��������� ����������
        ThisWorkbook.Sheets("Ranges").Activate
        '1 �������
        Ra1 = "K" & i
        y1 = Range(Ra1).Text
        Ra2 = "L" & i
        z1 = Range(Ra2).Text
        '2 �������
        Ra3 = "Q" & i
        y2 = Range(Ra3).Text
        Ra4 = "R" & i
        z2 = Range(Ra4).Text
        '3 �������
        Ra5 = "W" & i
        y3 = Range(Ra5).Text
        Ra6 = "X" & i
        z3 = Range(Ra6).Text
        '4 �������
        Ra7 = "AC" & i
        y4 = Range(Ra7).Text
        Ra8 = "AD" & i
        z4 = Range(Ra8).Text
        
    '������� ������ � ������
        '������� 1
        On Error Resume Next
        this.Activate
        Range(y1).Copy
        import.Activate
        Range(z1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 2
        On Error Resume Next
        this.Activate
        Range(y2).Copy
        import.Activate
        Range(z2).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 3
        On Error Resume Next
        this.Activate
        Range(y3).Copy
        import.Activate
        Range(z3).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        '������� 4
        On Error Resume Next
        this.Activate
        Range(y4).Copy
        import.Activate
        Range(z4).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    
    '���������� ����� �������
    import.Activate
    ActiveSheet.Protect Password:="tesla"
 
saver:

'������������ ��������
For i = 2 To a
    '����������� ����
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "A" & i
    y = Range(Range1).Text
    Set import = importWB.Sheets(y)
    Set this = ThisWorkbook.Sheets(y)
    
    '�������� ������
    import.Activate
    ActiveSheet.Unprotect Password:="tesla"
    this.Activate
    Range("A2").Copy
    import.Activate
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas
    import.Activate
    Range("A2").Copy
    ThisWorkbook.Sheets("Preferences").Activate
    Range("Q" & i + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    import.Activate
    Range("A2").Select
    With Selection
        .Clear
    End With
    import.Activate
    ActiveSheet.Protect Password:="tesla"
Next i
 
'���������� c ����������� � ������������ �����
    ThisWorkbook.Sheets("Preferences").Activate
    SaveName = ActiveSheet.Range("AC1").Text
    Folder = ActiveSheet.Range("AA1").Text

    Set object = CreateObject("Scripting.FileSystemObject")
    Path = ActiveWorkbook.Path

    If object.FolderExists(Path & "\" & Folder) Then
        Resume Next
    Else
        object.CreateFolder (Path & "\" & Folder)
    End If
     
    importWB.Activate
    importWB.SaveAs Filename:=Path & "\" & Folder & "\" & _
    SaveName & ".xlsm"

    importWB.Save
    importWB.Close

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets("Preferences").Activate
 ActiveWindow.Zoom = 100
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
           
End Sub


