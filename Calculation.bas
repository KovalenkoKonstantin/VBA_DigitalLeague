Attribute VB_Name = "Calculation"
Sub ������()

'===
'ActiveSheet.Protect PASSWORD:="tesla", AllowInsertingRows:=True
Worksheets("���_���").Unprotect Password:="tesla"
'=====
'   Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    ActiveSheet.DisplayPageBreaks = False
'    Application.DisplayStatusBar = False
'    Application.DisplayAlerts = False
    Application.Calculation = xlManual

'Worksheets("���_���").Protect Password:="tesla", DrawingObjects:=True, Contents:=True, Scenarios:=True _
'        , AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering _
'        :=True
'Worksheets("���_���").Unprotect Password:="tesla"


Worksheets("�-�").Unprotect Password:="tesla"
' ������� ������ ������ ��� ���������� ������ ���������
Dim row_t As Integer
For i = 20 To 1000
    If Worksheets("���_���").Cells(i, 2) = "��������� ������ �� ����� ��������� �������:" Then
        row_t = i + 1
        Exit For
    End If
Next

'������� ��������� ������ � ��������� ������ �� ���������
x = 1
Dim ppp As Integer
ppp = 140
While x
    i = row_t
    If Worksheets("���_���").Cells(i + 2, 1) = "" Then
        x = 0
    End If
    If x <> 0 Then
        '-------------------------------------------------------------
            If Left(Worksheets("���_���").Cells(i, 3), 2) = "60" Or Left(Worksheets("���_���").Cells(i, 3), 2) = "62" Or Left(Worksheets("���_���").Cells(i, 3), 2) = "76" And Not Worksheets("���_���").Cells(i, 3) = "76.01.1" Then
                If Worksheets("���_���").Cells(i, 14) <> 0 Or Worksheets("���_���").Cells(i, 15) <> 0 Or Worksheets("���_���").Cells(i, 16) <> 0 Or Worksheets("���_���").Cells(i, 17) <> 0 Then
                   Worksheets("�-�").Cells(ppp, 2) = Worksheets("���_���").Cells(i, 3)
                   Worksheets("�-�").Cells(ppp, 3) = Worksheets("���_���").Cells(i, 14)
                   Worksheets("�-�").Cells(ppp, 4) = Worksheets("���_���").Cells(i, 15)
                   Worksheets("�-�").Cells(ppp, 5) = Worksheets("���_���").Cells(i, 16)
                   Worksheets("�-�").Cells(ppp, 6) = Worksheets("���_���").Cells(i, 17)
                   Worksheets("�-�").Cells(ppp, 7) = Worksheets("���_���").Cells(i, 28)
                   Worksheets("�-�").Cells(ppp, 8) = Worksheets("���_���").Cells(i, 29)
                   Worksheets("�-�").Cells(ppp, 9) = Worksheets("���_���").Cells(i, 30)
                   Worksheets("�-�").Cells(ppp, 10) = Worksheets("���_���").Cells(i, 40)
                   Worksheets("�-�").Cells(ppp, 11) = Worksheets("���_���").Cells(i, 12)
                   Worksheets("�-�").Cells(ppp, 12) = Worksheets("���_���").Cells(i, 19)
                   Worksheets("�-�").Cells(ppp, 13) = Worksheets("���_���").Cells(i, 20)
                   Worksheets("�-�").Cells(ppp, 14) = Worksheets("���_���").Cells(i, 21)
                   Worksheets("�-�").Cells(ppp, 15) = Worksheets("���_���").Cells(i, 22)
                ppp = ppp + 1
                End If
            End If
        '-------------------------------------------------------------
        Worksheets("���_���").Rows(i).delete

    End If
Wend

Dim chet As String
Dim count As Integer
Dim dubl As Boolean
Dim formula  As String


Dim f_x_1, f_x_2, f_x_3, f_x_4 As Integer


For i = Application.Worksheets("���_���").Cells.SpecialCells(xlLastCell).Row - 3 To 15 Step -1
    If Worksheets("���_���").Cells(i, 2) = "��������� ������ �� ����� ��������� �������:" Then
        f_x_4 = i - 3
    End If
    If Worksheets("���_���").Cells(i, 2) = "������� �� ������ ������������� � ������:" Then
        f_x_3 = i + 1
        f_x_2 = i - 3
    End If
Next
f_x_1 = f_x_4 + 4 - 15
f_x_2 = 4

count = 0
formula = ""

For i = 30 To 120
    chet = Worksheets("�-�").Cells(i, 2)
    If chet = "" Then
        Exit For
    End If
    
    For j = 15 To row_t - 2
        If chet = Worksheets("���_���").Cells(j, 3) Then
            '��������� ���� �� ����� ����� ��������
            dubl = False
                        
            For Z = Application.Worksheets("���_���").Cells.SpecialCells(xlLastCell).Row - 3 To row_t Step -1
                If chet = Worksheets("���_���").Cells(Z, 3) Then
                    For x = 12 To 46
                        '28.01.2020 ������ ��� �������
                        'And x <> 33
                        If Worksheets("�-�").Cells(i, x) = "+" And x <> 33 And x <> 13 And x <> 15 Then
                            If Not Worksheets("���_���").Cells(Z, x) = Worksheets("���_���").Cells(j, x) Then
                                Exit For
                            End If
                        End If
                        If x = 46 Then
                            dubl = True
                        End If
                    Next
                End If
            Next
            
           
                
            '++++ 11.02.20 ������� ���
                '���� ����� ��������� ���
            If dubl = False Then
                Worksheets("���_���").Rows(row_t + count).Select
                Selection.Copy
                Selection.Insert Shift:=xlDown
                Application.CutCopyMode = False
                
                Worksheets("���_���").Cells(row_t + count, 3).Value = chet
                formula = "=SUMIFS(R[-" & f_x_1 + count & "]C:R[-" & f_x_2 + count & "]C,R[-" & f_x_1 + count & "]C[-7]:R[-" & f_x_2 + count & "]C[-7],RC[-7])"
            '�������  ���
                Formula2 = "=SUMIFS(R[-" & f_x_1 + count & "]C:R[-" & f_x_2 + count & "]C,R[-" & f_x_1 + count & "]C[-8]:R[-" & f_x_2 + count & "]C[-8],RC[-8])"

            '������� ������ ��������
                formula3 = "=SUMIFS(R[-" & f_x_1 + count & "]C:R[-" & f_x_2 + count & "]C,R[-" & f_x_1 + count & "]C[-6]:R[-" & f_x_2 + count & "]C[-6],RC[-6])"

                For x = 12 To 45
                        '28.01.2020 ������ ��� �������
                        'And x <> 33
                    If Worksheets("�-�").Cells(i, x) = "+" And x <> 33 And x <> 13 And x <> 15 Then
                        Worksheets("���_���").Cells(row_t + count, x) = Worksheets("���_���").Cells(j, x)
                        If Worksheets("���_���").Cells(j, x) = "" Then
                            formula = Left(formula, Len(formula) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 10 & "]:R[-" & f_x_2 + count & "]C[" & x - 10 & "],"""")"
                        '�������  ���
                            Formula2 = Left(Formula2, Len(Formula2) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 11 & "]:R[-" & f_x_2 + count & "]C[" & x - 11 & "],"""")"
'������� ������ ��������
                            formula3 = Left(formula3, Len(formula3) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 9 & "]:R[-" & f_x_2 + count & "]C[" & x - 9 & "],"""")"
                        Else
                            formula = Left(formula, Len(formula) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 10 & "]:R[-" & f_x_2 + count & "]C[" & x - 10 & "],RC[" & x - 10 & "])"
                        '�������  ���
                            Formula2 = Left(Formula2, Len(Formula2) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 11 & "]:R[-" & f_x_2 + count & "]C[" & x - 11 & "],RC[" & x - 11 & "])"
                            '������� ������ ��������
                            formula3 = Left(formula3, Len(formula3) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 9 & "]:R[-" & f_x_2 + count & "]C[" & x - 9 & "],RC[" & x - 9 & "])"
                        End If
                    End If
                Next
        '��������� ����������++++
                Worksheets("���_���").Cells(row_t + count, 10) = "=ROUND(" & Mid(formula, 2, Len(formula)) & ",2)"
                '������� ���
                Worksheets("���_���").Cells(row_t + count, 11) = "=ROUND(" & Mid(Formula2, 2, Len(Formula2)) & ",2)"
'������� ������ ��������
                Worksheets("���_���").Cells(row_t + count, 9) = "=ROUND(" & Mid(formula3, 2, Len(formula3)) & ",2)"
          
          
          '++++++
                count = count + 1
                
                
                '+++
                
                
            End If
            
                
        End If
    Next
Next
        

For i = Application.Worksheets("���_���").Cells.SpecialCells(xlLastCell).Row - 3 To row_t Step -1
    
    If Worksheets("���_���").Cells(i, 10) = 0 And Worksheets("���_���").Cells(i, 3) <> "" Then
        Worksheets("���_���").Rows(i).delete
    End If
Next
    
' ����������� ������ ��� ����������� ��_��
For i = Application.Worksheets("���_���").Cells.SpecialCells(xlLastCell).Row - 3 To row_t Step -1
    If Left(Worksheets("���_���").Cells(i, 3), 2) = "60" Or Left(Worksheets("���_���").Cells(i, 3), 2) = "62" Or Left(Worksheets("���_���").Cells(i, 3), 2) = "76" And Not Worksheets("���_���").Cells(i, 3) = "76.01.1" Then
        Application.Worksheets("���_���").Range(Application.Worksheets("���_���").Cells(i, 14), Application.Worksheets("���_���").Cells(i, 17)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColor = 65535
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Locked = False
        Selection.FormulaHidden = False
        
        Application.Worksheets("���_���").Range(Application.Worksheets("���_���").Cells(i, 13), Application.Worksheets("���_���").Cells(i, 17)).Select
        Selection.HorizontalAlignment = xlRight
        Selection.VerticalAlignment = xlBottom
        Worksheets("���_���").Cells(i, 13) = "=RC[-3]-SUM(RC[1]:RC[4])"

'������� ������ ��������
        Application.Worksheets("���_���").Range(Application.Worksheets("���_���").Cells(i, 19), Application.Worksheets("���_���").Cells(i, 22)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColor = 65535
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Locked = False
        Selection.FormulaHidden = False
        
        Application.Worksheets("���_���").Range(Application.Worksheets("���_���").Cells(i, 18), Application.Worksheets("���_���").Cells(i, 22)).Select
        Selection.HorizontalAlignment = xlRight
        Selection.VerticalAlignment = xlBottom
        Worksheets("���_���").Cells(i, 18) = "=RC[-9]-SUM(RC[1]:RC[4])"

        Application.Worksheets("���_���").Range(Application.Worksheets("���_���").Cells(i, 13), Application.Worksheets("���_���").Cells(i, 9)).Select
        Selection.HorizontalAlignment = xlRight
        Selection.VerticalAlignment = xlBottom
        
            '-------------------------------------------------------------
    '��������� ��� �� ���������
        For y = 140 To ppp
            If Worksheets("���_���").Cells(i, 3) = Worksheets("�-�").Cells(y, 2) And Worksheets("���_���").Cells(i, 28) = Worksheets("�-�").Cells(y, 7) _
            And Worksheets("���_���").Cells(i, 29) = Worksheets("�-�").Cells(y, 8) And Worksheets("���_���").Cells(i, 30) = Worksheets("�-�").Cells(y, 9) _
            And Worksheets("���_���").Cells(i, 40) = Worksheets("�-�").Cells(y, 10) And Worksheets("���_���").Cells(i, 12) = Worksheets("�-�").Cells(y, 11) Then
                    Worksheets("���_���").Cells(i, 14) = Worksheets("�-�").Cells(y, 3)
                    Worksheets("���_���").Cells(i, 15) = Worksheets("�-�").Cells(y, 4)
                    Worksheets("���_���").Cells(i, 16) = Worksheets("�-�").Cells(y, 5)
                    Worksheets("���_���").Cells(i, 17) = Worksheets("�-�").Cells(y, 6)
                    Worksheets("���_���").Cells(i, 19) = Worksheets("�-�").Cells(y, 12)
                    Worksheets("���_���").Cells(i, 20) = Worksheets("�-�").Cells(y, 13)
                    Worksheets("���_���").Cells(i, 21) = Worksheets("�-�").Cells(y, 14)
                    Worksheets("���_���").Cells(i, 22) = Worksheets("�-�").Cells(y, 15)
                    
                Exit For
            End If
            
        Next
    '-------------------------------------------------------------
        
    End If
    
Next
    
    
'������� �-�
For y = ppp To 140 Step -1
    Worksheets("�-�").Rows(y).delete
Next
    
    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'    ActiveSheet.DisplayPageBreaks = True
'    Application.DisplayStatusBar = True
'    Application.DisplayAlerts = True
    '==
    ActiveSheet.Protect Password:="tesla", DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering _
        :=True
    Worksheets("�-�").Protect Password:="tesla"
    Application.Worksheets("���_���").Range(Application.Worksheets("���_���").Cells(1, 1), Application.Worksheets("���_���").Cells(1, 1)).Select
    
End Sub


