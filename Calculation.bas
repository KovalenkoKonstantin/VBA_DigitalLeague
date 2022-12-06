Attribute VB_Name = "Calculation"
Sub Сальдо()

'===
'ActiveSheet.Protect PASSWORD:="tesla", AllowInsertingRows:=True
Worksheets("ПСД_ВГО").Unprotect Password:="tesla"
'=====
'   Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    ActiveSheet.DisplayPageBreaks = False
'    Application.DisplayStatusBar = False
'    Application.DisplayAlerts = False
    Application.Calculation = xlManual

'Worksheets("ПСД_ВГО").Protect Password:="tesla", DrawingObjects:=True, Contents:=True, Scenarios:=True _
'        , AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering _
'        :=True
'Worksheets("ПСД_ВГО").Unprotect Password:="tesla"


Worksheets("С-А").Unprotect Password:="tesla"
' находим первую строку для заполнения сальдо конечного
Dim row_t As Integer
For i = 20 To 1000
    If Worksheets("ПСД_ВГО").Cells(i, 2) = "Исходящее сальдо на конец отчетного периода:" Then
        row_t = i + 1
        Exit For
    End If
Next

'очищаем исходящее сальдо и формируем данные по просрочке
x = 1
Dim ppp As Integer
ppp = 140
While x
    i = row_t
    If Worksheets("ПСД_ВГО").Cells(i + 2, 1) = "" Then
        x = 0
    End If
    If x <> 0 Then
        '-------------------------------------------------------------
            If Left(Worksheets("ПСД_ВГО").Cells(i, 3), 2) = "60" Or Left(Worksheets("ПСД_ВГО").Cells(i, 3), 2) = "62" Or Left(Worksheets("ПСД_ВГО").Cells(i, 3), 2) = "76" And Not Worksheets("ПСД_ВГО").Cells(i, 3) = "76.01.1" Then
                If Worksheets("ПСД_ВГО").Cells(i, 14) <> 0 Or Worksheets("ПСД_ВГО").Cells(i, 15) <> 0 Or Worksheets("ПСД_ВГО").Cells(i, 16) <> 0 Or Worksheets("ПСД_ВГО").Cells(i, 17) <> 0 Then
                   Worksheets("С-А").Cells(ppp, 2) = Worksheets("ПСД_ВГО").Cells(i, 3)
                   Worksheets("С-А").Cells(ppp, 3) = Worksheets("ПСД_ВГО").Cells(i, 14)
                   Worksheets("С-А").Cells(ppp, 4) = Worksheets("ПСД_ВГО").Cells(i, 15)
                   Worksheets("С-А").Cells(ppp, 5) = Worksheets("ПСД_ВГО").Cells(i, 16)
                   Worksheets("С-А").Cells(ppp, 6) = Worksheets("ПСД_ВГО").Cells(i, 17)
                   Worksheets("С-А").Cells(ppp, 7) = Worksheets("ПСД_ВГО").Cells(i, 28)
                   Worksheets("С-А").Cells(ppp, 8) = Worksheets("ПСД_ВГО").Cells(i, 29)
                   Worksheets("С-А").Cells(ppp, 9) = Worksheets("ПСД_ВГО").Cells(i, 30)
                   Worksheets("С-А").Cells(ppp, 10) = Worksheets("ПСД_ВГО").Cells(i, 40)
                   Worksheets("С-А").Cells(ppp, 11) = Worksheets("ПСД_ВГО").Cells(i, 12)
                   Worksheets("С-А").Cells(ppp, 12) = Worksheets("ПСД_ВГО").Cells(i, 19)
                   Worksheets("С-А").Cells(ppp, 13) = Worksheets("ПСД_ВГО").Cells(i, 20)
                   Worksheets("С-А").Cells(ppp, 14) = Worksheets("ПСД_ВГО").Cells(i, 21)
                   Worksheets("С-А").Cells(ppp, 15) = Worksheets("ПСД_ВГО").Cells(i, 22)
                ppp = ppp + 1
                End If
            End If
        '-------------------------------------------------------------
        Worksheets("ПСД_ВГО").Rows(i).delete

    End If
Wend

Dim chet As String
Dim count As Integer
Dim dubl As Boolean
Dim formula  As String


Dim f_x_1, f_x_2, f_x_3, f_x_4 As Integer


For i = Application.Worksheets("ПСД_ВГО").Cells.SpecialCells(xlLastCell).Row - 3 To 15 Step -1
    If Worksheets("ПСД_ВГО").Cells(i, 2) = "Исходящее сальдо на конец отчетного периода:" Then
        f_x_4 = i - 3
    End If
    If Worksheets("ПСД_ВГО").Cells(i, 2) = "Обороты по счетам задолженности и оплата:" Then
        f_x_3 = i + 1
        f_x_2 = i - 3
    End If
Next
f_x_1 = f_x_4 + 4 - 15
f_x_2 = 4

count = 0
formula = ""

For i = 30 To 120
    chet = Worksheets("С-А").Cells(i, 2)
    If chet = "" Then
        Exit For
    End If
    
    For j = 15 To row_t - 2
        If chet = Worksheets("ПСД_ВГО").Cells(j, 3) Then
            'проверяем есть ли такой набор аналитик
            dubl = False
                        
            For Z = Application.Worksheets("ПСД_ВГО").Cells.SpecialCells(xlLastCell).Row - 3 To row_t Step -1
                If chet = Worksheets("ПСД_ВГО").Cells(Z, 3) Then
                    For x = 12 To 46
                        '28.01.2020 убрано доп условие
                        'And x <> 33
                        If Worksheets("С-А").Cells(i, x) = "+" And x <> 33 And x <> 13 And x <> 15 Then
                            If Not Worksheets("ПСД_ВГО").Cells(Z, x) = Worksheets("ПСД_ВГО").Cells(j, x) Then
                                Exit For
                            End If
                        End If
                        If x = 46 Then
                            dubl = True
                        End If
                    Next
                End If
            Next
            
           
                
            '++++ 11.02.20 колонка НДС
                'если такой аналитики нет
            If dubl = False Then
                Worksheets("ПСД_ВГО").Rows(row_t + count).Select
                Selection.Copy
                Selection.Insert Shift:=xlDown
                Application.CutCopyMode = False
                
                Worksheets("ПСД_ВГО").Cells(row_t + count, 3).Value = chet
                formula = "=SUMIFS(R[-" & f_x_1 + count & "]C:R[-" & f_x_2 + count & "]C,R[-" & f_x_1 + count & "]C[-7]:R[-" & f_x_2 + count & "]C[-7],RC[-7])"
            'колонка  НДС
                Formula2 = "=SUMIFS(R[-" & f_x_1 + count & "]C:R[-" & f_x_2 + count & "]C,R[-" & f_x_1 + count & "]C[-8]:R[-" & f_x_2 + count & "]C[-8],RC[-8])"

            'колонка валюта договора
                formula3 = "=SUMIFS(R[-" & f_x_1 + count & "]C:R[-" & f_x_2 + count & "]C,R[-" & f_x_1 + count & "]C[-6]:R[-" & f_x_2 + count & "]C[-6],RC[-6])"

                For x = 12 To 45
                        '28.01.2020 убрано доп условие
                        'And x <> 33
                    If Worksheets("С-А").Cells(i, x) = "+" And x <> 33 And x <> 13 And x <> 15 Then
                        Worksheets("ПСД_ВГО").Cells(row_t + count, x) = Worksheets("ПСД_ВГО").Cells(j, x)
                        If Worksheets("ПСД_ВГО").Cells(j, x) = "" Then
                            formula = Left(formula, Len(formula) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 10 & "]:R[-" & f_x_2 + count & "]C[" & x - 10 & "],"""")"
                        'колонка  НДС
                            Formula2 = Left(Formula2, Len(Formula2) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 11 & "]:R[-" & f_x_2 + count & "]C[" & x - 11 & "],"""")"
'колонка валюта договора
                            formula3 = Left(formula3, Len(formula3) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 9 & "]:R[-" & f_x_2 + count & "]C[" & x - 9 & "],"""")"
                        Else
                            formula = Left(formula, Len(formula) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 10 & "]:R[-" & f_x_2 + count & "]C[" & x - 10 & "],RC[" & x - 10 & "])"
                        'колонка  НДС
                            Formula2 = Left(Formula2, Len(Formula2) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 11 & "]:R[-" & f_x_2 + count & "]C[" & x - 11 & "],RC[" & x - 11 & "])"
                            'колонка валюта договора
                            formula3 = Left(formula3, Len(formula3) - 1) & ",R[-" & f_x_1 + count & "]C[" & x - 9 & "]:R[-" & f_x_2 + count & "]C[" & x - 9 & "],RC[" & x - 9 & "])"
                        End If
                    End If
                Next
        'добавлено округление++++
                Worksheets("ПСД_ВГО").Cells(row_t + count, 10) = "=ROUND(" & Mid(formula, 2, Len(formula)) & ",2)"
                'колонка НДС
                Worksheets("ПСД_ВГО").Cells(row_t + count, 11) = "=ROUND(" & Mid(Formula2, 2, Len(Formula2)) & ",2)"
'колонка валюта договора
                Worksheets("ПСД_ВГО").Cells(row_t + count, 9) = "=ROUND(" & Mid(formula3, 2, Len(formula3)) & ",2)"
          
          
          '++++++
                count = count + 1
                
                
                '+++
                
                
            End If
            
                
        End If
    Next
Next
        

For i = Application.Worksheets("ПСД_ВГО").Cells.SpecialCells(xlLastCell).Row - 3 To row_t Step -1
    
    If Worksheets("ПСД_ВГО").Cells(i, 10) = 0 And Worksheets("ПСД_ВГО").Cells(i, 3) <> "" Then
        Worksheets("ПСД_ВГО").Rows(i).delete
    End If
Next
    
' форматируем ячейки под расшифровку ДЗ_КЗ
For i = Application.Worksheets("ПСД_ВГО").Cells.SpecialCells(xlLastCell).Row - 3 To row_t Step -1
    If Left(Worksheets("ПСД_ВГО").Cells(i, 3), 2) = "60" Or Left(Worksheets("ПСД_ВГО").Cells(i, 3), 2) = "62" Or Left(Worksheets("ПСД_ВГО").Cells(i, 3), 2) = "76" And Not Worksheets("ПСД_ВГО").Cells(i, 3) = "76.01.1" Then
        Application.Worksheets("ПСД_ВГО").Range(Application.Worksheets("ПСД_ВГО").Cells(i, 14), Application.Worksheets("ПСД_ВГО").Cells(i, 17)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColor = 65535
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Locked = False
        Selection.FormulaHidden = False
        
        Application.Worksheets("ПСД_ВГО").Range(Application.Worksheets("ПСД_ВГО").Cells(i, 13), Application.Worksheets("ПСД_ВГО").Cells(i, 17)).Select
        Selection.HorizontalAlignment = xlRight
        Selection.VerticalAlignment = xlBottom
        Worksheets("ПСД_ВГО").Cells(i, 13) = "=RC[-3]-SUM(RC[1]:RC[4])"

'колонка валюта договора
        Application.Worksheets("ПСД_ВГО").Range(Application.Worksheets("ПСД_ВГО").Cells(i, 19), Application.Worksheets("ПСД_ВГО").Cells(i, 22)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColor = 65535
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Locked = False
        Selection.FormulaHidden = False
        
        Application.Worksheets("ПСД_ВГО").Range(Application.Worksheets("ПСД_ВГО").Cells(i, 18), Application.Worksheets("ПСД_ВГО").Cells(i, 22)).Select
        Selection.HorizontalAlignment = xlRight
        Selection.VerticalAlignment = xlBottom
        Worksheets("ПСД_ВГО").Cells(i, 18) = "=RC[-9]-SUM(RC[1]:RC[4])"

        Application.Worksheets("ПСД_ВГО").Range(Application.Worksheets("ПСД_ВГО").Cells(i, 13), Application.Worksheets("ПСД_ВГО").Cells(i, 9)).Select
        Selection.HorizontalAlignment = xlRight
        Selection.VerticalAlignment = xlBottom
        
            '-------------------------------------------------------------
    'проверяем был ли заполнены
        For y = 140 To ppp
            If Worksheets("ПСД_ВГО").Cells(i, 3) = Worksheets("С-А").Cells(y, 2) And Worksheets("ПСД_ВГО").Cells(i, 28) = Worksheets("С-А").Cells(y, 7) _
            And Worksheets("ПСД_ВГО").Cells(i, 29) = Worksheets("С-А").Cells(y, 8) And Worksheets("ПСД_ВГО").Cells(i, 30) = Worksheets("С-А").Cells(y, 9) _
            And Worksheets("ПСД_ВГО").Cells(i, 40) = Worksheets("С-А").Cells(y, 10) And Worksheets("ПСД_ВГО").Cells(i, 12) = Worksheets("С-А").Cells(y, 11) Then
                    Worksheets("ПСД_ВГО").Cells(i, 14) = Worksheets("С-А").Cells(y, 3)
                    Worksheets("ПСД_ВГО").Cells(i, 15) = Worksheets("С-А").Cells(y, 4)
                    Worksheets("ПСД_ВГО").Cells(i, 16) = Worksheets("С-А").Cells(y, 5)
                    Worksheets("ПСД_ВГО").Cells(i, 17) = Worksheets("С-А").Cells(y, 6)
                    Worksheets("ПСД_ВГО").Cells(i, 19) = Worksheets("С-А").Cells(y, 12)
                    Worksheets("ПСД_ВГО").Cells(i, 20) = Worksheets("С-А").Cells(y, 13)
                    Worksheets("ПСД_ВГО").Cells(i, 21) = Worksheets("С-А").Cells(y, 14)
                    Worksheets("ПСД_ВГО").Cells(i, 22) = Worksheets("С-А").Cells(y, 15)
                    
                Exit For
            End If
            
        Next
    '-------------------------------------------------------------
        
    End If
    
Next
    
    
'очищаем С-А
For y = ppp To 140 Step -1
    Worksheets("С-А").Rows(y).delete
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
    Worksheets("С-А").Protect Password:="tesla"
    Application.Worksheets("ПСД_ВГО").Range(Application.Worksheets("ПСД_ВГО").Cells(1, 1), Application.Worksheets("ПСД_ВГО").Cells(1, 1)).Select
    
End Sub


