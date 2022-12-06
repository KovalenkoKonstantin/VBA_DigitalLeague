Attribute VB_Name = "Transport"
Sub ПравильноВставить()

 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim importWB  As Workbook
 Dim SaveName As String
 Dim Folder As String
 Dim Path As String
 Dim import As Worksheet
 Dim this As Worksheet
 Dim MyRange As Range
 Dim MyCell As Range
 Dim flag As Integer
 Dim va As Integer
 Dim Response
 Dim Response1
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ErrHandler
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
 
 'интерактивочка
 ThisWorkbook.Activate
 СнятьЗащитуВсехЛистов
 If ThisWorkbook.Sheets("Preferences").Range("U1").value = False Then
            Style = vbYesNo
            Title = "ГИД"
            Response = MsgBox("Перед продолжением необходимо загрузить справочник ГИД. Загрузим справочник?", Style, Title)
                If Response = vbYes Then
                    инн
                    СнятьЗащитуВсехЛистов
                Else
                    MsgBox "Действие отменено."
                    GoTo ExitHandler
                End If
 End If

 If ThisWorkbook.Sheets("Preferences").Range("V1").value = False Then
            Style = vbYesNo
            Title = "ИНН"
            Response1 = MsgBox("Перед продолжением необходимо исправить некорректные ИНН. Будем исправлять?", Style, Title)
                If Response1 = vbNo Then
                    MsgBox "В шаблон будут добавлены заведомо некорректные ИНН."
                    MsgBox "Теперь надо выбрать шаблон куда будем вставлять данные."
                    Application.ScreenUpdating = False
                    Application.EnableEvents = False
                    ActiveSheet.DisplayPageBreaks = False
                    Application.DisplayStatusBar = False
                    Application.DisplayAlerts = False
                    GoTo proseed
                Else
                    MsgBox "Настало время Епифанцева."
                    рефактор_инн
                    Application.ScreenUpdating = False
                    Application.EnableEvents = False
                    ActiveSheet.DisplayPageBreaks = False
                    Application.DisplayStatusBar = False
                    Application.DisplayAlerts = False
                    СнятьЗащитуВсехЛистов
                    Application.ScreenUpdating = False
                    Application.EnableEvents = False
                    ActiveSheet.DisplayPageBreaks = False
                    Application.DisplayStatusBar = False
                    Application.DisplayAlerts = False
                    MsgBox "Теперь надо выбрать шаблон куда будем вставлять данные."
                End If
 End If
 
proseed:
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsm), *.xlsm", _
 MultiSelect:=True, Title:="Файл для вставки")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "Файл не выбран!"
 GoTo ExitHandler
 End If
 
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
  
 'содержание
 Set import = importWB.Sheets("Содержание")
 Set this = ThisWorkbook.Sheets("Preferences")
 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 
 'диапазоны содержания
import.Activate
Range("D2:E2").Select
Range("D2:E2").UnMerge

import.Activate
Range("D5:E12").Select
Range("D5:E12").UnMerge

import.Activate
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
 
 'счёт 90
 Set import = importWB.Sheets("90")
 Set this = ThisWorkbook.Sheets("90")
 ThisWorkbook.Sheets("Preferences").Activate
 va = ActiveSheet.Range("W4").value

import.Activate
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
'Application.DisplayAlerts = False

On Error GoTo Err1
    ActiveSheet.Protect Password:="tesla", AllowInsertingRows:=True
    ActiveSheet.Unprotect Password:="tesla"
    flag = va
        counter = 1
        Do While counter <= flag
        counter = counter + 1
        lastrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
        Rows(lastrow).AutoFill Rows(lastrow).Resize(2), xlFillDefault
        On Error Resume Next
        Rows(lastrow + 1).SpecialCells(xlConstants).ClearContents
        On Error GoTo 0
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
 
 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 'диапазоны счёта 90
 '1
 this.Activate
 Range("N8:V13").Copy
 import.Activate
 Range("C11:K16").Select
 Selection.PasteSpecial Paste:=xlPasteValues
    Range("C11:K16").Select
    Set MyRange = Selection
    For Each MyCell In MyRange
    If MyCell.value = 0 Then
    MyCell.value = Empty
    End If
    Next MyCell
 import.Activate
 ActiveSheet.Protect Password:="tesla"
 
 'счёт 90_ОРВГ
 Set import = importWB.Sheets("90_ОРВГ")
 Set this = ThisWorkbook.Sheets("90_ОРВГ")
 ThisWorkbook.Sheets("Preferences").Activate
 va = ActiveSheet.Range("X4").value

import.Activate
On Error GoTo Err1
    ActiveSheet.Protect Password:="tesla", AllowInsertingRows:=True
    ActiveSheet.Unprotect Password:="tesla"
    flag = va
        counter = 1
        Do While counter <= flag
        counter = counter + 1
        lastrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
        Rows(lastrow).AutoFill Rows(lastrow).Resize(2), xlFillDefault
        On Error Resume Next
        Rows(lastrow + 1).SpecialCells(xlConstants).ClearContents
        On Error GoTo 0
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
 
 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 'диапазоны счёта 90_ОРВГ
 '1
 this.Activate
 Range("Z6:Z56").Copy
 import.Activate
 Range("C10:C60").Select
 Selection.PasteSpecial Paste:=xlPasteValues
            Range("C10:C60").Select
            Set MyRange = Selection
            For Each MyCell In MyRange
            If MyCell.value = 0 Then
            MyCell.value = Empty
            End If
            Next MyCell
    '2
    this.Activate
    Range("AA6:AH56").Copy
    import.Activate
    Range("F10:M60").Select
    Selection.PasteSpecial Paste:=xlPasteValues
            Range("F10:M60").Select
            Set MyRange = Selection
            For Each MyCell In MyRange
            If MyCell.value = 0 Then
            MyCell.value = Empty
            End If
            Next MyCell
        '3
        this.Activate
        Range("Y6:Y56").Copy
        import.Activate
        Range("S10:S60").Select
        Selection.PasteSpecial Paste:=xlPasteValues
 import.Activate
 ActiveSheet.Protect Password:="tesla"
 
 'счёт 90_контр
 Set import = importWB.Sheets("90_контр")
 Set this = ThisWorkbook.Sheets("90 контр")
 ThisWorkbook.Sheets("Preferences").Activate
 va = ActiveSheet.Range("Y4").value

 import.Activate
On Error GoTo Err1
    ActiveSheet.Protect Password:="tesla", AllowInsertingRows:=True
    ActiveSheet.Unprotect Password:="tesla"
    flag = va
        counter = 1
        Do While counter <= flag
        counter = counter + 1
        lastrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
        Rows(lastrow).AutoFill Rows(lastrow).Resize(2), xlFillDefault
        On Error Resume Next
        Rows(lastrow + 1).SpecialCells(xlConstants).ClearContents
        On Error GoTo 0
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
 
 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 'диапазоны счёта 90_контр
 '1
 this.Activate
 Range("A9:I200").Copy
 import.Activate
 Range("C12:K203").Select
 Selection.PasteSpecial Paste:=xlPasteValues
            Range("C12:K203").Select
            Set MyRange = Selection
            For Each MyCell In MyRange
            If MyCell.value = 0 Then
            MyCell.value = Empty
            End If
            Next MyCell
 import.Activate
 ActiveSheet.Protect Password:="tesla"
 
 'счёт Зарубежная_выручка
 Set import = importWB.Sheets("Зарубежная_выручка")
 Set this = ThisWorkbook.Sheets("Зарубежная выручка")
 ThisWorkbook.Sheets("Preferences").Activate
 va = ActiveSheet.Range("Z4").value

 import.Activate
 On Error GoTo Err1
    ActiveSheet.Protect Password:="tesla", AllowInsertingRows:=True
    ActiveSheet.Unprotect Password:="tesla"
    flag = va
        counter = 1
        Do While counter <= flag
        counter = counter + 1
        lastrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
        Rows(lastrow).AutoFill Rows(lastrow).Resize(2), xlFillDefault
        On Error Resume Next
        Rows(lastrow + 1).SpecialCells(xlConstants).ClearContents
        On Error GoTo 0
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
 
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 'диапазоны счёта Зарубежная_выручка
 '1
 this.Activate
 Range("A4:N24").Copy
 import.Activate
 Range("C12:P32").Select
 Selection.PasteSpecial Paste:=xlPasteValues
            Range("C12:P32").Select
            Set MyRange = Selection
            For Each MyCell In MyRange
            If MyCell.value = 0 Then
            MyCell.value = Empty
            End If
            Next MyCell
 import.Activate
 ActiveSheet.Protect Password:="tesla"
 
    'проверка ГИД
    ThisWorkbook.Sheets("Preferences").Activate
    If ThisWorkbook.Sheets("Preferences").Range("T1").value = True Then
        importWB.Sheets("Проверка_ГИД").Activate
        ActiveSheet.Unprotect Password:="tesla"
'        importWB.Sheets("Проверка_ГИД").Select
        Range("E16").value = True
        ActiveSheet.Protect Password:="tesla"
    End If
    
    'проверка правильности переноса данных
'    importWB.Sheets("90_ОРВГ").Activate
'    ActiveSheet.Unprotect Password:="tesla"
'    importWB.Sheets("90_ОРВГ").Activate
'    Range("W10:X510").Copy
'    importWB.Sheets("90_ОРВГ").Range("W10:X510").Copy
'    ThisWorkbook.Sheets("Preferences").Range("W5:X505").Select
'    Selection.PasteSpecial Paste:=xlPasteValues
'    importWB.Sheets("90_ОРВГ").Activate
'    ActiveSheet.Protect Password:="tesla"

'    importWB.Sheets("90_контр").Activate
'    ActiveSheet.Unprotect Password:="tesla"
'    Range("P12:U512").Copy
'    ThisWorkbook.Sheets("Preferences").Range("Y11:AD511").Select
'    Selection.PasteSpecial Paste:=xlPasteValues
'    importWB.Sheets("90_контр").Activate
'    Range("Y12:Y512").Copy
'    ThisWorkbook.Sheets("Preferences").Range("AE11:AE511").Select
'    Selection.PasteSpecial Paste:=xlPasteValues
'    importWB.Sheets("90_контр").Activate
'    ActiveSheet.Protect Password:="tesla"
'
'    importWB.Sheets("Зарубежная_выручка").Activate
'    ActiveSheet.Unprotect Password:="tesla"
'    Range("R12:W512").Copy
'    ThisWorkbook.Sheets("Preferences").Range("AF20:AK520").Select
'    Selection.PasteSpecial Paste:=xlPasteValues
'    importWB.Sheets("Зарубежная_выручка").Activate
'    ActiveSheet.Protect Password:="tesla"
  
 importWB.Sheets("Содержание").Activate
 Range("E17").Select
 
    If Range("E17") = True Then
    ThisWorkbook.Sheets("Preferences").Activate
    SaveName = ActiveSheet.Range("M2").Text
    Folder = ActiveSheet.Range("L9").Text
    Application.DisplayAlerts = False

    create
    
    ThisWorkbook.Activate
    Path = ActiveWorkbook.Path
     
    importWB.Activate
    importWB.SaveAs Filename:=Path & "\" & Folder & "\" & _
    SaveName & ".xlsm"
'    importWB.SaveAs Filename:="c:\Users\kkovalenko\Desktop\" & _
'    SaveName & ".xlsm"
'    importWB.SaveAs Filename:="\\bs.phoenixit.ru\ФинРепорт_Росатом\РА\" & _
'    SaveName & ".xlsm"
    
'    Dim Style, Title
'    Style = vbExclamation = 48
'    Title = "Ура!"
'    MsgBox "ОК", Style, Title

    importWB.Save
    importWB.Close
 
    Else
    Dim Style1, Title1
    Style1 = vbCritical = 16
    Title1 = "Блин!"
    MsgBox "Были обнаружены ошибки. Файл не сохранён", Style1, Title1
    End If

    ЗаблокироватьВсеЛисты

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets("Preferences").Activate
 ЗаблокироватьВсеЛисты
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
 
Err1:
    ActiveSheet.Protect Password:="tesla", AllowInsertingRows:=False
           
End Sub

