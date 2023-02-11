Attribute VB_Name = "Transfer"
Sub ПравильноВставить()

 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim import, this As Worksheet
 Dim MyRange, MyCell As Range
 Dim SaveName, Range1, Range2, Range3, Range4, Range5, Range6, y, Ra1, Ra2, Folder, Path, Slash, j As String
 Dim x As Integer 'строки для добавления
 Dim object As Object
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ErrHandler
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsm), *.xlsm", _
 MultiSelect:=True, Title:="Файл для вставки")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "Файл не выбран!"
 GoTo ExitHandler
 End If
 
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
 
 'добавление строк
 On Error GoTo ExitHandler

For i = 2 To 8

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
        LastRow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
        Rows(LastRow).AutoFill Rows(LastRow).Resize(2), xlFillDefault
        On Error Resume Next
        Rows(LastRow + 1).SpecialCells(xlConstants).ClearContents
        On Error GoTo ExitHandler
        Rows(LastRow + 1).SpecialCells(xlCellTypeBlanks).Item(1).Activate
        Rows(LastRow - 1).Select
        Selection.Copy
        Rows(LastRow).Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        Loop
    ActiveSheet.Protect Password:="tesla", DrawingObjects:=True, Contents:=True, Scenarios:=True _
       , AllowFormattingCells:=False, AllowFormattingColumns:=True, _
         AllowFormattingRows:=True, AllowInsertingColumns:=False, AllowInsertingRows _
       :=False, AllowSorting:=False, AllowFiltering:=True, AllowUsingPivotTables _
       :=False
'        Loop
Next i
  
 'перенос содержимого в шаблон
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
 
 'счёт 60_01
 Set import = importWB.Sheets("60_01")
 Set this = ThisWorkbook.Sheets("60.01")

 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 'диапазоны счёта 60_01
 '1
 this.Activate
 Range("A10:B16").Copy
 import.Activate
 Range("C11:D17").Select
 Selection.PasteSpecial Paste:=xlPasteValues
    '2
    this.Activate
    Range("C9:C16").Copy
    import.Activate
    Range("E10:E17").Select
    Selection.PasteSpecial Paste:=xlPasteValues
        '3
        this.Activate
        Range("D9:D16").Copy
        import.Activate
        Range("H10:H17").Select
        Selection.PasteSpecial Paste:=xlPasteValues
            '4
            this.Activate
            Range("E9:J16").Copy
            import.Activate
            Range("K10:P17").Select
            Selection.PasteSpecial Paste:=xlPasteValues
' import.Activate
' ActiveSheet.Protect Password:="tesla"
 
 'счёт 60_02
 Set import = importWB.Sheets("60_02")
 Set this = ThisWorkbook.Sheets("60.02")

 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 'диапазоны счёта 60_02
 '1
 this.Activate
 Range("A14:C21").Copy
 import.Activate
 Range("C15:E22").Select
 Selection.PasteSpecial Paste:=xlPasteValues
    '2
    this.Activate
    Range("D9:D21").Copy
    import.Activate
    Range("F10:F22").Select
    Selection.PasteSpecial Paste:=xlPasteValues
        '3
        this.Activate
        Range("E9:G21").Copy
        import.Activate
        Range("I10:K22").Select
        Selection.PasteSpecial Paste:=xlPasteValues
' import.Activate
' ActiveSheet.Protect Password:="tesla"
 
 'счёт 62_01
 Set import = importWB.Sheets("62_01")
 Set this = ThisWorkbook.Sheets("62.01")

 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 'диапазоны счёта 62_01
 '1
 this.Activate
 Range("A8:B15").Copy
 import.Activate
 Range("C11:D18").Select
 Selection.PasteSpecial Paste:=xlPasteValues
    '2
    this.Activate
    Range("C7:C15").Copy
    import.Activate
    Range("E10:E18").Select
    Selection.PasteSpecial Paste:=xlPasteValues
        '3
        this.Activate
        Range("F7:M15").Copy
        import.Activate
        Range("H10:O18").Select
        Selection.PasteSpecial Paste:=xlPasteValues
' import.Activate
' ActiveSheet.Protect Password:="tesla"
 
 'счёт 62_02
 Set import = importWB.Sheets("62_02")
 Set this = ThisWorkbook.Sheets("62.02")

 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 'диапазоны счёта 62_02
 '1
 this.Activate
 Range("A8:D12").Copy
 import.Activate
 Range("C11:F15").Select
 Selection.PasteSpecial Paste:=xlPasteValues
    '2
    this.Activate
    Range("F7:G12").Copy
    import.Activate
    Range("H10:J15").Select
    Selection.PasteSpecial Paste:=xlPasteValues
' import.Activate
' ActiveSheet.Protect Password:="tesla"
'
' 'счёт 60,62,63,76_контр
' Set import = importWB.Sheets("60,62,63,76_контр")
' Set this = ThisWorkbook.Sheets("60-76")
'
' import.Activate
' ActiveSheet.Unprotect Password:="tesla"
' this.Activate
' ActiveSheet.Unprotect Password:="gfhjkm"
' 'диапазоны счёта 60,62,63,76_контр
' '1
' this.Activate
' Range("N8:N10").Copy
' import.Activate
' Range("P11:P13").Select
' Selection.PasteSpecial Paste:=xlPasteValues
'    '2
'    this.Activate
'    Range("T8:T10").Copy
'    import.Activate
'    Range("V11:V13").Select
'    Selection.PasteSpecial Paste:=xlPasteValues
'        '3
'        this.Activate
'        Range("AA11").Copy
'        import.Activate
'        Range("AC14").Select
'        Selection.PasteSpecial Paste:=xlPasteValues
'            '4
'            this.Activate
'            Range("BT12:CG480").Copy
'            import.Activate
'            Range("C15:P483").Select
'            Selection.PasteSpecial Paste:=xlPasteValues
'                '5
'                this.Activate
'                Range("CM12:CM480").Copy
'                import.Activate
'                Range("V15:V483").Select
'                Selection.PasteSpecial Paste:=xlPasteValues
'                    '6
'                    this.Activate
'                    Range("CT12:CT480").Copy
'                    import.Activate
'                    Range("AC15:AC483").Select
'                    Selection.PasteSpecial Paste:=xlPasteValues
'                        '7
'                        this.Activate
'                        Range("DA12:DB480").Copy
'                        import.Activate
'                        Range("AJ15:AK483").Select
'                        Selection.PasteSpecial Paste:=xlPasteValues
' import.Activate
' ActiveSheet.Protect Password:="tesla"
'
' 'счёт 68
' Set import = importWB.Sheets("68")
' Set this = ThisWorkbook.Sheets("68")
'
' import.Activate
' ActiveSheet.Unprotect Password:="tesla"
' this.Activate
' ActiveSheet.Unprotect Password:="gfhjkm"
' 'диапазоны счёта 68
' '1
' this.Activate
' Range("A6:H20").Copy
' import.Activate
' Range("C10:J24").Select
' Selection.PasteSpecial Paste:=xlPasteValues
'    '2
'    this.Activate
'    Range("K6:K20").Copy
'    import.Activate
'    Range("M10:M24").Select
'    Selection.PasteSpecial Paste:=xlPasteValues
'        '3
'        this.Activate
'        Range("N6:W20").Copy
'        import.Activate
'        Range("P10:Y24").Select
'        Selection.PasteSpecial Paste:=xlPasteValues
' import.Activate
' ActiveSheet.Protect Password:="tesla"
'
' 'счёт 68
' Set import = importWB.Sheets("68")
' Set this = ThisWorkbook.Sheets("68")
'
' import.Activate
' ActiveSheet.Unprotect Password:="tesla"
' this.Activate
' ActiveSheet.Unprotect Password:="gfhjkm"
' 'диапазоны счёта 68
' '1
' this.Activate
' Range("A6:H20").Copy
' import.Activate
' Range("C10:J24").Select
' Selection.PasteSpecial Paste:=xlPasteValues
'    '2
'    this.Activate
'    Range("K6:K20").Copy
'    import.Activate
'    Range("M10:M24").Select
'    Selection.PasteSpecial Paste:=xlPasteValues
'        '3
'        this.Activate
'        Range("N6:W20").Copy
'        import.Activate
'        Range("P10:Y24").Select
'        Selection.PasteSpecial Paste:=xlPasteValues
' import.Activate
' ActiveSheet.Protect Password:="tesla"
'
' 'счёт 76
' Set import = importWB.Sheets("76")
' Set this = ThisWorkbook.Sheets("76")
'
' import.Activate
' ActiveSheet.Unprotect Password:="tesla"
' this.Activate
' ActiveSheet.Unprotect Password:="gfhjkm"
' 'диапазоны счёта 76
' '1
' this.Activate
' Range("O26").Copy
' import.Activate
' Range("Q28").Select
' Selection.PasteSpecial Paste:=xlPasteValues
'    '2
'    this.Activate
'    Range("A34:G44").Copy
'    import.Activate
'    Range("C45:I55").Select
'    Selection.PasteSpecial Paste:=xlPasteValues
'    Range("C45:I55").Select
'    Set MyRange = Selection
'    For Each MyCell In MyRange
'    If MyCell.Value = 0 Then
'    MyCell.Value = Empty
'    End If
'    Next MyCell
'        '3
'        this.Activate
'        Range("J34:P44").Copy
'        import.Activate
'        Range("L45:R55").Select
'        Selection.PasteSpecial Paste:=xlPasteValues
'        Range("L45:R55").Select
'        Set MyRange = Selection
'        For Each MyCell In MyRange
'        If MyCell.Value = 0 Then
'        MyCell.Value = Empty
'        End If
'        Next MyCell
'            '4
'            this.Activate
'            Range("S34:S44").Copy
'            import.Activate
'            Range("U45:U55").Select
'            Selection.PasteSpecial Paste:=xlPasteValues
'            Range("U45:U55").Select
'            Set MyRange = Selection
'            For Each MyCell In MyRange
'            If MyCell.Value = 0 Then
'            MyCell.Value = Empty
'            End If
'            Next MyCell
' import.Activate
' ActiveSheet.Protect Password:="tesla"

'вставка видимых проверок
ThisWorkbook.Sheets("Preferences").Activate
Range("L2").Copy

On Error GoTo Saver

For i = 2 To 8
    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = "E" & i 'диапазоны проверок
    Range2 = "A" & i 'диапазоны отчётов
    j = Range(Range1).Text
    y = Range(Range2).Text
    importWB.Sheets(y).Activate
    Range(j).Select
    Set MyRange = Selection
        For Each MyCell In MyRange
            If MyCell.Value = True Or MyCell.Value = False Then
                MyCell.Select
                Selection.PasteSpecial Paste:=xlPasteFormats
            End If
        Next MyCell
Next i
 
'сохранение с добавлением файла в уже существующую папку/созданием папки при сотсутствии
Saver:
    ThisWorkbook.Sheets("Preferences").Activate
    SaveName = ActiveSheet.Range("AC1").Text
    Folder = ActiveSheet.Range("AA1").Text

    Set object = CreateObject("Scripting.FileSystemObject")
    Path = ActiveWorkbook.Path

    If object.FolderExists(Path & "\" & Folder) Then
        importWB.Activate
        importWB.SaveAs Filename:=Path & "\" & Folder & "\" & _
        SaveName & ".xlsm"
    Else
        object.CreateFolder (Path & "\" & Folder)
            importWB.Activate
            importWB.SaveAs Filename:=Path & "\" & Folder & "\" & _
            SaveName & ".xlsm"
    End If
    importWB.Save
    
 'добавление общей проверочной формулы
    importWB.Sheets("60_01").Activate
    Range("U6").Select
    Range("U6").FormulaLocal = "=И(U8:U500;Y8:Y500;AC8:AC500)"
        importWB.Sheets("60_02").Activate
        Range("P11").Select
        Range("P11").FormulaLocal = "=И(P13:P500;T13:T500;V13:V500)"
            importWB.Sheets("62_01").Activate
            Range("T6").Select
            Range("T6").FormulaLocal = "=И(T9:T500;X9:X500;Z9:Z500)"
                importWB.Sheets("62_02").Activate
                Range("T6").Select
                Range("O6").FormulaLocal = "=И(O9:Q500;U9:U500)"
 'перенос проверочной формулы в Preferences
    importWB.Sheets("60_01").Activate
    Range("U6").Copy
    ThisWorkbook.Sheets("Preferences").Activate
    Range("Q3").Select
    Selection.PasteSpecial Paste:=xlPasteValues
        importWB.Sheets("60_02").Activate
        Range("P11").Copy
        ThisWorkbook.Sheets("Preferences").Activate
        Range("Q4").Select
        Selection.PasteSpecial Paste:=xlPasteValues
            importWB.Sheets("62_01").Activate
            Range("T6").Copy
            ThisWorkbook.Sheets("Preferences").Activate
            Range("Q5").Select
            Selection.PasteSpecial Paste:=xlPasteValues
                importWB.Sheets("62_02").Activate
                Range("O6").Copy
                ThisWorkbook.Sheets("Preferences").Activate
                Range("Q6").Select
                Selection.PasteSpecial Paste:=xlPasteValues
    
'    importWB.Close
' ThisWorkbook.Sheets("Preferences").Activate
'вставка форматов проверочных формул для ФСД
 ThisWorkbook.Sheets("Preferences").Activate
 Range("L2").Copy
 Range("Q3:Q300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteFormats
 End With

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


