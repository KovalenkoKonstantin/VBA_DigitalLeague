Attribute VB_Name = "Test"
Sub тест_60_01()

 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim ws, this As Worksheet
 Dim pt As PivotTable
 Dim MyRange, MyCell As Range
 Dim x As Integer
 x = 25
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ErrHandler
 
Application.ScreenUpdating = False
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
 
  'проверка правильного выбора отчёта
  ThisWorkbook.Sheets("Test").Activate
  If Range("H1").Value > 1 Then
    MsgBox "Одновременно тестировать более чем один отчёт не могу"
    Range("G2").Value = False
    Range("G5").Value = False
    Range("G8").Value = False
    Range("G11").Value = False
    GoTo ExitHandler
  End If
  If Range("H1").Value = 0 Then
    MsgBox "Выберите отчёт для тестирования"
    GoTo ExitHandler
  End If
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Файл для вставки")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "Файл не выбран! Шеф Усё пропало!!!"
 GoTo ExitHandler
 End If
 
 'удаление предыдущих данных
 On Error Resume Next
 For i = 1 To x
 ThisWorkbook.Sheets(i).Activate
 Range("A1:BB3000").Select
 With Selection
        .Clear
 End With
 Next i

'вставка листов
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
For i = 1 To 1
On Error GoTo ExitHandler
 importWB.Sheets(i).Activate
 Range("A1:BB3000").Select
 Range("A1:BB3000").Copy
 ThisWorkbook.Sheets("100").Activate
 Range("A1:BB3000").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 Next i
 
 importWB.Close
 ThisWorkbook.Activate

 Dim Range1, Range2, Range3, Range4, Range5, Range6, Range7, Range8, y As String

 Dim object As Object
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ErrHandler
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsm), *.xlsm", _
 MultiSelect:=True, Title:="Файл для вставки")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "Файл не выбран! Тест завершён неудачно :("
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
Next i
 
 'определение диапазонов
'For i = 2 To 8
i = 2

    ThisWorkbook.Sheets("Ranges").Activate
    Range1 = Range("G" & i).Text
    Range2 = Range("H" & i).Text
    Range3 = Range("I" & i).Text
    Range4 = Range("J" & i).Text
    Range5 = Range("K" & i).Text
    Range6 = Range("L" & i).Text
    Range7 = Range("M" & i).Text
    Range8 = Range("N" & i).Text

    y = Range("A" & 2).Text
    importWB.Sheets(y).Activate
' Next i
 
 'счёт 60_01
 Set import = importWB.Sheets("60_01")
 Set this = ThisWorkbook.Sheets("100")

 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 'диапазоны счёта 60_01
 '1
 this.Activate
 Range(Range1).Copy
 import.Activate
 Range(Range2).Select
 Selection.PasteSpecial Paste:=xlPasteValues
    '2
    this.Activate
    Range(Range3).Copy
    import.Activate
    Range(Range4).Select
    Selection.PasteSpecial Paste:=xlPasteValues
        '3
        this.Activate
        Range(Range5).Copy
        import.Activate
        Range(Range6).Select
        Selection.PasteSpecial Paste:=xlPasteValues
            '4
            this.Activate
            Range(Range7).Copy
            import.Activate
            Range(Range8).Select
            Selection.PasteSpecial Paste:=xlPasteValues


'вставка видимых проверок
ThisWorkbook.Sheets("Preferences").Activate
Range("L2").Copy

On Error Resume Next

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

On Error GoTo ExitHandler
    importWB.Sheets("60_01").Activate
'    Range("U6").Select
    Range("U6").FormulaLocal = "=И(U8:U500;Y8:Y500;AC8:AC500)"
    
    If Range("U6").Value = True Then
        MsgBox "Все проверки прошли, ошибок не обнаружено"
    Else
        MsgBox "Проверки РА не пройдены. В отчёте допущены ошибки"
    End If
'    importWB.Close SaveChanges:=False

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Test").Activate
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
           
End Sub



