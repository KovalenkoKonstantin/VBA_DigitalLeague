Attribute VB_Name = "Transfer"
Sub Зачипись()

 Dim FilesToOpen
 Dim flag As Integer
 Dim flag1 As Integer
 Dim ThisWorkbook As Workbook
 Dim importWB  As Workbook
 Dim SaveName As String
 Dim Folder As String
 Dim Path As String
 Dim wb As Workbook
 Dim ws As Worksheet
 Dim MyRange As Range
 Dim MyCell As Range
 Dim range1 As String
 Dim range2 As String
 Dim range3 As String
 Dim range4 As String
 Dim range5 As String
 Dim range6 As String
 Dim range52 As String
 Dim range62 As String
 Dim variants As String
 Dim variants2 As String
 Dim operation As String
 
 On Error GoTo ErrHandler
 
 Set ThisWorkbook = ActiveWorkbook
 
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False
 
' ThisWorkbook.Activate
' СнятьЗащитуВсехЛистов
 
 ThisWorkbook.Sheets("Merge").Activate
 SaveName = ActiveSheet.Range("S1").Text
 Folder = ActiveSheet.Range("AB2").Text
 flag = ActiveSheet.Range("AE1").Value
 'области
 'счёт
 range1 = ActiveSheet.Range("AF1").Text
 range2 = ActiveSheet.Range("AF2").Text
 'кор.счёт
 range3 = ActiveSheet.Range("AG1").Text
 range4 = ActiveSheet.Range("AG2").Text
 'операции и аналитики1
 range5 = ActiveSheet.Range("AH1").Text
 range6 = ActiveSheet.Range("AH2").Text
 'операции и аналитики2
 range52 = ActiveSheet.Range("AI1").Text
 range62 = ActiveSheet.Range("AI2").Text
 'аналитики1
 variants = ActiveSheet.Range("AJ2").Text
 'аналитики2
 variants2 = ActiveSheet.Range("AK2").Text
 'операции
 operation = ActiveSheet.Range("AL2").Text
 
 FilesToOpen = Application.GetOpenFilename(MultiSelect:=True, Title:="Файл для вставки")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If

 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
 importWB.Sheets(1).Activate
' With importWB.Sheets(1).Cells
'    .EntireColumn.AutoFit
'    .EntireRow.AutoFit
' End With
 
 'добавление строк
 Set ws = importWB.ActiveSheet
 ws.Unprotect Password:="tesla"
 
    flag1 = 2
    lr = Cells(Rows.count, 1).End(xlUp).Row
    y = 1
While y <= lr
    If ws.Cells(y, 1).Value = "[END_" & CStr(flag1) & "]" Then
        lr = y - 1
        counter = 1
        Do While counter <= flag
            counter = counter + 1
            Rows(lr).Copy: Rows(lr).Offset(1).Resize(1).EntireRow.Insert
            For i = 1 To 100
                If Cells(lr + 1, i).Interior.ColorIndex = 19 Then
                    Cells(lr + 1, i).ClearContents
                End If
            Next i
            lr = lr + 1
        Loop
    End If
    y = y + 1
Wend
 

 'наименование компании
 importWB.Sheets(1).Activate
 ThisWorkbook.Sheets("Merge").Activate
 Range("AB1").Copy
 importWB.Sheets(1).Activate
 Range("C3").Select
 Selection.PasteSpecial Paste:=xlPasteValues
 
 'наименование контрагента
 importWB.Sheets(1).Activate
 ThisWorkbook.Sheets("Merge").Activate
 Range("AB2").Copy
 importWB.Sheets(1).Activate
 Range("G3").Select
 Selection.PasteSpecial Paste:=xlPasteValues
 
 'период
 importWB.Sheets(1).Activate
 ThisWorkbook.Sheets("Merge").Activate
 Range("AD1").Copy
 importWB.Sheets(1).Activate
 Range("G5").Select
 Selection.PasteSpecial Paste:=xlPasteValues

 'номер счета
 importWB.Sheets(1).Activate
 ThisWorkbook.Sheets("Lord").Activate
 Range(range1).Copy
 importWB.Sheets(1).Activate
 Range(range2).Select
' Selection.PasteSpecial Paste:=xlPasteValues
 With Selection
        .PasteSpecial Paste:=xlPasteValues
        .UnMerge
        .WrapText = False
        .MergeCells = False
 End With
  Range(range2).Select
    Set MyRange = Selection
    For Each MyCell In MyRange
    MyCell.Activate
    Next MyCell
 
 'номер кор счета
 importWB.Sheets(1).Activate
 ThisWorkbook.Sheets("Lord").Activate
 Range(range3).Copy
 importWB.Sheets(1).Activate
 Range(range4).Select
 Selection.PasteSpecial Paste:=xlPasteValues
 With Selection
        .PasteSpecial Paste:=xlPasteValues
        .UnMerge
        .WrapText = False
        .MergeCells = False
 End With
 Range(range4).Select
    Set MyRange = Selection
    For Each MyCell In MyRange
    MyCell.Activate
    Next MyCell
    
 'операции и аналитики1
 importWB.Sheets(1).Activate
 ThisWorkbook.Sheets("Lord").Activate
 Range(range5).Copy
 importWB.Sheets(1).Activate
 Range(range6).Select
 Selection.PasteSpecial Paste:=xlPasteValues
 With Selection
        .PasteSpecial Paste:=xlPasteValues
        .UnMerge
        .WrapText = False
        .MergeCells = False
 End With
 Range(range6).Select
    Set MyRange = Selection
    For Each MyCell In MyRange
        MyCell.Activate
    If MyCell.Value = 0 Then
    MyCell.Value = Empty
    End If
    Next MyCell
    
    'операции и аналитики2
 importWB.Sheets(1).Activate
 ThisWorkbook.Sheets("Lord").Activate
 Range(range52).Copy
 importWB.Sheets(1).Activate
 Range(range62).Select
 Selection.PasteSpecial Paste:=xlPasteValues
 With Selection
        .PasteSpecial Paste:=xlPasteValues
        .UnMerge
        .WrapText = False
        .MergeCells = False
 End With
 Range(range62).Select
    Set MyRange = Selection
    For Each MyCell In MyRange
        MyCell.Activate
    If MyCell.Value = 0 Then
    MyCell.Value = Empty
    End If
    Next MyCell
    
  'варианты1
 importWB.Sheets(1).Activate
 Range(variants).Select
 Set MyRange = Selection

 For Each MyCell In MyRange
    If MyCell.Value <> 0 Then
        MyCell.Select
                With Selection.Interior
                .Pattern = xlSolid
                .PatternColor = 65535
                .Color = 13434879
                .TintAndShade = 0
                .PatternTintAndShade = 0
                End With
        Selection.Locked = False
        Selection.FormulaHidden = False
        Selection.ShrinkToFit = False
    End If
 Next MyCell
 
 'варианты2
 importWB.Sheets(1).Activate
 Range(variants2).Select
 Set MyRange = Selection

 For Each MyCell In MyRange
    If MyCell.Value <> 0 Then
        MyCell.Select
                With Selection.Interior
                .Pattern = xlSolid
                .PatternColor = 65535
                .Color = 13434879
                .TintAndShade = 0
                .PatternTintAndShade = 0
                End With
        Selection.Locked = False
        Selection.FormulaHidden = False
        Selection.ShrinkToFit = False
    End If
 Next MyCell
 
 'операции
 importWB.Sheets(1).Activate
 Range(operation).Select
 Set MyRange = Selection

 For Each MyCell In MyRange
    If MyCell.Value <> 0 Then
        MyCell.Select
        Selection.Locked = False
        Selection.FormulaHidden = False
        Selection.ShrinkToFit = False
    End If
 Next MyCell
 
 Сальдо
' CheckAct

    'сохранение файла
    ThisWorkbook.Sheets("Merge").Activate
    create
    
    ThisWorkbook.Activate
    Path = ActiveWorkbook.Path
    importWB.Sheets(1).Activate
    importWB.Sheets(1).Protect Password:="tesla"
    importWB.SaveAs Filename:=Path & "\" & Folder & "\" & _
    SaveName & ".xlsm"

 importWB.Close

ExitHandler:
 ThisWorkbook.Sheets("Merge").Activate
 Application.ScreenUpdating = True
 Application.EnableEvents = True
 ActiveSheet.DisplayPageBreaks = True
 Application.DisplayStatusBar = True
 Application.DisplayAlerts = True
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
           
End Sub

