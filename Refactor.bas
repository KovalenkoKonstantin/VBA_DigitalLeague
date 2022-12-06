Attribute VB_Name = "Refactor"
Sub инн()

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
 MultiSelect:=True, Title:="Файл для вставки")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "Файл не выбран!"
 GoTo ExitHandler
 End If
 
 СнятьЗащитуВсехЛистов
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

ThisWorkbook.Sheets("ИНН").Activate
 Range("A1:BB400000").Select
 With Selection
        .Clear
 End With

 importWB.Sheets(1).Activate
 Range("A1:BB400000").Select
 Range("A1:BB400000").Copy
 ThisWorkbook.Sheets("ИНН").Activate
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
 ЗаблокироватьВсеЛисты
 
 MsgBox "Справочник ГИД (ИНН) успешно загружен"

ExitHandler:
 Application.ScreenUpdating = True
 ThisWorkbook.Sheets("Preferences").Activate
 ЗаблокироватьВсеЛисты
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub
Sub рефактор_инн()

 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim this As Worksheet
' Dim MyRange As Range
' Dim MyCell As Range
 
 Set ThisWorkbook = ActiveWorkbook
 
 On Error GoTo ErrHandler
 
 Application.ScreenUpdating = False
 Application.DisplayAlerts = False

 СнятьЗащитуВсехЛистов

ThisWorkbook.Sheets("90 контр").Activate
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
 ЗаблокироватьВсеЛисты
 
 MsgBox "ИНН исправлены"

ExitHandler:
 Application.ScreenUpdating = True
 ThisWorkbook.Sheets("Preferences").Activate
 ЗаблокироватьВсеЛисты
 ClearClipboard
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

