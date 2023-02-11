Attribute VB_Name = "Support"
Sub ОбновитьСводныеТаблицы()

Dim ws As Worksheet
Dim pt As PivotTable
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False

ThisWorkbook.Activate
For Each ws In ThisWorkbook.Worksheets
ws.Unprotect Password:="gfhjkm"
Next ws

For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws

 Application.ScreenUpdating = True
 Application.EnableEvents = True
 ActiveSheet.DisplayPageBreaks = True
 Application.DisplayStatusBar = True
 Application.DisplayAlerts = True

'For Each ws In ThisWorkbook.Worksheets
'ws.Protect PASSWORD:="gfhjkm"
'Next ws

End Sub
Public Sub ClearClipboard()
    With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    .SetText Empty: .PutInClipboard
    End With
End Sub

Sub Test()

 Dim x As Integer
 Dim ThisWorkbook As Workbook    'задаём имя открытой книге
 Set ThisWorkbook = ActiveWorkbook   'присваиваем перменную активной книге
 
Application.ScreenUpdating = False 'отключаем обновление экрана для быстродействия
 
If Range("O28") = True Then
Dim Style, Title
Style = vbExclamation = 48
Title = "Ура!"
MsgBox "Ошибок нет", Style, Title
Else
Dim Style1, Title1
Style1 = vbCritical = 16
Title1 = "Ошибка!"
MsgBox "Есть ошибки", Style1, Title1
End If
 
Application.ScreenUpdating = True  'включение (выключеного) обновления экрана

End Sub
Sub AddRow()

' Dim FilesToOpen
' Dim ThisWorkbook As Workbook
' Dim importWB  As Workbook
 Dim x As Integer 'строки для добавления
 Dim y, Range1, Range2 As String

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

' FilesToOpen = Application.GetOpenFilename _
' (FileFilter:="Microsoft Excel Files (*.xlsm), *.xlsm", _
' MultiSelect:=True, Title:="Файл для вставки")
'
' If TypeName(FilesToOpen) = "Boolean" Then
' MsgBox "Файл не выбран!"
' GoTo ExitHandler
' End If
' Set ThisWorkbook = ActiveWorkbook
' Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error GoTo ExitHandler
For i = 2 To 38

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
'        Loop
    ActiveSheet.Protect Password:="tesla", DrawingObjects:=True, Contents:=True, Scenarios:=True _
       , AllowFormattingCells:=False, AllowFormattingColumns:=True, _
         AllowFormattingRows:=True, AllowInsertingColumns:=False, AllowInsertingRows _
       :=False, AllowSorting:=False, AllowFiltering:=True, AllowUsingPivotTables _
       :=False
        Loop
Next i
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    Exit Sub
ExitHandler:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
  
End Sub

Sub ЗаблокироватьВсеЛисты()
Dim ws As Worksheet
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False
For Each ws In ActiveWorkbook.Worksheets
ws.Protect Password:="gfhjkm"
Next ws
' Application.ScreenUpdating = True
' Application.EnableEvents = True
' ActiveSheet.DisplayPageBreaks = True
' Application.DisplayStatusBar = True
' Application.DisplayAlerts = True
ThisWorkbook.Sheets("Preferences").Activate
End Sub
Sub СнятьЗащитуВсехЛистов()
Dim ws As Worksheet
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False
For Each ws In ActiveWorkbook.Worksheets
ws.Unprotect Password:="gfhjkm"
Next ws
 Application.ScreenUpdating = True
 Application.EnableEvents = True
 ActiveSheet.DisplayPageBreaks = True
 Application.DisplayStatusBar = True
 Application.DisplayAlerts = True
'ThisWorkbook.Sheets("Preferences").Activate
End Sub

Sub ЭйТи_Консалтинг()
 
 Dim ThisWorkbook As Workbook
 Set ThisWorkbook = ActiveWorkbook
 Dim ATK As Worksheet
 Set ATK = ThisWorkbook.Sheets("ЭйТи Консалтинг")
 Dim Inception As Worksheet
 Set Inception = ThisWorkbook.Sheets("Inception")
 Dim Parsing As Worksheet
 Set Parsing = ThisWorkbook.Sheets("Parsing")
 On Error GoTo ErrHandler
 Application.ScreenUpdating = False
  
' ATK.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
' Inception.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
' Parsing.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
 
 ATK.Activate
 Range("A1:N300").Copy

 Inception.Activate
 Range("A1:N300").Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
 
 Parsing.Activate
 Range("P5") = "ООО ЭЙТИ КОНСАЛТИНГ"
 
 Parsing.Activate
 Application.Run "ClearClipboard()"
    
' Inception.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"
' ATK.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"
' Parsing.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"

ExitHandler:
 Application.ScreenUpdating = True
 Parsing.Activate
 Exit Sub
 
ErrHandler:
MsgBox Err.Description
Resume ExitHandler

End Sub
Sub СиИкс_Лаб()
 
 Dim ThisWorkbook As Workbook
 Set ThisWorkbook = ActiveWorkbook
 Dim ATK As Worksheet
 Set ATK = ThisWorkbook.Sheets("СиИкс Лаб")
 Dim Inception As Worksheet
 Set Inception = ThisWorkbook.Sheets("Inception")
 Dim Parsing As Worksheet
 Set Parsing = ThisWorkbook.Sheets("Parsing")
 On Error GoTo ErrHandler
 Application.ScreenUpdating = False
  
' ATK.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
' Inception.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
' Parsing.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
  
 ATK.Activate
 Range("A1:N300").Copy

 Inception.Activate
 Range("A1:N300").Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
 Parsing.Activate
 Range("P5") = "СИ ИКС ЛАБ ООО"
 
 Parsing.Activate
 Application.Run "ClearClipboard()"
    
' Inception.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"
' ATK.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"
' Parsing.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"

ExitHandler:
 Application.ScreenUpdating = True
 Parsing.Activate
 Exit Sub
 
ErrHandler:
MsgBox Err.Description
Resume ExitHandler

End Sub
Sub Философия()
 
 Dim ThisWorkbook As Workbook
 Set ThisWorkbook = ActiveWorkbook
 Dim ATK As Worksheet
 Set ATK = ThisWorkbook.Sheets("Философия")
 Dim Inception As Worksheet
 Set Inception = ThisWorkbook.Sheets("Inception")
 Dim Parsing As Worksheet
 Set Parsing = ThisWorkbook.Sheets("Parsing")
 On Error GoTo ErrHandler
 Application.ScreenUpdating = False
  
' ATK.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
' Inception.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
' Parsing.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
  
 ATK.Activate
 Range("A1:N300").Copy

 Inception.Activate
 Range("A1:N300").Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
 Parsing.Activate
 Range("P5") = "ФИЛОСОФИЯ.ИТ ООО"
 
 Parsing.Activate
 Application.Run "ClearClipboard()"
    
' Inception.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"
' ATK.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"
' Parsing.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"

ExitHandler:
 Application.ScreenUpdating = True
 Parsing.Activate
 Exit Sub
 
ErrHandler:
MsgBox Err.Description
Resume ExitHandler

End Sub
Sub Группа()
 
 Dim ThisWorkbook As Workbook
 Set ThisWorkbook = ActiveWorkbook
 Dim ATK As Worksheet
 Set ATK = ThisWorkbook.Sheets("Группа ЭйТи")
 Dim Inception As Worksheet
 Set Inception = ThisWorkbook.Sheets("Inception")
 Dim Parsing As Worksheet
 Set Parsing = ThisWorkbook.Sheets("Parsing")
 On Error GoTo ErrHandler
 Application.ScreenUpdating = False
   
' ATK.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
' Inception.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
' Parsing.Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
  
 ATK.Activate
 Range("A1:N300").Copy

 Inception.Activate
 Range("A1:N300").Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
 
 Parsing.Activate
 Range("P5") = "ГРУППА ЭЙТИ АО"
 
 Parsing.Activate
 Application.Run "ClearClipboard()"
 
' Inception.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"
' ATK.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"
' Parsing.Activate
' ActiveSheet.Protect PASSWORD:="gfhjkm"

ExitHandler:
 Application.ScreenUpdating = True
 Parsing.Activate
 Exit Sub
 
ErrHandler:
MsgBox Err.Description
Resume ExitHandler

End Sub
