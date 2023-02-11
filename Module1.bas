Attribute VB_Name = "Module1"
Sub this()

 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim importWB  As Workbook
 Dim import As Worksheet
 Dim this As Worksheet
 Dim MyRange As Range
 Dim MyCell As Range
 
 On Error GoTo ErrHandler
 
 Application.ScreenUpdating = False
 Application.DisplayAlerts = False
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsm), *.xlsm", _
 MultiSelect:=True, Title:="Файл для вставки")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "Файл не выбран!"
 GoTo ExitHandler
 End If
 
 Set ThisWorkbook = ActiveWorkbook
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

'счёт 76
 Set import = importWB.Sheets("76")
 Set this = ThisWorkbook.Sheets("76")

 import.Activate
 ActiveSheet.Unprotect Password:="tesla"
 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"
 '1
 this.Activate
 Range("O26").Copy
 import.Activate
 Range("Q28").Select
 Selection.PasteSpecial Paste:=xlPasteValues
 'диапазоны счёта 76
    '2
    this.Activate
    Range("A34:G44").Copy
    import.Activate
    Range("C45:I55").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("C45:I55").Select
    Set MyRange = Selection
    For Each MyCell In MyRange
    If MyCell.Value = 0 Then
    MyCell.Value = Empty
    End If
    Next MyCell
        '3
        this.Activate
        Range("J34:P44").Copy
        import.Activate
        Range("L45:R55").Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Range("L45:R55").Select
        Set MyRange = Selection
        For Each MyCell In MyRange
        If MyCell.Value = 0 Then
        MyCell.Value = Empty
        End If
        Next MyCell
        
            '4
            this.Activate
            Range("S34:S44").Copy
            import.Activate
            Range("U45:U55").Select
            Selection.PasteSpecial Paste:=xlPasteValues
            Range("U45:U55").Select
            Set MyRange = Selection
            For Each MyCell In MyRange
            If MyCell.Value = 0 Then
            MyCell.Value = Empty
            End If
            Next MyCell
            


ExitHandler:
 Application.ScreenUpdating = True
 import.Activate
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub


'Dim flag As Integer
'Const PASSWORD = "gfhjkm"

Sub Добавление_файла()
 
 Dim FilesToOpen
 Dim x As Integer
 Dim ThisWorkbook As Workbook
 Dim importWB  As Workbook
 On Error GoTo ErrHandler
 Set ThisWorkbook = ActiveWorkbook
 
 Application.ScreenUpdating = False 'отключаем обновление экрана для быстродействия
 
 FilesToOpen = Application.GetOpenFilename(MultiSelect:=True, Title:="Файл для копирования") 'вызываем диалог выбора файлов для импорта
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "Файл не выбран!"   'сообщение при отсутствии файла
 GoTo ExitHandler
 End If
   
 x = 1  'задаём переменную
 While x <= UBound(FilesToOpen) 'пока не достигнуты рамки файла выполянется цикл
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(x))
  
' ThisWorkbook.Sheets("Актуальная").Activate
' ActiveSheet.Unprotect PASSWORD:="gfhjkm"
  
 importWB.Sheets(1).Activate
 Range("A1:BB300").Copy
  
 ThisWorkbook.Sheets("Актуальная").Activate
 Range("A1:BB300").Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False   'вставляем значения
' ActiveSheet.Protect PASSWORD:="gfhjkm"

 ThisWorkbook.Sheets("Parsing").Activate
 
 x = x + 1
 Application.Run "ClearClipboard()"  'вызов макроса из макроса
 importWB.Close
 Wend
 
ThisWorkbook.Sheets("Inception").Activate
If Range("O5") = True Then
Dim Style, Title
Style = vbExclamation = 48
Title = "Ура!"
MsgBox "Изменений нет", Style, Title
Else
Dim Style1, Title1
Style1 = vbCritical = 16
Title1 = "Блин!"
MsgBox "Были внесены изменения", Style1, Title1
End If
ThisWorkbook.Sheets("Parsing").Activate

ExitHandler:    'обработка выхода
 Application.ScreenUpdating = True  'включение (выключеного) обновления экрана
 ThisWorkbook.Sheets("Parsing").Activate
 Exit Sub
 
ErrHandler: 'обработка ошибки
 MsgBox Err.Description
 Resume ExitHandler

End Sub



