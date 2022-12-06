Attribute VB_Name = "FiftyEighth"
Sub FiftyEights()
 
 Dim ws As Worksheet
 Dim pt As PivotTable
 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim importWB  As Workbook
 Dim rCell As Range
 On Error GoTo ErrHandler
 Set ThisWorkbook = ActiveWorkbook
 
 Application.ScreenUpdating = False
 
 FilesToOpen = Application.GetOpenFilename(MultiSelect:=True, Title:="‘айл дл€ копировани€")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 MsgBox "‘айл не выбран!"
 GoTo ExitHandler
 End If
   
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
 
 ThisWorkbook.Activate
 —н€ть«ащиту¬сехЋистов
   
 importWB.Sheets(1).Activate
 Range("A1:BB300").Copy
  
 ThisWorkbook.Sheets("58").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(2).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("58н").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(3).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("58контр").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

ThisWorkbook.Sheets("Processing 58").Activate

 
    For Each ws In ThisWorkbook.Worksheets
    For Each pt In ws.PivotTables
    pt.RefreshTable
    Next pt
    Next ws
 
    With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    .SetText Empty: .PutInClipboard
    End With
    
 importWB.Close

For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws
 
If Range("BB3") = True Then
Dim Style, Title
Style = vbExclamation = 48
Title = "”ра!"
Else
Dim Style1, Title1
Style1 = vbCritical = 16
Title1 = "Ѕлин!"
MsgBox "...", Style1, Title1
End If

' ThisWorkbook.Activate
' «аблокировать¬сеЋисты

ExitHandler:
 Application.ScreenUpdating = True
 ThisWorkbook.Activate
' «аблокировать¬сеЋисты
 ThisWorkbook.Sheets("Processing 58").Activate
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub


