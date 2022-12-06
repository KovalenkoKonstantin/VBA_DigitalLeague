Attribute VB_Name = "Support"
Sub create()
Dim Folder As String
Dim Path As String
Dim Slash As String
Dim object As Object
Dim ThisWorkbook As Workbook

Set ThisWorkbook = ActiveWorkbook
Set object = CreateObject("Scripting.FileSystemObject")

ThisWorkbook.Sheets("Preferences").Activate
Folder = ActiveSheet.Range("L9").Text
Path = ActiveWorkbook.Path

    If object.FolderExists(Path & "\" & Folder) Then
        object.DeleteFolder (Path & "\" & Folder)
        object.CreateFolder (Path & "\" & Folder)
    Else
        object.CreateFolder (Path & "\" & Folder)
    End If

End Sub
Sub ќбновить—водные“аблицы()

Dim ws As Worksheet
Dim pt As PivotTable
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook

'ThisWorkbook.Activate
'For Each ws In ThisWorkbook.Worksheets
'ws.Unprotect PASSWORD:="gfhjkm"
'Next ws

For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws

'For Each ws In ThisWorkbook.Worksheets
'ws.Protect PASSWORD:="gfhjkm"
'Next ws

End Sub

Public Sub ClearClipboard()
    With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    .SetText Empty: .PutInClipboard
    End With
End Sub

Sub «аблокировать¬сеЋисты()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Protect Password:="gfhjkm"
Next ws
ActiveWorkbook.Sheets("—писок_компаний").Unprotect Password:="gfhjkm"
ActiveWorkbook.Protect Password:="gfhjkm"
ThisWorkbook.Sheets("Preferences").Activate
End Sub

Sub —н€ть«ащиту¬сехЋистов()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Unprotect Password:="gfhjkm"
Next ws
ActiveWorkbook.Unprotect Password:="gfhjkm"
ThisWorkbook.Sheets("Preferences").Activate
End Sub

Sub »справить()
 Dim ThisWorkbook As Workbook
 Dim ws As Worksheet
 Dim pt As PivotTable
 Dim MyRange As Range
 Dim MyCell As Range
 Dim this As Worksheet
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ErrHandler
 Application.ScreenUpdating = False
 Application.DisplayAlerts = False
 
 Set this = ThisWorkbook.Sheets("Preferences")

 this.Activate
 ActiveSheet.Unprotect Password:="gfhjkm"

 this.Activate
 Range("AS1").Copy
 this.Activate
 Range("AS2").Select
 Selection.PasteSpecial Paste:=xlPasteValues

 
ExitHandler:
 Application.ScreenUpdating = True
 ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
 
End Sub

Sub delete()
 Dim ws As Worksheet
 Dim pt As PivotTable
 Dim ThisWorkbook As Workbook
 Dim rCell As Range
 On Error GoTo ErrHandler
 Set ThisWorkbook = ActiveWorkbook
 
 Application.ScreenUpdating = False
 
 —н€ть«ащиту¬сехЋистов
  
 ThisWorkbook.Sheets("Data90").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("Data90-1").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("Data90-2").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("Preferences").Activate
 Range("AS2").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("»ЌЌ").Activate
 Range("A1:BB400000").Select
 With Selection
        .Clear
 End With


For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws

For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws

 ThisWorkbook.Sheets("Preferences").Activate
 «аблокировать¬сеЋисты
 
' MsgBox "All clear"
 
ExitHandler:
 Application.ScreenUpdating = True
 ThisWorkbook.Sheets("Preferences").Activate
 «аблокировать¬сеЋисты
 Exit Sub

ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub



