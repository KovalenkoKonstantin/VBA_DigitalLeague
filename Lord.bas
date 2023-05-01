Attribute VB_Name = "Lord"
Sub Òûê()
 
 Dim ws As Worksheet
 Dim pt As PivotTable
 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim importWB  As Workbook
 Dim rCell As Range
 On Error GoTo ErrHandler
 Set ThisWorkbook = ActiveWorkbook
 
 Application.ScreenUpdating = False
 
 FilesToOpen = Application.GetOpenFilename(MultiSelect:=True, Title:="Ôàéë äëÿ êîïèðîâàíèÿ")
 
 If TypeName(FilesToOpen) = "Boolean" Then
' MsgBox "Ôàéë íå âûáðàí!"
 GoTo ExitHandler
 End If
   
 Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
 
 ThisWorkbook.Activate
 ÑíÿòüÇàùèòóÂñåõËèñòîâ
   
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
 ThisWorkbook.Sheets("58í").Activate
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
 ThisWorkbook.Sheets("58êîíòð").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(4).Activate
 Range("A1:BB300").Copy
  
 ThisWorkbook.Sheets("60").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(5).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("60í").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(6).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("60êîíòð").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With
 
 importWB.Sheets(7).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("62").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

 importWB.Sheets(8).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("62í").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

 importWB.Sheets(9).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("62êîíòð").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

 importWB.Sheets(10).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("66").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

 importWB.Sheets(11).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("66í").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

 importWB.Sheets(12).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("66êîíòð").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

 importWB.Sheets(13).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("76").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

 importWB.Sheets(14).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("76í").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

 importWB.Sheets(15).Activate
 Range("A1:BB300").Copy
 ThisWorkbook.Sheets("76êîíòð").Activate
 Range("A1:BB300").Select
 With Selection
        .PasteSpecial Paste:=xlPasteAll
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
 End With

ThisWorkbook.Sheets("Merge").Activate

 
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
 
If Range("W3") = True Then
Dim Style, Title
Style = vbExclamation = 48
Title = "Óðà!"
'MsgBox "Âñ¸ âñòàëî íà ñâîè ìåñòà", Style, Title
Else
Dim Style1, Title1
Style1 = vbCritical = 16
Title1 = "Áëèí!"
'MsgBox "...", Style1, Title1
End If

ExitHandler:
 Application.ScreenUpdating = True
' ThisWorkbook.Activate
' ÇàáëîêèðîâàòüÂñåËèñòû
 ThisWorkbook.Sheets("Merge").Activate
 Exit Sub
 
' ÇàáëîêèðîâàòüÂñåËèñòû
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub


