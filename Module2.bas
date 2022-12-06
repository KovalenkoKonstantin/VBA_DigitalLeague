Attribute VB_Name = "Module2"
'����������...

Const TEMPLATE_KEY = "������"
Const MAX_LINES = 65000
Const MSG_TITLE = "���_���"
Const STR_DEBIT = "�����"
Const STR_CREDIT = "������"
Const ROW_TITLE = 12
Const COL_AMOUNT_DEBIT = 11
Const COL_AMOUNT_CREDIT = 13
Const COL_SETTLEMENT_ACC = 3
Const COL_CORR_ACC = 5
Const SHT_5100 = "5100"
Const COL_5100_DEBIT = 3
Const COL_5100_CREDIT = 5
Const STR_KEY_DYN = "RC"
Const SHEET_CHECK_SETTIGS = "System"
Const COL_CHECK_SETTINGS = 7 '����� ������� � ������� �������� �������� ����� � ����������
Const SHEET_PARAMS = "System"
Const CHK_TYPE_FIXED = "F"
Const COLCALC = "COLCALC"
Const ROWCALC = "ROWCALC"
Const PAYSTR = "PAYSTR"
Const INVSTR = "INVSTR"
Const BALSTR = "BALSTR"
Const STR_COM_ERR = "<���_���>"
Public NA
'������ ������ ��������� ������ �� �����
Public Function GetLastRow(sht As Worksheet) As Long
 '   If sht.Cells.Count > 0 Then
 Set cell = sht.Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
 If Not cell Is Nothing Then
        GetLastRow = cell.Row
    End If
End Function
'������ ������ ��������� ������� �� �����
Public Function GetLastCol(sht As Worksheet) As Long
'    If sht.Cells.Count > 0 Then
Set cell = sht.Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
If Not cell Is Nothing Then
GetLastCol = cell.Column

    End If
End Function
Public Sub CleanAct()

Dim c As Range

For Each c In ActiveSheet.UsedRange

   If c.Interior.Pattern = xlHorizontal Then
      
      c.Interior.Pattern = xlSolid
      
   End If
   
Next c

End Sub
Public Function IsSheetExists(Name As String) As Boolean
On Error GoTo NoValid
Dim sht As Worksheet
Set sht = ActiveWorkbook.Worksheets(Name)
IsSheetExists = True

Exit Function
NoValid:
IsSheetExists = False
End Function

Public Function IsSheetExists2(wkbook As Workbook, Name As String) As Boolean
On Error GoTo NoValid
Dim sht As Worksheet
Set sht = wkbook.Worksheets(Name)
IsSheetExists2 = True

Exit Function
NoValid:
IsSheetExists2 = False
End Function
Private Function GetParValStr(strPar As String) As String
Dim cell As Range
Dim sht As Worksheet
If IsSheetExists(SHEET_PARAMS) = True Then
Set sht = ThisWorkbook.Worksheets(SHEET_PARAMS)
Set cell = sht.Columns(1).Find(strPar)
If Not cell Is Nothing Then
GetParValStr = cell.Offset(0, 1).Value
Else
GetParValStr = ""
End If
Else
GetParValStr = ""
End If

End Function

Private Function GetParValInt(strPar As String) As String
Dim str As String
str = GetParValStr(strPar)
If str = "" Then
GetParValInt = 0
Else
GetParValInt = CInt(str)
End If
End Function
Public Function CheckAnalytics(ByRef pList As Worksheet, pAcc As String, pColumn As Integer) As Boolean

' �������� ����� �� ����� ������������ ��� ��������� �����
Const pRangeToFind = "B1:B500"

CheckAnalytics = False

With pList.Range(pRangeToFind)
    Set c = .Find(pAcc, LookIn:=xlValues)
    If Not c Is Nothing Then
    
       If (pList.Cells(c.Row, pColumn) <> Empty) Then
          
          CheckAnalytics = True
       
       End If
    
    End If
    
End With

End Function
'������� ��������� ������� �� ��������� ���� ����/��������� �� �������� ���������
Private Function GetAccExceptions(strAcc As String, j As Integer) As Boolean
Dim str As String
str = GetParValStr(strAcc)
If InStr(1, str, CStr(j)) > 0 Then
    GetAccExceptions = True
    Else
    GetAccExceptions = False
    End If

End Function
'�������� ���� ���������
Public Function CheckByFormulas() As Integer

On Error GoTo catcherr:
'���� ���_���
Dim sht As Worksheet
'���� � �����������
Dim shtSettings As Worksheet
Dim nrowstart As Integer
Dim nRow As Long
Dim pColBal As Collection
Dim pColInv As Collection
Dim pColPay As Collection
Dim pColFix As Collection
Dim cell As Range
Dim pCheck As CVGOCheckParams
Dim nErr As Integer '���������� ��� ��������
'������ �������� ����� �� ������
Set pColPay = New Collection
'������ �������� ����� �� ����������
Set pColInv = New Collection
'������ �������� ����� ��
Set pColBal = New Collection
'������ �������� ����� ������������� ��������
Set pColFix = New Collection

Dim sBlockBegBalance   As String
'�������� ������� ��
sBlockBegBalance = GetParValStr("BALSTR")
If sBlockBegBalance = "" Then
    sBlockBegBalance = "�������� ������ �� ������ ��������� �������:"
End If
Dim sBlockInvoice   As String
'��������� �������� ����� �� ����������
sBlockInvoice = GetParValStr("INVSTR")
If sBlockInvoice = "" Then
    sBlockInvoice = "������� �� ������ ������������� � ������:"
End If
Dim sBlockPayment    As String
'��������� �������� ����� �� ������

sBlockPayment = GetParValStr("PAYSTR")
If sBlockPayment = "" Then
    sBlockPayment = "������:"
End If

'������� ����� ��������� ����������� �� �������
ClearCommetnts

nrowstart = 1
If IsSheetExists(SHEET_CHECK_SETTIGS) = True Then
    Set shtSettings = ThisWorkbook.Worksheets(SHEET_CHECK_SETTIGS)
    If shtSettings.Cells(nrowstart, COL_CHECK_SETTINGS).Value <> "" Then
        Set cell = shtSettings.Columns(COL_CHECK_SETTINGS).Find("*", shtSettings.Cells(nrowstart, COL_CHECK_SETTINGS))
        If Not cell Is Nothing Then
            '��������� ��������� ��������
            While cell.Row <> nrowstart
                Set pCheck = New CVGOCheckParams
                pCheck.WkbookName = ThisWorkbook.Name
                pCheck.Num = shtSettings.Cells(cell.Row, cell.Column - 1).Value
                pCheck.BlockName = cell.Value
                pCheck.CheckType = shtSettings.Cells(cell.Row, cell.Column + 1).Value
                pCheck.Name = shtSettings.Cells(cell.Row, cell.Column + 2).Value
                pCheck.Description = shtSettings.Cells(cell.Row, cell.Column + 3)
                pCheck.formula = shtSettings.Cells(cell.Row, cell.Column + 4).Value
                If shtSettings.Cells(cell.Row, cell.Column + 5).Value <> "" Then
                    pCheck.Col = CInt(shtSettings.Cells(cell.Row, cell.Column + 5).Value)
                Else
                    pCheck.Col = 0
                End If
                If shtSettings.Cells(cell.Row, cell.Column + 6).Value <> "" Then
                    pCheck.Row = CInt(shtSettings.Cells(cell.Row, cell.Column + 6).Value)
                Else
                    pCheck.Row = 0
                End If
                '���������, ��� ��� ������������ ��������� �������� ���� �������. ����� ���������� ��������
                If pCheck.formula <> "" And pCheck.BlockName <> "" And pCheck.Description <> "" And _
                    pCheck.Col > 0 And ((pCheck.Row > 0 And pCheck.CheckType = "F") Or pCheck.CheckType <> "F") _
                Then
                    If pCheck.CheckType <> "F" Then
                        If pCheck.BlockName = "BAL" Then
                            pColBal.Add pCheck
                        Else
                            If pCheck.BlockName = "INV" Then
                                pColInv.Add pCheck
                            Else
                                If pCheck.BlockName = "PAY" Then
                                    pColPay.Add pCheck
                                End If
                            End If
                        End If
                    Else
                        pColFix.Add pCheck
                    End If
                End If
                Set cell = shtSettings.Columns(COL_CHECK_SETTINGS).Find("*", cell)
            Wend
            '�������� ����
            Set sht = ThisWorkbook.Worksheets("���_���")
            nLastRow = GetLastRow(sht)
            nErr = 0
            nErr = nErr + CheckBlockByFormula(sht, pColFix, 1)

            nRow = 1
            Do
                nRow = FindNextBlockRow(sht, sBlockBegBalance, nRow)
                If nRow > 0 Then
                    nErr = nErr + CheckBlockByFormula(sht, pColBal, nRow)
                    nRow = FindNextBlockRow(sht, sBlockInvoice, nRow + 1)
                    If nRow > 0 Then
                        nErr = nErr + CheckBlockByFormula(sht, pColInv, nRow)
                        'nRow = FindNextBlockRow(sht, sBlockPayment, nRow + 1)
                        'If nRow > 0 Then
                        '    nErr = nErr + CheckBlockByFormula(sht, pColPay, nRow)
                        'End If
                    End If
                End If
                If nRow > 0 Then
                    nRow = nRow + 1
                End If
            Loop While nRow > 0
        Else
            MsgBox "����������� ���������"
        End If
    Else
        MsgBox "����� ������� �������� � ������� B1 ������ ���� ���������!"
    End If
Else
    MsgBox "����������� ���� ���������"
End If

Set pColPay = Nothing
Set pColInv = Nothing
Set pColBal = Nothing
Set pColFix = Nothing
CheckByFormulas = nErr
Exit Function

catcherr:
 MsgBox Err.Description

End Function

Public Sub CheckAct()
'____ ������� ������ � ������
ActiveSheet.Protect Password:="tesla", AllowInsertingRows:=True
ActiveSheet.Unprotect Password:="tesla"
'____
CleanAct

'Worksheets("���_���").Cells(1, 9) = "��������"

Const pSheetACC_AN = "�-�" ' ��� ���� ��� ����� ��������� ������� ������������ ������ - ����������
Const pSheetLISTS = "������" ' ��� ���� ��� ����� ��������� ������� ������������ ������ - ����������
Const pAddressOfCurrency = "C5" ' ������ ��� �������� ������
Dim pRANGEwithACCOUNTS As String
pRANGEwithACCOUNTS = GetParValStr("RANGEwithACCOUNTS")
If RANGEwithACCOUNTS = "" Then
pRANGEwithACCOUNTS = "M4:M30" '�������� ��� �������� ����� ��������
End If
Const pColumnAcc1 = 3 ' ������� ��� �������� ������ ���� - ���� ��������
Const pColumnAcc2 = 6 ' ������� ��� �������� ������ ���� - ���� ��������
Const pColumnOperation = 8 ' ������� ��� �������� ��������

Const pOperationStartBlock = "������ ����� - ����� ����� ��� �� �������� ������������ ������" ' ��������
Const pOperationOpBal = "�������� ������ �� ������ ��������� �������:" ' �������� ���������� ������
Const pOperationClBal = "��������� ������" ' �������� ��������� ������

'Const pOperationForContrCheck1 = "����������� �� ������ ��������� � ������������"
'Const pOperationForContrCheck2 = "�������� ������ � ��������������"
'Const pOperationForContrCheck3 = "������� �� ������������ ����� ������������� � �������������"
'Const pOperationForContrCheck4 = "����� ������������� �������������� ������ �������� ������� � ��������� �� ������������"
'Const pOperationForContrCheck5 = "��������� ������������� ��������� �� ����������� �������"
'Const pOperationForContrCheck6 = "������� �� ������������� ����� ������������� � ������������"
'Const pOperationForContrCheck7 = "������ ������� ������ ����� (��������, � ������� ������ ������������ �������������, �� ������ �������)"

Const pColumnSumForeign1 = 9 ' ������� � ������� �������� �����
Const pColumnSumForeign2 = 9
Const pColumnSum1 = 10
Const pColumnSum2 = 10
Const pColumnSumVAT1 = 11
Const pColumnSumVAT2 = 11

Const pColumnAnStart = 14 '������ � ��������� ������� � �����������
Dim pColumnAnEnd As Integer
pColumnAnEnd = 37
Const pColumnCurrency = 12
Const pColumnContractor = 13

Dim Act As Worksheet
Dim Matrix As Worksheet
Dim Lists As Worksheet

Set Act = ActiveWorkbook.ActiveSheet
Set Matrix = ActiveWorkbook.Worksheets(pSheetACC_AN)
Set Lists = ActiveWorkbook.Worksheets(pSheetLISTS)

Const pStartRow = 14
Dim pEndRow As Long
Dim shtAct As Worksheet
Set shtAct = ThisWorkbook.Worksheets("���_���")

pEndRow = GetLastRow(shtAct) '.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1


Dim i, j As Integer
Dim r As Range
Dim cellTest As Range
Dim strRangeOperContrCheck As String
strRangeOperContrCheck = GetParValStr("RangeOperContrCheck")
If strRangeOperContrCheck = "" Then
strRangeOperContrCheck = "'" & SHEET_CHECK_SETTIGS & "'!Z1:Z100"
End If

pColumnAnEnd = GetParValInt("ColumnAnEnd")
If pColumnAnEnd = 0 Then
pColumnAnEnd = 41
End If

Dim pErrors As Integer

pErrors = 0
Act.Range("H8").Value = Empty

Dim pResult, pResult1, pResult2 As Boolean

Dim pAcc As String
'Act.Unprotect PASSWORD

'���������� 22.09.2020-----
Application.Calculation = xlManual
'------

For i = pStartRow To pEndRow

    '���������� �������� ��� ������ 008 � 009
    ' ��������� ������ �������, � ������� �������� ���� ��������, ��������� ������� ����������
    If (Act.Cells(i, pColumnAcc1).Value <> Empty) And (Act.Cells(i, pColumnOperation).Value <> pOperationClBal) And Act.Cells(i, pColumnAcc1).Value <> "008" And Act.Cells(i, pColumnAcc1).Value <> "009" Then
       
        ' ���� ������� ����� - ������ �����
        If (Act.Cells(i, pColumnSum1).Value + Act.Cells(i, pColumnSum2).Value <> 0) Then
       
            ' �������� ��� �������� ���. ����
              
            If (Act.Cells(i, pColumnAcc2).Value = Empty) Then
                Act.Cells(i, pColumnAcc2).Interior.Pattern = xlHorizontal
                pErrors = pErrors + 1
            End If
            
            ' �������� ��� ��������� ��������
            
            If (Act.Cells(i, pColumnOperation).Value = Empty) Then
                Act.Cells(i, pColumnOperation).Interior.Pattern = xlHorizontal
                pErrors = pErrors + 1
            End If
            
            ' �������� �� ������
            ' ������ ������ ���� ���������
            
            If (Act.Cells(i, pColumnCurrency).Value = Empty) Then
                Act.Cells(i, pColumnCurrency).Interior.Pattern = xlHorizontal
                pErrors = pErrors + 1
            End If
            
            ' ���� ������ ���������� �� ������ ��������, �� ����� � ������ ������ ���� ���������
            
            If (Act.Cells(i, pColumnSum2).Value = 0) Or (Act.Cells(i, pColumnSumForeign2).Value = 0) And Act.Range(pAddressOfCurrency).Value <> "{RUB} ���������� �����" Then
                Act.Cells(i, pColumnSumForeign2).Interior.Pattern = xlHorizontal
                pErrors = pErrors + 1
            End If
            
            
            If (Act.Cells(i, pColumnCurrency).Value = Empty) Then
            
                Act.Cells(i, pColumnCurrency).Interior.Pattern = xlHorizontal
                pErrors = pErrors + 1
            
            Else
            
                If (Act.Cells(i, pColumnCurrency).Value <> Act.Range(pAddressOfCurrency).Value) Then
            
                                           
                        
                        If (Act.Cells(i, pColumnSum1).Value <> 0) And (Act.Cells(i, pColumnSumForeign1).Value = 0) Then
                            '17.09.20 �������� �21 �� �������� �������---------
                            If Not (Act.Cells(i, 22).Value = "{1200100} ������ � ���� �������� ������ �� ���. � �.�." Or Act.Cells(i, 22).Value = "{2200100} ������� � ���� �������� ������ �� ���. � �.�." Or Act.Cells(i, 22).Value = "{2200100} ������� � ���� �������� ������ �� ���. � �.�." Or Act.Cells(i, 22).Value = "{1200200} ������ � ���� �������� ������ �� �������������� � �������, ���������� � ����������� ������" Or Act.Cells(i, 22).Value = "{2200200} ������� � ���� �������� ������ �� �������������� � �������, ���������� � ����������� ������") Then
                            '---------
                                Act.Cells(i, pColumnSumForeign1).Interior.Pattern = xlHorizontal
                                pErrors = pErrors + 1
                            '---------
                            End If
                            '-------
                        End If
                        
                        
                        
                     
                        If (Act.Cells(i, pColumnSum1).Value = 0) And (Act.Cells(i, pColumnSumForeign1).Value <> 0) Then
                           Act.Cells(i, pColumnSumForeign1).Interior.Pattern = xlHorizontal
                           pErrors = pErrors + 1
                        End If
                    
                        'If (Act.Cells(i, pColumnSum2).Value = 0) And (Act.Cells(i, pColumnSumForeign2).Value > 0) Then
                        '    Act.Cells(i, pColumnSumForeign2).Interior.Pattern = xlHorizontal
                        '    pErrors = pErrors + 1
                        'End If
            
                 End If
                 
             End If
           ' ��������� ��������� ����������
           
            ' ������� ���� �������� ���.����
             If (Act.Cells(i, pColumnAcc2).Value <> Empty) Then
                   
                   pAcc = Act.Cells(i, pColumnAcc2).Value
                   
                   With Lists.Range(pRANGEwithACCOUNTS)
                        Set c = .Find(pAcc, LookIn:=xlValues, LookAt:=xlWhole) '�������� ��� ������� ����������
                        If Not c Is Nothing Then
                        ' ���� ����������� ���� �������� ������ ��������
                           Set r = Application.Range(strRangeOperContrCheck)
                           Set cellTest = r.Find(Act.Cells(i, pColumnOperation).Value)
                                If Not cellTest Is Nothing Then
                                
                                    If (Act.Cells(i, pColumnContractor).Value <> Empty) Then

                                        '28.01.2020 �������� ������
                                        'Act.Cells(i, pColumnContractor).Interior.Pattern = xlHorizontal
                                        'pErrors = pErrors + 1

                                    End If
                                    
                                Else

                                    If (Act.Cells(i, pColumnContractor).Value = Empty) Then
                                        '28.01.2020 �������� ������
                                        'Act.Cells(i, pColumnContractor).Interior.Pattern = xlHorizontal
                                        'pErrors = pErrors + 1

                                    End If

                                End If
                         
                          End If
                    
                   End With
                End If
            
             
           ' ��������� ���������
           
           For j = pColumnAnStart To pColumnAnEnd
              
              pAcc = Act.Cells(i, pColumnAcc1).Value
             
              pResult1 = CheckAnalytics(Matrix, pAcc, j)
              
              pAcc = Act.Cells(i, pColumnAcc2).Value
              
              pResult2 = CheckAnalytics(Matrix, pAcc, j)
              
              pResult = pResult1 Or pResult2
              
              If (Act.Cells(i, j).Value <> Empty) And (pResult = False) Then
                '19.05.2011 ��������� ���������� � ��������
                pResult1 = GetAccExceptions(pAcc, j)
                If pResult1 = False Then
                  Act.Cells(i, j).Interior.Pattern = xlHorizontal
                  pErrors = pErrors + 1
                End If
              
              End If
        
              If (Act.Cells(i, j).Value = Empty) And (pResult = True) Then
                '19.05.2011 ��������� ���������� � ��������
                pResult1 = GetAccExceptions(pAcc, j)
                         
        '07.01.17
        '���� ��������� 33 ������� �
                '������ 2 ������� � ������� C = "76" �� �������� �� ����������.
              ' If (j = 33 And Left(Act.Cells(i, 3), 2) = "76") Then pResult1 = True
                               
              '�����
                If j = 33 And Act.Cells(i, 8) = "�������� ������" Then
                    pResult1 = True
                End If
                               
                
                
                If pResult1 = False Then
                  Act.Cells(i, j).Interior.Pattern = xlHorizontal
                  pErrors = pErrors + 1
                 End If
             
              End If
        
           Next j
           
        '�������� ���������, ��� ������ �� ��������� �������, ����� ����� �� �������
        Else
        
            If (Act.Cells(i, pColumnOperation).Value <> pOperationStartBlock) Then
        
                ' �������� ��� �������� ���. ����
                  
                If (Act.Cells(i, pColumnAcc2).Value <> Empty) And (Act.Cells(i, pColumnOperation).Value <> pOperationOpBal) Then
                    Act.Cells(i, pColumnAcc2).Interior.Pattern = xlHorizontal
                    pErrors = pErrors + 1
                End If
                
                ' �������� ��� ��������� ��������
                
                If ((Act.Cells(i, pColumnOperation).Value <> Empty) And (Act.Cells(i, pColumnOperation).Value <> pOperationOpBal) And (Act.Cells(i, pColumnOperation).Value <> pOperationClBal)) Then
                    Act.Cells(i, pColumnOperation).Interior.Pattern = xlHorizontal
                    pErrors = pErrors + 1
                End If
                
                If (Act.Cells(i, pColumnSumForeign1).Value <> 0) Then
                    Act.Cells(i, pColumnSumForeign1).Interior.Pattern = xlHorizontal
                    pErrors = pErrors + 1
                End If
                
                'If (Act.Cells(i, pColumnSumForeign2).Value <> 0) Then
                '    Act.Cells(i, pColumnSumForeign2).Interior.Pattern = xlHorizontal
                '    pErrors = pErrors + 1
                'End If
                
                If (Act.Cells(i, pColumnSumVAT1).Value <> 0) Then
                    Act.Cells(i, pColumnSumVAT1).Interior.Pattern = xlHorizontal
                    pErrors = pErrors + 1
                End If
                
                If (Act.Cells(i, pColumnSumVAT2).Value <> 0) Then
                    Act.Cells(i, pColumnSumVAT2).Interior.Pattern = xlHorizontal
                    pErrors = pErrors + 1
                End If
                
                If ((Act.Cells(i, pColumnContractor).Value <> Empty) And (Act.Cells(i, pColumnOperation).Value <> pOperationOpBal) And (Act.Cells(i, pColumnOperation).Value <> pOperationClBal)) Then
                    Act.Cells(i, pColumnContractor).Interior.Pattern = xlHorizontal
                    pErrors = pErrors + 1
                End If
                
                For j = pColumnAnStart To pColumnAnEnd
                
                    If Act.Cells(i, j).Value <> Empty Then
                       Act.Cells(i, j).Interior.Pattern = xlHorizontal
                       pErrors = pErrors + 1
                    End If
                
                Next j
                
             End If
                
         End If
    
    End If

Next i
'�������������� �������� ���������
 pErrors = pErrors + CheckByFormulas()

'pErrors = pErrors + DopProv()

'����������---22.09.2020

Application.Calculation = xlAutomatic

'-----


'MsgBox ("�������� ���� ���������. ���������� " & pErrors & " ���������")

If pErrors > 0 Then
  Act.Range("H8").Value = "���������� " & pErrors & " ���������"
Else
'  Act.Range("H8").Value = Empty
  Act.Range("H8").Value = "�� �������"
End If
'---- �������� �����
Worksheets("���_���").Cells(1, 9) = ""

ActiveSheet.Protect Password:="tesla", DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering _
        :=True
 
 
 'Act.Protect PASSWORD:=PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
 '       , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
 '       AllowFormattingRows:=True, AllowInsertingColumns:=False, AllowInsertingRows _
 '       :=False, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables _
 '       :=True
 '  Act.EnableSelection = xlNoRestrictions

End Sub

