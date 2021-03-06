Attribute VB_Name = "myLib"
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long

'
' The state of the Ctrl key
Function CtrlKey() As Boolean
    CtrlKey = (GetAsyncKeyState(vbKeyControl) And &H8000)
End Function
'
' The state of either Shift keys
Function ShiftKey() As Boolean
    ShiftKey = (GetAsyncKeyState(vbKeyShift) And &H8000)
End Function
'
' The state of the Alt key
Function AltKey() As Boolean
    AltKey = (GetAsyncKeyState(vbKeyMenu) And &H8000)
End Function
'


Function Replace_symbols(ByVal Txt As String) As String
    st$ = "~!@/\#$%^:?&*=|`;"""
    For f_i% = 1 To Len(st$)
        Txt = Replace(Txt, Mid(st$, f_i, 1), "_")
        Txt = Replace(Txt, Chr(10), "_")
    Next
    Replace_symbols = Txt
End Function

Sub VBA_Start()
With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
    '.DisplayPageBreaks = False
    .DisplayAlerts = False
End With
End Sub

Sub VBA_End()
With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
    .DisplayStatusBar = True
    .DisplayAlerts = True
End With
End Sub

Sub CreateSh(cr_sh As String)
For Each Sh In ThisWorkbook.Worksheets
    If Sh.name = cr_sh Then
    chek_name = 1
    End If
Next Sh
    If chek_name <> 1 Then
    Set Sh = Worksheets.Add()
    Sh.name = cr_sh
    End If
End Sub

Function OpenFile(ByRef patch As String, nm_sh As String, Optional stMessage As Boolean = True) As String
Dim result$

    If Dir(patch) = "" Then
        If stMessage Then MsgBox ("FileNotFound")
    Else
        Workbooks.Open Filename:=patch, Notify:=False
        
        result = ActiveWorkbook.name
        Sheets(nm_sh).Select
        ActiveSheet.AutoFilterMode = False
    End If

OpenFile = result
End Function

Sub openFileCSV(ByRef patch As String)
Dim result$
If Dir(patch) = "" Then
    MsgBox ("FileNotFound")
Else
    Workbooks.OpenText Filename:=patch, _
        Origin:=65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, _
        Comma:=False, Space:=False, Other:=False, TrailingMinusNumbers:=True
End If
End Sub

Function GetQuartal(num_month&) As String
Dim result As String
result = Empty
    Select Case num_month
        Case 1, 2, 3
        result = "1Q"
        Case 4, 5, 6
        result = "2Q"
        Case 7, 8, 9
        result = "3Q"
        Case 10, 11, 12
        result = "4Q"
    End Select
GetQuartal = result
End Function

Function GetMonth_form_00(num_month As Integer) As String
Dim result As String
result = Empty

    If num_month < 10 Then
        result = "0" & num_month
    Else
        result = num_month
    End If

GetMonth_form_00 = result
End Function

Function GetPatchHistTR(nmBrand As String, ThisYear As Integer, VarYear As Integer, ThisMonth As Integer, VarMonth As Integer) As String
Dim result As String
result = Empty
month00 = GetMonth_form_00(VarMonth)

If VarMonth = 12 Then
    result = "p:\DPP\Business development\Book commercial\" & nmBrand & "\Top Russia Total " & VarYear & " " & nmBrand & ".xlsm"
    ElseIf VarMonth & VarYear = ThisMonth & ThisYear Then
        result = "p:\DPP\Business development\Book commercial\" & nmBrand & "\Top Russia Total " & VarYear & " " & nmBrand & ".xlsm"
        Else
            result = "p:\DPP\Business development\Book commercial\" & nmBrand & "\" & VarYear & "\History " & VarYear & "\Top Russia Total " & VarYear & "." & month00 & " " & nmBrand & ".xlsm"
End If

GetPatchHistTR = result
End Function

Function GetLastRow() As Long
Dim result As Long
result = Empty
    With ActiveWorkbook.ActiveSheet.UsedRange
    result = .row + .Rows.Count - 1
    End With
GetLastRow = result
End Function

Function GetLastColumn() As Long
Dim result As Long
result = Empty
    
    With ActiveWorkbook.ActiveSheet.UsedRange
    result = .Column + .Columns.Count - 1
    End With

GetLastColumn = result
End Function

Function GetClntType(in_data$, i&) As String
Dim result
Dim ar_type_clients(1 To 4, 1 To 12)
Dim f_sl&

    ar_type_clients(1, 1) = "�����"
    ar_type_clients(2, 1) = "salon"
    ar_type_clients(3, 1) = "salon"
    ar_type_clients(4, 1) = "single"

    ar_type_clients(1, 2) = "���� �������"
    ar_type_clients(2, 2) = "chain_salons"
    ar_type_clients(3, 2) = "salon"
    ar_type_clients(4, 2) = "chain"

    ar_type_clients(1, 3) = "�/�"
    ar_type_clients(2, 3) = "hdres"
    ar_type_clients(3, 3) = "salon"
    ar_type_clients(4, 3) = "single"

    ar_type_clients(1, 4) = "���� ���������"
    ar_type_clients(2, 4) = "chain_shops"
    ar_type_clients(3, 4) = "shop"
    ar_type_clients(4, 4) = "chain"

    ar_type_clients(1, 5) = "�������"
    ar_type_clients(2, 5) = "shop"
    ar_type_clients(3, 5) = "shop"
    ar_type_clients(4, 5) = "single"

    ar_type_clients(1, 6) = "�����-���."
    ar_type_clients(2, 6) = "salon"
    ar_type_clients(3, 6) = "salon"
    ar_type_clients(4, 6) = "single"

    ar_type_clients(1, 7) = "(�����)"
    ar_type_clients(2, 7) = "other"
    ar_type_clients(3, 7) = "other"
    ar_type_clients(4, 7) = "single"

    ar_type_clients(1, 8) = "�����"
    ar_type_clients(2, 8) = "school"
    ar_type_clients(3, 8) = "school"
    ar_type_clients(4, 8) = "single"

    ar_type_clients(1, 9) = "������"
    ar_type_clients(2, 9) = "other"
    ar_type_clients(3, 9) = "other"
    ar_type_clients(4, 9) = "single"

    ar_type_clients(1, 10) = "����-���"
    ar_type_clients(2, 10) = "nails_bar"
    ar_type_clients(3, 10) = "nails"
    ar_type_clients(4, 10) = "single"

    ar_type_clients(1, 11) = "���� ����-�����"
    ar_type_clients(2, 11) = "chain_nails"
    ar_type_clients(3, 11) = "nails"
    ar_type_clients(4, 11) = "chain"

    ar_type_clients(1, 12) = "e-commerce"
    ar_type_clients(2, 12) = "e-commerce"
    ar_type_clients(3, 12) = "e-commerce"
    ar_type_clients(4, 12) = "single"

For f_sl = 1 To 12
    
If StrComp(ar_type_clients(1, f_sl), LCase(in_data), vbTextCompare) Then
    
    result = ar_type_clients(i, f_sl)
    Exit For
    Else
    result = Empty
End If
Next f_sl

GetClntType = result
End Function

Function GetMregWhitoutBrand(in_data$) As String
Dim result$
Dim ar_nmBran()
Select Case in_data
        Case Empty: result = Empty
        Case Else: If Len(in_data) > 3 And Mid(in_data, 3, 1) = " " Then result = Right(in_data, Len(in_data) - 3) Else result = in_data
End Select
GetMregWhitoutBrand = result
End Function


Function GetMregExt(in_data_mreg$, in_data_reg$) As String
Dim result$
Dim extPos&
textPos = 0
If LCase(in_data_mreg) = LCase("Moscou GR") Then
    textPos = InStr(in_data_reg, "MSK")
    textPos = InStr(in_data_reg, "Moscou") + textPos
        If textPos > 0 Then
        result = "Moscou"
        Else
        result = "GR"
        End If
Else
    result = in_data_mreg
End If
GetMregExt = result
End Function

Function GetMregLat(in_data_mreg As String) As String
Dim result$
Dim f_mr&
Dim ar_nmMregEN(), ar_nmMregLT()
result = Empty
ar_nmMregEN = Array("MOSCOW", "GR", "NORTHWEST", "CENTER", "VOLGA", "SOUTH", "URAL", "SIBERIA", "FAR EAST")
ar_nmMregLT = Array("Moscou", "GR", "Nord-Ouest", "Centre", "Volga-Centre", "Sud", "Oural", "Siberie", "EO")

For f_mr = 0 To UBound(ar_nmMregLT)
    If ar_nmMregLT(f_mr) = in_data_mreg Then
        result = ar_nmMregEN(f_mr)
        Exit For
    End If
Next f_mr

GetMregLat = result

End Function

Function GetClientName(in_clnt_nm$, in_clnt_addres$, in_city$) As String
Dim result$

result = Trim(Replace_symbols(Left(in_clnt_nm, 30) & ". " & Left(in_clnt_addres, 50) & " " & Left(in_city, 50)))

GetClientName = result
End Function

Function GetMonthNumeric(in_data$) As Integer
Dim result&
Dim f_m&, num_month&

ar_nm_month_qnc_rus = Array("������", "�������", "����", "������", "���", "����", "����", "������", "��������", "�������", "������", "�������")
result = 1
    For f_m = 0 To 11
    If ar_nm_month_qnc_rus(f_m) = in_data Then
    result = f_m + 1
    Exit For
    End If
    Next f_m

GetMonthNumeric = result
End Function
'----------------------------------------
Function GetNameMonthEN(in_data%) As String
Dim result$
Dim f_m&, num_month&
ar_month_eng = Array(0, "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
result = Empty
If IsNumeric(in_data) Then
    Select Case in_data
        Case Is > 0, Is < 13
        result = ar_month_eng(in_data)
        Case Else
        result = Empty
    End Select
End If
GetNameMonthEN = result
End Function


Function GetMonthEng(month$) As String
Dim result$
Dim f_m&

ar_month_rus = Array("������", "�������", "����", "������", "���", "����", "����", "������", "��������", "�������", "������", "�������")
ar_month_eng = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    For f_m = 0 To 11
        If ar_month_rus(f_m) = month Then
        result = ar_month_eng(f_m)
        Exit For
        End If
    Next f_m
    
GetMonthEng = result
End Function


Function GetYearType(ThisYear As Integer, in_data As Integer, i&) As Variant
Dim result1&, result2$, result3$
    
    If in_data >= 2008 And in_data <= ThisYear Then result1 = in_data Else result1 = 2008
    
    
        Select Case result1
            Case ThisYear: result2 = "TY": result3 = "CNQ_TY"
            Case ThisYear - 1: result2 = "PY": result3 = "CNQ_PY"
            Case Else: result2 = "PPY": result3 = "PPY"
        End Select

Select Case i
Case 1: GetYearType = result1
Case 2: GetYearType = result2
Case 3: GetYearType = result3
Case Else
    GetYearType = Empty
End Select
End Function


Function GetMag(in_min_price As Long, in_max_price As Long, in_place As Long, mag_type As String) As Variant

Dim result As Variant
Dim mag_avg_price&
        
If IsNumeric(in_min_price) And IsNumeric(in_max_price) Then
    mag_avg_price = Application.WorksheetFunction.Average(in_min_price, in_max_price)
Else
    mag_avg_price = in_min_price + in_max_price
End If

Select Case LCase(mag_type)
    Case "avg_price"
        result = mag_avg_price

    Case "hair"
        Select Case mag_avg_price
            Case 100 To 799
                result = "D"
            Case 800 To 1199
                result = "C"
            Case 1200 To 2000
                result = "B"
            Case Is > 2000
                result = "A"
            Case Else
                result = Empty
        End Select
    
    Case "nail"
        Select Case mag_avg_price
            Case 10 To 319
                result = "D"
            Case 320 To 479
                result = "C"
            Case 480 To 799
                result = "B"
            Case Is > 800
                result = "A"
            Case Else
                result = Empty
        End Select
    
    Case "skin"
        Select Case mag_avg_price
            Case 100 To 799
                result = "D"
            Case 800 To 1199
                result = "C"
            Case 1200 To 2000
                result = "B"
            Case Is > 2000
                result = "A"
            Case Else
                result = Empty
        End Select

    Case "place"
        If IsNumeric(in_place) Then
        in_place = Round(in_place, 0)
        End If
        Select Case in_place
            Case 1 To 2
            result = "1"
            Case 3 To 4
            result = "2"
            Case Is > 4
            result = "3"
            Case Else
            result = Empty
        End Select

    End Select
    
GetMag = result
End Function

Function GetTypeBusiness(in_brand$) As String
Dim result$
Select Case in_brand
        Case "LP", "MX", "KR", "RD"
        result = "Hair"
        Case "ES"
        result = "Nails"
        Case "DE", "CR"
        result = "Skin"
End Select
GetTypeBusiness = result
End Function

Function GetTypeDN(in_data As Integer) As String
Dim result$

Select Case in_data
    Case 1
        result = "Active"
    Case 0
        result = "Closed"
End Select
GetTypeDN = result
End Function

Function GetRoundNum(in_data As Variant) As Double
Dim result&
If IsNumeric(in_data) And Len(in_data) > 0 Then
    result = Round(in_data, 0)
Else
    result = 0
End If
GetRoundNum = result
End Function

Function GetNum2num0(in_data As Variant) As Double
Dim result&
If Len(in_data) > 0 And IsNumeric(in_data) Then
result = in_data
Else
result = 0
End If
GetNum2num0 = result
End Function

Function GetNum2numNull(in_data) As Variant
Dim result As Variant
If Len(in_data) > 0 And in_data <> 0 Then
result = in_data
Else
result = Empty
End If
GetNum2numNull = result
End Function

Function GetNmChainTop(inNmChain$, inCdChain&, inNmTypeClnt$) As String
Dim result$
If Left(inCdChain, 2) = "92" And GetClntType(inNmTypeClnt, 4) = "chain" Then
result = inNmChain
Else
result = Empty
End If
GetNmChainTop = result
End Function


Public Function GetLTM(ByRef wks As Worksheet, ByVal in_row As Long, ThisMonth As Integer, typeFN As String) As Variant
Dim result$
Dim f_a&, f_avg&, sum_CA_LTM&, AVG_CA_LTM&, frqOrder&
Dim MinVal!, MaxVal!
Dim ML13 As Variant
Dim val As Variant
Dim str_clm As Integer, end_clm As Integer

ar_DataMonthPRTN = Array(0, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77)
ar_nmAVG_Order = Array(2.5, 5, 10, 15, 20, 25, 30, 40, 50, 60, 80, 100, 100000)
sum_CA_LTM = 0
frqOrder = 0
str_clm = ThisMonth + 12
end_clm = ThisMonth + 1
For f_a = str_clm To end_clm Step -1
    val = wks.Cells(in_row, ar_DataMonthPRTN(f_a))
    If IsNumeric(val) And val > 0 Then
        frqOrder = frqOrder + 1
        sum_CA_LTM = sum_CA_LTM + val
    End If
Next f_a
ML13 = wks.Cells(in_row, ar_DataMonthPRTN(end_clm - 1))

AVG_CA_LTM = Round(sum_CA_LTM / 12 / 1000, 1)

Select Case typeFN
Case "avg_ca"
    If sum_CA_LTM <> 0 Then
    result = AVG_CA_LTM
    Else
    result = Empty
    End If

Case "frqOrders"
    result = frqOrder & "\12"
    
Case "type_avg_ca"
    MinVal = 0
    MaxVal = 0
    
        
        Select Case AVG_CA_LTM
        Case 0
            result = "0"
        Case Is >= 70
            result = ">70"
        Case Is < 70
            For f_avg = 0 To UBound(ar_nmAVG_Order)
                MaxVal = ar_nmAVG_Order(f_avg)
                If AVG_CA_LTM <= MaxVal And AVG_CA_LTM > MinVal Then result = "'" & MinVal & "-" & MaxVal: Exit For
                
                MinVal = MaxVal
            Next f_avg
        Case Else
        result = Empty
        End Select
Case "LostLTM"
    result = IIf(sum_CA_LTM = 0 And ML13 <> 0, 1, 0)
End Select
GetLTM = result
End Function

Function GetVectoreEV(in_data As Double) As String
Dim result$

If IsNumeric(in_data) Then
    Select Case in_data
    Case Is > 0
        result = "+"
    Case Is < 0
        result = "-"
    Case Else
        result = Empty
    End Select
Else
result = Null
End If

GetVectoreEV = result
End Function


Function GetMonthlyCA&(in_row&, in_month&, in_thisMonth&, in_typeY$, in_typeVal$, in_type_period$)
Dim result&, val&
Dim typeF$
Dim clm_PY_LOR_VAL%, clm_TY_LOR_VAL%, clm_PY_PRTN_VAL%, clm_TY_PRTN_VAL%
Dim ar_Matrix(1 To 2, 1 To 2)

val = Empty
typeF = in_typeY & "_" & in_typeVal
Select Case typeF
    Case "PY_LOR": clm = 106
    Case "TY_LOR": clm = 93
    Case "PY_PRTN": clm = 79
    Case "TY_PRTN": clm = 66
    Case Else
        Exit Function
End Select

Select Case in_type_period
    Case "Total"
        in_thisMonth = 12
    Case "YTD"
        in_thisMonth = in_thisMonth
End Select

Select Case in_month
    Case Is <= in_thisMonth
        val = GetNum2num0(Cells(in_row, clm + in_month - 1))
        If val = 0 Then val = Empty Else val = val / 1000
    Case Else
        val = Empty
End Select

result = val
GetMonthlyCA = result
End Function


Function GetCA_Cnq(in_monthQnc&)

        Case cd_ThisYear - 1
        fst_order_LOR_PY = Cells(f_i, clm_PYper_LOR_VAL + cd_month_qnc - 1) / 1000
        fst_order_PRTN_PY = Cells(f_i, clm_PYper_PRTN_VAL + cd_month_qnc - 1) / 1000
        
            If cd_month_qnc = cd_ThisMonth Then
            fst_order_LOR_M_PY = Cells(f_i, clm_PYper_LOR_VAL + cd_month_qnc - 1) / 1000
            End If
                            
        Case cd_ThisYear
        fst_order_LOR_TY = Cells(f_i, clm_TYper_LOR_VAL + cd_month_qnc - 1) / 1000
        fst_order_PRTN_TY = Cells(f_i, clm_TYper_PRTN_VAL + cd_month_qnc - 1) / 1000

            If cd_month_qnc = cd_ThisMonth Then
            fst_order_LOR_M_TY = Cells(f_i, clm_TYper_LOR_VAL + cd_month_qnc - 1) / 1000
            End If

        End Select

End Function


Function avgCA(in_data&, in_month&) As String
Dim result&

If Not IsEmpty(in_data) And IsNumeric(in_data) Then
result = in_data / in_month
Else
result = Empty

End If
avgCA = result
End Function


Function GetSREP_type(nm_Srep$, nm_FLSM$) As String
Dim result$
If Trim(LCase(nm_Srep)) = Trim(LCase(nm_FLSM)) Then
    result = "FLSMasSREP"
    ElseIf InStr(1, LCase(nm_Srep), "�����", vbTextCompare) <> 0 Then
        result = "vacancy"
        ElseIf InStr(1, LCase(nm_Srep), "vakan", vbTextCompare) <> 0 Then
            result = "vacancy"
            Else
            result = "active"
End If
GetSREP_type = result

End Function

    
Sub IsOpenTRtoClsd()
Dim wbBook As Workbook
For Each wbBook In Workbooks
    If wbBook.name <> ThisWorkbook.name Then
        If Windows(wbBook.name).Visible Then
            If wbBook.name Like "Top Russia*" Then wbBook.Close: Exit For
        End If
    End If
Next wbBook
End Sub


Sub CloseNoMotherBook(ByVal ShIn As String)
    If ActiveWorkbook.name <> ShIn Then

    ActiveWindow.Close
    Application.DisplayAlerts = False
        End If
End Sub

 
Function GetDateEmpty(in_date As Variant) As Variant
Dim result As Variant
If IsDate(in_date) Then
    result = in_date
Else
    result = Empty
End If
ifDateTheDate = result
End Function
 

Function GetLast4quartal(in_date As Variant, in_ActiveM%, in_ActiveY%) As String
Dim result$
Dim ActDate As Date, in_dateN As Date
Dim count_month As Integer


If Not IsDate(in_date) Then
    in_dateN = CDate(in_date)
    Else
    in_dateN = in_date
End If

If IsNumeric(in_ActiveY) And IsNumeric(in_ActiveM) And IsDate(in_date) Then
ActDate = DateSerial(in_ActiveY, in_ActiveM, 1)
count_qurtal = DateDiff("q", in_dateN, ActDate)
End If
Select Case count_qurtal
Case 1: result = "-1Q"
Case 2: result = "-2Q"
Case 3: result = "-3Q"
    Case 4: result = "-4Q"
    Case Else: result = "OLD"
End Select
GetLast4quartal = result
End Function

 
Sub sheetActivateCleer(in_sh$)
Sheets(in_sh).Select
ActiveSheet.UsedRange.Cells.ClearContents
End Sub

Function GetHash(ByVal Txt$) As String
    Dim oUTF8, oMD5, abyt, i&, k&, hi&, lo&, chHi$, chLo$
    Set oUTF8 = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    abyt = oMD5.ComputeHash_2(oUTF8.GetBytes_4(Txt$))
    For i = 1 To LenB(abyt)
        k = AscB(MidB(abyt, i, 1))
        lo = k Mod 16: hi = (k - lo) / 16
        If hi > 9 Then chHi = Chr(Asc("a") + hi - 10) Else chHi = Chr(Asc("0") + hi)
        If lo > 9 Then chLo = Chr(Asc("a") + lo - 10) Else chLo = Chr(Asc("0") + lo)
        GetHash = GetHash & chHi & chLo
    Next
    Set oUTF8 = Nothing: Set oMD5 = Nothing
End Function

 
Function GetStatus(in_data As String) As String
Dim result$
Select Case Trim(LCase(in_data))
    Case "�������", "�������", "partner": result = "partner"
    Case "�������", "loreal", "l'oreal", "�'������", "��� �'������": result = "loreal"
    Case "ancore", "ancor", "�����", "inter", "���������": result = "inter"
    Case Else: result = in_data
End Select
GetStatus = result
End Function

Function fixError(in_data As Variant) As Variant
Dim result As Variant
If IsError(in_data) Then
result = Empty
Else
result = in_data
End If
fixError = result
End Function

Function selectFile() As String
nameOfFile = ""
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .InitialFileName = "*.*"
    .title = "Select a file"
    .Show
    If .SelectedItems.Count = 1 Then nameOfFile = .SelectedItems(1)
End With
selectFile = nameOfFile
End Function

Function GetCol(n As Integer, text As String) As Integer
    result = 0
    For i = 1 To GetLastColumn()
        If Cells(n, i) = text Then
            result = i
            Exit For
        End If
    Next i
    GetCol = result
End Function

Function GetRow(n As Integer, text As String) As Integer
    result = 0
    For i = 1 To GetLastRow()
        If Cells(i, n) = text Then
            result = i
            Exit For
        End If
    Next i
    GetRow = result
End Function

Public Function GetExtension(Filepath As String)
    Dim FilenameParts() As String
    FilenameParts = VBA.Split(Filepath, ".")
    GetExtension = FilenameParts(UBound(FilenameParts))
End Function

Public Function DeleteFile(Filepath As String)
    If FileExists(Filepath) Then
        SetAttr Filepath, vbNormal
        Kill Filepath
    End If
End Function

Public Function FullPath(RelativePath As String) As String
    FullPath = ThisWorkbook.Path & Application.PathSeparator & VBA.Replace$(RelativePath, "/", Application.PathSeparator)
End Function

Public Function GetFilename(Filepath As String) As String
    Dim FilepathParts() As String
    FilepathParts = VBA.Split(Filepath, Application.PathSeparator)
    GetFilename = FilepathParts(UBound(FilepathParts))
End Function

Public Function RemoveExtension(Filename As String) As String
    Dim FilenameParts() As String
    FilenameParts = VBA.Split(Filename, ".")
    If UBound(FilenameParts) > LBound(FilenameParts) Then
        ReDim Preserve FilenameParts(UBound(FilenameParts) - 1)
    End If
    RemoveExtension = VBA.Join(FilenameParts, ".")
End Function

Public Function FileExists(Filepath As String) As Boolean
    FileExists = VBA.Len(VBA.Dir(Filepath)) <> 0
End Function

Public Function getNumInThrousend(ByVal in_data As Double) As Variant
    Dim result As Variant
    If IsNumeric(in_data) And in_data <> 0 Then
        result = in_data / 1000
    Else
        result = Null
    End If
    getNumInThrousend = result
End Function

Public Function getUniversCode(brand As String, row As Long, cd_Univers As Variant) As Variant
Dim result As Variant
    If Len(cd_Univers) <> 9 Then
        result = brand & row
    Else
        result = cd_Univers
    End If
getUniversCode = result
End Function

Function getResponse(ByVal text As String, request As String, key As String, secret_key As String) As String

    Dim result As String
    Dim objHTTP As Object
    
    result = ""
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    With objHTTP
    Select Case request
        Case "address"
            URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/address"
            send_txt = Chr(34) & "query" & Chr(34) & ": " & Chr(34) & text & Chr(34) & ", " & Chr(34) & "Count" & Chr(34) & ": 1"
        
        Case "fias"
            URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/address"
            send_txt = Chr(34) & "query" & Chr(34) & ": " & Chr(34) & text & Chr(34)
    End Select
        .Open "POST", URL, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Authorization", "Token " & key
        '.setProxy 2, "128.114.0.21:8080", ""
        .send ("{" & send_txt & "}")
        
    End With
    
    result = objHTTP.responseText
    result = Replace(result, "[", "")
    result = Replace(result, "]", "")
    
    If result = "{""suggestions"":}" Then result = "{""suggestions"":null}"
    
    getResponse = result
    
End Function


Sub letHead(sts_head As Boolean, clmn As Long, name As String)
If sts_head Then
    Cells(1, clmn) = name
End If

End Sub

Sub OpenInChromeOrDefaultBrowser(ByVal url As String)

    Dim wholeContent As String
    chrome = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

    If Dir(chrome, vbNormal) <> "" Then
        '' 32 bit Chrome
        wholeContent = """" & chrome & """ -url " & url
    Else
        '' 64 bit Chrome
        chrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        wholeContent = """" & chrome & """ -url " & url
    End If

    If Dir(chrome, vbNormal) <> "" Then
        '' Chrome is found, execute
        Shell wholeContent
    Else
        '' Chrome was not found, open url using the default program
        Dim Sh As Object
        Set Sh = CreateObject("WScript.Shell")
        Sh.Run url
    End If
End Sub

                                      
Sub CreateFolderWithSubfolders(ByVal PatchCreateFolder$)
 
   If Len(Dir(PatchCreateFolder$, vbDirectory)) = 0 Then
       SHCreateDirectoryEx Application.hwnd, PatchCreateFolder$, ByVal 0&
   End If
End Sub

Option Compare Database
Option Explicit

Public Function fxVerifyEMail(strEMail As String) As Boolean

Dim strEval As String
Dim strTemp As String

strTemp = strEMail

If InStr(1, strTemp, "@") Then
    
    strEval = Left(strEMail, (InStr(1, strEMail, "@")) - 1)
    
    If Len(strEval) < 1 Then
        fxVerifyEMail = False
        Exit Function
    End If
    
    strEval = Mid(strTemp, (InStr(1, strTemp, "@")) + 1, Len(strTemp))

    If (Len(strEval) - 1) < 1 Then
        fxVerifyEMail = False
        Exit Function
    End If

    If InStr(1, strEval, ".") Then

        strEval = Mid(strEval, InStr(1, strEval, ".") + 1, Len(strTemp))
        strEval = Right(strEval, 3)
    
        If (Len(strEval)) < 2 Then
            fxVerifyEMail = False
            Exit Function
        End If

    Else
        
        fxVerifyEMail = False
        Exit Function
    
    End If


Else
        
        fxVerifyEMail = False
        Exit Function
    
End If
    
fxVerifyEMail = True


Exit Function

ErrorTrap:

    fxVerifyEMail = False

End Function

Public Function fxVerifyURL(ByVal strURL As String) As Boolean

Dim Response As Long

Dim FX As Double
Dim Url As String
Dim IEDir As String

On Error GoTo ErrorTrap

    IEDir = "C:\Program Files\Internet Explorer"
    
    ' the file exists, returns "" (empty string) if the file doesn't exist.
    
    If Dir(IEDir & "\iexplore.exe") <> "" Then
        FX = Shell(IEDir & "\IEXPLORE.EXE " & strURL, 1)
    Else
        MsgBox "Explorer was not found in the expected folder!", vbOKOnly, "Error!"
    End If

    fxVerifyURL = True

ErrorTrap:

    fxVerifyURL = False

End Function

Public Function NumericValidation(ByVal ValueChecked As String, ByVal KeyAscii As Integer)

Dim bYesDecPoint As Boolean

If KeyAscii = 8 Then              '   Backspace
    NumericValidation = KeyAscii
    Exit Function
End If
If KeyAscii = 47 Then              '   Slash
    NumericValidation = 0
    Exit Function
End If
If InStr(ValueChecked, ".") > 0 Then bYesDecPoint = True
If Not (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = 0
    Beep
End If
'If KeyAscii = 46 And bYesDecPoint Then  '   Second Dec. Point
'    KeyAscii = 0
'    Beep
'End If
NumericValidation = KeyAscii

End Function


