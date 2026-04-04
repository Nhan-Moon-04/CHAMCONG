Attribute VB_Name = "Module2"
Option Explicit

'==================================================================================================
' MODULE 2: XU LY CHAM CONG VA TINH LUONG (PHAN 2)
' - Doc du lieu tu sheet 1 (CSV).
' - Tham chieu cac sheet cau hinh: Loai Ca, Bang Luong, Lich Lam, Tang Ca, Ngay OFF. (Tien an lay tu Loai Ca, khong co sheet rieng)
' - Tao sheet "Cham Cong" chi tiet voi 16 cot theo yeu cau.
' - Tao sheet luong chi tiet cho tung nhan vien ("BL_[MaNV]").
' - Khong tao sheet "Tong Hop".
'==================================================================================================

Private Const NGAY_TINH_LUONG As Long = 26

'==================================================================================================
' MAIN SUBROUTINE
'==================================================================================================

Sub XuLyChamCongVaTaoBangLuong()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' 0. Xoa sheet Tong Hop cu neu co
    Dim wsToDelete As Worksheet
    On Error Resume Next
    Set wsToDelete = wb.Sheets("Tong Hop")
    If Not wsToDelete Is Nothing Then wsToDelete.Delete
    Set wsToDelete = Nothing
    On Error GoTo ErrorHandler
    
    ' 1. DOC DU LIEU CSV
    Dim csvWs As Worksheet
    Set csvWs = wb.Sheets(1)
    Dim lastRowCsv As Long
    lastRowCsv = csvWs.Cells(csvWs.Rows.Count, "A").End(xlUp).Row
    If lastRowCsv < 2 Then
        MsgBox "Sheet 1 (CSV) khong co du lieu!", vbExclamation
        GoTo CleanUp
    End If
    
    ' 2. LOAD TAT CA CAC DICTIONARY CAU HINH
    Dim dictLoaiCa As Object, dictLuong As Object, dictOFF As Object, dictNghiViec As Object
    Dim dictLichLam As Object, dictTangCa As Object, dictDonPhep As Object
    Dim salaryDivisor As Double
    
    Set dictLoaiCa = LoadData_LoaiCa(wb)
    Set dictLuong = LoadData_BangLuong(wb)
    Set dictOFF = LoadData_NgayOFF(wb)
    Set dictNghiViec = LoadData_NghiViec(wb)
    Set dictLichLam = LoadData_LichLam(wb)
    Set dictTangCa = LoadData_TangCa(wb)
    Set dictDonPhep = LoadData_DonPhep(wb)
    salaryDivisor = LoadSalaryDivisor(wb)
    
    ' 3. GOM DU LIEU CHAM CONG
    Dim dictChamCong As Object
    Set dictChamCong = AggregateChamCongData(csvWs, lastRowCsv, dictLichLam)
    Dim dictDateRange As Object
    Set dictDateRange = BuildDateRangeDict(dictChamCong, dictDonPhep)
    Set dictChamCong = ExpandDateRangeRows(dictChamCong, dictDateRange, wb)
    Set dictChamCong = AddScheduleRowsToData(dictChamCong, dictLichLam, wb, dictDateRange)
    
    ' 4. TAO SHEET "CHAM CONG"
    On Error Resume Next
    wb.Sheets("Cham Cong").Delete
    On Error GoTo ErrorHandler
    Dim wsChamCong As Worksheet
    Set wsChamCong = wb.Sheets.Add(After:=csvWs)
    wsChamCong.Name = "Cham Cong"
    Set dictChamCong = AddLeaveRowsToData(dictChamCong, dictDonPhep, wb, dictDateRange)
    Call CreateChamCongSheet(wsChamCong, dictChamCong, dictLoaiCa, dictLuong, dictOFF, dictNghiViec, dictLichLam, dictTangCa, dictDonPhep, salaryDivisor)
    
    ' 5. TAO SHEET LUONG CHI TIET CHO TUNG NHAN VIEN (DA BI VO HIEU HOA)
    ' Call CreateIndividualPayslips(wb, wsChamCong)

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Err.Number = 0 Then
        MsgBox "Hoan thanh xu ly cham cong va tao bang luong chi tiet!", vbInformation
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Da co loi xay ra: " & vbCrLf & Err.Description, vbCritical, "Loi Xu Ly"
    GoTo CleanUp
End Sub

Private Sub ApplyEmployeeBanding(ws As Worksheet, startRow As Long, endRow As Long, useColor As Boolean)
    If endRow < startRow Then Exit Sub
    ws.Range("A" & startRow & ":P" & endRow).Interior.ColorIndex = xlNone
End Sub


'==================================================================================================
' PART 1: LOAD DATA FUNCTIONS
'==================================================================================================

Private Function LoadData_LoaiCa(wb As Workbook) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: On Error Resume Next
    Set ws = wb.Sheets("Loai Ca")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    
    Dim r As Long: r = 2
    Do While ws.Cells(r, 1).Value <> ""
        Dim key As String: key = UCase(Trim(CStr(ws.Cells(r, 1).Value)))
        If Not dict.Exists(key) Then
            ' Array: TenCa, SoGio, HeSoLuong, TienAn, GioBatDau, GioKetThuc
            dict.Add key, Array( _
                CStr(ws.Cells(r, 2).Value), _
                CDbl(ws.Cells(r, 3).Value), _
                CDbl(ws.Cells(r, 6).Value), _
                CDbl(ws.Cells(r, 7).Value), _
                ws.Cells(r, 4).Value, _
                ws.Cells(r, 5).Value)
        End If
        r = r + 1
    Loop
    Set LoadData_LoaiCa = dict
End Function

Private Function LoadData_BangLuong(wb As Workbook) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: On Error Resume Next
    Set ws = wb.Sheets("Bang Luong")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    
    Dim r As Long: r = 2
    Do While ws.Cells(r, 2).Value <> ""
        Dim key As Long: key = CLng(ws.Cells(r, 2).Value)
        If Not dict.Exists(key) And IsNumeric(ws.Cells(r, 4).Value) Then
            ' Value: LuongThang
            dict.Add key, CDbl(ws.Cells(r, 4).Value)
        End If
        r = r + 1
    Loop
    Set LoadData_BangLuong = dict
End Function

Private Function LoadSalaryDivisor(wb As Workbook) As Double
    LoadSalaryDivisor = NGAY_TINH_LUONG

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Bang Luong")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    If IsNumeric(ws.Range("F2").Value) Then
        If CDbl(ws.Range("F2").Value) > 0 Then
            LoadSalaryDivisor = CDbl(ws.Range("F2").Value)
        End If
    End If
End Function

Private Function LoadData_NgayOFF(wb As Workbook) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: On Error Resume Next
    Set ws = wb.Sheets("Ngay OFF")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    
    Dim r As Long: r = 2
    Do While ws.Cells(r, 2).Value <> ""
        If UCase(Trim(CStr(ws.Cells(r, 6).Value))) = "X" Then
            Dim key As String: key = Format(CDate(ws.Cells(r, 2).Value), "yyyy-mm-dd")
            If Not dict.Exists(key) Then
                ' Value: Loai (Le/CN)
                dict.Add key, UCase(Trim(CStr(ws.Cells(r, 4).Value)))
            End If
        End If
        r = r + 1
    Loop
    Set LoadData_NgayOFF = dict
End Function

Private Function LoadData_NghiViec(wb As Workbook) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    ' No sheet "Nghi Viec" in new design, but keep function for compatibility
    Set LoadData_NghiViec = dict
End Function

Private Function LoadData_LichLam(wb As Workbook) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Lich Lam")
    If ws Is Nothing Then Set ws = wb.Sheets("Ngay Nghi") ' Backward compatibility
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Dim c As Long, r As Long
    For c = 3 To lastCol
        Dim nvID As Long: nvID = GetNvIdFromHeader(ws, c)
        If nvID > 0 Then
            For r = 3 To lastRow
                Dim val As String: val = UCase(Trim(CStr(ws.Cells(r, c).Value)))
                Dim dayVal As Variant: dayVal = ParseDateValue(ws.Cells(r, 2).Value)
                If val <> "" And IsDate(dayVal) Then
                    Dim key As String: key = nvID & "|" & Format(dayVal, "yyyy-mm-dd")
                    If Not dict.Exists(key) Then dict.Add key, val
                End If
            Next r
        End If
    Next c
    Set LoadData_LichLam = dict
End Function

Private Function LoadData_TangCa(wb As Workbook) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: On Error Resume Next
    Set ws = wb.Sheets("Tang Ca")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Dim c As Long, r As Long
    For c = 3 To lastCol
        Dim nvID As Long: nvID = GetNvIdFromHeader(ws, c)
        If nvID > 0 Then
            For r = 3 To lastRow
                Dim dayVal As Variant: dayVal = ParseDateValue(ws.Cells(r, 2).Value)
                If IsNumeric(ws.Cells(r, c).Value) And CDbl(ws.Cells(r, c).Value) > 0 And IsDate(dayVal) Then
                    Dim key As String: key = nvID & "|" & Format(dayVal, "yyyy-mm-dd")
                    If Not dict.Exists(key) Then dict.Add key, CDbl(ws.Cells(r, c).Value)
                End If
            Next r
        End If
    Next c
    Set LoadData_TangCa = dict
End Function

'==================================================================================================
' DON NGHI / PHEP
'==================================================================================================
Private Function LoadData_DonPhep(wb As Workbook) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Don Phep")
    If ws Is Nothing Then Set ws = wb.Sheets("Don Nghiphep")
    If ws Is Nothing Then Set ws = wb.Sheets("Don Nghiphep")
    If ws Is Nothing Then Set ws = wb.Sheets("Don Nghi")
    On Error GoTo 0
    If ws Is Nothing Then
        Set LoadData_DonPhep = dict
        Exit Function
    End If
    Dim r As Long: r = 2
    Do While ws.Cells(r, 1).Value <> ""
        Dim dayVal As Variant: dayVal = ParseDateValue(ws.Cells(r, 2).Value)
        If IsNumeric(ws.Cells(r, 1).Value) And IsDate(dayVal) Then
            Dim nvID As Long: nvID = CLng(ws.Cells(r, 1).Value)
            Dim ngayVal As Date: ngayVal = CDate(dayVal)
            Dim loai As String: loai = UCase(Trim(CStr(ws.Cells(r, 3).Value)))
            If loai <> "" Then
                Dim key As String: key = nvID & "|" & Format(ngayVal, "yyyy-mm-dd")
                If Not dict.Exists(key) Then dict.Add key, loai
            End If
        End If
        r = r + 1
    Loop
    Set LoadData_DonPhep = dict
End Function

' Extract Ma NV from either the numeric helper row (row 2) or the header text (row 1: "ID 2 - PHUNG VAN GHET")
Private Function GetNvIdFromHeader(ws As Worksheet, colIdx As Long) As Long
    If IsNumeric(ws.Cells(2, colIdx).Value) Then
        GetNvIdFromHeader = CLng(ws.Cells(2, colIdx).Value)
        Exit Function
    End If
    Dim raw As String: raw = CStr(ws.Cells(1, colIdx).Value)
    Dim digits As String: digits = ""
    Dim i As Long
    For i = 1 To Len(raw)
        Dim ch As String: ch = Mid$(raw, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf digits <> "" Then
            Exit For ' stop once we've passed the first number block
        End If
    Next i
    If digits <> "" Then GetNvIdFromHeader = CLng(digits)
End Function

' Parse date whether stored as real date or text dd/mm/yyyy; returns 0 if cannot parse
Private Function ParseDateValue(val As Variant) As Variant
    If IsDate(val) Then
        ParseDateValue = CDate(val)
        Exit Function
    End If
    Dim s As String: s = Trim(CStr(val))
    If s = "" Then Exit Function
    Dim sep As String
    If InStr(s, "/") > 0 Then
        sep = "/"
    ElseIf InStr(s, "-") > 0 Then
        sep = "-"
    ElseIf InStr(s, ".") > 0 Then
        sep = "."
    End If
    If sep <> "" Then
        Dim parts() As String: parts = Split(s, sep)
        If UBound(parts) = 2 Then
            Dim d As Long, m As Long, y As Long
            d = Val(parts(0)): m = Val(parts(1)): y = Val(parts(2))
            If d >= 1 And d <= 31 And m >= 1 And m <= 12 And y > 1900 Then
                ParseDateValue = DateSerial(y, m, d)
            End If
        End If
    End If
End Function

Private Function GetBaseShiftCode(rawCode As String) As String
    Dim s As String: s = UCase$(Trim$(rawCode))
    If InStr(1, s, ",", vbTextCompare) > 0 Then
        s = Left$(s, InStr(1, s, ",", vbTextCompare) - 1)
    End If
    GetBaseShiftCode = Trim$(s)
End Function

Private Sub ParseShiftCodeAndLeave(rawCode As String, ByRef baseCode As String, ByRef leaveHours As Double)
    Dim s As String: s = Trim$(rawCode)
    baseCode = GetBaseShiftCode(s)
    leaveHours = 0

    If InStr(1, s, ",", vbTextCompare) > 0 Then
        Dim arr() As String: arr = Split(s, ",")
        If UBound(arr) >= 1 Then
            Dim leaveText As String: leaveText = Trim$(arr(1))
            If IsNumeric(leaveText) Then
                leaveHours = CDbl(leaveText)
                If leaveHours < 0 Then leaveHours = 0
                If leaveHours > 7 Then leaveHours = 7
            End If
        End If
    End If
End Sub

Private Function GetRequiredHoursByCode(baseCode As String, dictLoaiCa As Object) As Double
    Select Case UCase$(Trim$(baseCode))
        Case "X", "XVP", "CN", "L", "P"
            GetRequiredHoursByCode = 8
        Case "S", "C"
            GetRequiredHoursByCode = 4
        Case "D", "CND"
            GetRequiredHoursByCode = 8
        Case "N", "OFF", ""
            GetRequiredHoursByCode = 0
        Case Else
            If Not dictLoaiCa Is Nothing Then
                If dictLoaiCa.Exists(baseCode) Then
                    Dim caInfo As Variant: caInfo = dictLoaiCa(baseCode)
                    GetRequiredHoursByCode = CDbl(caInfo(1))
                End If
            End If
    End Select
End Function

Private Function GetDefaultOvertimeHours(baseCode As String, hasScan As Boolean, gioThuc As Double) As Double
    If Not hasScan Then Exit Function

    Select Case UCase$(Trim$(baseCode))
        Case "D", "CND"
            GetDefaultOvertimeHours = 4
        Case "X", "XVP"
            If gioThuc >= 11 Then GetDefaultOvertimeHours = 3
    End Select
End Function

' Lookup employee name from CSV sheet 1; cache for repeated queries
Private Function LookupEmployeeName(wb As Workbook, nvID As Long, cache As Object) As String
    If cache.Exists(nvID) Then
        LookupEmployeeName = cache(nvID)
        Exit Function
    End If
    Dim csvWs As Worksheet
    On Error Resume Next
    Set csvWs = wb.Sheets(1)
    On Error GoTo 0
    If csvWs Is Nothing Then Exit Function
    Dim lastRow As Long: lastRow = csvWs.Cells(csvWs.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        Dim rawID As String: rawID = Replace(CStr(csvWs.Cells(i, 1).Value), "'", "")
        If IsNumeric(rawID) Then
            If CLng(rawID) = nvID Then
                cache.Add nvID, CStr(csvWs.Cells(i, 2).Value)
                LookupEmployeeName = cache(nvID)
                Exit Function
            End If
        End If
    Next i
End Function

'==================================================================================================
' PART 2: AGGREGATE CHAM CONG DATA
'==================================================================================================

Private Function AggregateChamCongData(csvWs As Worksheet, lastRow As Long, dictLichLam As Object) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    
    ' === PASS 1: Gom tat ca quet theo NGAY CALENDAR (khong dich ngay) ===
    Dim i As Long
    For i = 2 To lastRow
        Dim nvID As Long, shiftDate As Date, timeStamp As Date
        
        On Error Resume Next
        nvID = CLng(Replace(CStr(csvWs.Cells(i, 1).Value), "'", ""))
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextRow
        End If
        On Error GoTo 0
        
        timeStamp = CDate(csvWs.Cells(i, 4).Value)
        shiftDate = DateValue(timeStamp) ' Luon dung ngay calendar
        
        Dim key As String: key = nvID & "|" & Format(shiftDate, "yyyy-mm-dd")
        
        If Not dict.Exists(key) Then
            ' Array: MaNV, HoTen, Ngay, GioVao (datetime), GioRa (datetime), CoQuetNgay, CoQuetDem
            dict.Add key, Array(nvID, CStr(csvWs.Cells(i, 2).Value), shiftDate, timeStamp, timeStamp, False, False)
        End If
        
        Dim arr As Variant: arr = dict(key)
        If timeStamp < arr(3) Then arr(3) = timeStamp
        If timeStamp > arr(4) Then arr(4) = timeStamp
        
        Dim h As Long: h = Hour(timeStamp)
        If h >= 6 And h < 18 Then arr(5) = True
        If h >= 18 Or h < 6 Then arr(6) = True
        
        dict(key) = arr
NextRow:
    Next i
    
    ' === PASS 2: Xu ly ca dem - merge quet 0h-6h sang ngay X+1 vao ngay X neu co ca dem ===
    If Not dictLichLam Is Nothing And dictLichLam.Count > 0 Then
        ' Tim cac ngay co ca dem (D, CND) trong lich lam
        Dim keysToMerge As Object: Set keysToMerge = CreateObject("Scripting.Dictionary")
        Dim k As Variant
        For Each k In dict.keys
            Dim parts() As String: parts = Split(k, "|")
            Dim nv As Long: nv = CLng(parts(0))
            Dim dt As Date: dt = CDate(parts(1))
            
            ' Check: ngay nay co quet nao truoc 6h khong?
            Dim rowArr As Variant: rowArr = dict(k)
            Dim gioVaoCheck As Date: gioVaoCheck = rowArr(3)
            If Hour(gioVaoCheck) < 6 Then
                ' Ngay hom truoc (dt-1) co ca dem trong lich lam khong?
                Dim prevKey As String: prevKey = nv & "|" & Format(dt - 1, "yyyy-mm-dd")
                Dim lichKey As String: lichKey = nv & "|" & Format(dt - 1, "yyyy-mm-dd")
                If dictLichLam.Exists(lichKey) Then
                    Dim lichVal As String: lichVal = GetBaseShiftCode(CStr(dictLichLam(lichKey)))
                    If lichVal = "D" Or lichVal = "CND" Then
                        ' Danh dau can merge: quet ngay nay (truoc 6h) -> gop vao ngay hom truoc
                        If Not keysToMerge.Exists(k) Then keysToMerge.Add k, prevKey
                    End If
                End If
            End If
        Next k
        
        ' Thuc hien merge
        Dim mergeKey As Variant
        For Each mergeKey In keysToMerge.keys
            Dim targetKey As String: targetKey = keysToMerge(mergeKey)
            Dim srcArr As Variant: srcArr = dict(mergeKey)
            
            If dict.Exists(targetKey) Then
                ' Merge: cap nhat gio ra cua ngay hom truoc
                Dim tgtArr As Variant: tgtArr = dict(targetKey)
                If srcArr(4) > tgtArr(4) Then tgtArr(4) = srcArr(4)
                If srcArr(3) < tgtArr(3) Then tgtArr(3) = srcArr(3)
                tgtArr(6) = True ' Co quet dem
                dict(targetKey) = tgtArr
            Else
                ' Ngay hom truoc chua co data -> tao moi
                dict.Add targetKey, Array(CLng(Split(targetKey, "|")(0)), srcArr(1), CDate(Split(targetKey, "|")(1)), srcArr(3), srcArr(4), False, True)
            End If
            
            ' Xoa ngay hien tai neu chi co quet truoc 6h (toan bo la ca dem hom truoc)
            ' Chi xoa neu KHONG co quet nao >= 6h trong ngay do
            If srcArr(5) = False Then
                ' Khong co quet ngay (chi co quet dem) -> xoa
                dict.Remove mergeKey
            Else
                ' Co ca quet ngay lan quet dem -> chi cap nhat lai gio vao (bo phan truoc 6h)
                ' Can re-scan CSV de tim gio som nhat >= 6h -> phuc tap, de nguyen
            End If
        Next mergeKey
    End If
    
    ' Sort keys
    Dim sortedKeys As Variant: sortedKeys = SortDictionaryKeys(dict)
    
    Dim sortedDict As Object: Set sortedDict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(sortedKeys)
        sortedDict.Add sortedKeys(i), dict(sortedKeys(i))
    Next i
    
    Set AggregateChamCongData = sortedDict
End Function

Private Function MergeDictWithNewKeys(dictBase As Object, dictNew As Object, wb As Workbook, Optional dictRange As Object = Nothing, Optional addName As Boolean = True) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In dictBase.keys
        dict.Add k, dictBase(k)
    Next k
    If Not dictNew Is Nothing Then
        Dim nameCache As Object: Set nameCache = CreateObject("Scripting.Dictionary")
        Dim key As Variant
        For Each key In dictNew.keys
            If Not dict.Exists(key) Then
                Dim parts() As String: parts = Split(key, "|")
                If UBound(parts) = 1 Then
                    Dim nvID As Long: nvID = CLng(parts(0))
                    Dim ngayLam As Date: ngayLam = CDate(parts(1))
                    If ShouldIncludeDate(nvID, ngayLam, dictRange) Then
                        Dim hoTen As String
                        If addName Then hoTen = LookupEmployeeName(wb, nvID, nameCache)
                        dict.Add key, Array(nvID, hoTen, ngayLam, 0, 0, False, False)
                    End If
                End If
            End If
        Next key
    End If
    Dim sortedKeys As Variant: sortedKeys = SortDictionaryKeys(dict)
    Dim sortedDict As Object: Set sortedDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To UBound(sortedKeys)
        sortedDict.Add sortedKeys(i), dict(sortedKeys(i))
    Next i
    Set MergeDictWithNewKeys = sortedDict
End Function

' Add scheduled days (Lich Lam) so absences on working days are visible (within date range)
Private Function AddScheduleRowsToData(dictData As Object, dictLichLam As Object, wb As Workbook, dictRange As Object) As Object
    If dictLichLam Is Nothing Or dictLichLam.Count = 0 Then
        Set AddScheduleRowsToData = dictData
        Exit Function
    End If
    Set AddScheduleRowsToData = MergeDictWithNewKeys(dictData, dictLichLam, wb, dictRange, False)
End Function

' Merge leave-only days into aggregated data so days without scans still appear (within date range)
Private Function AddLeaveRowsToData(dictData As Object, dictDonPhep As Object, wb As Workbook, dictRange As Object) As Object
    If dictDonPhep Is Nothing Or dictDonPhep.Count = 0 Then
        Set AddLeaveRowsToData = dictData
        Exit Function
    End If
    Set AddLeaveRowsToData = MergeDictWithNewKeys(dictData, dictDonPhep, wb, dictRange)
End Function

Private Function SortDictionaryKeys(dict As Object) As Variant
    Dim keys As Variant: keys = dict.keys
    Dim i As Long, j As Long, temp As Variant
    
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            Dim key1 As Variant: key1 = Split(keys(i), "|")
            Dim key2 As Variant: key2 = Split(keys(j), "|")
            
            Dim shouldSwap As Boolean: shouldSwap = False
            If CLng(key1(0)) > CLng(key2(0)) Then
                shouldSwap = True
            ElseIf CLng(key1(0)) = CLng(key2(0)) Then
                If CDate(key1(1)) > CDate(key2(1)) Then
                    shouldSwap = True
                End If
            End If
            
            If shouldSwap Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    SortDictionaryKeys = keys
End Function

' Build min/max date per NV from existing logs and leave requests (dictData + dictDonPhep)
Private Function BuildDateRangeDict(dictData As Object, dictDonPhep As Object) As Object
    Dim dictRange As Object: Set dictRange = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    Dim globalMin As Date: globalMin = 0
    Dim globalMax As Date: globalMax = 0
    For Each key In dictData.keys
        Dim parts() As String: parts = Split(key, "|")
        If UBound(parts) = 1 Then
            Dim nvID As Long: nvID = CLng(parts(0))
            Dim ngayVal As Date: ngayVal = CDate(parts(1))
            Call UpdateRange(dictRange, nvID, ngayVal)
            If globalMin = 0 Or ngayVal < globalMin Then globalMin = ngayVal
            If globalMax = 0 Or ngayVal > globalMax Then globalMax = ngayVal
        End If
    Next key
    If Not dictDonPhep Is Nothing Then
        For Each key In dictDonPhep.keys
            Dim p() As String: p = Split(key, "|")
            If UBound(p) = 1 Then
                Dim nvID2 As Long: nvID2 = CLng(p(0))
                Dim ngay2 As Date: ngay2 = CDate(p(1))
                Call UpdateRange(dictRange, nvID2, ngay2)
                If globalMin = 0 Or ngay2 < globalMin Then globalMin = ngay2
                If globalMax = 0 Or ngay2 > globalMax Then globalMax = ngay2
            End If
        Next key
    End If
    ' Ensure every NV range spans global min-max so all NV fill đầy đủ ngày
    If globalMin <> 0 And globalMax <> 0 Then
        Dim nvIDKey As Variant
        For Each nvIDKey In dictRange.keys
            dictRange(nvIDKey) = Array(globalMin, globalMax)
        Next nvIDKey
    End If
    Set BuildDateRangeDict = dictRange
End Function

Private Sub UpdateRange(dictRange As Object, nvID As Long, ngayVal As Date)
    If Not dictRange.Exists(nvID) Then
        dictRange.Add nvID, Array(ngayVal, ngayVal)
    Else
        Dim arr As Variant: arr = dictRange(nvID)
        If ngayVal < arr(0) Then arr(0) = ngayVal
        If ngayVal > arr(1) Then arr(1) = ngayVal
        dictRange(nvID) = arr
    End If
End Sub

Private Function ShouldIncludeDate(nvID As Long, ngayVal As Date, dictRange As Object) As Boolean
    If dictRange Is Nothing Then
        ShouldIncludeDate = True
        Exit Function
    End If
    If Not dictRange.Exists(nvID) Then
        ' If no range, skip to avoid flooding
        ShouldIncludeDate = False
    Else
        Dim arr As Variant: arr = dictRange(nvID)
        ShouldIncludeDate = (ngayVal >= arr(0) And ngayVal <= arr(1))
    End If
End Function

' Fill missing dates within each NV's min-max range so Cham Cong hiển đầy đủ ngày
Private Function ExpandDateRangeRows(dictData As Object, dictRange As Object, wb As Workbook) As Object
    If dictRange Is Nothing Or dictRange.Count = 0 Then
        Set ExpandDateRangeRows = dictData
        Exit Function
    End If
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In dictData.keys
        dict.Add k, dictData(k)
    Next k
    Dim nameCache As Object: Set nameCache = CreateObject("Scripting.Dictionary")
    Dim nvID As Variant
    For Each nvID In dictRange.keys
        Dim arr As Variant: arr = dictRange(nvID)
        Dim d As Date
        For d = arr(0) To arr(1)
            Dim key As String: key = nvID & "|" & Format(d, "yyyy-mm-dd")
            If Not dict.Exists(key) Then
                Dim hoTen As String: hoTen = LookupEmployeeName(wb, CLng(nvID), nameCache)
                dict.Add key, Array(CLng(nvID), hoTen, d, 0, 0, False, False)
            End If
        Next d
    Next nvID
    Dim sortedKeys As Variant: sortedKeys = SortDictionaryKeys(dict)
    Dim sortedDict As Object: Set sortedDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To UBound(sortedKeys)
        sortedDict.Add sortedKeys(i), dict(sortedKeys(i))
    Next i
    Set ExpandDateRangeRows = sortedDict
End Function


'==================================================================================================
' PART 3: CREATE "CHAM CONG" SHEET
'==================================================================================================

Private Sub CreateChamCongSheet(ws As Worksheet, dictData As Object, dictLoaiCa As Object, dictLuong As Object, dictOFF As Object, dictNghiViec As Object, dictLichLam As Object, dictTangCa As Object, dictDonPhep As Object, salaryDivisor As Double)
    
    ' --- HEADERS ---
    Dim headers As Variant
    headers = Array("STT", "Ma NV", "Ho Ten", "Ngay", "Ten Ca", "Gio Vao", "Gio Ra", "Gio Thuc", "Chenh Lech", "Tang Ca", "Tong Gio", "Status Code", "So Gio", "Luong Theo Ngay", "Ghi Chu", "Tien An Theo Ngay")
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    ' --- WRITE DATA ---
    Dim r As Long: r = 2
    Dim stt As Long: stt = 1
    Dim prevMaNV As Long: prevMaNV = -1
    Dim prevHoTen As String
    Dim bandAlt As Boolean
    ' Summary row sits above each employee block
    Dim groupSummaryRow As Long: groupSummaryRow = r
    Dim groupDataStartRow As Long: groupDataStartRow = r + 1
    r = groupDataStartRow
    Dim tongLuongDot1 As Double, tongLuongDot2 As Double, tongAnDot1 As Double, tongAnDot2 As Double
    Dim cntNoPerm As Long, cntBoCa As Long
    
    Dim key As Variant
    For Each key In dictData.keys
        Dim arr As Variant: arr = dictData(key)
        Dim maNV As Long: maNV = arr(0)
        Dim hoTen As String: hoTen = arr(1)
        Dim ngayLam As Date: ngayLam = arr(2)
        
        ' --- Write summary row for previous employee ---
        If maNV <> prevMaNV And prevMaNV <> -1 Then
            If r - 1 >= groupDataStartRow Then Call ApplyEmployeeBanding(ws, groupDataStartRow, r - 1, bandAlt)
            bandAlt = Not bandAlt
            Call WriteSummaryRow(ws, groupSummaryRow, prevMaNV, prevHoTen, tongLuongDot1 + tongLuongDot2, tongAnDot1, tongAnDot2, cntNoPerm, cntBoCa)
            ' Move pointers for next employee: summary row first, data starts one row below
            groupSummaryRow = r
            groupDataStartRow = r + 1
            r = groupDataStartRow
            tongLuongDot1 = 0: tongLuongDot2 = 0: tongAnDot1 = 0: tongAnDot2 = 0
            cntNoPerm = 0
            cntBoCa = 0
        End If
        prevMaNV = maNV
        prevHoTen = hoTen
        
        ' --- Get data for current row ---
        Dim gioVao As Date: gioVao = arr(3)
        Dim gioRa As Date: gioRa = arr(4)
        Dim coQuetNgay As Boolean: coQuetNgay = arr(5)
        Dim coQuetDem As Boolean: coQuetDem = arr(6)
        Dim hasScan As Boolean: hasScan = (coQuetNgay Or coQuetDem Or gioVao <> 0 Or gioRa <> 0)
        
        Dim luongThang As Double: luongThang = 0
        Dim luongNgayCoBan As Double: luongNgayCoBan = 0
        If dictLuong.Exists(maNV) Then luongThang = dictLuong(maNV)
        If salaryDivisor <= 0 Then salaryDivisor = NGAY_TINH_LUONG
        If luongThang > 0 Then luongNgayCoBan = luongThang / salaryDivisor
        
        Dim gioThuc As Double: gioThuc = (gioRa - gioVao) * 24
        If gioThuc < 0.1 Then gioThuc = 0
        
        Dim gioTangCa As Double: gioTangCa = 0 ' Tang ca theo dung NV + ngay
        Dim lookupKey As String: lookupKey = maNV & "|" & Format(ngayLam, "yyyy-mm-dd")
        If dictTangCa.Exists(lookupKey) Then gioTangCa = dictTangCa(lookupKey)
        
        ' --- Determine Status Code ---
        Dim statusCode As String, ghiChu As String, ca As String
        Dim isSunday As Boolean: isSunday = (Weekday(ngayLam, vbSunday) = 1)
        Dim boldRow As Boolean: boldRow = False ' True = co lich lam nhung khong quet the
        
        Dim offKey As String: offKey = Format(ngayLam, "yyyy-mm-dd")
        If isSunday Then
            If dictOFF.Exists(offKey) And dictOFF(offKey) = "LE" Then
                statusCode = "L"
                ghiChu = "Ngay le"
            ElseIf dictLichLam.Exists(lookupKey) Then
                statusCode = dictLichLam(lookupKey)
                If hasScan Then
                    ghiChu = ""
                ElseIf dictDonPhep.Exists(lookupKey) Then
                    statusCode = dictDonPhep(lookupKey)
                    ghiChu = ""
                Else
                    boldRow = True
                    ghiChu = ""
                End If
            Else
                statusCode = "OFF"
                ghiChu = ""
            End If
        Else
            If hasScan Then
                If dictLichLam.Exists(lookupKey) Then
                    statusCode = dictLichLam(lookupKey)
                    ghiChu = ""
                ElseIf dictOFF.Exists(offKey) Then
                    If dictOFF(offKey) = "LE" Then
                        statusCode = "L"
                        ghiChu = "Ngay le"
                    Else
                        statusCode = "X"
                        ghiChu = "Ngay OFF, co log"
                    End If
                Else
                    statusCode = "X"
                    ghiChu = "Khong co lich"
                End If
            Else
                If dictDonPhep.Exists(lookupKey) Then
                    statusCode = dictDonPhep(lookupKey)
                    ghiChu = ""
                ElseIf dictLichLam.Exists(lookupKey) Then
                    statusCode = dictLichLam(lookupKey)
                    boldRow = True
                    ghiChu = ""
                ElseIf dictOFF.Exists(offKey) Then
                    If dictOFF(offKey) = "LE" Then
                        statusCode = "L"
                        ghiChu = "Ngay le"
                    Else
                        statusCode = "OFF"
                        ghiChu = ""
                    End If
                Else
                    statusCode = "N"
                    ghiChu = ""
                End If
            End If
        End If
        
        Dim noteIssue As String: noteIssue = ""
        Dim isBoCa As Boolean: isBoCa = False
        If Not hasScan Then
            If dictLichLam.Exists(lookupKey) Then
                noteIssue = "Bo ca"
                isBoCa = True
            ElseIf Not dictOFF.Exists(offKey) Then
                noteIssue = "Khong quet the"
            End If
        ElseIf gioVao = gioRa Then
            If Hour(gioVao) < 12 Then
                noteIssue = "Quen checkout"
            Else
                noteIssue = "Quen checkin"
            End If
        End If

        Dim statusBase As String, leaveHours As Double
        Call ParseShiftCodeAndLeave(statusCode, statusBase, leaveHours)

        ' --- Build ghi chu from status code (clean labels) ---
        If ghiChu = "" Then
            Select Case UCase(statusBase)
                Case "P": ghiChu = "Nghi phep"
                Case "S": ghiChu = "Nghi sang"
                Case "C": ghiChu = "Nghi chieu"
                Case "N": ghiChu = "Nghi khong phep"
                Case "OFF": ghiChu = "OFF"
                Case "L": ghiChu = "Ngay le"
                Case "X": ghiChu = ""
                Case "XVP": ghiChu = ""
                Case "CN": ghiChu = ""
                Case "D": ghiChu = ""
                Case "CND": ghiChu = ""
                Case Else: ghiChu = ""
            End Select
        End If

        If noteIssue <> "" Then
            If ghiChu = "" Then
                ghiChu = noteIssue
            Else
                ghiChu = ghiChu & "; " & noteIssue
            End If
        End If

        If leaveHours > 0 Then
            If ghiChu = "" Then
                ghiChu = "Nghi " & leaveHours & " gio"
            Else
                ghiChu = ghiChu & "; Nghi " & leaveHours & " gio"
            End If
        End If

        ' --- Calculate based on Status Code ---
        Dim soGio As Double: soGio = 0
        Dim luongTheoNgay As Double, tienAn As Double: tienAn = 0
        Dim heSo As Double: heSo = 0
        Dim caNames As Collection: Set caNames = New Collection
        If dictLoaiCa.Exists(statusBase) Then
            Dim caInfo As Variant: caInfo = dictLoaiCa(statusBase)
            soGio = CDbl(caInfo(1))
            heSo = CDbl(caInfo(2))
            tienAn = CDbl(caInfo(3))
            caNames.Add CStr(caInfo(0))
        End If

        Dim requiredHours As Double: requiredHours = GetRequiredHoursByCode(statusBase, dictLoaiCa)
        If requiredHours = 0 Then requiredHours = soGio

        Dim defaultOT As Double: defaultOT = GetDefaultOvertimeHours(statusBase, hasScan, gioThuc)
        If gioTangCa <= 0 Then gioTangCa = defaultOT

        Dim payableHours As Double: payableHours = requiredHours + gioTangCa - leaveHours
        If payableHours < 0 Then payableHours = 0

        If luongThang > 0 Then
            luongTheoNgay = (luongThang / salaryDivisor / 8) * payableHours
        Else
            luongTheoNgay = 0
        End If

        Dim chenhLech As Double: chenhLech = gioThuc - payableHours
        Dim tongGio As Double: tongGio = gioThuc + gioTangCa - leaveHours
        If tongGio < 0 Then tongGio = 0

        ' --- Counters ---
        If statusBase = "N" Or statusBase = "VANG" Then cntNoPerm = cntNoPerm + 1
        If isBoCa Then cntBoCa = cntBoCa + 1
        
        ' --- Accumulate totals ---
        If Day(ngayLam) <= 15 Then
            tongLuongDot1 = tongLuongDot1 + luongTheoNgay
            tongAnDot1 = tongAnDot1 + tienAn
        Else
            tongLuongDot2 = tongLuongDot2 + luongTheoNgay
            tongAnDot2 = tongAnDot2 + tienAn
        End If
        
        ' --- Write row data ---
        ws.Cells(r, 1).Value = stt
        ws.Cells(r, 2).Value = maNV
        ws.Cells(r, 3).Value = hoTen
        ws.Cells(r, 4).Value = ngayLam
        If caNames.Count > 0 Then
            Dim nameArr() As String
            ReDim nameArr(0 To caNames.Count - 1)
            Dim idx As Long: idx = 0
            Dim nm As Variant
            For Each nm In caNames
                nameArr(idx) = nm
                idx = idx + 1
            Next nm
            ca = Join(nameArr, " + ")
        Else
            ca = statusCode ' fallback to code when not defined in Loai Ca
        End If
        ws.Cells(r, 5).Value = ca
        ws.Cells(r, 6).Value = gioVao
        ws.Cells(r, 7).Value = gioRa
        ws.Cells(r, 8).Value = gioThuc
        ws.Cells(r, 9).Value = chenhLech
        ws.Cells(r, 10).Value = gioTangCa
        ws.Cells(r, 11).Value = tongGio
        ws.Cells(r, 12).Value = statusCode
        ws.Cells(r, 13).Value = payableHours
        ws.Cells(r, 14).Value = luongTheoNgay
        ws.Cells(r, 15).Value = ghiChu
        ws.Cells(r, 16).Value = tienAn
        
        ' Bold row if has schedule but no scan (quen cham cong)
        If boldRow Then
            ws.Range("A" & r & ":P" & r).Font.Bold = True
        End If
        
        stt = stt + 1
        r = r + 1
    Next key
    
    ' --- Write summary for the very last employee ---
    If prevMaNV <> -1 Then
        If r - 1 >= groupDataStartRow Then Call ApplyEmployeeBanding(ws, groupDataStartRow, r - 1, bandAlt)
        Call WriteSummaryRow(ws, groupSummaryRow, prevMaNV, prevHoTen, tongLuongDot1 + tongLuongDot2, tongAnDot1, tongAnDot2, cntNoPerm, cntBoCa)
    End If
    
    ' --- Final Formatting ---
    Call FormatChamCongSheet(ws, r - 1)
End Sub

Private Sub WriteSummaryRow(ws As Worksheet, r As Long, maNV As Long, hoTen As String, tongLuong As Double, an1 As Double, an2 As Double, cntNoPerm As Long, cntBoCa As Long)
    ws.Range("Q" & r & ":R" & (r + 5)).Interior.ColorIndex = xlNone
    ws.Range("Q" & r & ":R" & (r + 5)).Borders.LineStyle = xlContinuous
    ws.Range("Q" & r & ":R" & (r + 5)).Font.Bold = False
    
    ws.Cells(r, 17).Value = ">> TONG KET: " & hoTen
    ws.Cells(r, 17).Font.Bold = True
    ws.Cells(r, 18).Font.Bold = True
    
    ws.Cells(r + 1, 17).Value = "Tong Luong:"
    ws.Cells(r + 1, 18).Value = tongLuong
    ws.Cells(r + 2, 17).Value = "Tien An D1:"
    ws.Cells(r + 2, 18).Value = an1
    ws.Cells(r + 3, 17).Value = "Tien An D2:"
    ws.Cells(r + 3, 18).Value = an2
    ws.Cells(r + 4, 17).Value = "Bo ca:"
    ws.Cells(r + 4, 18).Value = cntBoCa
    
    If cntNoPerm > 0 Then
        ws.Cells(r + 5, 17).Value = "Nghi khong phep:"
        ws.Cells(r + 5, 18).Value = cntNoPerm
        ws.Cells(r + 5, 18).Font.Color = RGB(255, 0, 0)
    End If
    
End Sub

Private Sub FormatChamCongSheet(ws As Worksheet, lastRow As Long)
    With ws.Range("A1:P1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(79, 129, 189)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("D2:D" & lastRow).NumberFormat = "dd/mm/yyyy"
    ws.Range("F2:G" & lastRow).NumberFormat = "hh:mm"
    ws.Range("H2:K" & lastRow).NumberFormat = "0.0"
    ws.Range("M2:M" & lastRow).NumberFormat = "0.0"
    ws.Range("N2:N" & lastRow & ", P2:P" & lastRow).NumberFormat = "#,##0"
    ws.Range("R2:R" & lastRow).NumberFormat = "#,##0"
    
    ws.Columns.AutoFit
    ws.Columns("C").ColumnWidth = 25
    ws.Columns("O").ColumnWidth = 20
    ws.Columns("Q:R").ColumnWidth = 18
    
    ' Thin borders for all data
    With ws.Range("A1:P" & lastRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").AutoFilter

    ' === ROW COLORING + EMPLOYEE SEPARATOR + CHENH LECH HIGHLIGHT ===
    Dim r As Long
    Dim prevMaNV As Long: prevMaNV = 0
    
    For r = 2 To lastRow
        Dim rowRg As Range: Set rowRg = ws.Range("A" & r & ":P" & r)
        rowRg.Interior.ColorIndex = xlColorIndexNone
        
        Dim currentMaNV As Long: currentMaNV = 0
        If IsNumeric(ws.Cells(r, 2).Value) And ws.Cells(r, 2).Value <> "" Then
            currentMaNV = CLng(ws.Cells(r, 2).Value)
        End If
        
        ' --- Thick border between different employees ---
        If currentMaNV <> prevMaNV And prevMaNV <> 0 And currentMaNV <> 0 Then
            ' Draw thick top border on current row = separator line
            With ws.Range("A" & r & ":P" & r).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .Color = RGB(0, 0, 0)
            End With
        End If
        If currentMaNV <> 0 Then prevMaNV = currentMaNV
        
        ' --- Row color by status ---
        Dim sc As String: sc = GetBaseShiftCode(CStr(ws.Cells(r, 12).Value))
        Dim soGioVal As Double: soGioVal = Val(ws.Cells(r, 13).Value)
        Dim isSunday As Boolean: isSunday = False
        If IsDate(ws.Cells(r, 4).Value) Then isSunday = (Weekday(ws.Cells(r, 4).Value, vbSunday) = 1)
        
        Dim rowNote As String: rowNote = UCase(Trim(CStr(ws.Cells(r, 15).Value)))
        If InStr(1, rowNote, "BO CA", vbTextCompare) > 0 Then
            rowRg.Interior.ThemeColor = xlThemeColorAccent2
            rowRg.Interior.TintAndShade = 0
        ElseIf sc = "N" Or sc = "VANG" Then
            rowRg.Interior.Color = RGB(255, 199, 206) ' Do nhat - nghi khong phep
        ElseIf sc = "P" Then
            rowRg.Interior.Color = RGB(198, 224, 180) ' Xanh la nhat - nghi phep ca ngay
        ElseIf sc = "S" Or sc = "C" Then
            rowRg.Interior.Color = RGB(226, 239, 218) ' Xanh la rat nhat - nghi nua ngay phep
        ElseIf sc = "OFF" Then
            rowRg.Interior.Color = RGB(242, 242, 242) ' Xam nhat - OFF
        ElseIf sc = "L" Then
            rowRg.Interior.Color = RGB(255, 230, 153) ' Vang dam - ngay le
        ElseIf isSunday Then
            rowRg.Interior.Color = RGB(221, 235, 247) ' Xanh duong nhat - Chu Nhat
        ElseIf soGioVal > 0 And soGioVal < 8 Then
            rowRg.Interior.Color = RGB(255, 242, 204) ' Vang nhat - lam nua ngay
        End If
        
        ' --- Chenh lech < -1 => to do o chenh lech (cot I) ---
        Dim chenhLechVal As Double
        If IsNumeric(ws.Cells(r, 9).Value) Then
            chenhLechVal = CDbl(ws.Cells(r, 9).Value)
            If chenhLechVal < -1 And soGioVal > 0 And InStr(1, rowNote, "BO CA", vbTextCompare) = 0 Then
                ws.Cells(r, 9).Interior.Color = RGB(255, 199, 206) ' Do nhat
                ws.Cells(r, 9).Font.Color = RGB(156, 0, 6) ' Do dam
                ws.Cells(r, 9).Font.Bold = True
            End If
        End If
    Next r
End Sub


'==================================================================================================
' PART 4: CREATE INDIVIDUAL PAYSLIPS
'==================================================================================================

Private Sub CreateIndividualPayslips(wb As Workbook, wsData As Worksheet)
    ' --- Delete old payslip sheets ---
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Left(ws.Name, 3) = "BL_" Then ws.Delete
    Next ws
    
    ' --- Get unique employee list from Cham Cong sheet ---
    Dim dictNV As Object: Set dictNV = CreateObject("Scripting.Dictionary")
    Dim lastDataRow As Long: lastDataRow = wsData.Cells(wsData.Rows.Count, "B").End(xlUp).Row
    Dim nvID As Long ' *** FIX: Declare nvID once here
    
    Dim r As Long
    For r = 2 To lastDataRow
        If IsNumeric(wsData.Cells(r, 2).Value) Then
            nvID = CLng(wsData.Cells(r, 2).Value) ' *** FIX: Remove Dim
            If Not dictNV.Exists(nvID) Then
                dictNV.Add nvID, wsData.Cells(r, 3).Value
            End If
        End If
    Next r
    
    ' --- Create sheet for each employee ---
    Dim key As Variant
    For Each key In dictNV.keys
        nvID = CLng(key) ' *** FIX: Remove Dim
        Dim hoTen As String: hoTen = dictNV(key)
        
        Dim wsNew As Worksheet
        Set wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsNew.Name = "BL_" & nvID
        
        ' --- Copy headers ---
        wsData.Range("A1:P1").Copy wsNew.Range("A1")
        
        ' --- Copy data rows and summary rows ---
        Dim destRow As Long: destRow = 2
        For r = 2 To lastDataRow
            ' Copy data row
            If wsData.Cells(r, 2).Value = nvID Then
                wsData.Rows(r).Copy wsNew.Rows(destRow)
                destRow = destRow + 1
            End If
            ' Copy summary row
            If InStr(1, wsData.Cells(r, 3).Value, "NV: " & nvID) > 0 Then
                wsData.Rows(r).Copy wsNew.Rows(destRow)
                wsData.Rows(r + 1).Copy wsNew.Rows(destRow + 1)
                destRow = destRow + 2
            End If
        Next r
        
        ' --- Final Formatting ---
        wsNew.Columns.AutoFit
        wsNew.Activate
        ActiveWindow.FreezePanes = False
        wsNew.Range("A2").Select
        ActiveWindow.FreezePanes = True
    Next key
End Sub
