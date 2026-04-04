Attribute VB_Name = "Mod1_TaoSheetCauHinh"
'==================================================================================================
' MODULE 1: TAO CAC SHEET CAU HINH
' - Chay 1 lan de khoi tao hoac khi co thay doi ve danh sach nhan vien/nam.
' - Tao ra cac sheet nen tang cho viec tinh luong:
'   1. Loai Ca: Dinh nghia cac loai ca lam viec, he so luong, gio chuan.
'   2. Bang Luong: De nguoi dung nhap luong co ban cho tung nhan vien.
'   3. Lich Lam: Cham cong + quan ly phep nam (12 ngay/nam, 1 ngay/thang).
'   4. Tang Ca: Cham cong gio lam them.
'   5. Ngay OFF: Danh dau cac ngay nghi le, nghi toan cong ty.
'   6. Tien An: Quan ly tien an theo tung nhan vien theo ngay.
'==================================================================================================

Option Explicit

Private Const PHEP_NAM As Double = 12 ' 12 ngay phep/nam = 1 ngay/thang
Private Const LICH_FILTER_HELPER_COL As String = "ZY"

'==================================================================================================
' MAIN SUBROUTINE
'==================================================================================================

Sub TaoTatCaSheetCauHinh()
    If MsgBox("Tao tat ca 6 sheet cau hinh?", vbYesNo + vbQuestion, "Xac Nhan") = vbNo Then Exit Sub
    RunCreateSheets 1, 2, 3, 4, 5, 6
    MsgBox "Da tao xong tat ca cac sheet cau hinh!", vbInformation, "Hoan Thanh"
End Sub

Sub TaoSheetTheoLuaChon()
    Dim prompt As String
    prompt = "Chon sheet can tao (nhap so, phan cach boi dau phay):" & vbCrLf & _
             "1: Loai Ca" & vbCrLf & _
             "2: Bang Luong" & vbCrLf & _
             "3: Lich Lam (co quan ly phep)" & vbCrLf & _
             "4: Tang Ca" & vbCrLf & _
             "5: Ngay OFF" & vbCrLf & _
             "6: Tien An" & vbCrLf & _
             "7: Tat ca" & vbCrLf & _
             "(vd: 3 hoac 1,3,5)"
    Dim choice As String
    choice = Application.InputBox(prompt, "Lua chon sheet", Type:=2)
    If choice = "False" Or Trim(choice) = "" Then Exit Sub
    
    Dim ids As Collection
    Set ids = New Collection
    Dim i As Long
    For i = 1 To 6
        If ShouldCreate(choice, i) Then ids.Add i
    Next i
    If ids.Count = 0 Then Exit Sub
    
    Dim arr() As Variant
    ReDim arr(0 To ids.Count - 1)
    For i = 1 To ids.Count
        arr(i - 1) = ids(i)
    Next i
    RunCreateSheets arr
    MsgBox "Da tao xong cac sheet duoc chon!", vbInformation, "Hoan Thanh"
End Sub

'==================================================================================================
' QUICK MACROS: TAO TUNG SHEET RIENG LE
'==================================================================================================

Sub GenSheet_LoaiCa()
    Application.ScreenUpdating = False
    On Error GoTo ErrSingle
    CreateSheet_LoaiCa
    Application.ScreenUpdating = True
    Exit Sub
ErrSingle:
    Application.ScreenUpdating = True
    MsgBox "Loi: " & Err.Description, vbCritical
End Sub

Sub GenSheet_BangLuong()
    Application.ScreenUpdating = False
    On Error GoTo ErrSingle
    CreateSheet_BangLuong
    Application.ScreenUpdating = True
    Exit Sub
ErrSingle:
    Application.ScreenUpdating = True
    MsgBox "Loi: " & Err.Description, vbCritical
End Sub

Sub GenSheet_LichLam()
    Application.ScreenUpdating = False
    On Error GoTo ErrSingle
    CreateSheet_LichLam
    Application.ScreenUpdating = True
    Exit Sub
ErrSingle:
    Application.ScreenUpdating = True
    MsgBox "Loi: " & Err.Description, vbCritical
End Sub

Sub GenSheet_TangCa()
    Application.ScreenUpdating = False
    On Error GoTo ErrSingle
    CreateSheet_TangCa
    Application.ScreenUpdating = True
    Exit Sub
ErrSingle:
    Application.ScreenUpdating = True
    MsgBox "Loi: " & Err.Description, vbCritical
End Sub

Sub GenSheet_NgayOFF()
    Application.ScreenUpdating = False
    On Error GoTo ErrSingle
    CreateSheet_NgayOFF
    Application.ScreenUpdating = True
    Exit Sub
ErrSingle:
    Application.ScreenUpdating = True
    MsgBox "Loi: " & Err.Description, vbCritical
End Sub

Sub GenSheet_TienAn()
    Application.ScreenUpdating = False
    On Error GoTo ErrSingle
    CreateSheet_TienAn
    Application.ScreenUpdating = True
    Exit Sub
ErrSingle:
    Application.ScreenUpdating = True
    MsgBox "Loi: " & Err.Description, vbCritical
End Sub

Sub GenSheet_CauHinh_TatCa()
    If MsgBox("Tao tat ca 6 sheet cau hinh?", vbYesNo + vbQuestion, "Xac Nhan") = vbNo Then Exit Sub
    RunCreateSheets 1, 2, 3, 4, 5, 6
End Sub

'==================================================================================================
' KIEM TRA PHEP: Quet sheet Lich Lam, bao loi neu NV nao vuot 1 ngay phep/thang
' Goi truoc khi chay cham cong, hoac chay thu cong bat ky luc nao.
'==================================================================================================
Sub KiemTraPhep()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Lich Lam")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Khong tim thay sheet 'Lich Lam'.", vbExclamation
        Exit Sub
    End If
    
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Dim violations As String: violations = ""
    Dim violationCount As Long: violationCount = 0
    
    Dim col As Long
    For col = 3 To lastCol
        Dim nvID As String: nvID = CStr(ws.Cells(2, col).Value)
        Dim nvName As String: nvName = CStr(ws.Cells(1, col).Value)
        
        ' Scan monthly summary rows
        Dim r As Long
        For r = 3 To lastRow
            Dim cellB As String: cellB = CStr(ws.Cells(r, 2).Value)
            If Left(cellB, 11) = "Phep thang " Then
                Dim monthNum As String: monthNum = Mid(cellB, 12)
                Dim phepVal As Double: phepVal = 0
                If IsNumeric(ws.Cells(r, col).Value) Then phepVal = CDbl(ws.Cells(r, col).Value)
                
                If phepVal > 1 Then
                    violationCount = violationCount + 1
                    violations = violations & "  - " & nvName & ": Thang " & monthNum & " nghi " & phepVal & " ngay phep (max 1.0)" & vbCrLf
                    ' Highlight the cell
                    ws.Cells(r, col).Interior.Color = RGB(255, 0, 0)
                    ws.Cells(r, col).Font.Color = RGB(255, 255, 255)
                End If
            End If
        Next r
    Next col
    
    If violationCount > 0 Then
        MsgBox "CANH BAO: Co " & violationCount & " truong hop vuot phep thang (max 1 ngay/thang):" & vbCrLf & vbCrLf & violations & vbCrLf & _
               "Vui long sua lai sheet 'Lich Lam' truoc khi chay cham cong.", vbExclamation, "Vuot Phep"
    Else
        MsgBox "OK! Tat ca nhan vien deu trong han muc phep (max 1 ngay/thang).", vbInformation, "Kiem Tra Phep"
    End If
End Sub

' Ham kiem tra phep, tra ve True neu OK, False neu co vi pham (dung trong lan3.bas truoc khi chay)
Public Function KiemTraPhepTruocKhiChay() As Boolean
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Lich Lam")
    On Error GoTo 0
    If ws Is Nothing Then
        KiemTraPhepTruocKhiChay = True ' Khong co sheet thi bo qua
        Exit Function
    End If
    
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Dim violations As String: violations = ""
    Dim col As Long, r As Long
    For col = 3 To lastCol
        For r = 3 To lastRow
            If Left(CStr(ws.Cells(r, 2).Value), 11) = "Phep thang " Then
                If IsNumeric(ws.Cells(r, col).Value) Then
                    If CDbl(ws.Cells(r, col).Value) > 1 Then
                        violations = violations & "  - " & ws.Cells(1, col).Value & " T" & Mid(CStr(ws.Cells(r, 2).Value), 12) & ": " & ws.Cells(r, col).Value & " ngay" & vbCrLf
                    End If
                End If
            End If
        Next r
    Next col
    
    If violations <> "" Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("CANH BAO vuot phep thang:" & vbCrLf & violations & vbCrLf & "Van tiep tuc chay cham cong?", vbYesNo + vbExclamation, "Vuot Phep")
        KiemTraPhepTruocKhiChay = (ans = vbYes)
    Else
        KiemTraPhepTruocKhiChay = True
    End If
End Function

Private Sub RunCreateSheets(ParamArray ids() As Variant)
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    Dim flat As Collection
    Set flat = New Collection
    Dim item As Variant
    For Each item In ids
        If IsArray(item) Then
            Dim inner As Variant
            For Each inner In item
                flat.Add inner
            Next inner
        Else
            flat.Add item
        End If
    Next item
    
    If flat.Count = 0 Then GoTo CleanUp
    
    Dim id As Variant
    For Each id In flat
        If IsNumeric(id) Then
            Select Case CLng(id)
                Case 1: CreateSheet_LoaiCa
                Case 2: CreateSheet_BangLuong
                Case 3: CreateSheet_LichLam
                Case 4: CreateSheet_TangCa
                Case 5: CreateSheet_NgayOFF
                Case 6: CreateSheet_TienAn
            End Select
        End If
    Next id

CleanUp:
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    MsgBox "Da co loi xay ra: " & Err.Description, vbCritical, "Loi"
    Resume CleanUp
End Sub


'==================================================================================================
' SUB 1: TAO SHEET "LOAI CA"
'==================================================================================================
Private Sub CreateSheet_LoaiCa()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Loai Ca")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "Loai Ca"
    Else
        ws.Cells.Clear
    End If
    
    ' --- HEADERS ---
    Dim headers As Variant
    headers = Array("Ma Ca", "Ten Ca", "So Gio", "Gio Bat Dau", "Gio Ket Thuc", "He So Luong", "Tien An", "Ghi Chu")
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    ' --- DATA ---
    Dim dataArr As Variant
    ReDim dataArr(0 To 23)
    dataArr(0) = Array("X", "Ca Ngay Thuong", 8, "07:00", "16:00", 1, 25000, "Ca ngay 8h (7h-16h)")
    dataArr(1) = Array("XVP", "Ca Van Phong", 8, "08:00", "17:00", 1, 25000, "Ca van phong 8h-17h")
    dataArr(2) = Array("CN", "Ca Ngay Chu Nhat", 8, "07:00", "16:00", 2, 25000, "He so x2")
    dataArr(3) = Array("D", "Ca Dem Thuong", 8, "18:00", "06:00", 1.5, 30000, "Ca dem: 8h cong + 4h OT mac dinh")
    dataArr(4) = Array("CND", "Ca Dem Chu Nhat", 8, "18:00", "06:00", 3, 30000, "Ca dem CN: 8h cong + 4h OT mac dinh")
    dataArr(5) = Array("N3", "Ca Nu Sang", 8, "07:00", "16:00", 1, 25000, "Ca nu sang")
    dataArr(6) = Array("N4", "Ca Nu Toi", 8, "18:00", "22:00", 1, 25000, "Ca nu toi")
    dataArr(7) = Array("L", "Ngay Le", 8, "", "", 3, 0, "He so x3, khong co tien an")
    dataArr(8) = Array("N", "Nghi Ca Ngay", 0, "", "", 0, 0, "Khong tinh luong")
    dataArr(9) = Array("S", "Nghi Sang", 4, "", "", 0, 0, "Lam 4h chieu, khong luong cho 4h sang, tru 0.5 phep")
    dataArr(10) = Array("C", "Nghi Chieu", 4, "", "", 0, 0, "Lam 4h sang, khong luong cho 4h chieu, tru 0.5 phep")
    dataArr(11) = Array("P", "Nghi Phep", 8, "", "", 1, 0, "Nghi co huong luong, khong tien an, tru 1 phep")
    dataArr(12) = Array("T1", "Lam Thay 1 Gio", 1, "", "", 1, 0, "Lam thay 1 gio")
    dataArr(13) = Array("T2", "Lam Thay 2 Gio", 2, "", "", 1, 0, "Lam thay 2 gio")
    dataArr(14) = Array("T3", "Lam Thay 3 Gio", 3, "", "", 1, 0, "Lam thay 3 gio")
    dataArr(15) = Array("T4", "Lam Thay 4 Gio", 4, "", "", 1, 0, "Lam thay 4 gio")
    dataArr(16) = Array("T5", "Lam Thay 5 Gio", 5, "", "", 1, 0, "Lam thay 5 gio")
    dataArr(17) = Array("T6", "Lam Thay 6 Gio", 6, "", "", 1, 0, "Lam thay 6 gio")
    dataArr(18) = Array("T7", "Lam Thay 7 Gio", 7, "", "", 1, 0, "Lam thay 7 gio")
    dataArr(19) = Array("T8", "Lam Thay 8 Gio", 8, "", "", 1, 0, "Lam thay 8 gio")
    dataArr(20) = Array("T9", "Lam Thay 9 Gio", 9, "", "", 1, 0, "Lam thay 9 gio")
    dataArr(21) = Array("T10", "Lam Thay 10 Gio", 10, "", "", 1, 0, "Lam thay 10 gio")
    dataArr(22) = Array("T11", "Lam Thay 11 Gio", 11, "", "", 1, 0, "Lam thay 11 gio")
    dataArr(23) = Array("T12", "Lam Thay 12 Gio", 12, "", "", 1, 0, "Lam thay 12 gio")
    
    Dim i As Long
    For i = 0 To UBound(dataArr)
        ws.Range("A" & i + 2).Resize(1, UBound(dataArr(i)) + 1).Value = dataArr(i)
    Next i
    
    ' --- FORMATTING ---
    ws.Range("F2:F" & UBound(dataArr) + 2).NumberFormat = "0.0"
    ws.Range("C2:C" & UBound(dataArr) + 2).NumberFormat = "0"
    ws.Range("G2:G" & UBound(dataArr) + 2).NumberFormat = "#,##0"
    
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(79, 129, 189)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Columns.AutoFit
    ws.Columns("H").ColumnWidth = 35
    
    With ws.Range("A1:H" & UBound(dataArr) + 2).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Range("A1:H" & UBound(dataArr) + 2).AutoFilter
    
    ws.Activate
    MsgBox "Da tao/cap nhat sheet 'Loai Ca'.", vbInformation
End Sub

'==================================================================================================
' SUB 2: TAO SHEET "BANG LUONG"
'==================================================================================================
Private Sub CreateSheet_BangLuong()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim dictNV As Object
    Set dictNV = GetUniqueEmployeeList(wb)
    If dictNV Is Nothing Then Exit Sub
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Bang Luong")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "Bang Luong"
    Else
        ws.Cells.Clear
    End If
    
    ' --- HEADERS ---
    ws.Range("A1:D1").Value = Array("STT", "Ma NV", "Ho Ten", "Muc Luong Thang")
    ws.Range("F1").Value = "He so chia luong (ngay)"
    ws.Range("F2").Value = 26
    
    ' --- DATA ---
    Dim arrIDs As Variant
    arrIDs = SortKeys(dictNV)
    
    Dim i As Long
    For i = 0 To UBound(arrIDs)
        ws.Cells(i + 2, 1).Value = i + 1
        ws.Cells(i + 2, 2).Value = arrIDs(i)
        ws.Cells(i + 2, 3).Value = dictNV(arrIDs(i))
    Next i
    
    ' --- FORMATTING ---
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 176, 80)
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("F1:F2")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    ws.Range("F1").Font.Bold = True
    ws.Range("F1").Interior.Color = RGB(226, 239, 218)
    ws.Range("F2").NumberFormat = "0.00"
    ws.Range("F2").Interior.Color = RGB(255, 242, 204)
    
    ws.Range("D2:D" & UBound(arrIDs) + 2).NumberFormat = "#,##0"
    ws.Range("D2:D" & UBound(arrIDs) + 2).Interior.Color = RGB(226, 239, 218)
    
    ws.Columns.AutoFit
    ws.Columns("C").ColumnWidth = 25
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 22
    
    With ws.Range("A1:D" & UBound(arrIDs) + 2).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Range("A1:D" & UBound(arrIDs) + 2).AutoFilter
    
    ws.Activate
    MsgBox "Da tao/cap nhat sheet 'Bang Luong'. Vui long nhap muc luong thang cho tung nhan vien.", vbInformation
End Sub

'==================================================================================================
' SUB 3: TAO SHEET "LICH LAM" (CO QUAN LY PHEP NAM)
'==================================================================================================
Private Sub CreateSheet_LichLam()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim dictNV As Object
    Set dictNV = GetUniqueEmployeeList(wb)
    If dictNV Is Nothing Then Exit Sub
    
    Dim targetYear As Long
    targetYear = GetTargetYear("Lich Lam")
    If targetYear = 0 Then Exit Sub
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Lich Lam")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "Lich Lam"
    Else
        ws.Cells.Clear
    End If
    
    ' --- HEADERS ---
    ws.Range("A1:B1").Value = Array("STT", "Ngay")
    ws.Cells(2, 2).Value = "Ma NV ->"
    
    Dim arrIDs As Variant
    arrIDs = SortKeys(dictNV)
    
    Dim i As Long, col As Long
    For i = 0 To UBound(arrIDs)
        col = i + 3
        ws.Cells(1, col).Value = "ID " & arrIDs(i) & " - " & dictNV(arrIDs(i))
        ws.Cells(2, col).Value = arrIDs(i) ' Ma NV so cho parser
    Next i
    
    Dim lastCol As Long: lastCol = UBound(arrIDs) + 3
    
    ' --- DAYS ---
    Dim r As Long: r = 3
    Dim d As Date
    Dim currentMonth As Long: currentMonth = 0
    
    For d = DateSerial(targetYear, 1, 1) To DateSerial(targetYear, 12, 31)
        ' --- Insert monthly summary row at start of each new month ---
        If Month(d) <> currentMonth Then
            If currentMonth > 0 Then
                ' Write summary for previous month
                Call WriteMonthlySummaryRow(ws, r, currentMonth, targetYear, arrIDs, lastCol)
                r = r + 1
            End If
            currentMonth = Month(d)
        End If
        
        ws.Cells(r, 1).Value = r - 2
        ws.Cells(r, 2).Value = d
        If Weekday(d, vbSunday) = 1 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(255, 242, 204) ' Yellow for Sunday
        End If
        r = r + 1
    Next d
    
    ' --- Summary for last month (Dec) ---
    If currentMonth > 0 Then
        Call WriteMonthlySummaryRow(ws, r, currentMonth, targetYear, arrIDs, lastCol)
        r = r + 1
    End If
    
    ' --- YEARLY SUMMARY ROWS ---
    ' Row: Tong phep da dung ca nam
    ws.Cells(r, 1).Value = ""
    ws.Cells(r, 2).Value = "TONG PHEP DA DUNG"
    ws.Cells(r, 2).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(255, 199, 206) ' Light red
    For i = 0 To UBound(arrIDs)
        col = i + 3
        Dim sumFormula As String
        sumFormula = BuildSumPhepFormula(ws, col, 3, r - 1)
        If sumFormula <> "" Then
            ws.Cells(r, col).formula = sumFormula
        Else
            ws.Cells(r, col).Value = 0
        End If
        ws.Cells(r, col).NumberFormat = "0.0"
        ws.Cells(r, col).Font.Bold = True
    Next i
    Dim rowTongPhep As Long: rowTongPhep = r
    r = r + 1
    
    ' Row: Phep nam (12)
    ws.Cells(r, 2).Value = "PHEP NAM (" & PHEP_NAM & " ngay)"
    ws.Cells(r, 2).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(198, 224, 180) ' Light green
    For i = 0 To UBound(arrIDs)
        ws.Cells(r, i + 3).Value = PHEP_NAM
        ws.Cells(r, i + 3).NumberFormat = "0.0"
        ws.Cells(r, i + 3).Font.Bold = True
    Next i
    Dim rowPhepNam As Long: rowPhepNam = r
    r = r + 1
    
    ' Row: Phep con lai
    ws.Cells(r, 2).Value = "PHEP CON LAI"
    ws.Cells(r, 2).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(180, 198, 231) ' Light blue
    For i = 0 To UBound(arrIDs)
        col = i + 3
        ws.Cells(r, col).formula = "=" & ws.Cells(rowPhepNam, col).Address(False, False) & "-" & ws.Cells(rowTongPhep, col).Address(False, False)
        ws.Cells(r, col).NumberFormat = "0.0"
        ws.Cells(r, col).Font.Bold = True
        ' Conditional Formatting: phep con lai < 0 => do
        ws.Cells(r, col).FormatConditions.Delete
        ws.Cells(r, col).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        With ws.Cells(r, col).FormatConditions(1)
            .Interior.Color = RGB(255, 0, 0)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With
    Next i
    r = r + 1
    
    ' --- VALIDATION: Danh sach hop le (gom ma ca thuong + ma ca kem gio nghi 1..7) ---
    Dim baseCodes As Variant
    baseCodes = Array("X", "XVP", "CN", "D", "CND", "N3", "N4", "L", "N", "S", "C", "P", "T1", "T2", "T3", "T4", "T5", "T6", "T7", "T8", "T9", "T10", "T11", "T12")

    Dim helperCol As String: helperCol = "ZZ"
    ws.Columns(helperCol).ClearContents

    Dim listRow As Long: listRow = 1
    For i = LBound(baseCodes) To UBound(baseCodes)
        ws.Cells(listRow, helperCol).Value = baseCodes(i)
        listRow = listRow + 1
    Next i

    Dim h As Long
    For i = LBound(baseCodes) To UBound(baseCodes)
        For h = 1 To 7
            ws.Cells(listRow, helperCol).Value = baseCodes(i) & "," & h
            listRow = listRow + 1
        Next h
    Next i

    ws.Columns(helperCol).Hidden = True

    Dim dataRow As Long
    For dataRow = 3 To r - 1
        If IsDate(ws.Cells(dataRow, 2).Value) Then
            On Error Resume Next
            ws.Range(ws.Cells(dataRow, 3), ws.Cells(dataRow, lastCol)).Validation.Delete
            ws.Range(ws.Cells(dataRow, 3), ws.Cells(dataRow, lastCol)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$" & helperCol & "$1:$" & helperCol & "$" & (listRow - 1)
            ws.Range(ws.Cells(dataRow, 3), ws.Cells(dataRow, lastCol)).Validation.IgnoreBlank = True
            ws.Range(ws.Cells(dataRow, 3), ws.Cells(dataRow, lastCol)).Validation.InCellDropdown = True
            ws.Range(ws.Cells(dataRow, 3), ws.Cells(dataRow, lastCol)).Validation.ShowInput = True
            ws.Range(ws.Cells(dataRow, 3), ws.Cells(dataRow, lastCol)).Validation.ShowError = True
            ws.Range(ws.Cells(dataRow, 3), ws.Cells(dataRow, lastCol)).Validation.ErrorTitle = "Ma ca khong hop le"
            ws.Range(ws.Cells(dataRow, 3), ws.Cells(dataRow, lastCol)).Validation.ErrorMessage = "Chi nhap ma hop le: X, XVP, CN, D, CND, N3, N4, L, N, S, C, P, T1..T12 hoac dang maCa,1 den maCa,7 (vd: N4,2; khong co khoang trang)."
            On Error GoTo 0
        End If
    Next dataRow
    
    ' --- CONDITIONAL FORMATTING: Phep thang > 1.0 => to do canh bao ---
    ' Apply to monthly summary rows ("Phep thang X")
    Dim cfRow As Long
    For cfRow = 3 To r - 1
        Dim cellB As String: cellB = CStr(ws.Cells(cfRow, 2).Value)
        If Left(cellB, 11) = "Phep thang " Then
            For i = 0 To UBound(arrIDs)
                col = i + 3
                ws.Cells(cfRow, col).FormatConditions.Delete
                ' > 1.0 => do dam (vuot phep thang)
                ws.Cells(cfRow, col).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="1"
                With ws.Cells(cfRow, col).FormatConditions(1)
                    .Interior.Color = RGB(255, 0, 0)
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Bold = True
                End With
                ' = 1.0 => vang (het phep thang)
                ws.Cells(cfRow, col).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="1"
                With ws.Cells(cfRow, col).FormatConditions(2)
                    .Interior.Color = RGB(255, 235, 156)
                    .Font.Bold = True
                End With
            Next i
        End If
    Next cfRow
    
    ' --- FORMATTING ---
    ws.Range("B3:B" & r - 1).NumberFormat = "dd/mm/yyyy"
    
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(218, 150, 148)
    End With
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With
    
    ws.Columns.AutoFit
    ws.Columns("B").ColumnWidth = 22
    ws.Range(ws.Cells(3, 3), ws.Cells(r - 1, lastCol)).HorizontalAlignment = xlCenter

    ' Helper col for month filter: keep monthly/yearly leave summary rows visible when filtering a month
    ws.Columns(LICH_FILTER_HELPER_COL).ClearContents
    ws.Cells(1, LICH_FILTER_HELPER_COL).Value = "__LOC_THANG__"
    ws.Cells(2, LICH_FILTER_HELPER_COL).Value = ""
    For dataRow = 3 To r - 1
        If IsDate(ws.Cells(dataRow, 2).Value) Then
            ws.Cells(dataRow, LICH_FILTER_HELPER_COL).Value = Month(CDate(ws.Cells(dataRow, 2).Value))
        Else
            Dim bVal As String: bVal = UCase(Trim(CStr(ws.Cells(dataRow, 2).Value)))
            If Left(bVal, 11) = "PHEP THANG " Or bVal = "TONG PHEP DA DUNG" Or Left(bVal, 8) = "PHEP NAM" Or bVal = "PHEP CON LAI" Then
                ws.Cells(dataRow, LICH_FILTER_HELPER_COL).Value = "PHEP"
            Else
                ws.Cells(dataRow, LICH_FILTER_HELPER_COL).Value = ""
            End If
        End If
    Next dataRow
    ws.Columns(LICH_FILTER_HELPER_COL).Hidden = True
    
    With ws.Range(ws.Cells(1, 1), ws.Cells(r - 1, lastCol)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Guide text
    ws.Cells(r, 1).Value = "Ma ca: X, XVP, CN, D, CND, N3, N4, T1..T12, L, N, S(-0.5P), C(-0.5P), P(-1P). Nghi theo gio nhap dang maCa,1 den maCa,7 (vd: N4,2), khong co khoang trang; gio nghi se quy doi tru phep theo ty le gio/8. Moi thang toi da 1 ngay phep. Vuot phep => o DO."
    ws.Cells(r, 1).Font.Italic = True
    
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("C3").Select
    ActiveWindow.FreezePanes = True

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(r - 1, ws.Columns(LICH_FILTER_HELPER_COL).Column)).AutoFilter

    Call CreateSheet_TongPhepThang(wb, ws, arrIDs, lastCol)
    
    MsgBox "Da tao/cap nhat sheet 'Lich Lam' cho nam " & targetYear & "." & vbCrLf & _
           "Phep nam: " & PHEP_NAM & " ngay/nam (1 ngay/thang)." & vbCrLf & _
               "P = -1 ngay phep, S/C = -0.5 ngay phep, nhap maCa,1 den maCa,7 (vd: N4,2) se tru phep theo gio/8 (khong khoang trang)." & vbCrLf & _
           "Vuot 1 ngay phep/thang => o DO canh bao.", vbInformation
End Sub

Private Sub CreateSheet_TongPhepThang(wb As Workbook, wsLich As Worksheet, arrIDs As Variant, lastColLich As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Tong Phep")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "Tong Phep"
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1:B1").Value = Array("STT", "Thang")

    Dim i As Long, col As Long
    For i = 0 To UBound(arrIDs)
        col = i + 3
        ws.Cells(1, col).Value = wsLich.Cells(1, col).Value
    Next i

    Dim r As Long
    For r = 2 To 13
        ws.Cells(r, 1).Value = r - 1
        ws.Cells(r, 2).Value = r - 1
        For i = 0 To UBound(arrIDs)
            col = i + 3
            Dim colLetter As String: colLetter = ColumnLetter(col)
            ws.Cells(r, col).Formula = "=IFERROR(INDEX('" & wsLich.Name & "'!" & colLetter & ":" & colLetter & ",MATCH(""Phep thang ""&B" & r & ",'" & wsLich.Name & "'!$B:$B,0)),0)"
            ws.Cells(r, col).NumberFormat = "0.0"
        Next i
    Next r

    ws.Cells(14, 2).Value = "Tong phep da dung"
    ws.Cells(15, 2).Value = "Phep nam"
    ws.Cells(16, 2).Value = "Phep con lai"
    For i = 0 To UBound(arrIDs)
        col = i + 3
        ws.Cells(14, col).Formula = "=SUM(" & ws.Cells(2, col).Address(False, False) & ":" & ws.Cells(13, col).Address(False, False) & ")"
        ws.Cells(15, col).Value = PHEP_NAM
        ws.Cells(16, col).Formula = "=" & ws.Cells(15, col).Address(False, False) & "-" & ws.Cells(14, col).Address(False, False)
        ws.Range(ws.Cells(14, col), ws.Cells(16, col)).NumberFormat = "0.0"
    Next i

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastColLich))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 112, 192)
    End With

    ws.Range(ws.Cells(14, 1), ws.Cells(14, lastColLich)).Interior.Color = RGB(255, 199, 206)
    ws.Range(ws.Cells(15, 1), ws.Cells(15, lastColLich)).Interior.Color = RGB(198, 224, 180)
    ws.Range(ws.Cells(16, 1), ws.Cells(16, lastColLich)).Interior.Color = RGB(180, 198, 231)
    ws.Range(ws.Cells(14, 1), ws.Cells(16, lastColLich)).Font.Bold = True

    ws.Columns.AutoFit
    ws.Columns("B").ColumnWidth = 10
    ws.Range(ws.Cells(2, 2), ws.Cells(13, 2)).HorizontalAlignment = xlCenter

    With ws.Range(ws.Cells(1, 1), ws.Cells(16, lastColLich)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(16, lastColLich)).AutoFilter
End Sub

Private Function ColumnLetter(colNum As Long) As String
    ColumnLetter = Split(Cells(1, colNum).Address(False, False), "1")(0)
End Function

Sub LocLichLamTheoThang_GiuDongPhep()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Lich Lam")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim m As Variant
    m = Application.InputBox("Nhap thang can loc (1-12):", "Loc Lich Lam", Type:=1)
    If m = False Then Exit Sub
    If Not IsNumeric(m) Then Exit Sub
    If CLng(m) < 1 Or CLng(m) > 12 Then Exit Sub

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    Dim filterField As Long: filterField = ws.Columns(LICH_FILTER_HELPER_COL).Column

    If ws.AutoFilterMode = False Then
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, filterField)).AutoFilter
    End If

    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, filterField)).AutoFilter Field:=filterField, Criteria1:=CStr(CLng(m)), Operator:=xlOr, Criteria2:="PHEP"
End Sub

' Write a monthly summary row showing leave used in that month
Private Sub WriteMonthlySummaryRow(ws As Worksheet, r As Long, monthNum As Long, targetYear As Long, arrIDs As Variant, lastCol As Long)
    Dim monthName As String
    monthName = "Thang " & monthNum
    
    ws.Cells(r, 1).Value = ""
    ws.Cells(r, 2).Value = "Phep thang " & monthNum
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Italic = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(242, 220, 219) ' Light pink
    
    ' Find the range of rows for this month (scan upwards from r)
    Dim firstDayRow As Long: firstDayRow = 0
    Dim lastDayRow As Long: lastDayRow = r - 1
    Dim scanRow As Long
    For scanRow = r - 1 To 3 Step -1
        If IsDate(ws.Cells(scanRow, 2).Value) Then
            If Month(CDate(ws.Cells(scanRow, 2).Value)) = monthNum And Year(CDate(ws.Cells(scanRow, 2).Value)) = targetYear Then
                If firstDayRow = 0 Or scanRow < firstDayRow Then firstDayRow = scanRow
            End If
        End If
    Next scanRow
    
    If firstDayRow = 0 Then Exit Sub
    
    ' For each NV column, count P (=1), S/C (=0.5) and ma kem gio nghi (vd X,1 => 1/8 ngay phep)
    Dim i As Long, col As Long
    For i = 0 To UBound(arrIDs)
        col = i + 3
        ' COUNTIF for P + 0.5*COUNTIF for S + 0.5*COUNTIF for C + gio nghi/8
        Dim rng As String
        rng = ws.Cells(firstDayRow, col).Address(False, False) & ":" & ws.Cells(lastDayRow, col).Address(False, False)
        ws.Cells(r, col).formula = "=COUNTIF(" & rng & ",""P"")+COUNTIF(" & rng & ",""S"")*0.5+COUNTIF(" & rng & ",""C"")*0.5+SUMPRODUCT(IFERROR(VALUE(TRIM(MID(" & rng & ",FIND("",""," & rng & ")+1,10))),0))/8"
        ws.Cells(r, col).NumberFormat = "0.0"
        ws.Cells(r, col).Font.Bold = True
    Next i
End Sub

' Build a SUM formula that adds up all monthly summary cells for a given column
Private Function BuildSumPhepFormula(ws As Worksheet, col As Long, startRow As Long, endRow As Long) As String
    Dim addresses As String: addresses = ""
    Dim r As Long
    For r = startRow To endRow
        Dim cellVal As String: cellVal = CStr(ws.Cells(r, 2).Value)
        If Left(cellVal, 11) = "Phep thang " Then
            If addresses <> "" Then addresses = addresses & ","
            addresses = addresses & ws.Cells(r, col).Address(False, False)
        End If
    Next r
    If addresses <> "" Then
        BuildSumPhepFormula = "=" & addresses
        ' Use SUM if multiple
        If InStr(addresses, ",") > 0 Then
            BuildSumPhepFormula = "=SUM(" & addresses & ")"
        Else
            BuildSumPhepFormula = "=" & addresses
        End If
    End If
End Function


'==================================================================================================
' SUB 4: TAO SHEET "TANG CA"
'==================================================================================================
Private Sub CreateSheet_TangCa()
    Call CreateGenericGridSheet("Tang Ca", "", "Nhap so gio tang ca (vd: 2.5).", RGB(0, 112, 192), True)
    Call ApplyAutoTangCaFromLichLam
End Sub

Private Sub ApplyAutoTangCaFromLichLam()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsTang As Worksheet, wsLich As Worksheet

    On Error Resume Next
    Set wsTang = wb.Sheets("Tang Ca")
    Set wsLich = wb.Sheets("Lich Lam")
    On Error GoTo 0

    If wsTang Is Nothing Or wsLich Is Nothing Then Exit Sub

    Dim lastCol As Long
    lastCol = wsTang.Cells(2, wsTang.Columns.Count).End(xlToLeft).Column
    If lastCol < 3 Then Exit Sub

    Dim lastRow As Long
    lastRow = wsTang.Cells(wsTang.Rows.Count, "B").End(xlUp).Row
    If lastRow < 3 Then Exit Sub

    Dim lastDateRow As Long: lastDateRow = 2
    Dim r As Long
    For r = 3 To lastRow
        If IsDate(wsTang.Cells(r, 2).Value) Then
            lastDateRow = r
        Else
            Exit For
        End If
    Next r
    If lastDateRow < 3 Then Exit Sub

    Dim targetRg As Range
    Set targetRg = wsTang.Range(wsTang.Cells(3, 3), wsTang.Cells(lastDateRow, lastCol))

    targetRg.Formula = "=GetAutoTangCaHoursFromCode(IFERROR(INDEX('" & wsLich.Name & "'!$A:$XFD,MATCH($B3,'" & wsLich.Name & "'!$B:$B,0),MATCH(C$2,'" & wsLich.Name & "'!$2:$2,0)),""""))"
    targetRg.NumberFormat = "0.##"

    With wsTang.Range(wsTang.Cells(3, 3), wsTang.Cells(lastDateRow, lastCol)).Validation
        .Delete
        .Add Type:=xlValidateDecimal, Operator:=xlGreaterEqual, Formula1:="0"
        .IgnoreBlank = True
        .ErrorTitle = "Du lieu khong hop le"
        .ErrorMessage = "Vui long nhap so >= 0."
    End With

    wsTang.Range("A" & lastDateRow + 1).Value = "Tu dong: Lich Lam N4 = 4h, N3 = 3h; neu nhap dang N4,2 thi tru 2 gio. Ket qua toi thieu = 0 (khong am)."
    wsTang.Range("A" & lastDateRow + 1).Font.Italic = True
End Sub

Public Function GetAutoTangCaHoursFromCode(rawCode As String) As Double
    Dim s As String: s = UCase$(Trim$(rawCode))
    If s = "" Then Exit Function

    Dim baseCode As String: baseCode = s
    Dim leaveHours As Double: leaveHours = 0

    Dim p As Long: p = InStr(1, s, ",", vbTextCompare)
    If p > 0 Then
        baseCode = Trim$(Left$(s, p - 1))
        Dim leaveText As String: leaveText = Trim$(Mid$(s, p + 1))
        If IsNumeric(leaveText) Then leaveHours = CDbl(leaveText)
    End If

    Dim baseOT As Double
    Select Case baseCode
        Case "N4": baseOT = 4
        Case "N3": baseOT = 3
        Case Else: baseOT = 0c:\Users\Admin\Downloads\ModChamCong (2).bas
    End Select

    GetAutoTangCaHoursFromCode = baseOT - leaveHours
    If GetAutoTangCaHoursFromCode < 0 Then GetAutoTangCaHoursFromCode = 0
End Function

'==================================================================================================
' SUB 5: TAO SHEET "NGAY OFF"
'==================================================================================================
Private Sub CreateSheet_NgayOFF()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim targetYear As Long
    targetYear = GetTargetYear("Ngay OFF")
    If targetYear = 0 Then Exit Sub
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Ngay OFF")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "Ngay OFF"
    Else
        ws.Cells.Clear
    End If
    
    ' --- HEADERS ---
    ws.Range("A1:F1").Value = Array("STT", "Ngay", "Thu", "Loai", "Mo Ta", "OFF")
    
    ' --- DATA ---
    Dim r As Long: r = 2
    Dim stt As Long: stt = 1
    Dim d As Date
    
    ' Ngay le mac dinh
    Dim dictHoliday As Object
    Set dictHoliday = CreateObject("Scripting.Dictionary")
    dictHoliday.Add Format(DateSerial(targetYear, 1, 1), "yyyymmdd"), "Tet Duong Lich"
    dictHoliday.Add Format(DateSerial(targetYear, 4, 30), "yyyymmdd"), "Ngay Giai Phong 30/4"
    dictHoliday.Add Format(DateSerial(targetYear, 5, 1), "yyyymmdd"), "Quoc Te Lao Dong 1/5"
    dictHoliday.Add Format(DateSerial(targetYear, 9, 2), "yyyymmdd"), "Quoc Khanh 2/9"
    
    For d = DateSerial(targetYear, 1, 1) To DateSerial(targetYear, 12, 31)
        Dim dKey As String: dKey = Format(d, "yyyymmdd")
        If Weekday(d, vbMonday) = 7 Or dictHoliday.Exists(dKey) Then
            ws.Cells(r, 1).Value = stt
            ws.Cells(r, 2).Value = d
            ws.Cells(r, 3).Value = Format(d, "dddd")
            
            If dictHoliday.Exists(dKey) Then
                ws.Cells(r, 4).Value = "Le"
                ws.Cells(r, 5).Value = dictHoliday(dKey)
                ws.Range("A" & r & ":F" & r).Interior.Color = RGB(255, 199, 206)
            Else
                ws.Cells(r, 4).Value = "CN"
                ws.Cells(r, 5).Value = "Chu Nhat"
                ws.Range("A" & r & ":F" & r).Interior.Color = RGB(255, 242, 204)
            End If
            
            ws.Cells(r, 6).Value = "X"
            stt = stt + 1
            r = r + 1
        End If
    Next d
    
    ' --- FORMATTING ---
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(112, 48, 160)
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("F2:F" & r - 1).Validation.Add Type:=xlValidateList, Formula1:="X,"
    ws.Columns.AutoFit
    ws.Columns("E").ColumnWidth = 30
    With ws.Range("A1:F" & r - 1).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Range("A1:F" & r - 1).AutoFilter
    
    ws.Activate
    MsgBox "Da tao/cap nhat sheet 'Ngay OFF'.", vbInformation
End Sub

'==================================================================================================
' SUB 6: TAO SHEET "TIEN AN"
'==================================================================================================
Private Sub CreateSheet_TienAn()
    Call CreateGenericGridSheet("Tien An", "", "Nhap so tien an theo ngay (vd: 25000).", RGB(255, 153, 0), True)
End Sub


'==================================================================================================
' HELPER FUNCTIONS
'==================================================================================================

' Ham lay danh sach nhan vien duy nhat tu sheet 1
Private Function GetUniqueEmployeeList(wb As Workbook) As Object
    Dim csvWs As Worksheet
    On Error Resume Next
    Set csvWs = wb.Sheets(1)
    On Error GoTo 0
    
    If csvWs Is Nothing Then
        MsgBox "Khong tim thay Sheet 1 (chua du lieu CSV).", vbCritical
        Set GetUniqueEmployeeList = Nothing
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = csvWs.Cells(csvWs.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Sheet 1 khong co du lieu.", vbExclamation
        Set GetUniqueEmployeeList = Nothing
        Exit Function
    End If
    
    Dim dictNV As Object
    Set dictNV = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To lastRow
        Dim rawID As String
        rawID = Replace(CStr(csvWs.Cells(i, 1).Value), "'", "")
        If IsNumeric(rawID) And rawID <> "" Then
            Dim nvID As Long
            nvID = CLng(rawID)
            If Not dictNV.Exists(nvID) Then
                dictNV.Add nvID, CStr(csvWs.Cells(i, 2).Value)
            End If
        End If
    Next i
    
    Set GetUniqueEmployeeList = dictNV
End Function

' Ham sap xep key cua dictionary (Bubble Sort)
Private Function SortKeys(dict As Object) As Variant
    Dim arrKeys() As Variant
    ReDim arrKeys(0 To dict.Count - 1)
    
    Dim i As Long, j As Long, idx As Long
    idx = 0
    Dim k As Variant
    For Each k In dict.keys
        arrKeys(idx) = k
        idx = idx + 1
    Next k
    
    Dim temp As Variant
    For i = 0 To UBound(arrKeys) - 1
        For j = i + 1 To UBound(arrKeys)
            If CLng(arrKeys(i)) > CLng(arrKeys(j)) Then
                temp = arrKeys(i)
                arrKeys(i) = arrKeys(j)
                arrKeys(j) = temp
            End If
        Next j
    Next i
    
    SortKeys = arrKeys
End Function

' Ham lay nam muc tieu tu nguoi dung
Private Function GetTargetYear(sheetName As String) As Long
    Dim strYear As String
    strYear = InputBox("Nhap nam can tao sheet '" & sheetName & "' (vd: " & Year(Now) & "):", "Chon Nam", CStr(Year(Now)))
    If strYear = "" Or Not IsNumeric(strYear) Then
        GetTargetYear = 0
    Else
        GetTargetYear = CLng(strYear)
    End If
End Function

' Xac dinh co tao sheet hay khong dua tren chuoi lua chon
Private Function ShouldCreate(choice As String, targetId As Long) As Boolean
    Dim normalized As String
    normalized = Replace(choice, " ", "")
    
    Dim parts() As String
    parts = Split(normalized, ",")
    
    Dim part As Variant
    For Each part In parts
        If part = "" Then GoTo NextPart
        If part = "7" Then ' 7 = tat ca
            ShouldCreate = True
            Exit Function
        End If
        If IsNumeric(part) Then
            If CLng(part) = targetId Then
                ShouldCreate = True
                Exit Function
            End If
        End If
NextPart:
    Next part
End Function

' Ham chung de tao cac sheet dang luoi (Tang Ca, Tien An)
Private Sub CreateGenericGridSheet(sheetName As String, validationList As String, guideText As String, headerColor As Long, Optional isNumberOnly As Boolean = False)
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim dictNV As Object
    Set dictNV = GetUniqueEmployeeList(wb)
    If dictNV Is Nothing Then Exit Sub
    
    Dim targetYear As Long
    targetYear = GetTargetYear(sheetName)
    If targetYear = 0 Then Exit Sub
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    
    ' --- HEADERS ---
    ws.Range("A1:B1").Value = Array("STT", "Ngay")
    ws.Cells(2, 2).Value = "Ma NV ->"
    
    Dim arrIDs As Variant
    arrIDs = SortKeys(dictNV)
    
    Dim i As Long, col As Long
    For i = 0 To UBound(arrIDs)
        col = i + 3
        ws.Cells(1, col).Value = "ID " & arrIDs(i) & " - " & dictNV(arrIDs(i))
        ws.Cells(2, col).Value = arrIDs(i)
    Next i
    
    ' --- DAYS ---
    Dim r As Long: r = 3
    Dim d As Date
    For d = DateSerial(targetYear, 1, 1) To DateSerial(targetYear, 12, 31)
        ws.Cells(r, 1).Value = r - 2
        ws.Cells(r, 2).Value = d
        If Weekday(d, vbSunday) = 1 Then
            ws.Range("A" & r & ":B" & r).Interior.Color = RGB(255, 242, 204)
        End If
        r = r + 1
    Next d
    
    ' --- FORMATTING ---
    Dim lastCol As Long: lastCol = UBound(arrIDs) + 3
    ws.Range("B3:B" & r - 1).NumberFormat = "dd/mm/yyyy"
    
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = headerColor
    End With
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With
    
    If validationList <> "" Then
        ws.Range("C3").Resize(r - 3, lastCol - 2).Validation.Add Type:=xlValidateList, Formula1:=validationList
    ElseIf isNumberOnly Then
        With ws.Range("C3").Resize(r - 3, lastCol - 2).Validation
            .Delete
            .Add Type:=xlValidateDecimal, Operator:=xlGreater, Formula1:="0"
            .IgnoreBlank = True
            .ErrorTitle = "Du lieu khong hop le"
            .ErrorMessage = "Vui long chi nhap so."
        End With
    End If
    
    ws.Columns.AutoFit
    ws.Columns("B").ColumnWidth = 12
    
    With ws.Range(ws.Cells(1, 1), ws.Cells(r - 1, lastCol)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ws.Range("A" & r).Value = guideText
    ws.Range("A" & r).Font.Italic = True
    
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("C3").Select
    ActiveWindow.FreezePanes = True
    ws.Range(ws.Cells(1, 1), ws.Cells(r - 1, lastCol)).AutoFilter
    
    MsgBox "Da tao/cap nhat sheet '" & sheetName & "' cho nam " & targetYear & ".", vbInformation
End Sub

'==================================================================================================
' VALIDATION FUNCTION FOR ANNUAL LEAVE CHECK
'==================================================================================================
Private Function KiemTraPhepNhap(cell As Range, annualLimit As Double) As Boolean
    Dim cellValue As String
    cellValue = UCase$(Trim$(CStr(cell.Value)))

    If cellValue = "" Then
        KiemTraPhepNhap = True
        Exit Function
    End If

    Select Case cellValue
        Case "P", "S", "C"
            Dim usedLeave As Double
            usedLeave = Application.WorksheetFunction.CountIf(cell.Worksheet.Range(cell.Worksheet.Cells(3, cell.Column), cell), "P") + _
                        Application.WorksheetFunction.CountIf(cell.Worksheet.Range(cell.Worksheet.Cells(3, cell.Column), cell), "S") * 0.5 + _
                        Application.WorksheetFunction.CountIf(cell.Worksheet.Range(cell.Worksheet.Cells(3, cell.Column), cell), "C") * 0.5
            KiemTraPhepNhap = (usedLeave <= annualLimit)
        Case Else
            KiemTraPhepNhap = True
    End Select
End Function
