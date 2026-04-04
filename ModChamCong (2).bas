
Attribute VB_Name = "ModChamCong"
Sub XuLyChamCong()
    
    ' Thu thap Thang Nam tu nguoi dung
    Dim inputStr As String
    inputStr = InputBox("Nhap thang va nam can tao (MM-YYYY):" & vbCrLf & "Vi du: 04-2026", "Chon Thang Cham Cong", Format(Date, "mm-yyyy"))
    If inputStr = "" Then Exit Sub
    
    Dim targetMonth As Integer, targetYear As Integer
    Dim parts() As String
    If InStr(inputStr, "-") > 0 Then
        parts = Split(inputStr, "-")
    ElseIf InStr(inputStr, "/") > 0 Then
        parts = Split(inputStr, "/")
    Else
        MsgBox "Dinh dang khong hop le. Vi du: 04-2026", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    targetMonth = Val(parts(0))
    targetYear = Val(parts(1))
    On Error GoTo 0
    
    If targetMonth < 1 Or targetMonth > 12 Or targetYear < 2000 Then
        MsgBox "Thang hoac nam khong hop le!", vbExclamation
        Exit Sub
    End If
    
    Dim startDate As Date, endDate As Date
    startDate = DateSerial(targetYear, targetMonth, 1)
    endDate = DateSerial(targetYear, targetMonth + 1, 0)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Doc du lieu tu Sheet 1 cua file dang mo
    Dim thisWb As Workbook
    Set thisWb = ThisWorkbook
    
    Dim csvWs As Worksheet
    Set csvWs = thisWb.Sheets(1)
    
    Dim lastRow As Long
    lastRow = csvWs.Cells(csvWs.Rows.Count, "A").End(xlUp).Row
    
    ' Read data into arrays
    Dim dataCount As Long
    dataCount = lastRow - 1
    
    If dataCount < 1 Then
        MsgBox "Sheet 1 khong co du lieu!", vbExclamation
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    ' Store raw data
    ReDim personID(1 To dataCount) As String
    ReDim personName(1 To dataCount) As String
    ReDim dept(1 To dataCount) As String
    ReDim timeStamp(1 To dataCount) As Date
    ReDim dateOnly(1 To dataCount) As Date
    
    Dim i As Long
    For i = 2 To lastRow
        Dim idx As Long
        idx = i - 1
        personID(idx) = CStr(csvWs.Cells(i, 1).Value)
        personName(idx) = CStr(csvWs.Cells(i, 2).Value)
        dept(idx) = CStr(csvWs.Cells(i, 3).Value)
        
        Dim rawTime As Variant
        rawTime = csvWs.Cells(i, 4).Value
        If IsDate(rawTime) Then
            timeStamp(idx) = CDate(rawTime)
        Else
            timeStamp(idx) = CDate(CStr(rawTime))
        End If
        dateOnly(idx) = DateSerial(Year(timeStamp(idx)), Month(timeStamp(idx)), Day(timeStamp(idx)))
    Next i
    
    ' Process: find first/last per person per day
    ' Use Dictionary
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim dictUniquePersons As Object
    Set dictUniquePersons = CreateObject("Scripting.Dictionary")
    
    For i = 1 To dataCount
        Dim eID As String, eName As String, eDept As String
        eID = personID(i)
        eName = personName(i)
        eDept = dept(i)
        
        Dim uKey As String
        uKey = eID & "|" & eName
        If Not dictUniquePersons.Exists(uKey) Then
            dictUniquePersons.Add uKey, Array(eID, eName, eDept)
        End If
        
        Dim key As String
        key = eID & "|" & eName & "|" & Format(dateOnly(i), "yyyy-mm-dd")
        
        If Not dict.Exists(key) Then
            dict.Add key, Array(timeStamp(i), timeStamp(i), personID(i), personName(i), dept(i), dateOnly(i))
        Else
            Dim arr As Variant
            arr = dict(key)
            If timeStamp(i) < CDate(arr(0)) Then arr(0) = timeStamp(i)
            If timeStamp(i) > CDate(arr(1)) Then arr(1) = timeStamp(i)
            dict(key) = arr
        End If
    Next i
    
    ' === CHAM CONG SHEET (Sheet 2) ===
    Dim wsChamCong As Worksheet
    On Error Resume Next
    Set wsChamCong = thisWb.Sheets("Cham Cong")
    On Error GoTo 0
    If Not wsChamCong Is Nothing Then
        wsChamCong.Delete
    End If
    Set wsChamCong = thisWb.Sheets.Add(After:=csvWs)
    wsChamCong.Name = "Cham Cong"
    
    ' Headers
    Dim headers As Variant
    headers = Array("STT", "Ma NV", "Ho Ten", "Ngay", "Gio Vao", "Gio Ra", "So Gio Lam")
    
    Dim c As Long
    For c = 0 To 6
        wsChamCong.Cells(1, c + 1).Value = headers(c)
    Next c
    
    ' Format headers
    With wsChamCong.Range("A1:G1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 11
        .Font.Name = "Arial"
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' Sort keys
    Dim keys As Variant
    keys = dictUniquePersons.keys
    
    ' Sort unique persons by employee code
    Dim tempKey As String
    Dim j As Long
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            Dim arrI As Variant, arrJ As Variant
            arrI = dictUniquePersons(keys(i))
            arrJ = dictUniquePersons(keys(j))
            Dim idI As Long, idJ As Long
            idI = CLng(Val(Replace(CStr(arrI(0)), "'", "")))
            idJ = CLng(Val(Replace(CStr(arrJ(0)), "'", "")))
            If idI > idJ Then
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i
    
    ' Write data
    Dim r As Long
    r = 2
    Dim sttCount As Long
    sttCount = 1
    
    For i = 0 To UBound(keys)
        Dim empArr As Variant
        empArr = dictUniquePersons(keys(i))
        eID = CStr(empArr(0))
        eName = CStr(empArr(1))
        
        Dim d As Date
        For d = startDate To endDate
            wsChamCong.Cells(r, 1).Value = sttCount
            sttCount = sttCount + 1
            wsChamCong.Cells(r, 2).Value = CLng(Val(Replace(eID, "'", "")))
            wsChamCong.Cells(r, 3).Value = eName
            wsChamCong.Cells(r, 4).Value = d
            wsChamCong.Cells(r, 4).NumberFormat = "dd/mm/yyyy"
            
            Dim wkDay As Integer
            wkDay = Weekday(d, vbSunday)
            Dim isSunday As Boolean
            isSunday = (wkDay = 1)
            
            Dim searchKey As String
            searchKey = eID & "|" & eName & "|" & Format(d, "yyyy-mm-dd")
            
            If dict.Exists(searchKey) Then
                ' Worked
                Dim dataArr As Variant
                dataArr = dict(searchKey)
                
                wsChamCong.Cells(r, 5).Value = CDate(dataArr(0))
                wsChamCong.Cells(r, 5).NumberFormat = "hh:mm"
                wsChamCong.Cells(r, 6).Value = CDate(dataArr(1))
                wsChamCong.Cells(r, 6).NumberFormat = "hh:mm"
                
                Dim hoursWorked As Double
                hoursWorked = (CDate(dataArr(1)) - CDate(dataArr(0))) * 24
                If hoursWorked < 0 Then hoursWorked = 0
                wsChamCong.Cells(r, 7).Value = Round(hoursWorked, 1)
                wsChamCong.Cells(r, 7).NumberFormat = "0.0"
                
                ' Light red if Sunday
                If isSunday Then
                    wsChamCong.Range("A" & r & ":G" & r).Interior.Color = RGB(255, 204, 204)
                End If
                
            Else
                ' Absent
                If isSunday Then
                    ' Chu nhat dc nghi, van de trong thoi gian
                    wsChamCong.Range("A" & r & ":G" & r).Interior.Color = RGB(255, 204, 204)
                    wsChamCong.Cells(r, 5).Value = ""
                    wsChamCong.Cells(r, 6).Value = ""
                    wsChamCong.Cells(r, 7).Value = ""
                Else
                    ' Mon-Sat absent -> to do, ghi ngay vao cot ranh
                    wsChamCong.Range("A" & r & ":G" & r).Interior.Color = RGB(255, 0, 0)
                    wsChamCong.Range("A" & r & ":G" & r).Font.Color = RGB(255, 255, 255)
                    wsChamCong.Cells(r, 5).Value = "'" & Format(d, "dd/mm/yyyy") ' "ghi ngày trong ô đó luôn"
                    wsChamCong.Cells(r, 5).HorizontalAlignment = xlCenter
                    wsChamCong.Cells(r, 6).Value = "'" & Format(d, "dd/mm/yyyy")
                    wsChamCong.Cells(r, 6).HorizontalAlignment = xlCenter
                    wsChamCong.Cells(r, 7).Value = ""
                End If
            End If
            
            r = r + 1
        Next d
        
        ' Them dong trong giua moi nguoi
        If i < UBound(keys) Then
            r = r + 1
        End If
    Next i
    
    Dim totalRows As Long
    totalRows = r - 1

    ' Format data area
    With wsChamCong.Range("A2:G" & totalRows)
        .Font.Name = "Arial"
        .Font.Size = 11
        .VerticalAlignment = xlCenter
    End With
    
    wsChamCong.Range("A2:A" & totalRows).HorizontalAlignment = xlCenter
    wsChamCong.Range("B2:B" & totalRows).HorizontalAlignment = xlCenter
    wsChamCong.Range("D2:G" & totalRows).HorizontalAlignment = xlCenter
    
    ' Borders
    With wsChamCong.Range("A1:G" & totalRows).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Alternating row colors by person and clear blank row borders
    Dim prevName As String
    Dim colorToggle As Boolean
    prevName = ""
    colorToggle = False
    For i = 2 To totalRows
        If wsChamCong.Cells(i, 3).Value <> "" Then
            If wsChamCong.Cells(i, 3).Value <> prevName Then
                colorToggle = Not colorToggle
                prevName = wsChamCong.Cells(i, 3).Value
            End If
            If colorToggle Then
                ' Only color if it has no specific background color yet (not Sunday or Absent)
                If wsChamCong.Range("A" & i).Interior.ColorIndex = xlNone Or wsChamCong.Range("A" & i).Interior.Color = 16777215 Then
                    wsChamCong.Range("A" & i & ":G" & i).Interior.Color = RGB(221, 235, 247)
                End If
            End If
        Else
            ' Remove borders and background color from the blank spacer rows
            wsChamCong.Range("A" & i & ":G" & i).Interior.ColorIndex = xlNone
            wsChamCong.Range("A" & i & ":G" & i).Borders.LineStyle = xlNone
        End If
    Next i
    
    ' Column widths
    wsChamCong.Columns("A").ColumnWidth = 6
    wsChamCong.Columns("B").ColumnWidth = 10
    wsChamCong.Columns("C").ColumnWidth = 28
    wsChamCong.Columns("D").ColumnWidth = 14
    wsChamCong.Columns("E").ColumnWidth = 12
    wsChamCong.Columns("F").ColumnWidth = 12
    wsChamCong.Columns("G").ColumnWidth = 14
    
    ' Freeze panes
    wsChamCong.Activate
    wsChamCong.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    ' Auto filter
    wsChamCong.Range("A1:G" & totalRows).AutoFilter
    
    ' === TONG HOP SHEET (Sheet 3) ===
    Dim wsTongHop As Worksheet
    On Error Resume Next
    Set wsTongHop = thisWb.Sheets("Tong Hop")
    On Error GoTo 0
    If Not wsTongHop Is Nothing Then
        wsTongHop.Delete
    End If
    Set wsTongHop = thisWb.Sheets.Add(After:=wsChamCong)
    wsTongHop.Name = "Tong Hop"
    
    ' Get unique persons
    Dim dictPerson As Object
    Set dictPerson = CreateObject("Scripting.Dictionary")
    
    Dim dKeys As Variant
    dKeys = dict.keys
    
    For i = 0 To UBound(dKeys)
        Dim dArr As Variant
        dArr = dict(dKeys(i))
        Dim pKey As String
        pKey = CStr(dArr(2)) & "|" & CStr(dArr(3))
        If Not dictPerson.Exists(pKey) Then
            dictPerson.Add pKey, Array(CStr(dArr(2)), CStr(dArr(3)), 0, 0#)
        End If
        Dim pArr As Variant
        pArr = dictPerson(pKey)
        pArr(2) = CLng(pArr(2)) + 1
        Dim hrs As Double
        hrs = (CDate(dArr(1)) - CDate(dArr(0))) * 24
        If hrs < 0 Then hrs = 0
        pArr(3) = CDbl(pArr(3)) + hrs
        dictPerson(pKey) = pArr
    Next i
    
    ' Headers
    Dim thHeaders As Variant
    thHeaders = Array("STT", "Ma NV", "Ho Ten", "Tong Ngay Cong", "Tong Gio Lam", "TB Gio/Ngay")
    
    For c = 0 To 5
        wsTongHop.Cells(1, c + 1).Value = thHeaders(c)
    Next c
    
    With wsTongHop.Range("A1:F1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 11
        .Font.Name = "Arial"
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' Sort person keys by employee code
    Dim pKeys As Variant
    pKeys = dictPerson.keys
    For i = 0 To UBound(pKeys) - 1
        For j = i + 1 To UBound(pKeys)
            Dim pArrI As Variant, pArrJ As Variant
            pArrI = dictPerson(pKeys(i))
            pArrJ = dictPerson(pKeys(j))
            Dim pidI As Long, pidJ As Long
            pidI = CLng(Val(Replace(CStr(pArrI(0)), "'", "")))
            pidJ = CLng(Val(Replace(CStr(pArrJ(0)), "'", "")))
            If pidI > pidJ Then
                Dim tmpKey As String
                tmpKey = pKeys(i)
                pKeys(i) = pKeys(j)
                pKeys(j) = tmpKey
            End If
        Next j
    Next i
    
    r = 2
    For i = 0 To UBound(pKeys)
        pArr = dictPerson(pKeys(i))
        wsTongHop.Cells(r, 1).Value = i + 1
        wsTongHop.Cells(r, 2).Value = Replace(CStr(pArr(0)), "'", "")
        wsTongHop.Cells(r, 3).Value = CStr(pArr(1))
        wsTongHop.Cells(r, 4).Value = CLng(pArr(2))
        wsTongHop.Cells(r, 5).Value = Round(CDbl(pArr(3)), 1)
        wsTongHop.Cells(r, 5).NumberFormat = "0.0"
        If CLng(pArr(2)) > 0 Then
            wsTongHop.Cells(r, 6).Value = Round(CDbl(pArr(3)) / CLng(pArr(2)), 1)
        Else
            wsTongHop.Cells(r, 6).Value = 0
        End If
        wsTongHop.Cells(r, 6).NumberFormat = "0.0"
        r = r + 1
    Next i
    
    Dim totalRowsTH As Long
    totalRowsTH = r - 1
    
    With wsTongHop.Range("A2:F" & totalRowsTH)
        .Font.Name = "Arial"
        .Font.Size = 11
        .VerticalAlignment = xlCenter
    End With
    
    wsTongHop.Range("A2:A" & totalRowsTH).HorizontalAlignment = xlCenter
    wsTongHop.Range("B2:B" & totalRowsTH).HorizontalAlignment = xlCenter
    wsTongHop.Range("D2:F" & totalRowsTH).HorizontalAlignment = xlCenter
    
    With wsTongHop.Range("A1:F" & totalRowsTH).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Alternating rows
    For i = 2 To totalRowsTH
        If i Mod 2 = 0 Then
            wsTongHop.Range("A" & i & ":F" & i).Interior.Color = RGB(221, 235, 247)
        End If
    Next i
    
    wsTongHop.Columns("A").ColumnWidth = 6
    wsTongHop.Columns("B").ColumnWidth = 10
    wsTongHop.Columns("C").ColumnWidth = 28
    wsTongHop.Columns("D").ColumnWidth = 16
    wsTongHop.Columns("E").ColumnWidth = 14
    wsTongHop.Columns("F").ColumnWidth = 14
    
    wsChamCong.Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Hoan thanh!" & vbCrLf & _
           "- Sheet 'Cham Cong': " & dict.Count & " dong (gio vao/ra)" & vbCrLf & _
           "- Sheet 'Tong Hop': " & dictPerson.Count & " nhan vien", vbInformation, "Xu Ly Cham Cong"
    
End Sub
