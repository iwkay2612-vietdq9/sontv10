' ==================================================================================
' HƯỚNG DẪN CÀI ĐẶT:
' 1. Mở Excel, tạo một file mới.
' 2. Lưu file này dạng "Excel Macro-Enabled Workbook (.xlsm)".
' 3. Đổi tên Sheet1 thành "Control".
' 4. Thiết kế Sheet "Control" như sau:
'    - Ô B2: Đường dẫn file nguồn (Ví dụ: C:\Data\Source.xlsx)
'    - Ô B3: Thư mục xuất file (Ví dụ: C:\Data\Output\)
'    - Ô B5: "Tên Cột (Header)"
'    - Ô C5: "Giá trị (cách nhau dấu phẩy)"
'    - Từ B6 đến B10: Nhập tên cột muốn lọc.
'    - Từ C6 đến C10: Nhập giá trị muốn lọc.
'    - Ô B12: Tên cột để Tách file (nếu chọn chế độ tách)
'    - Ô B13: Chế độ (Nhập 1 để Gộp, 2 để Tách Tự Động)
' 5. Nhấn Alt + F11 -> Insert -> Module -> Dán toàn bộ code bên dưới vào.
' 6. Vẽ một nút bấm (Shape) ở sheet Control, chuột phải chọn Assign Macro -> MainRun.
' ==================================================================================

Option Explicit

Sub MainRun()
    Dim wsControl As Worksheet
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim sourcePath As String, outputFolder As String
    Dim splitColName As String
    Dim mode As Integer
    Dim lastRow As Long, i As Integer
    Dim dictFilters As Object
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wsControl = ThisWorkbook.Sheets("Control")
    Set dictFilters = CreateObject("Scripting.Dictionary")
    
    ' ======= 1. ĐỌC CẤU HÌNH =======
    sourcePath = wsControl.Range("B2").Value
    outputFolder = wsControl.Range("B3").Value
    splitColName = wsControl.Range("B12").Value
    mode = wsControl.Range("B13").Value ' 1=Merge, 2=Split
    
    ' Validate inputs
    If Dir(sourcePath) = "" Then
        MsgBox "File nguồn không tồn tại!", vbCritical
        Exit Sub
    End If
    If Right(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"
    If Dir(outputFolder, vbDirectory) = "" Then
        MsgBox "Thư mục xuất không tồn tại!", vbCritical
        Exit Sub
    End If
    
    ' Đọc Filter (B6:C10)
    For i = 6 To 10
        Dim colName As String, valStr As String
        colName = Trim(wsControl.Cells(i, 2).Value)
        valStr = Trim(wsControl.Cells(i, 3).Value)
        
        If colName <> "" And valStr <> "" Then
            dictFilters.Add colName, Split(valStr, ",")
        End If
    Next i
    
    ' ======= 2. MỞ FILE NGUỒN =======
    Set wbSource = Workbooks.Open(sourcePath)
    Set wsSource = wbSource.Sheets(1) ' Giả sử dữ liệu ở Sheet 1
    
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    ' Map Header Name -> Column Index
    Dim headers As Range
    Dim headerCell As Range
    Dim dictCols As Object
    Set dictCols = CreateObject("Scripting.Dictionary")
    
    Set headers = wsSource.Range("A1", wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft))
    For Each headerCell In headers
        dictCols(headerCell.Value) = headerCell.Column
    Next headerCell
    
    ' ======= 3. ÁP DỤNG FILTER =======
    Dim key As Variant, arrVal As Variant
    For Each key In dictFilters.Keys
        If dictCols.Exists(key) Then
            arrVal = dictFilters(key)
            ' AutoFilter supports array of strings for exact match (xlFilterValues)
            ' Trim spaces
            Dim j As Integer
            For j = LBound(arrVal) To UBound(arrVal)
                arrVal(j) = Trim(arrVal(j))
            Next j
            
            wsSource.Range("A1").AutoFilter Field:=dictCols(key), Criteria1:=arrVal, Operator:=xlFilterValues
        End If
    Next key
    
    ' ======= 4. XỬ LÝ E XUẤT =======
    
    If mode = 1 Then
        ' MODE 1: GỘP (MERGE) - Chỉ cần Copy Visible cells ra file mới
        SaveVisibleData wsSource, outputFolder & "Filtered_Output.xlsx"
        MsgBox "Đã xuất file gộp thành công!", vbInformation
        
    ElseIf mode = 2 Then
        ' MODE 2: TÁCH TỰ ĐỘNG (SPLIT)
        If Not dictCols.Exists(splitColName) Then
            MsgBox "Không tìm thấy cột để tách: " & splitColName, vbCritical
            wbSource.Close False
            Exit Sub
        End If
        
        Dim splitColIdx As Integer
        splitColIdx = dictCols(splitColName)
        
        ' Lấy danh sách giá trị duy nhất trong cột tách (chỉ lấy cell visible)
        Dim rngVisible As Range, cell As Range
        Dim uniqueVals As Object
        Set uniqueVals = CreateObject("Scripting.Dictionary")
        
        On Error Resume Next
        Set rngVisible = wsSource.Range(wsSource.Cells(2, splitColIdx), wsSource.Cells(lastRow, splitColIdx)).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        If rngVisible Is Nothing Then
            MsgBox "Không có dữ liệu nào sau khi lọc!", vbExclamation
            wbSource.Close False
            Exit Sub
        End If
        
        For Each cell In rngVisible
            Dim v As String
            v = Trim(CStr(cell.Value))
            If v <> "" And Not uniqueVals.Exists(v) Then uniqueVals.Add v, 1
        Next cell
        
        ' Loop từng giá trị tách -> Filter thêm lần nữa -> Save
        Dim splitVal As Variant
        For Each splitVal In uniqueVals.Keys
            ' Vẫn giữ các filter cũ, chỉ thêm filter cho cột split
            wsSource.Range("A1").AutoFilter Field:=splitColIdx, Criteria1:=splitVal
            
            ' Save
            Dim safeName As String
            safeName = Replace(splitVal, "/", "_")
            safeName = Replace(safeName, "\", "_")
            SaveVisibleData wsSource, outputFolder & safeName & ".xlsx"
        Next splitVal
        
        MsgBox "Đã tách thành công " & uniqueVals.Count & " files!", vbInformation
    End If
    
    wbSource.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub SaveVisibleData(ws As Worksheet, savePath As String)
    Dim wbNew As Workbook
    Set wbNew = Workbooks.Add
    
    ' Copy Visible cells & Formats
    ' Phương pháp: Copy toàn bộ sheet, sau đó xóa row ẩn? 
    ' Hoặc Copy SpecialCells(xlCellTypeVisible) -> Paste. 
    ' Copy Visible thường an toàn nhất cho Format.
    
    ws.Cells.SpecialCells(xlCellTypeVisible).Copy
    wbNew.Sheets(1).Range("A1").PasteSpecial xlPasteAll ' Paste formulas, formats
    wbNew.Sheets(1).Range("A1").PasteSpecial xlPasteColumnWidths
    
    wbNew.SaveAs savePath
    wbNew.Close False
End Sub
