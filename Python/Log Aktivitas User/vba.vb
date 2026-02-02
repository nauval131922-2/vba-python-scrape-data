Sub ScrapeData_LogAktivitasUser()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    ' ========================================
    ' CEK FILE SUDAH DISIMPAN
    ' ========================================
    If ThisWorkbook.Path = "" Then
        MsgBox "Simpan file Excel terlebih dahulu!", vbCritical
        GoTo CleanUp
    End If

    Dim wsTarget As Worksheet
    Set wsTarget = ActiveSheet

    ' ========================================
    ' AMBIL TANGGAL DARI CELL D2 & D3
    ' ========================================
    Dim startDate As Date, endDate As Date
    
    ' Validasi cell D2 (Start Date)
    If Not IsDate(wsTarget.Range("D2").Value) Then
        MsgBox "Cell D2 harus berisi tanggal yang valid!", vbCritical
        Exit Sub
    End If
    
    ' Validasi cell D3 (End Date)
    If Not IsDate(wsTarget.Range("D3").Value) Then
        MsgBox "Cell D3 harus berisi tanggal yang valid!", vbCritical
        Exit Sub
    End If
    
    startDate = wsTarget.Range("D2").Value
    endDate = wsTarget.Range("D3").Value
    
    ' Validasi range tanggal
    If startDate > endDate Then
        MsgBox "Tanggal awal (D2) tidak boleh lebih besar dari tanggal akhir (D3)!", vbCritical
        Exit Sub
    End If

    ' ========================================
    ' CLEAR SHEET
    ' ========================================
    Dim lastRowDst As Long
    lastRowDst = Cells(Rows.Count, "B").End(xlUp).Row
    Dim barisDst As Long
    barisDst = 5
    
    ' Clear data lama (dynamic column)
    If lastRowDst > barisDst Then
        Dim lastColDst As Long
        lastColDst = wsTarget.Cells(barisDst, Columns.Count).End(xlToLeft).Column
        wsTarget.Range(wsTarget.Cells(barisDst, 2), wsTarget.Cells(lastRowDst, lastColDst)).Clear
    End If

    ' ========================================
    ' PATH SETUP
    ' ========================================
    Dim wbPath As String
    wbPath = ThisWorkbook.Path

    Dim pythonPath As String
    Dim scriptPath As String
    Dim outputXlsx As String

    pythonPath = "python"
    scriptPath = wbPath & "\Python\Log Aktivitas User\scraper.py"
    outputXlsx = wbPath & "\Python\Log Aktivitas User\temp.xlsx"

    ' Cek script Python
    If Dir(scriptPath) = "" Then
        MsgBox "File Python tidak ditemukan!" & vbCrLf & _
               "Path: " & scriptPath, vbCritical
        GoTo CleanUp
    End If

    ' Hapus file output lama
    On Error Resume Next
    Kill outputXlsx
    On Error GoTo ErrorHandler

    ' ========================================
    ' JALANKAN PYTHON dengan parameter tanggal
    ' ========================================
    Dim customStart As String, customEnd As String
    customStart = Format(startDate, "yyyy-mm-dd")
    customEnd = Format(endDate, "yyyy-mm-dd")
    
    Dim cmd As String
    ' Gunakan choice=9 (custom) dan pass tanggal
    cmd = "cmd /c START /WAIT ""Scraper Log"" """ & pythonPath & """ """ & _
          scriptPath & """ 9 " & customStart & " " & customEnd

    CreateObject("WScript.Shell").Run cmd, 1, True

    ' ========================================
    ' CEK FILE HASIL
    ' ========================================
    Dim waitCount As Integer
    waitCount = 0
    Do While Dir(outputXlsx) = "" And waitCount < 60
        Application.Wait Now + TimeValue("0:00:01")
        waitCount = waitCount + 1
        DoEvents
    Loop

    If Dir(outputXlsx) = "" Then
        MsgBox "File hasil tidak ditemukan setelah menunggu " & waitCount & " detik!" & vbCrLf & _
               "Path: " & outputXlsx, vbCritical
        GoTo CleanUp
    End If

    ' ========================================
    ' IMPORT DATA
    ' ========================================
    Dim wbSrc As Workbook
    Set wbSrc = Workbooks.Open(outputXlsx, ReadOnly:=True)
    wbSrc.Sheets(1).UsedRange.Copy
    wsTarget.Range("B" & barisDst).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    wbSrc.Close SaveChanges:=False

    ' ========================================
    ' FORMAT SHEET
    ' ========================================
    wsTarget.UsedRange.Columns.AutoFit

    Dim col As Range
    For Each col In wsTarget.UsedRange.Columns
        If col.ColumnWidth > 50 Then col.ColumnWidth = 50
    Next col

    ' ========================================
    ' BUAT TABLE
    ' ========================================
    Dim lastRowDataScrape As Long
    lastRowDataScrape = wsTarget.Cells(wsTarget.Rows.Count, "D").End(xlUp).Row

    ' Hapus table lama jika ada
    On Error Resume Next
    wsTarget.ListObjects("TableLogAktivitasUser").Delete
    On Error GoTo ErrorHandler

    ' Buat table baru
    Dim lastColTable As Long
    lastColTable = wsTarget.Cells(barisDst, Columns.Count).End(xlToLeft).Column
    
    wsTarget.ListObjects.Add( _
        xlSrcRange, _
        wsTarget.Range(wsTarget.Cells(barisDst, 2), wsTarget.Cells(lastRowDataScrape, lastColTable)), _
        , xlYes _
    ).Name = "TableLogAktivitasUser"

    ' ========================================
    ' COPY FORMULA DARI MASTER (jika ada)
    ' ========================================
    On Error Resume Next
    If Not Sheets("Master") Is Nothing Then
        ' Copy master
        Sheets("Master").Range("i22:i22").Copy
        Sheets("Log Aktivitas User").Range("i" & barisDst).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        
        Sheets("Log Aktivitas User").Range("i" & barisDst + 1).FormulaR1C1 = _
            Sheets("Master").Range("i23").FormulaR1C1
        
        
        Columns("i:i").EntireColumn.AutoFit
        
        ' Format tanggal
        With Sheets("Log Aktivitas User")
            .Range("i" & barisDst + 1, .Cells(.Rows.Count, "I").End(xlUp)).NumberFormat = _
                "dd mmm yyyy"

        End With
    End If
    On Error GoTo ErrorHandler

    ' Hapus file hasil
    Kill outputXlsx

    MsgBox "Data berhasil diperbarui!" & vbCrLf & vbCrLf & _
           "Periode: " & Format(startDate, "dd-mm-yyyy") & " s/d " & Format(endDate, "dd-mm-yyyy"), _
           vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
