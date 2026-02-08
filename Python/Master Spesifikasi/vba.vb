Sub ScrapeData()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler
    
    'setting data
    'ws dest
    Dim wsDest As Worksheet
    Set wsDest = SheetMasterSpesifikasi
    
    'kolom lastrow sheet dst
    Dim kolomLastRowDest As String
    kolomLastRowDest = "a"
    
    'baris pertama sheet dest
    Dim barisPertamaDest As Long
    barisPertamaDest = 1
    
    'folder python
    Dim folderPython As String
    folderPython = "Master Spesifikasi"
    
    'kolom paste sheet dest
    Dim kolomPasteSheetDest As String
    kolomPasteSheetDest = "a"
    
    'kolom lastrow data src
    Dim kolomLastRowDataSrc As String
    kolomLastRowDataSrc = "a"
    
    'nama table sheet dest
    Dim namaTableSheetDest As String
    namaTableSheetDest = "TableMasterSpesifikasi"
    
    
    ' ========================================
    ' CEK FILE SUDAH DISIMPAN
    ' ========================================
    If ThisWorkbook.Path = "" Then
        MsgBox "Simpan file Excel terlebih dahulu!", vbCritical
        GoTo CleanUp
    End If

    ' ========================================
    ' CLEAR SHEET
    ' ========================================
    Dim lastRowDest As Long
    lastRowDest = wsDest.Cells(Rows.Count, kolomLastRowDest).End(xlUp).Row
    
    ' Clear data lama (dynamic column)
    If lastRowDest > barisPertamaDest Then
        Dim lastColDst As Long
        lastColDst = wsDest.Cells(barisPertamaDest, Columns.Count).End(xlToLeft).Column
        wsDest.Range(wsDest.Cells(barisPertamaDest, kolomLastRowDest), wsDest.Cells(lastRowDest, lastColDst)).Clear
    End If

    ' ========================================
    ' PATH SETUP
    ' ========================================
    Dim wbPath As String
    wbPath = ThisWorkbook.path

    Dim pythonPath As String
    Dim scriptPath As String
    Dim outputXlsx As String

    pythonPath = "python"
    scriptPath = wbPath & "\Python\" & folderPython & "\scraper.py"
    outputXlsx = wbPath & "\Python\" & folderPython & "\temp.xlsx"

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

    ' ===============
    ' JALANKAN PYTHON
    ' ===============
    Dim cmd As String
    cmd = "cmd /c START /WAIT ""Scraper Log"" """ & pythonPath & """ """ & scriptPath & """"

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
    wsDest.Range(kolomPasteSheetDest & barisPertamaDest).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    wbSrc.Close SaveChanges:=False

    ' ========================================
    ' FORMAT SHEET
    ' ========================================
    wsDest.UsedRange.Columns.AutoFit

    Dim col As Range
    For Each col In wsDest.UsedRange.Columns
        If col.ColumnWidth > 50 Then col.ColumnWidth = 50
    Next col

    ' ========================================
    ' BUAT TABLE
    ' ========================================
    Dim lastRowDataScrape As Long
    lastRowDataScrape = wsDest.Cells(wsDest.Rows.Count, kolomLastRowDataSrc).End(xlUp).Row

    ' Hapus table lama jika ada
    On Error Resume Next
    wsDest.ListObjects(namaTableSheetDest).Delete
    On Error GoTo ErrorHandler

    ' Buat table baru
    Dim lastColTable As Long
    lastColTable = wsDest.Cells(barisPertamaDest, Columns.Count).End(xlToLeft).Column
    
    wsDest.ListObjects.Add( _
        xlSrcRange, _
        wsDest.Range(wsDest.Cells(barisPertamaDest, kolomPasteSheetDest), wsDest.Cells(lastRowDataScrape, lastColTable)), _
        , xlYes _
    ).Name = namaTableSheetDest

    ' ========================================
    ' COPY FORMULA DARI MASTER (jika ada)
    ' ========================================
    On Error Resume Next
    If Not Sheets("Master") Is Nothing Then
        ' Copy master
        Sheets("Master").Range("i22:i22").Copy
        Sheets("Master Spesifikasi").Range("i" & barisPertamaDest).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        
        Sheets("Master Spesifikasi").Range("i" & barisPertamaDest + 1).FormulaR1C1 = _
            Sheets("Master").Range("i23").FormulaR1C1
        
        
        Columns("i:i").EntireColumn.AutoFit
        
        ' Format tanggal
        With Sheets("Master Spesifikasi")
            .Range("i" & barisPertamaDest + 1, .Cells(.Rows.Count, "I").End(xlUp)).NumberFormat = _
                "dd mmm yyyy"

        End With
    End If
    On Error GoTo ErrorHandler

    ' Hapus file hasil
    Kill outputXlsx

'    MsgBox "Data berhasil diperbarui!", _
'           vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description, vbCritical
End Sub