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
    ' MENU PILIHAN PERIODE
    ' ========================================
    Dim menuMsg As String
    menuMsg = "Pilih periode data:" & vbCrLf & vbCrLf & _
              "1 = Hari ini" & vbCrLf & _
              "2 = Kemarin" & vbCrLf & _
              "3 = Bulan ini" & vbCrLf & _
              "4 = Bulan lalu" & vbCrLf & _
              "5 = Semester ini" & vbCrLf & _
              "6 = Semester lalu" & vbCrLf & _
              "7 = Tahun ini" & vbCrLf & _
              "8 = Tahun lalu" & vbCrLf & _
              "9 = Custom" & vbCrLf & _
              "Ketik angka 1-9:"

    Dim choice As String
    choice = InputBox(menuMsg, "Pilih Periode", "")

    ' Validasi input
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then
        MsgBox "Input harus berupa angka!", vbExclamation
        Exit Sub
    End If
    If Val(choice) < 1 Or Val(choice) > 9 Then
        MsgBox "Pilihan harus 1-9!", vbExclamation
        Exit Sub
    End If

    ' ========================================
    ' INPUT CUSTOM DATE
    ' ========================================
    Dim customStart As String, customEnd As String

    If Val(choice) = 9 Then
        customStart = InputDateDMY("Masukkan tanggal AWAL")
        If customStart = "" Then Exit Sub

        Dim displayStart As String
        displayStart = Format(CDate(Replace(customStart, "-", "/")), "dd-mm-yyyy")

        customEnd = InputDateDMY( _
            "Masukkan tanggal AKHIR" & vbCrLf & _
            "Tanggal awal: " & displayStart _
        )
        If customEnd = "" Then Exit Sub

        If CDate(Replace(customStart, "-", "/")) > CDate(Replace(customEnd, "-", "/")) Then
            MsgBox "Tanggal awal tidak boleh lebih besar dari tanggal akhir!", vbCritical
            Exit Sub
        End If
    End If

    ' ========================================
    ' CLEAR SHEET
    ' ========================================
    Dim lastRowDst As Long
    lastRowDst = Cells(Rows.Count, "b").End(xlUp).Row
    
    Dim barisDst As Long
    barisDst = 5
    
    If lastRowDst > barisDst Then
        wsTarget.Range("b" & barisDst & ":i" & lastRowDst).Clear
    End If
    
    'tampilkan start date dan enddate
    Dim startDate As Date, endDate As Date
    Dim today As Date
    today = Date
    
    Select Case Val(choice)
        Case 1 ' Hari ini
            startDate = today
            endDate = today
    
        Case 2 ' Kemarin
            startDate = today - 1
            endDate = today - 1
    
        Case 3 ' Bulan ini
            startDate = DateSerial(Year(today), Month(today), 1)
            endDate = today
    
        Case 4 ' Bulan lalu
            startDate = DateSerial(Year(today), Month(today) - 1, 1)
            endDate = DateSerial(Year(today), Month(today), 0)
    
        Case 5 ' Semester ini
            If Month(today) <= 6 Then
                startDate = DateSerial(Year(today), 1, 1)
            Else
                startDate = DateSerial(Year(today), 7, 1)
            End If
            endDate = today
    
        Case 6 ' Semester lalu
            If Month(today) <= 6 Then
                startDate = DateSerial(Year(today) - 1, 7, 1)
                endDate = DateSerial(Year(today) - 1, 12, 31)
            Else
                startDate = DateSerial(Year(today), 1, 1)
                endDate = DateSerial(Year(today), 6, 30)
            End If
    
        Case 7 ' Tahun ini
            startDate = DateSerial(Year(today), 1, 1)
            endDate = today
    
        Case 8 ' Tahun lalu
            startDate = DateSerial(Year(today) - 1, 1, 1)
            endDate = DateSerial(Year(today) - 1, 12, 31)
    
        Case 9 ' Custom
            startDate = CDate(Replace(customStart, "-", "/"))
            endDate = CDate(Replace(customEnd, "-", "/"))
    End Select
    
    ' Tampilkan ke sheet
    With wsTarget
        .Range("D2").Value = startDate
        .Range("D3").Value = endDate
        .Range("D2:D3").NumberFormat = "dd mmm yyyy"
    End With


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
    ' JALANKAN PYTHON
    ' ========================================
    Dim cmd As String

    If Val(choice) = 9 Then
        cmd = "cmd /c START /WAIT ""Scraper Log"" """ & pythonPath & """ """ & _
              scriptPath & """ 9 " & customStart & " " & customEnd
    Else
        cmd = "cmd /c START /WAIT ""Scraper Log"" """ & pythonPath & """ """ & _
              scriptPath & """ " & choice
        
    End If

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

    wsTarget.ListObjects.Add( _
        xlSrcRange, _
        wsTarget.Range("$B$" & barisDst & ":$H$" & lastRowDataScrape), , xlYes _
    ).Name = "TableLogAktivitasUser"

    ' Copy master
    Sheets("Master").Range("i22:i22").Copy
    Sheets("Log Aktivitas User").Range("i" & barisDst).PasteSpecial xlPasteAll
    
    Application.CutCopyMode = False

    Sheets("Log Aktivitas User").Range("i" & barisDst + 1).FormulaR1C1 = _
        Sheets("Master").Range("i23").FormulaR1C1
        
    Columns("i:i").EntireColumn.AutoFit

    ' Format tanggal
    With Sheets("Log Aktivitas User")
        
        .Range("i" & barisDst + 1, .Cells(.Rows.Count, "i").End(xlUp)).NumberFormat = _
            "dd mmm yyyy"
    End With

    ' Hapus file hasil
    Kill outputXlsx

    MsgBox "Data berhasil diperbarui!", vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description, vbCritical

End Sub