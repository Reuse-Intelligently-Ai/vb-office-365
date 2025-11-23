Option Explicit

Private Function BuatSheetJurnal() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Jurnal")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Jurnal"

        ws.Range("A1:F1").Value = Array("No", "Tanggal", "Akun", "Akun", "Debit", "Kredit")
        With ws.Range("A1:F1")
            .Interior.Color = RGB(79, 129, 189)
            .Font.Bold = True
            .Font.Color = vbWhite
            .Borders.LineStyle = xlContinuous
        End With
    End If

    Set BuatSheetJurnal = ws
End Function

Private Function FungsiNomorUrutTerakhir(ws As Worksheet) As Long
    Dim i As Long
    
    'Mulai dari baris terakhir yang ada datanya
    i = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Cari ke atas sampai ketemu angka
    Do While i > 1 And Not IsNumeric(ws.Cells(i, 1).Value)
        i = i - 1
    Loop

    'Jika belum ada nomor sama sekali (hanya header)
    If i <= 1 Then
        FungsiNomorUrutTerakhir = 1
    Else
        FungsiNomorUrutTerakhir = ws.Cells(i, 1).Value + 1
    End If
End Function


Private Sub TambahkanKeterangan(ws As Worksheet, ByVal ket As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ' ========= MERGE KOLOM A & B DUA BARIS DI ATAS ============
    If lastRow > 3 Then
        With ws.Range("A" & lastRow & ":A" & lastRow - 2)
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With ws.Range("B" & lastRow & ":B" & lastRow - 2)
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If

    ' ========= MERGE KETERANGAN PADA C-D SAJA ==================
    With ws.Range("C" & lastRow & ":D" & lastRow)
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Value = "(" & ket & ")"
    End With

    ' ========= BORDER ==========================================
    ws.Range("A" & lastRow & ":F" & lastRow).Borders.LineStyle = xlContinuous
End Sub


Private Sub FormatBaris(ws As Worksheet, rowN As Long)
    ws.Range("A" & rowN & ":F" & rowN).Borders.LineStyle = xlContinuous
    ws.Columns("B").NumberFormat = "dd-mm-yyyy"
    ws.Columns("E:F").NumberFormat = "#,##0"
End Sub

Private Sub UserForm_Initialize()
    CmbAkun.Clear
    With CmbAkun
        .AddItem "Kas"
        .AddItem "Piutang Usaha"
        .AddItem "Perlengkapan"
        .AddItem "Pendapatan"
        .AddItem "Beban Gaji"
        .AddItem "Beban Perlengkapan"
    End With
End Sub


Private Sub btnBersih_Click()
    txtTanggal.Value = ""
    CmbAkun.Value = ""
    txtDebit.Value = ""
    txtKredit.Value = ""
    txtKet.Value = ""
    txtTanggal.SetFocus
End Sub


Private Sub btnSimpan_Click()
    Dim ws As Worksheet
    Dim lastNum As Variant

    Dim rowN As Long
    Dim noUrut As Long
    Dim Tgl As Date
    Dim akun As String, ket As String
    Dim Debit As Variant, Kredit As Variant

    '--- Validasi awal ---
    If Trim(txtTanggal.Value) = "" Then
        MsgBox "Isi tanggal!", vbExclamation: Exit Sub
    End If

    If Not IsDate(txtTanggal.Value) Then
        MsgBox "Format tanggal salah!", vbExclamation: Exit Sub
    End If

    Set ws = BuatSheetJurnal() 'call function

    akun = Trim(CmbAkun.Value)
    ket = Trim(txtKet.Value)
    Debit = Trim(txtDebit.Value)
    Kredit = Trim(txtKredit.Value)

    '--- Kasus KETERANGAN ---
    If ket <> "" And Debit = "" And Kredit = "" Then
        TambahkanKeterangan ws, ket
        btnBersih_Click
        MsgBox "Keterangan ditambahkan!", vbInformation
        Exit Sub
    End If

    '--- Validasi Debit/Kredit ---
    If akun = "" Then MsgBox "Pilih akun!", vbExclamation: Exit Sub

    If Debit <> "" And Kredit <> "" Then
        MsgBox "Hanya boleh isi Debit atau Kredit, bukan dua-duanya!", vbExclamation
        Exit Sub
    End If

    If Debit = "" And Kredit = "" Then
        MsgBox "Isi Debit atau Kredit!", vbExclamation
        Exit Sub
    End If

    If Debit <> "" And Not IsNumeric(Debit) Then
        MsgBox "Debit harus angka!", vbExclamation: Exit Sub
    End If

    If Kredit <> "" And Not IsNumeric(Kredit) Then
        MsgBox "Kredit harus angka!", vbExclamation: Exit Sub
    End If

    
    '--- Tentukan baris ---
    rowN = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    '--- Nomor otomatis ---
    noUrut = FungsiNomorUrutTerakhir(ws)


    '--- Isi data utama ---
    ws.Cells(rowN, 1).Value = noUrut
    ws.Cells(rowN, 2).Value = CDate(txtTanggal.Value)

    If Debit <> "" Then
        ws.Cells(rowN, 3).Value = akun   ' akun di C
        ws.Cells(rowN, 5).Value = CDbl(Debit) ' nominal di E
    Else
        ws.Cells(rowN, 4).Value = akun   ' akun di D
        ws.Cells(rowN, 6).Value = CDbl(Kredit) ' nominal di F
    End If

    FormatBaris ws, rowN

    MsgBox "Data berhasil disimpan!", vbInformation
    btnBersih_Click
End Sub