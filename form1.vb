Imports System.Data.Odbc

Public Class form1
    Public conn As OdbcConnection
    Public cmd As OdbcCommand
    Public dr As OdbcDataReader

    Sub koneksi()
        conn = New OdbcConnection("Dsn=konek_dblatihan;Database=db_latihan;Uid=root;Pwd=")
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
                MsgBox("koneksi berhasil")
            End If
        Catch ex As Exception
            MsgBox("Koneksi Gagal.." & ex.Message)
        End Try
    End Sub

    Sub IsiComboBox()
        Try
            Dim query As String = "SELECT kode_pelanggan FROM tbl_pelanggan"
            cmd = New OdbcCommand(query, conn)
            dr = cmd.ExecuteReader()
            cbopelanggan.Items.Clear()
            While dr.Read()
                cbopelanggan.Items.Add(dr("kode_pelanggan").ToString())
            End While
            dr.Close()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        Try
            Dim query As String = "SELECT Nama_Barang FROM tblBarang"
            cmd = New OdbcCommand(query, conn)
            dr = cmd.ExecuteReader()
            cbobarang.Items.Clear()
            While dr.Read()
                cbobarang.Items.Add(dr("Nama_Barang").ToString())
            End While
            dr.Close()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub TampilkanSemuaDataPenjualan()
        Dim dt As New DataTable()
        Dim query As String = "SELECT * FROM tbl_penjualan_rinci"

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            cmd = New OdbcCommand(query, conn)

            ' Menggunakan OdbcDataAdapter untuk mengisi DataTable
            Using adapter As New OdbcDataAdapter(cmd)
                adapter.Fill(dt)
            End Using

            ' Mengisi DataGridView dengan DataTable
            dataGridView1.DataSource = dt

        Catch ex As Exception
            MsgBox("Kesalahan saat menampilkan data: " & ex.Message)
        End Try
    End Sub

    Sub hitungTotal()
        ' Pastikan koneksi terbuka
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        ' Menghitung total kolom subtotal berdasarkan faktur
        Dim totalQuery As String = "SELECT SUM(sub_total) AS TotalFaktur FROM tbl_penjualan_rinci WHERE faktur_penjualan = ?"
        cmd = New OdbcCommand(totalQuery, conn)
        cmd.Parameters.AddWithValue("?", txtfaktur.Text)

        Try
            ' Menggunakan ExecuteScalar untuk mengambil total secara langsung
            Dim total As Object = cmd.ExecuteScalar()

            ' Memeriksa apakah total bukan null, lalu menampilkan di txttotal
            If total IsNot DBNull.Value Then
                txttotal.Text = total
            Else
                txttotal.Text = ""
            End If

        Catch ex As Exception
            MsgBox("Kesalahan saat menghitung total: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Sub AutoIncrementFaktur()
        ' Mengambil nomor faktur terakhir dari database
        Dim query As String = "SELECT faktur_penjualan FROM tbl_penjualan ORDER BY faktur_penjualan DESC LIMIT 1"
        Dim lastFaktur As String = ""

        Try
            If conn.State = ConnectionState.Closed Then conn.Open()
            cmd = New OdbcCommand(query, conn)
            dr = cmd.ExecuteReader()

            If dr.Read() Then
                lastFaktur = dr("faktur_penjualan").ToString()
            End If
            dr.Close()

            ' Mengatur nomor faktur baru berdasarkan nomor terakhir
            If lastFaktur <> "" Then
                Dim num As Integer = Integer.Parse(lastFaktur.Substring(lastFaktur.Length - 4)) + 1
                txtfaktur.Text = "F" & DateTime.Now.ToString("yyMMdd") & num.ToString("D4")
            Else
                ' Jika belum ada faktur, mulai dari awal
                txtfaktur.Text = "F" & DateTime.Now.ToString("yyMMdd") & "0001"
            End If

        Catch ex As Exception
            MsgBox("Kesalahan saat mengisi nomor faktur: " & ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        TampilkanSemuaDataPenjualan()
        IsiComboBox()
        AutoIncrementFaktur()
    End Sub

    Private Sub cbopelanggan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbopelanggan.SelectedIndexChanged
        Try
            Dim selectedKode As String = cbopelanggan.SelectedItem.ToString()
            Dim query As String = "SELECT nama_pelanggan FROM tbl_pelanggan WHERE kode_pelanggan = ?"
            cmd = New OdbcCommand(query, conn)
            cmd.Parameters.Add(New OdbcParameter("namaPelanggan", selectedKode))
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                txtpelanggan.Text = dr("nama_pelanggan").ToString()
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox("Error:" & ex.Message)
        End Try
    End Sub

    Private Sub cbobarang_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbobarang.SelectedIndexChanged
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Dim selectedBarang As String = cbobarang.SelectedItem.ToString()
            Dim query As String = "SELECT harga_beli, kode_barang, jenis, harga_jual FROM tblbarang WHERE Nama_Barang = ?"
            cmd = New OdbcCommand(query, conn)
            cmd.Parameters.Add(New OdbcParameter("namaBarang", selectedBarang))
            dr = cmd.ExecuteReader()

            If dr.Read() Then
                txtkodebarang.Text = dr("kode_barang").ToString()
                txtjenis.Text = dr("jenis").ToString()
                txthargakotor.Text = dr("harga_beli").ToString()
                txthargabersih.Text = dr("harga_jual").ToString()
            End If

            dr.Close()
            If txtjumlahbeli.Text IsNot "" Then
                Dim jumlahbeli As Integer = Integer.Parse(txtjumlahbeli.Text)
                Dim hargabeli As Double = Double.Parse(txthargabersih.Text)
                Dim subtotal As Double = jumlahbeli * hargabeli
                txtsubtotal.Text = subtotal
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub txtjumlahbeli_TextChanged(sender As Object, e As EventArgs) Handles txtjumlahbeli.TextChanged
        If txtjumlahbeli.Text IsNot "" Then
            Try
                Dim jumlahbeli As Integer = Integer.Parse(txtjumlahbeli.Text)
                Dim hargabeli As Double = Double.Parse(txthargabersih.Text)
                Dim subtotal As Double = jumlahbeli * hargabeli
                txtsubtotal.Text = subtotal
            Catch ex As Exception
                MsgBox("Error: Masukkan Angka!")
            End Try
        Else
            txtsubtotal.Text = ""
        End If
    End Sub

    Sub buatpenjualan()
        Dim nofaktur = txtfaktur.Text
        Dim tgl = txttanggal.Value
        Dim kodepelanggan = cbopelanggan.Text
        Dim totalpembelian = txttotal.Text
        Dim query As String = "INSERT INTO tbl_penjualan (faktur_penjualan, tgl_penjualan, kode_pelanggan, total) VALUES (?, ?, ?, ?)"
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmd = New OdbcCommand(query, conn)
            cmd.Parameters.AddWithValue("?", nofaktur)
            cmd.Parameters.AddWithValue("?", tgl)
            cmd.Parameters.AddWithValue("?", kodepelanggan)
            cmd.Parameters.AddWithValue("?", totalpembelian)
            cmd.ExecuteNonQuery()
            MsgBox("Data penjualan berhasil disimpan ke dalam tabel.")
        Catch ex As Exception
            MsgBox("Kesalahan saat menyimpan data: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub
    Sub simpandetail()
        Dim nofaktur = txtfaktur.Text
        Dim kodebarang = txtkodebarang.Text
        Dim hargajual = txthargabersih.Text
        Dim jumlahbeli = txtjumlahbeli.Text
        Dim subtotal = txtsubtotal.Text
        Dim query As String = "INSERT INTO tbl_penjualan_rinci (faktur_penjualan, kode_barang, harga_jual, jumlah, sub_total) VALUES (?, ?, ?, ?, ?)"

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmd = New OdbcCommand(query, conn)
            cmd.Parameters.AddWithValue("?", nofaktur)
            cmd.Parameters.AddWithValue("?", kodebarang)
            cmd.Parameters.AddWithValue("?", hargajual)
            cmd.Parameters.AddWithValue("?", jumlahbeli)
            cmd.Parameters.AddWithValue("?", subtotal)
            cmd.ExecuteNonQuery()
            MsgBox("Data detail penjualan berhasil disimpan ke dalam tabel.")
        Catch ex As Exception
            MsgBox("Kesalahan saat menyimpan data detail penjualan: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Sub simpanpenjualan()
        Dim nofaktur = txtfaktur.Text
        Dim totalpembelian = txttotal.Text

        Dim query As String = "UPDATE tbl_penjualan SET total = ? WHERE faktur_penjualan = ?"
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            cmd = New OdbcCommand(query, conn)
            cmd.Parameters.AddWithValue("?", totalpembelian)
            cmd.Parameters.AddWithValue("?", nofaktur)
            cmd.ExecuteNonQuery()

            MsgBox("Total penjualan berhasil diperbarui di tabel.")
        Catch ex As Exception
            MsgBox("Kesalahan saat memperbarui data: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub btnsimpan_Click(sender As Object, e As EventArgs) Handles btnsimpan.Click
        Call simpanpenjualan()
        TampilkanSemuaDataPenjualan()
        AutoIncrementFaktur()
        hitungTotal()
        Tambah.Text = "Buat"
        txttanggal.Enabled = True
        cbopelanggan.Enabled = True

        cbopelanggan.Text = ""
        cbobarang.Text = ""
        txtjumlahbeli.Clear()
        txtpelanggan.Clear()
        txtkodebarang.Clear()
        txtjenis.Clear()
        txthargabersih.Clear()
        txthargakotor.Clear()
    End Sub

    Private Sub ProdukToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProdukToolStripMenuItem.Click
        produkform.Show()
        Me.Hide()
    End Sub

    Private Sub Tambah_Click(sender As Object, e As EventArgs) Handles Tambah.Click
        If String.IsNullOrWhiteSpace(txtpelanggan.Text) Then
            MessageBox.Show("Nama pelanggan tidak boleh  kosong.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        If Tambah.Text Is "Buat" Then
            Call buatpenjualan()
            Call simpandetail()
            TampilkanSemuaDataPenjualan()
            hitungTotal()
            Tambah.Text = "Tambah"
            txttanggal.Enabled = False
            cbopelanggan.Enabled = False
        Else
            Call simpandetail()
            TampilkanSemuaDataPenjualan()
            hitungTotal()
        End If
    End Sub

End Class
