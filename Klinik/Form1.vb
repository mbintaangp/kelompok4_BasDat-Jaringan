Imports System.Data.SqlClient
Public Class Form1
    Private connectionString As String = ("Data Source=LAPTOP-13U1C95E; " &
                "user id = userol; password=123456; Integrated Security=True; " &
                "database=Klinik")
    Private connection As SqlConnection
    Private command As SqlCommand
    Private adapter As SqlDataAdapter
    Private dataTable As DataTable

    Private Sub OpenConnection()
        connection = New SqlConnection(connectionString)
        connection.Open()
    End Sub

    Private Sub CloseConnection()
        connection.Close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OpenConnection()

        LoadDataPasien()
        LoadDataPemeriksaan()
        LoadDataResep()

        CloseConnection()
    End Sub

    ' Load data pasien dari database ke DataGridView
    Private Sub LoadDataPasien()
        dataTable = New DataTable()
        adapter = New SqlDataAdapter("SELECT * FROM Pasien", connection)
        adapter.Fill(dataTable)

        DataGridViewPasien.DataSource = dataTable
    End Sub

    ' Load data pemeriksaan kesehatan dari database ke DataGridView
    Private Sub LoadDataPemeriksaan()
        dataTable = New DataTable()
        adapter = New SqlDataAdapter("SELECT * FROM Pemeriksaan", connection)
        adapter.Fill(dataTable)

        DataGridViewPemeriksaan.DataSource = dataTable
    End Sub

    ' Load data resep dari database ke DataGridView
    Private Sub LoadDataResep()
        dataTable = New DataTable()
        adapter = New SqlDataAdapter("SELECT * FROM Resep", connection)
        adapter.Fill(dataTable)

        DataGridViewResep.DataSource = dataTable
    End Sub

    ' Tambah data pasien ke database
    Private Sub TambahDataPasien(nama As String, umur As Integer, alamat As String)
        command = New SqlCommand("INSERT INTO Pasien (Nama, Umur, Alamat) VALUES (@Nama, @Umur, @Alamat)", connection)
        command.Parameters.AddWithValue("@Nama", nama)
        command.Parameters.AddWithValue("@Umur", umur)
        command.Parameters.AddWithValue("@Alamat", alamat)

        command.ExecuteNonQuery()
    End Sub

    ' Edit data pasien di database
    Private Sub EditDataPasien(id As Integer, nama As String, umur As Integer, alamat As String)
        command = New SqlCommand("UPDATE Pasien SET Nama = @Nama, Umur = @Umur, Alamat = @Alamat WHERE ID = @ID", connection)
        command.Parameters.AddWithValue("@Nama", nama)
        command.Parameters.AddWithValue("@Umur", umur)
        command.Parameters.AddWithValue("@Alamat", alamat)
        command.Parameters.AddWithValue("@ID", id)

        command.ExecuteNonQuery()
    End Sub

    ' Hapus data pasien dari database
    Private Sub HapusDataPasien(id As Integer)
        command = New SqlCommand("DELETE FROM Pasien WHERE ID = @ID", connection)
        command.Parameters.AddWithValue("@ID", id)

        command.ExecuteNonQuery()
    End Sub

    ' Tambah transaksi pemeriksaan ke database
    Private Sub TambahTransaksiPemeriksaan(idPasien As Integer, tanggal As DateTime, keterangan As String)
        command = New SqlCommand("INSERT INTO Pemeriksaan (IDPasien, Tanggal, Keterangan) VALUES (@IDPasien, @Tanggal, @Keterangan)", connection)
        command.Parameters.AddWithValue("@IDPasien", idPasien)
        command.Parameters.AddWithValue("@Tanggal", tanggal)
        command.Parameters.AddWithValue("@Keterangan", keterangan)

        command.ExecuteNonQuery()
    End Sub

    ' Edit transaksi pemeriksaan di database
    Private Sub EditTransaksiPemeriksaan(idPemeriksaan As Integer, idPasien As Integer, tanggal As DateTime, keterangan As String)
        command = New SqlCommand("UPDATE Pemeriksaan SET IDPasien = @IDPasien, Tanggal = @Tanggal, Keterangan = @Keterangan WHERE IDPemeriksaan = @IDPemeriksaan", connection)
        command.Parameters.AddWithValue("@IDPasien", idPasien)
        command.Parameters.AddWithValue("@Tanggal", tanggal)
        command.Parameters.AddWithValue("@Keterangan", keterangan)
        command.Parameters.AddWithValue("@IDPemeriksaan", idPemeriksaan)

        command.ExecuteNonQuery()
    End Sub

    ' Hapus transaksi pemeriksaan dari database
    Private Sub HapusTransaksiPemeriksaan(idPemeriksaan As Integer)
        command = New SqlCommand("DELETE FROM Pemeriksaan WHERE IDPemeriksaan = @IDPemeriksaan", connection)
        command.Parameters.AddWithValue("@IDPemeriksaan", idPemeriksaan)

        command.ExecuteNonQuery()
    End Sub

    ' Tambah transaksi resep ke database
    Private Sub TambahTransaksiResep(idPasien As Integer, tanggal As DateTime, obat As String, dosis As String)
        command = New SqlCommand("INSERT INTO Resep (IDPasien, Tanggal, Obat, Dosis) VALUES (@IDPasien, @Tanggal, @Obat, @Dosis)", connection)
        command.Parameters.AddWithValue("@IDPasien", idPasien)
        command.Parameters.AddWithValue("@Tanggal", tanggal)
        command.Parameters.AddWithValue("@Obat", obat)
        command.Parameters.AddWithValue("@Dosis", dosis)

        command.ExecuteNonQuery()
    End Sub

    ' Edit transaksi resep di database
    Private Sub EditTransaksiResep(idResep As Integer, idPasien As Integer, tanggal As DateTime, obat As String, dosis As String)
        command = New SqlCommand("UPDATE Resep SET IDPasien = @IDPasien, Tanggal = @Tanggal, Obat = @Obat, Dosis = @Dosis WHERE IDResep = @IDResep", connection)
        command.Parameters.AddWithValue("@IDPasien", idPasien)
        command.Parameters.AddWithValue("@Tanggal", tanggal)
        command.Parameters.AddWithValue("@Obat", obat)
        command.Parameters.AddWithValue("@Dosis", dosis)
        command.Parameters.AddWithValue("@IDResep", idResep)

        command.ExecuteNonQuery()
    End Sub

    ' Hapus transaksi resep dari database
    Private Sub HapusTransaksiResep(idResep As Integer)
        command = New SqlCommand("DELETE FROM Resep WHERE IDResep = @IDResep", connection)
        command.Parameters.AddWithValue("@IDResep", idResep)

        command.ExecuteNonQuery()
    End Sub

    'Report Pemeriksaan
    Private Sub GenerateLaporanBulananPemeriksaan(bulan As Integer, tahun As Integer)
        OpenConnection()

        ' Menggunakan SQL query untuk mengambil data pemeriksaan resep berdasarkan bulan dan tahun
        Dim query As String = "SELECT * FROM Pemeriksaan WHERE MONTH(Tanggal) = @Bulan AND YEAR(Tanggal) = @Tahun"
        command = New SqlCommand(query, connection)
        command.Parameters.AddWithValue("@Bulan", bulan)
        command.Parameters.AddWithValue("@Tahun", tahun)

        adapter = New SqlDataAdapter(command)
        dataTable = New DataTable()
        adapter.Fill(dataTable)

        ' Menampilkan hasil laporan di DataGridView
        DataGridViewLaporanPemeriksaan.DataSource = dataTable

        ' Menyesuaikan tampilan kolom di DataGridView
        DataGridViewLaporanPemeriksaan.AutoResizeColumns()

        ' Menutup koneksi setelah selesai
        CloseConnection()


    End Sub

    'Report Resep
    Private Sub GenerateLaporanBulanan(bulan As Integer, tahun As Integer)
        OpenConnection()

        ' Menggunakan SQL query untuk mengambil data pemeriksaan resep berdasarkan bulan dan tahun
        Dim query As String = "SELECT * FROM Resep WHERE MONTH(Tanggal) = @Bulan AND YEAR(Tanggal) = @Tahun"
        command = New SqlCommand(query, connection)
        command.Parameters.AddWithValue("@Bulan", bulan)
        command.Parameters.AddWithValue("@Tahun", tahun)

        adapter = New SqlDataAdapter(command)
        dataTable = New DataTable()
        adapter.Fill(dataTable)

        ' Menampilkan hasil laporan di DataGridView
        DataGridViewLaporan.DataSource = dataTable

        ' Menyesuaikan tampilan kolom di DataGridView
        DataGridViewLaporan.AutoResizeColumns()

        ' Menutup koneksi setelah selesai
        CloseConnection()


    End Sub





    ' Tombol Tambah Pasien
    Private Sub btnTambahPasien_Click(sender As Object, e As EventArgs) Handles btnTambahPasien.Click
        OpenConnection()

        Dim nama As String = txtNamaPasien.Text
        Dim umur As Integer = Convert.ToInt32(txtUmurPasien.Text)
        Dim alamat As String = txtAlamatPasien.Text

        TambahDataPasien(nama, umur, alamat)
        LoadDataPasien()

        CloseConnection()
    End Sub

    ' Tombol Edit Pasien
    Private Sub btnEditPasien_Click(sender As Object, e As EventArgs) Handles btnEditPasien.Click
        OpenConnection()

        Dim id As Integer = Convert.ToInt32(txtIDPasien.Text)
        Dim nama As String = txtNamaPasien.Text
        Dim umur As Integer = Convert.ToInt32(txtUmurPasien.Text)
        Dim alamat As String = txtAlamatPasien.Text

        EditDataPasien(id, nama, umur, alamat)
        LoadDataPasien()

        CloseConnection()
    End Sub

    ' Tombol Hapus Pasien
    Private Sub btnHapusPasien_Click(sender As Object, e As EventArgs) Handles btnHapusPasien.Click
        OpenConnection()

        Dim id As Integer = Convert.ToInt32(txtIDPasien.Text)

        HapusDataPasien(id)
        LoadDataPasien()

        CloseConnection()
    End Sub

    ' Tombol Tambah Transaksi Pemeriksaan
    Private Sub btnTambahPemeriksaan_Click(sender As Object, e As EventArgs) Handles btnTambahPemeriksaan.Click
        OpenConnection()

        Dim idPasien As Integer = Convert.ToInt32(txtIDPemeriksaan.Text)
        Dim tanggal As DateTime = DateTimePickerPemeriksaan.Value
        Dim keterangan As String = txtKeteranganPemeriksaan.Text

        TambahTransaksiPemeriksaan(idPasien, tanggal, keterangan)
        LoadDataPemeriksaan()

        CloseConnection()
    End Sub

    ' Tombol Edit Transaksi Pemeriksaan
    Private Sub btnEditPemeriksaan_Click(sender As Object, e As EventArgs) Handles btnEditPemeriksaan.Click
        OpenConnection()

        Dim idPemeriksaan As Integer = Convert.ToInt32(txtIDPemeriksaan.Text)
        Dim idPasien As Integer = Convert.ToInt32(txtIDPemeriksaan.Text)
        Dim tanggal As DateTime = DateTimePickerPemeriksaan.Value
        Dim keterangan As String = txtKeteranganPemeriksaan.Text

        EditTransaksiPemeriksaan(idPemeriksaan, idPasien, tanggal, keterangan)
        LoadDataPemeriksaan()

        CloseConnection()
    End Sub

    ' Tombol Hapus Transaksi Pemeriksaan
    Private Sub btnHapusPemeriksaan_Click(sender As Object, e As EventArgs) Handles btnHapusPemeriksaan.Click
        OpenConnection()

        Dim idPemeriksaan As Integer = Convert.ToInt32(txtIDPemeriksaan.Text)

        HapusTransaksiPemeriksaan(idPemeriksaan)
        LoadDataPemeriksaan()

        CloseConnection()
    End Sub

    ' Tombol Tambah Transaksi Resep
    Private Sub btnTambahResep_Click(sender As Object, e As EventArgs) Handles btnTambahResep.Click
        OpenConnection()

        Dim idPasien As Integer = Convert.ToInt32(txtIDResep.Text)
        Dim tanggal As DateTime = DateTimePickerResep.Value
        Dim obat As String = txtObatResep.Text
        Dim dosis As String = txtDosisResep.Text

        TambahTransaksiResep(idPasien, tanggal, obat, dosis)
        LoadDataResep()

        CloseConnection()
    End Sub

    Private Sub btnEditResep_Click(sender As Object, e As EventArgs) Handles btnEditResep.Click
        OpenConnection()

        Dim idResep As Integer = Convert.ToInt32(txtIDResep.Text)
        Dim idPasien As Integer = Convert.ToInt32(txtIDResep.Text)
        Dim tanggal As DateTime = DateTimePickerResep.Value
        Dim obat As String = txtObatResep.Text
        Dim dosis As String = txtDosisResep.Text

        EditTransaksiResep(idResep, idPasien, tanggal, obat, dosis)
        LoadDataResep()

        CloseConnection()

    End Sub
    ' Tombol Hapus Transaksi Resep
    Private Sub BtnHapusResep_Click(sender As Object, e As EventArgs) Handles BtnHapusResep.Click
        OpenConnection()

        Dim idResep As Integer = Convert.ToInt32(txtIDResep.Text)

        HapusTransaksiResep(idResep)
        LoadDataResep()

        CloseConnection()
    End Sub

    Private Sub btnPemeriksaanReport_Click(sender As Object, e As EventArgs) Handles btnPemeriksaanReport.Click

        Dim tanggal As DateTime = DateTimePickerPemeriksaan.Value
        Dim bulan As Integer = tanggal.Month
        Dim tahun As Integer = tanggal.Year

        GenerateLaporanBulananPemeriksaan(bulan, tahun)

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim tanggal As DateTime = DateTimePickerPemeriksaan.Value
        Dim bulan As Integer = tanggal.Month
        Dim tahun As Integer = tanggal.Year

        GenerateLaporanBulanan(bulan, tahun)
    End Sub
End Class
