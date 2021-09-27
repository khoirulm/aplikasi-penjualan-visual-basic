Imports System.Data.Odbc
Public Class FormTransRetur
    Dim TglMySql As String
    Sub KondisiAwal()
        LBLNamaPlg.Text = ""
        LBLAlamat.Text = ""
        LBLKodeAdmin.Text = ""
        LBLKodePlg.Text = ""
        'LBLTelepon.Text = ""
        'LBLTanggal.Text = Today
        'LBLAdmin.Text = FormMenuUtama.STLabel4.Text
        'LBLKembali.Text = ""
        TextBox2.Text = ""
        LBLNamaBarang.Text = ""
        LBLHargaBarang.Text = ""
        TextBox3.Text = ""
        TextBox3.Enabled = False
        'LBLItem.Text = ""
        'Call MunculKodePelanggan()
        'Call NomorOtomatis()
        'Call BuatKolom()
        'Label14.Text = "0"
        TextBox1.Text = ""
        'ComboBox1.Text = ""


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        LBLJam.Text = TimeOfDay
        LBLTgl.Text = Today
    End Sub

    Sub NomorOtomatis()
        Call Koneksi()
        Cmd = New OdbcCommand("Select * from tbl_retur where noretur in (select max(noretur) from tbl_retur)", Conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Not Rd.HasRows Then
            UrutanKode = "R" + Format(Now, "yyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 9) + 1
            UrutanKode = "R" + Format(Now, "yyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        LBLNoRetur.Text = UrutanKode
    End Sub


    Private Sub FormTransRetur_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call NomorOtomatis()
        Call KondisiAwal()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call MunculkanData()

        End If
    End Sub
    Sub MunculkanData()
        Call Koneksi()
        Da = New OdbcDataAdapter("Select kodebarang,namabarang,hargajual,jumlahjual,subtotal  From tbl_detailjual where NoJual='" & TextBox1.Text & "'", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "tbl_detailjual")
        DataGridView2.DataSource = Ds.Tables("tbl_detailjual")
        DataGridView2.ReadOnly = True

        Call Koneksi()
        Cmd = New OdbcCommand("Select * From tbl_jual where noJual = '" & TextBox1.Text & "'", Conn)
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Not Rd.HasRows Then
            MsgBox("No Jual Tidak Ada")
        Else
            LBLKodePlg.Text = Rd.Item("KodePelanggan")
            LBLKodeAdmin.Text = Rd.Item("KodeAdmin")
           
        End If

        Call Koneksi()
        Cmd = New OdbcCommand("Select * From tbL_pelanggan where KodePelanggan = '" & LBLKodePlg.Text & "'", Conn)
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Not Rd.HasRows Then
            MsgBox("Kode Pelanggan Tidak Ada")
        Else
            LBLNamaPlg.Text = Rd.Item("NamaPelanggan")
            LBLAlamat.Text = Rd.Item("AlamatPelanggan")

        End If

    End Sub


End Class