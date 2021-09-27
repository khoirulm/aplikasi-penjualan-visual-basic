Public Class FormLapJual

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        AxCrystalReport1.SelectionFormula = "totext({tbl_Jual.TglJual})='" & DateTimePicker1.Value & "'"
        AxCrystalReport1.ReportFileName = "LaporanHarian.rpt"
        AxCrystalReport1.WindowState = Crystal.WindowStateConstants.crptMaximized
        AxCrystalReport1.RetrieveDataFiles()
        AxCrystalReport1.Action = 1
    End Sub

    Private Sub FormLapJual_Load(sender As Object, e As EventArgs) Handles MyBase.Load
      

        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("01")
        ComboBox1.Items.Add("02")
        ComboBox1.Items.Add("03")
        ComboBox1.Items.Add("04")
        ComboBox1.Items.Add("05")
        ComboBox1.Items.Add("06")
        ComboBox1.Items.Add("07")
        ComboBox1.Items.Add("08")
        ComboBox1.Items.Add("09")
        ComboBox1.Items.Add("11")
        ComboBox1.Items.Add("12")




        ComboBox2.Items.Clear()
        ComboBox2.Text = Date.Now.Year
        For i As Integer = 0 To 5
            ComboBox2.Items.Add(Date.Now.Year - i)
        Next
        'Label7.Text = "2021, 09, 14"
        'Label8.Text = "2021, 09, 16"

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim dateBulan As Integer
        Dim dateTahun As Integer
      
        dateBulan = ComboBox1.Text
        dateTahun = ComboBox2.Text

        'Label7.Text = dateBulan
        'Label8.Text = dateTahun

      

        If ComboBox1.Text = "" Or ComboBox2.Text = "" Then
            MsgBox("silahkan isi bulan dan tahunnya terlebih dahulu!")
        Else


            AxCrystalReport1.SelectionFormula = "Month({tbl_jual.tgljual})=" & dateBulan & "and year({tbl_jual.tgljual})=" & dateTahun & ""
            'AxCrystalReport1.SelectionFormula = "month(tgljual)='" & dateBulan & " and year(tgljual)='" & dateTahun
            AxCrystalReport1.ReportFileName = "laporanbulanan.rpt"
            AxCrystalReport1.WindowState = Crystal.WindowStateConstants.crptMaximized
            AxCrystalReport1.RetrieveDataFiles()
            AxCrystalReport1.Action = 1
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim tglAwal As String
        Dim tglAkhir As String

        tglAwal = Format(DateTimePicker2.Value, "yyyy, MM, dd")
        tglAkhir = Format(DateTimePicker3.Value, "yyyy, MM, dd")


        AxCrystalReport1.SelectionFormula = "{tbl_Jual.TglJual}in date (" & tglAwal & ") to date (" & tglAkhir & ")"
        AxCrystalReport1.ReportFileName = "LaporanMingguan.rpt"
        AxCrystalReport1.WindowState = Crystal.WindowStateConstants.crptMaximized
        AxCrystalReport1.RetrieveDataFiles()
        AxCrystalReport1.Action = 1
    End Sub


End Class