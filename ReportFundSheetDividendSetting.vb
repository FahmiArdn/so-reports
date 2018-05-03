Imports simpi.GlobalUtilities
Imports simpi.GlobalConnection
Public Class ReportFundSheetDividendSetting
    Public frm As ReportFundSheetDividend
    Dim reportSection As String = "REPORT FUND SHEET DIVIDEND"

    Public Sub FormLoad()
        If frm.pdfLayout.layoutType = "DEFAULT" Then
            rbDefault.Checked = True
        ElseIf frm.pdfLayout.layoutType = "OPTION1" Then
            rbOption1.Checked = True
        Else
            rbOption2.Checked = True
        End If
    End Sub

#Region "setting"
    Private Sub colorSet()
        If cd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            If rbReportTitle.Checked Then
                txtColorReportTitle.BackColor = cd.Color
                ReportTitle_R.Text = RGBWrite("R", cd.Color.R)
                ReportTitle_G.Text = RGBWrite("G", cd.Color.G)
                ReportTitle_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbTujuanInvestasi.Checked Then
                txtColorTujuanInvestasi.BackColor = cd.Color
                TujuanInvestasi_R.Text = RGBWrite("R", cd.Color.R)
                TujuanInvestasi_G.Text = RGBWrite("G", cd.Color.G)
                TujuanInvestasi_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbInformasiReksaDana.Checked Then
                txtColorInformasiReksaDana.BackColor = cd.Color
                InformasiReksaDana_R.Text = RGBWrite("R", cd.Color.R)
                InformasiReksaDana_G.Text = RGBWrite("G", cd.Color.G)
                InformasiReksaDana_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbInvestasiDanBiaya.Checked Then
                txtColorInvestasiDanBiaya.BackColor = cd.Color
                InvestasiDanBiaya_R.Text = RGBWrite("R", cd.Color.R)
                InvestasiDanBiaya_G.Text = RGBWrite("G", cd.Color.G)
                InvestasiDanBiaya_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbStatistikReksadana.Checked Then
                txtColorStatistikReksadana.BackColor = cd.Color
                StatistikReksadana_R.Text = RGBWrite("R", cd.Color.R)
                StatistikReksadana_G.Text = RGBWrite("G", cd.Color.G)
                StatistikReksadana_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbRisikoInvestasi.Checked Then
                txtColorRisikoInvestasi.BackColor = cd.Color
                RisikoInvestasi_R.Text = RGBWrite("R", cd.Color.R)
                RisikoInvestasi_G.Text = RGBWrite("G", cd.Color.G)
                RisikoInvestasi_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbKlasifikasiRisiko.Checked Then
                txtColorKlasifikasiRisiko.BackColor = cd.Color
                KlasifikasiRisiko_R.Text = RGBWrite("R", cd.Color.R)
                KlasifikasiRisiko_G.Text = RGBWrite("G", cd.Color.G)
                KlasifikasiRisiko_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbChartTitle.Checked Then
                txtColorChartTitle.BackColor = cd.Color
                ChartTitle_R.Text = RGBWrite("R", cd.Color.R)
                ChartTitle_G.Text = RGBWrite("G", cd.Color.G)
                ChartTitle_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbKinerjaKumulatif.Checked Then
                txtColorKinerjaKumulatif.BackColor = cd.Color
                KinerjaKumulatif_R.Text = RGBWrite("R", cd.Color.R)
                KinerjaKumulatif_G.Text = RGBWrite("G", cd.Color.G)
                KinerjaKumulatif_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbKebijakanInvestasi.Checked Then
                txtColorKebijakanInvestasi.BackColor = cd.Color
                KebijakanInvestasi_R.Text = RGBWrite("R", cd.Color.R)
                KebijakanInvestasi_G.Text = RGBWrite("G", cd.Color.G)
                KebijakanInvestasi_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbEfekPortfolio.Checked Then
                txtColorEfekPortfolio.BackColor = cd.Color
                EfekPortfolio_R.Text = RGBWrite("R", cd.Color.R)
                EfekPortfolio_G.Text = RGBWrite("G", cd.Color.G)
                EfekPortfolio_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbInformasiDividend.Checked Then
                txtColorInformasiDividend.BackColor = cd.Color
                InformasiDividend_R.Text = RGBWrite("R", cd.Color.R)
                InformasiDividend_G.Text = RGBWrite("G", cd.Color.G)
                InformasiDividend_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbReportLine.Checked Then
                txtColorReportLine.BackColor = cd.Color
                ReportLine_R.Text = RGBWrite("R", cd.Color.R)
                ReportLine_G.Text = RGBWrite("G", cd.Color.G)
                ReportLine_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbItemWarna.Checked Then
                txtColorItemWarna.BackColor = cd.Color
                ItemWarna_R.Text = RGBWrite("R", cd.Color.R)
                ItemWarna_G.Text = RGBWrite("G", cd.Color.G)
                ItemWarna_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbChartBorder.Checked Then
                txtColorChartBorder.BackColor = cd.Color
                ChartBorder_R.Text = RGBWrite("R", cd.Color.R)
                ChartBorder_G.Text = RGBWrite("G", cd.Color.G)
                ChartBorder_B.Text = RGBWrite("B", cd.Color.B)
            End If
        End If
    End Sub

#Region "Radio Button"

    Private Sub rbReportTitle_Click(sender As Object, e As EventArgs) Handles rbReportTitle.Click
        colorSet()
    End Sub

    Private Sub rbTujuanInvestasi_CheckedChanged(sender As Object, e As EventArgs) Handles rbTujuanInvestasi.Click
        colorSet()
    End Sub

    Private Sub rbInformasiReksaDana_CheckedChanged(sender As Object, e As EventArgs) Handles rbInformasiReksaDana.Click
        colorSet()
    End Sub

    Private Sub rbInvestasiDanBiaya_CheckedChanged(sender As Object, e As EventArgs) Handles rbInvestasiDanBiaya.Click
        colorSet()
    End Sub

    Private Sub rbStatistikReksadana_CheckedChanged(sender As Object, e As EventArgs) Handles rbStatistikReksadana.Click
        colorSet()
    End Sub

    Private Sub rbRisikoInvestasi_CheckedChanged(sender As Object, e As EventArgs) Handles rbRisikoInvestasi.Click
        colorSet()
    End Sub

    Private Sub rbKlasifikasiRisiko_CheckedChanged(sender As Object, e As EventArgs) Handles rbKlasifikasiRisiko.Click
        colorSet()
    End Sub

    Private Sub rbChartTitle_CheckedChanged(sender As Object, e As EventArgs) Handles rbChartTitle.Click
        colorSet()
    End Sub

    Private Sub rbKinerjaKumulatif_CheckedChanged(sender As Object, e As EventArgs) Handles rbKinerjaKumulatif.Click
        colorSet()
    End Sub

    Private Sub rbKebijakanInvestasi_CheckedChanged(sender As Object, e As EventArgs) Handles rbKebijakanInvestasi.Click
        colorSet()
    End Sub

    Private Sub rbEfekPortfolio_CheckedChanged(sender As Object, e As EventArgs) Handles rbEfekPortfolio.Click
        colorSet()
    End Sub

    Private Sub rbInformasiDividend_CheckedChanged(sender As Object, e As EventArgs) Handles rbInformasiDividend.Click
        colorSet()
    End Sub

    Private Sub rbReportLine_CheckedChanged(sender As Object, e As EventArgs) Handles rbReportLine.Click
        colorSet()
    End Sub

    Private Sub rbItemWarna_CheckedChanged(sender As Object, e As EventArgs) Handles rbItemWarna.Click
        colorSet()
    End Sub

    Private Sub rbChartBorder_CheckedChanged(sender As Object, e As EventArgs) Handles rbChartBorder.Click
        colorSet()
    End Sub
#End Region


#End Region

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Close()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If rbDefault.Checked Then
            iniSave("DEFAULT")
        ElseIf rbOption1.Checked Then
            iniSave("OPTION1")
        ElseIf rbOption2.Checked Then
            iniSave("OPTION2")
        End If
    End Sub

    Private Sub rbDefault_CheckedChanged(sender As Object, e As EventArgs) Handles rbDefault.CheckedChanged
        iniCheck()
    End Sub

    Private Sub rbOption1_CheckedChanged(sender As Object, e As EventArgs) Handles rbOption1.CheckedChanged
        iniCheck()
    End Sub

    Private Sub rbOption2_CheckedChanged(sender As Object, e As EventArgs) Handles rbOption2.CheckedChanged
        iniCheck()
    End Sub

    Private Sub iniCheck()
        If rbDefault.Checked Then
            iniLoad("DEFAULT")
        ElseIf rbOption1.Checked Then
            iniLoad("OPTION1")
        ElseIf rbOption2.Checked Then
            iniLoad("OPTION2")
        End If
    End Sub

    Private Sub iniLoad(ByVal iniType As String)
        Try
            If iniType.Trim = "DEFAULT" Then
                _default()
            Else
                Dim strFile As String = simpiFile("simpi.ini")
                If GlobalFileWindows.FileExists(strFile) Then
                    Dim r, g, b As Integer
                    Dim file As New GlobalINI(strFile)
                    r = file.GetInteger(reportSection, iniType & " Report Title R", 0)
                    g = file.GetInteger(reportSection, iniType & " Report Title G", 0)
                    b = file.GetInteger(reportSection, iniType & " Report Title B", 0)
                    ReportTitle_R.Text = RGBWrite("R", r)
                    ReportTitle_G.Text = RGBWrite("G", g)
                    ReportTitle_B.Text = RGBWrite("B", b)
                    txtColorReportTitle.BackColor = Color.FromArgb(r, g, b)
                    txtReportTitle.Text = file.GetString(reportSection, iniType & " Report Title", "")

                    r = file.GetInteger(reportSection, iniType & " Tujuan Investasi R", 0)
                    g = file.GetInteger(reportSection, iniType & " Tujuan Investasi G", 0)
                    b = file.GetInteger(reportSection, iniType & " Tujuan Investasi B", 0)
                    TujuanInvestasi_R.Text = RGBWrite("R", r)
                    TujuanInvestasi_G.Text = RGBWrite("G", g)
                    TujuanInvestasi_B.Text = RGBWrite("B", b)
                    txtColorTujuanInvestasi.BackColor = Color.FromArgb(r, g, b)
                    txtTujuanInvestasi.Text = file.GetString(reportSection, iniType & " Tujuan Investasi", "")

                    r = file.GetInteger(reportSection, iniType & " Informasi Reksa Dana R", 0)
                    g = file.GetInteger(reportSection, iniType & " Informasi Reksa Dana G", 0)
                    b = file.GetInteger(reportSection, iniType & " Informasi Reksa Dana B", 0)
                    InformasiReksaDana_R.Text = RGBWrite("R", r)
                    InformasiReksaDana_G.Text = RGBWrite("G", g)
                    InformasiReksaDana_B.Text = RGBWrite("B", b)
                    txtColorInformasiReksaDana.BackColor = Color.FromArgb(r, g, b)
                    txtInformasiReksaDana.Text = file.GetString(reportSection, iniType & " Informasi Reksa Dana", "")
                    txtJenisReksa.Text = file.GetString(reportSection, iniType & " Jenis Reksa", "")
                    txtTanggalPeluncuran.Text = file.GetString(reportSection, iniType & " Tanggal Peluncuran", "")
                    txtDanaKelolaan.Text = file.GetString(reportSection, iniType & " Dana Kelolaan", "")
                    txtMataUang.Text = file.GetString(reportSection, iniType & " Mata Uang", "")
                    txtFrekuensiValuasi.Text = file.GetString(reportSection, iniType & " Frekuensi Valuasi", "")
                    txtBankKustodian.Text = file.GetString(reportSection, iniType & " Bank Kustodian", "")
                    txtTolakUkur.Text = file.GetString(reportSection, iniType & " Tolak Ukur", "")
                    txtNabUnit.Text = file.GetString(reportSection, iniType & " Nab Unit", "")

                    r = file.GetInteger(reportSection, iniType & " Investasi dan Biaya R", 0)
                    g = file.GetInteger(reportSection, iniType & " Investasi dan Biaya G", 0)
                    b = file.GetInteger(reportSection, iniType & " Investasi dan Biaya B", 0)
                    InvestasiDanBiaya_R.Text = RGBWrite("R", r)
                    InvestasiDanBiaya_G.Text = RGBWrite("G", g)
                    InvestasiDanBiaya_B.Text = RGBWrite("B", b)
                    txtColorInvestasiDanBiaya.BackColor = Color.FromArgb(r, g, b)
                    txtInvestasiDanBiaya.Text = file.GetString(reportSection, iniType & " Investasi dan Biaya", "")
                    txtMinInvestasiawal.Text = file.GetString(reportSection, iniType & " Minimal Investasi Awal (Rp)", "")
                    txtMinInvestasiSelanjutnya.Text = file.GetString(reportSection, iniType & " Minimal Investasi Selanjutnya (Rp)", "")
                    txtBiayaPembelian.Text = file.GetString(reportSection, iniType & " Biaya Pembelian (%)", "")
                    txtBiayaPenjualan.Text = file.GetString(reportSection, iniType & " Biaya Penjualan (%)", "")
                    txtBiayaPengalihan.Text = file.GetString(reportSection, iniType & " Biaya Pengalihan (%)", "")
                    txtBiayaJasaPengelola.Text = file.GetString(reportSection, iniType & " Biaya Jasa Pengelolaan MI (%)", "")
                    txtBiayaJasaBank.Text = file.GetString(reportSection, iniType & " Biaya Jasa Bank Kustodian (%)", "")

                    r = file.GetInteger(reportSection, iniType & " Statistik Reksadana R", 0)
                    g = file.GetInteger(reportSection, iniType & " Statistik Reksadana G", 0)
                    b = file.GetInteger(reportSection, iniType & " Statistik Reksadana B", 0)
                    StatistikReksadana_R.Text = RGBWrite("R", r)
                    StatistikReksadana_G.Text = RGBWrite("G", g)
                    StatistikReksadana_B.Text = RGBWrite("B", b)
                    txtColorStatistikReksadana.BackColor = Color.FromArgb(r, g, b)
                    txtStatistikReksadana.Text = file.GetString(reportSection, iniType & " Statistik Reksadana", "")

                    r = file.GetInteger(reportSection, iniType & " Risiko Investasi R", 0)
                    g = file.GetInteger(reportSection, iniType & " Risiko Investasi G", 0)
                    b = file.GetInteger(reportSection, iniType & " Risiko Investasi B", 0)
                    RisikoInvestasi_R.Text = RGBWrite("R", r)
                    RisikoInvestasi_G.Text = RGBWrite("G", g)
                    RisikoInvestasi_B.Text = RGBWrite("B", b)
                    txtColorRisikoInvestasi.BackColor = Color.FromArgb(r, g, b)
                    txtRisikoInvestasi.Text = file.GetString(reportSection, iniType & " Risiko Investasi", "")

                    r = file.GetInteger(reportSection, iniType & " Klasifikasi Risiko R", 0)
                    g = file.GetInteger(reportSection, iniType & " Klasifikasi Risiko G", 0)
                    b = file.GetInteger(reportSection, iniType & " Klasifikasi Risiko B", 0)
                    KlasifikasiRisiko_R.Text = RGBWrite("R", r)
                    KlasifikasiRisiko_G.Text = RGBWrite("G", g)
                    KlasifikasiRisiko_B.Text = RGBWrite("B", b)
                    txtColorKlasifikasiRisiko.BackColor = Color.FromArgb(r, g, b)
                    txtKlasifikasiRisiko.Text = file.GetString(reportSection, iniType & " Klasifikasi Risiko", "")

                    r = file.GetInteger(reportSection, iniType & " Chart Title R", 0)
                    g = file.GetInteger(reportSection, iniType & " Chart Title G", 0)
                    b = file.GetInteger(reportSection, iniType & " Chart Title B", 0)
                    ChartTitle_R.Text = RGBWrite("R", r)
                    ChartTitle_G.Text = RGBWrite("G", g)
                    ChartTitle_B.Text = RGBWrite("B", b)
                    txtColorChartTitle.BackColor = Color.FromArgb(r, g, b)
                    txtChartTitle.Text = file.GetString(reportSection, iniType & " Chart Title", "")

                    r = file.GetInteger(reportSection, iniType & " Kinerja Kumulatif R", 0)
                    g = file.GetInteger(reportSection, iniType & " Kinerja Kumulatif G", 0)
                    b = file.GetInteger(reportSection, iniType & " Kinerja Kumulatif B", 0)
                    KinerjaKumulatif_R.Text = RGBWrite("R", r)
                    KinerjaKumulatif_G.Text = RGBWrite("G", g)
                    KinerjaKumulatif_B.Text = RGBWrite("B", b)
                    txtColorKinerjaKumulatif.BackColor = Color.FromArgb(r, g, b)
                    txtKinerjaKumulatif.Text = file.GetString(reportSection, iniType & " Kinerja Kumulatif", "")

                    r = file.GetInteger(reportSection, iniType & " Efek Portfolio R", 0)
                    g = file.GetInteger(reportSection, iniType & " Efek Portfolio G", 0)
                    b = file.GetInteger(reportSection, iniType & " Efek Portfolio B", 0)
                    EfekPortfolio_R.Text = RGBWrite("R", r)
                    EfekPortfolio_G.Text = RGBWrite("G", g)
                    EfekPortfolio_B.Text = RGBWrite("B", b)
                    txtColorEfekPortfolio.BackColor = Color.FromArgb(r, g, b)
                    txtEfekPortfolio.Text = file.GetString(reportSection, iniType & " Efek Portfolio", "")

                    r = file.GetInteger(reportSection, iniType & " Informasi Dividend R", 0)
                    g = file.GetInteger(reportSection, iniType & " Informasi Dividend G", 0)
                    b = file.GetInteger(reportSection, iniType & " Informasi Dividend B", 0)
                    InformasiDividend_R.Text = RGBWrite("R", r)
                    InformasiDividend_G.Text = RGBWrite("G", g)
                    InformasiDividend_B.Text = RGBWrite("B", b)
                    txtColorInformasiDividend.BackColor = Color.FromArgb(r, g, b)
                    txtInformasiDividend.Text = file.GetString(reportSection, iniType & " Informasi Dividend", "")

                    r = file.GetInteger(reportSection, iniType & " Report Line R", 0)
                    g = file.GetInteger(reportSection, iniType & " Report Line G", 0)
                    b = file.GetInteger(reportSection, iniType & " Report Line B", 0)
                    ReportLine_R.Text = RGBWrite("R", r)
                    ReportLine_G.Text = RGBWrite("G", g)
                    ReportLine_B.Text = RGBWrite("B", b)
                    txtColorReportLine.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " Report Item R", 0)
                    g = file.GetInteger(reportSection, iniType & " Report Item G", 0)
                    b = file.GetInteger(reportSection, iniType & " Report Item B", 0)
                    ItemWarna_R.Text = RGBWrite("R", r)
                    ItemWarna_G.Text = RGBWrite("G", g)
                    ItemWarna_B.Text = RGBWrite("B", b)
                    txtColorReportLine.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " Chart Border R", 0)
                    g = file.GetInteger(reportSection, iniType & " Chart Border G", 0)
                    b = file.GetInteger(reportSection, iniType & " Chart Border B", 0)
                    ChartBorder_R.Text = RGBWrite("R", r)
                    ChartBorder_G.Text = RGBWrite("G", g)
                    ChartBorder_B.Text = RGBWrite("B", b)
                    txtColorReportLine.BackColor = Color.FromArgb(r, g, b)
                    If file.GetBoolean(reportSection, iniType & " Chart Border", False) Then chkChartBorder.Checked = True Else chkChartBorder.Checked = False

                    'r = file.GetInteger(reportSection, iniType & " Table Item R", 0)
                    'g = file.GetInteger(reportSection, iniType & " Table Item G", 0)
                    'b = file.GetInteger(reportSection, iniType & " Table Item B", 0)
                    'TableItem_R.Text = RGBWrite("R", r)
                    'TableItem_G.Text = RGBWrite("G", g)
                    'TableItem_B.Text = RGBWrite("B", b)
                    'txtColorTableItem.BackColor = Color.FromArgb(r, g, b)
                    'txtTableItemReturn.Text = file.GetString(reportSection, iniType & " Return", "")
                    'txtTableItemBenchmark.Text = file.GetString(reportSection, iniType & " Benchmark", "")
                    'txtTableItem1Bulan.Text = file.GetString(reportSection, iniType & " 1 Bulan", "")
                    'txtTableItem3Bulan.Text = file.GetString(reportSection, iniType & " 3 Bulan", "")
                    'txtTableItem6Bulan.Text = file.GetString(reportSection, iniType & " 6 Bulan", "")
                    'txtTableItem1Tahun.Text = file.GetString(reportSection, iniType & " 1 Tahun", "")
                    'txtTableItemDariAwalTahun.Text = file.GetString(reportSection, iniType & " Dari Awal", "")
                    'txtTableItemSejakPembentukan.Text = file.GetString(reportSection, iniType & " Sejak Pembentukan", "")
                End If
            End If
        Catch ex As Exception
            _default()
        End Try
    End Sub

    Private Sub _default()
        With frm.pdfLayout
            frm.pdfColorDefault()
            ReportTitle_R.Text = "R: " & .ReportTitle_R
            ReportTitle_G.Text = "G: " & .ReportTitle_G
            ReportTitle_B.Text = "B: " & .ReportTitle_B
            txtColorReportTitle.BackColor = Color.FromArgb(.ReportTitle_R, .ReportTitle_G, .ReportTitle_B)
            txtReportTitle.Text = .ReportTitle

            TujuanInvestasi_R.Text = "R: " & .TujuanInvestasi_R
            TujuanInvestasi_G.Text = "G: " & .TujuanInvestasi_G
            TujuanInvestasi_B.Text = "B: " & .TujuanInvestasi_B
            txtColorTujuanInvestasi.BackColor = Color.FromArgb(.TujuanInvestasi_R, .TujuanInvestasi_G, .TujuanInvestasi_B)
            txtTujuanInvestasi.Text = .TujuanInvestasi

            InformasiReksaDana_R.Text = "R: " & .InformasiReksaDana_R
            InformasiReksaDana_G.Text = "G: " & .InformasiReksaDana_G
            InformasiReksaDana_B.Text = "B: " & .InformasiReksaDana_B
            txtColorInformasiReksaDana.BackColor = Color.FromArgb(.InformasiReksaDana_R, .InformasiReksaDana_G, .InformasiReksaDana_B)
            txtInformasiReksaDana.Text = .InformasiReksaDana
            txtJenisReksa.Text = .JenisReksa
            txtTanggalPeluncuran.Text = .TanggalPeluncuran
            txtDanaKelolaan.Text = .DanaKelolaan
            txtMataUang.Text = .MataUang
            txtFrekuensiValuasi.Text = .FrekuensiValuasi
            txtBankKustodian.Text = .BankKustodian
            txtTolakUkur.Text = .TolakUkur
            txtNabUnit.Text = .NabUnit

            InvestasiDanBiaya_R.Text = "R: " & .InvestasiDanBiaya_R
            InvestasiDanBiaya_G.Text = "G: " & .InvestasiDanBiaya_G
            InvestasiDanBiaya_B.Text = "B: " & .InvestasiDanBiaya_B
            txtColorInvestasiDanBiaya.BackColor = Color.FromArgb(.InvestasiDanBiaya_R, .InvestasiDanBiaya_G, .InvestasiDanBiaya_B)
            txtInvestasiDanBiaya.Text = .InvestasiDanBiaya
            txtMinInvestasiawal.Text = .MinInvestasiawal
            txtMinInvestasiSelanjutnya.Text = .MinInvestasiSelanjutnya
            txtBiayaPembelian.Text = .BiayaPembelian
            txtBiayaPenjualan.Text = .BiayaPenjualan
            txtBiayaPengalihan.Text = .BiayaPengalihan
            txtBiayaJasaPengelola.Text = .BiayaJasaPengelola
            txtBiayaJasaBank.Text = .BiayaJasaBank

            StatistikReksadana_R.Text = "R: " & .StatistikReksadana_R
            StatistikReksadana_G.Text = "G: " & .StatistikReksadana_G
            StatistikReksadana_B.Text = "B: " & .StatistikReksadana_B
            txtColorStatistikReksadana.BackColor = Color.FromArgb(.StatistikReksadana_R, .StatistikReksadana_G, .StatistikReksadana_B)
            txtStatistikReksadana.Text = .StatistikReksadana

            RisikoInvestasi_R.Text = "R: " & .RisikoInvestasi_R
            RisikoInvestasi_G.Text = "G: " & .RisikoInvestasi_G
            RisikoInvestasi_B.Text = "B: " & .RisikoInvestasi_B
            txtColorRisikoInvestasi.BackColor = Color.FromArgb(.RisikoInvestasi_R, .RisikoInvestasi_G, .RisikoInvestasi_B)
            txtRisikoInvestasi.Text = .RisikoInvestasi

            KlasifikasiRisiko_R.Text = "R: " & .KlasifikasiRisiko_R
            KlasifikasiRisiko_G.Text = "G: " & .KlasifikasiRisiko_G
            KlasifikasiRisiko_B.Text = "B: " & .KlasifikasiRisiko_B
            txtColorKlasifikasiRisiko.BackColor = Color.FromArgb(.KlasifikasiRisiko_R, .KlasifikasiRisiko_G, .KlasifikasiRisiko_B)
            txtKlasifikasiRisiko.Text = .KlasifikasiRisiko

            ChartTitle_R.Text = "R: " & .ChartTitle_R
            ChartTitle_G.Text = "G: " & .ChartTitle_G
            ChartTitle_B.Text = "B: " & .ChartTitle_B
            txtColorChartTitle.BackColor = Color.FromArgb(.ChartTitle_R, .ChartTitle_G, .ChartTitle_B)
            txtChartTitle.Text = .ChartTitle

            KinerjaKumulatif_R.Text = "R: " & .KinerjaKumulatif_R
            KinerjaKumulatif_G.Text = "G: " & .KinerjaKumulatif_G
            KinerjaKumulatif_B.Text = "B: " & .KinerjaKumulatif_B
            txtColorKinerjaKumulatif.BackColor = Color.FromArgb(.KinerjaKumulatif_R, .KinerjaKumulatif_G, .KinerjaKumulatif_B)
            txtKinerjaKumulatif.Text = .KinerjaKumulatif

            KebijakanInvestasi_R.Text = "R: " & .KebijakanInvestasi_R
            KebijakanInvestasi_G.Text = "G: " & .KebijakanInvestasi_G
            KebijakanInvestasi_B.Text = "B: " & .KebijakanInvestasi_B
            txtColorKebijakanInvestasi.BackColor = Color.FromArgb(.KebijakanInvestasi_R, .KebijakanInvestasi_G, .KebijakanInvestasi_B)
            txtKebijakanInvestasi.Text = .KebijakanInvestasi

            EfekPortfolio_R.Text = "R: " & .EfekPortfolio_R
            EfekPortfolio_G.Text = "G: " & .EfekPortfolio_G
            EfekPortfolio_B.Text = "B: " & .EfekPortfolio_B
            txtColorEfekPortfolio.BackColor = Color.FromArgb(.EfekPortfolio_R, .EfekPortfolio_G, .EfekPortfolio_B)
            txtEfekPortfolio.Text = .EfekPortfolio

            InformasiDividend_R.Text = "R: " & .InformasiDividend_R
            InformasiDividend_G.Text = "G: " & .InformasiDividend_G
            InformasiDividend_B.Text = "B: " & .InformasiDividend_B
            txtColorInformasiDividend.BackColor = Color.FromArgb(.InformasiDividend_R, .InformasiDividend_G, .InformasiDividend_B)
            txtInformasiDividend.Text = .InformasiDividend

            ReportLine_R.Text = "R: " & .ReportLine_R
            ReportLine_G.Text = "G: " & .ReportLine_G
            ReportLine_B.Text = "B: " & .ReportLine_B
            txtColorReportLine.BackColor = Color.FromArgb(.ReportLine_R, .ReportLine_G, .ReportLine_B)

            ItemWarna_R.Text = "R: " & .ItemWarna_R
            ItemWarna_G.Text = "G: " & .ItemWarna_G
            ItemWarna_B.Text = "B: " & .ItemWarna_B
            txtColorItemWarna.BackColor = Color.FromArgb(.ItemWarna_R, .ItemWarna_G, .ItemWarna_B)

            ChartBorder_R.Text = "R: " & .ChartBorder_R
            ChartBorder_G.Text = "G: " & .ChartBorder_G
            ChartBorder_B.Text = "B: " & .ChartBorder_B
            txtColorReportLine.BackColor = Color.FromArgb(.ChartBorder_R, .ChartBorder_G, .ChartBorder_B)
        End With
    End Sub

    Private Sub iniSave(ByVal iniType As String)
        Try
            Dim strFile As String = simpiFile("simpi.ini")
            Dim file As New GlobalINI(strFile)
            file.WriteString(reportSection, "LAYOUT", iniType)
            If iniType.Trim <> "DEFAULT" Then
                file.WriteInteger(reportSection, iniType & " Report Title R", RGBRead(ReportTitle_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Report Title G", RGBRead(ReportTitle_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Report Title B", RGBRead(ReportTitle_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Report Title", txtReportTitle.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Tujuan Investasi R", RGBRead(TujuanInvestasi_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tujuan Investasi G", RGBRead(TujuanInvestasi_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tujuan Investasi B", RGBRead(TujuanInvestasi_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Tujuan Investasi", txtTujuanInvestasi.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Informasi Reksa Dana R", RGBRead(InformasiReksaDana_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Informasi Reksa Dana G", RGBRead(InformasiReksaDana_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Informasi Reksa Dana B", RGBRead(InformasiReksaDana_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Informasi Reksa Dana", txtInformasiReksaDana.Text.Trim)
                file.WriteString(reportSection, iniType & " Jenis Reksa", txtJenisReksa.Text.Trim)
                file.WriteString(reportSection, iniType & " Dana Kelolaan", txtDanaKelolaan.Text.Trim)
                file.WriteString(reportSection, iniType & " Mata Uang", txtMataUang.Text.Trim)
                file.WriteString(reportSection, iniType & " Frekuensi Valuasi", txtFrekuensiValuasi.Text.Trim)
                file.WriteString(reportSection, iniType & " Bank Kustodian", txtBankKustodian.Text.Trim)
                file.WriteString(reportSection, iniType & " Tolak Ukur", txtBankKustodian.Text.Trim)
                file.WriteString(reportSection, iniType & " Nab Unit", txtNabUnit.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Investasi dan Biaya R", RGBRead(InvestasiDanBiaya_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Investasi dan Biaya G", RGBRead(InvestasiDanBiaya_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Investasi dan Biaya B", RGBRead(InvestasiDanBiaya_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Investasi dan Biaya", txtInvestasiDanBiaya.Text.Trim)
                file.WriteString(reportSection, iniType & " Minimal Investasi Awal (Rp)", txtMinInvestasiawal.Text.Trim)
                file.WriteString(reportSection, iniType & " Minimal Investasi Selanjutnya (Rp)", txtMinInvestasiSelanjutnya.Text.Trim)
                file.WriteString(reportSection, iniType & " Biaya Pembelian (%)", txtBiayaPembelian.Text.Trim)
                file.WriteString(reportSection, iniType & " Biaya Penjualan (%)", txtBiayaPenjualan.Text.Trim)
                file.WriteString(reportSection, iniType & " Biaya Pengalihan (%)", txtBiayaPengalihan.Text.Trim)
                file.WriteString(reportSection, iniType & " Biaya Jasa Pengelolaan MI (%)", txtBiayaJasaPengelola.Text.Trim)
                file.WriteString(reportSection, iniType & " Biaya Jasa Bank Kustodian (%)", txtBiayaJasaBank.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Statistik Reksadana R", RGBRead(StatistikReksadana_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Statistik Reksadana G", RGBRead(StatistikReksadana_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Statistik Reksadana B", RGBRead(StatistikReksadana_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Statistik Reksadana", txtStatistikReksadana.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Risiko Investasi R", RGBRead(RisikoInvestasi_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Risiko Investasi G", RGBRead(RisikoInvestasi_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Risiko Investasi B", RGBRead(RisikoInvestasi_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Risiko Investasi", txtRisikoInvestasi.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Klasifikasi Risiko R", RGBRead(KlasifikasiRisiko_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Klasifikasi Risiko G", RGBRead(KlasifikasiRisiko_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Klasifikasi Risiko B", RGBRead(KlasifikasiRisiko_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Klasifikasi Risiko", txtKlasifikasiRisiko.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Chart Title R", RGBRead(ChartTitle_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Title G", RGBRead(ChartTitle_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Title B", RGBRead(ChartTitle_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Chart Title", txtChartTitle.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Kinerja Kumulatif R", RGBRead(KinerjaKumulatif_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kinerja Kumulatif G", RGBRead(KinerjaKumulatif_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kinerja Kumulatif B", RGBRead(KinerjaKumulatif_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Kinerja Kumulatif", txtKinerjaKumulatif.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Kebijakan Investasi R", RGBRead(KebijakanInvestasi_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kebijakan Investasi G", RGBRead(KebijakanInvestasi_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kebijakan Investasi B", RGBRead(KebijakanInvestasi_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Kebijakan Investasi", txtKebijakanInvestasi.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Efek Portfolio R", RGBRead(EfekPortfolio_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Efek Portfolio G", RGBRead(EfekPortfolio_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Efek Portfolio B", RGBRead(EfekPortfolio_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Efek Portfolio", txtEfekPortfolio.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Informasi Dividend R", RGBRead(InformasiDividend_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Informasi Dividend G", RGBRead(InformasiDividend_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Informasi Dividend B", RGBRead(InformasiDividend_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Informasi Dividend", txtInformasiDividend.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Report Line R", RGBRead(ReportLine_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Report Line G", RGBRead(ReportLine_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Report Line B", RGBRead(ReportLine_R.Text.Trim))

                file.WriteInteger(reportSection, iniType & " Item Warna R", RGBRead(ItemWarna_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Item Warna G", RGBRead(ItemWarna_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Item Warna B", RGBRead(ItemWarna_B.Text.Trim))

                file.WriteInteger(reportSection, iniType & " Chart Border R", RGBRead(ChartBorder_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Border G", RGBRead(ChartBorder_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Border B", RGBRead(ChartBorder_B.Text.Trim))
            End If
            frm.pdfSetting()
            Close()
        Catch ex As Exception
            ExceptionMessage.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class