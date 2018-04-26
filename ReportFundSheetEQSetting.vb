Imports simpi.GlobalUtilities
Imports simpi.GlobalConnection

Public Class ReportFundSheetEQSetting
    Public frm As ReportFundSheetEQ
    Dim reportSection As String = "REPORT FUND SHEET EQ"

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

        If ColorDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            If rbReportTitle.Checked Then
                txtColorReportTitle.BackColor = ColorDialog1.Color
                ReportTitle_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                ReportTitle_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                ReportTitle_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbNAB.Checked Then
                txtColorNAB.BackColor = ColorDialog1.Color
                NAB_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                NAB_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                NAB_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbBankKustodian.Checked Then
                txtColorBank.BackColor = ColorDialog1.Color
                Bank_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                Bank_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                Bank_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbTanggalPeluncuran.Checked Then
                txtColorTanggal.BackColor = ColorDialog1.Color
                Tanggal_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                Tanggal_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                Tanggal_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbTotal.Checked Then
                txtColorTotal.BackColor = ColorDialog1.Color
                Total_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                Total_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                Total_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbMataUang.Checked Then
                txtColorMatUang.BackColor = ColorDialog1.Color
                MataUang_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                MataUang_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                MataUang_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbImbalManajer.Checked Then
                txtColorImbalManajer.BackColor = ColorDialog1.Color
                ImbalManajer_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                ImbalManajer_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                ImbalManajer_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbImbalBank.Checked Then
                txtColorImbalBank.BackColor = ColorDialog1.Color
                ImbalBank_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                ImbalBank_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                ImbalBank_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbBiayaBeli.Checked Then
                txtColorBiayaBeli.BackColor = ColorDialog1.Color
                BiayaBeli_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                BiayaBeli_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                BiayaBeli_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbBiayaJual.Checked Then
                txtColorBiayaJual.BackColor = ColorDialog1.Color
                BiayaJual_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                BiayaJual_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                BiayaJual_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbInvestAwal.Checked Then
                txtColorMinInvest.BackColor = ColorDialog1.Color
                InvestAwal_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                InvestAwal_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                InvestAwal_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbInvestSelanjutnya.Checked Then
                txtColorMinInvestSelanjutnya.BackColor = ColorDialog1.Color
                InvestSelanjutnya_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                InvestSelanjutnya_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                InvestSelanjutnya_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbKode.Checked Then
                txtColorKode.BackColor = ColorDialog1.Color
                Kode_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                Kode_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                Kode_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbKinerjaSejakDiluncurkan.Checked Then
                txtColorKinerja.BackColor = ColorDialog1.Color
                KinerjaAwal_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                KinerjaAwal_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                KinerjaAwal_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbKinerjaTerbaik.Checked Then
                txtColorKinerjaTerbaik.BackColor = ColorDialog1.Color
                KinerjaTerbaik_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                KinerjaTerbaik_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                KinerjaTerbaik_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbKinerjaTerburuk.Checked Then
                txtColorKinerjaTerburuk.BackColor = ColorDialog1.Color
                KinerjaTerburuk_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                KinerjaTerburuk_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                KinerjaTerburuk_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbFaktorRisikUtama.Checked Then
                txtColorFaktorRisiko.BackColor = ColorDialog1.Color
                FaktorRisiko_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                FaktorRisiko_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                FaktorRisiko_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbPeriodeInvestasi.Checked Then
                txtColorPeriode.BackColor = ColorDialog1.Color
                PeriodeInvestasi_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                PeriodeInvestasi_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                PeriodeInvestasi_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbTingkatRisiko.Checked Then
                txtColorTingkatRisiko.BackColor = ColorDialog1.Color
                TingkatRisiko_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                TIngkatRisiko_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                TingkatRisiko_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbTujuanInvestasi.Checked Then
                txtColorTujuanInvestasi.BackColor = ColorDialog1.Color
                TujuanInvestasi_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                TujuanInvestasi_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                TujuanInvestasi_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbInvestasi.Checked Then
                txtColorInvestasi.BackColor = ColorDialog1.Color
                Investasi_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                Investasi_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                Investasi_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbPortofolio.Checked Then
                txtColorKomposisPortofolio.BackColor = ColorDialog1.Color
                Portofolio_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                Portofolio_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                Portofolio_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbChartTitle.Checked Then
                txtColorChartTitle.BackColor = ColorDialog1.Color
                KinerjaReksaDana_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                KinerjaReksaDana_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                KinerjaReksaDana_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbChartBorder.Checked Then
                txtColorChartBorder.BackColor = ColorDialog1.Color
                ChartBorder_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                ChartBorder_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                ChartBorder_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbTitleKepemilikan.Checked Then
                txtColorTitleKepemilikan.BackColor = ColorDialog1.Color
                TitleKepemilikan_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                TitleKepemilikan_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                TitleKepemilikan_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbFilter.Checked Then
                txtColorFilter.BackColor = ColorDialog1.Color
                Filter_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                Filter_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                Filter_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbChartTitlePie.Checked Then
                txtColorChartTitlePie.BackColor = ColorDialog1.Color
                ChartTitle_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                ChartTitle_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                ChartTitle_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbChartBorderPie.Checked Then
                txtColorChartBorderPie.BackColor = ColorDialog1.Color
                ChartBorderPie_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                ChartBorderPie_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                ChartBorderPie_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbTableTitle.Checked Then
                txtColorTableTitle.BackColor = ColorDialog1.Color
                TableTitle_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                TableTitle_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                TableTitle_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbTableItem.Checked Then
                txtColorTableItem.BackColor = ColorDialog1.Color
                TableItem_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                TableItem_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                TableItem_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbOutlook.Checked Then
                txtColorOutlook.BackColor = ColorDialog1.Color
                OutlookHolding_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                OutlookHolding_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                OutlookHolding_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            ElseIf rbTentang.Checked Then
                txtColorTentang.BackColor = ColorDialog1.Color
                TentangHolding_R.Text = RGBWrite("R", ColorDialog1.Color.R)
                TentangHolding_G.Text = RGBWrite("G", ColorDialog1.Color.G)
                TentangHolding_B.Text = RGBWrite("B", ColorDialog1.Color.B)
            End If
        End If
    End Sub
    Private Sub rbReportTitle_Click(sender As Object, e As EventArgs) Handles rbReportTitle.Click
        colorSet()
    End Sub
    Private Sub rbNAB_Click(sender As Object, e As EventArgs) Handles rbNAB.Click
        colorSet()
    End Sub
    Private Sub rbBankKustodian_Click(sender As Object, e As EventArgs) Handles rbBankKustodian.Click
        colorSet()
    End Sub
    Private Sub rbTanggalPeluncuran_Click(sender As Object, e As EventArgs) Handles rbTanggalPeluncuran.Click
        colorSet()
    End Sub
    Private Sub rbTotal_Click(sender As Object, e As EventArgs) Handles rbTotal.Click
        colorSet()
    End Sub
    Private Sub rbMataUang_Click(sender As Object, e As EventArgs) Handles rbMataUang.Click
        colorSet()
    End Sub
    Private Sub rbImbalManajer_Click(sender As Object, e As EventArgs) Handles rbImbalManajer.Click
        colorSet()
    End Sub
    Private Sub rbImbalBank_Click(sender As Object, e As EventArgs) Handles rbImbalBank.Click
        colorSet()
    End Sub
    Private Sub rbBiayaBeli_Click(sender As Object, e As EventArgs) Handles rbBiayaBeli.Click
        colorSet()
    End Sub
    Private Sub rbBiayaJual_Click(sender As Object, e As EventArgs) Handles rbBiayaJual.Click
        colorSet()
    End Sub
    Private Sub rbInvestAwal_Click(sender As Object, e As EventArgs) Handles rbInvestAwal.Click
        colorSet()
    End Sub
    Private Sub rbInvestSelajutnya_Click(sender As Object, e As EventArgs) Handles rbInvestSelanjutnya.Click
        colorSet()
    End Sub
    Private Sub rbKode_Click(sender As Object, e As EventArgs) Handles rbKode.Click
        colorSet()
    End Sub
    Private Sub rbKinerjaSejakDiluncurkan_Click(sender As Object, e As EventArgs) Handles rbKinerjaSejakDiluncurkan.Click
        colorSet()
    End Sub
    Private Sub rbKinerjaTerbaik_Click(sender As Object, e As EventArgs) Handles rbKinerjaTerbaik.Click
        colorSet()
    End Sub
    Private Sub rbKinerjaTerburuk_Click(sender As Object, e As EventArgs) Handles rbKinerjaTerburuk.Click
        colorSet()
    End Sub
    Private Sub rbFaktorRisikoUtama_Click(sender As Object, e As EventArgs) Handles rbFaktorRisikUtama.Click
        colorSet()
    End Sub
    Private Sub rbPeriodeInvestasi_Click(sender As Object, e As EventArgs) Handles rbPeriodeInvestasi.Click
        colorSet()
    End Sub
    Private Sub rbTingkatRisiko_Click(sender As Object, e As EventArgs) Handles rbTingkatRisiko.Click
        colorSet()
    End Sub
    Private Sub rbTujuanInvestasi_Click(sender As Object, e As EventArgs) Handles rbTujuanInvestasi.Click
        colorSet()
    End Sub
    Private Sub rbInvestasi_Click(sender As Object, e As EventArgs) Handles rbInvestasi.Click
        colorSet()
    End Sub
    Private Sub rbPortofolio_Click(sender As Object, e As EventArgs) Handles rbPortofolio.Click
        colorSet()
    End Sub
    Private Sub rbKinerjaReksaDana_Click(sender As Object, e As EventArgs) Handles rbChartTitle.Click
        colorSet()
    End Sub
    Private Sub rbChartBorder_Click(sender As Object, e As EventArgs) Handles rbChartBorder.Click
        colorSet()
    End Sub
    Private Sub chkChartBorder_Click(sender As Object, e As EventArgs) Handles chkChartBorder.Click
        If chkChartBorder.Checked Then chkChartBorder.Text = "Y" Else chkChartBorder.Text = "N"
    End Sub
    Private Sub rbKepemilikan_Click(sender As Object, e As EventArgs) Handles rbTitleKepemilikan.Click
        colorSet()
    End Sub
    Private Sub rbFilter_Click(sender As Object, e As EventArgs) Handles rbFilter.Click
        colorSet()
    End Sub
    Private Sub rbChartTitlePie_Click(sender As Object, e As EventArgs) Handles rbChartTitlePie.Click
        colorSet()
    End Sub
    Private Sub rbChartBorderPie_Click(sender As Object, e As EventArgs) Handles rbChartBorderPie.Click
        colorSet()
    End Sub
    Private Sub chkChartBorderpie_Click(sender As Object, e As EventArgs) Handles chkChartBorderPie.Click
        If chkChartBorderPie.Checked Then chkChartBorderPie.Text = "Y" Else chkChartBorderPie.Text = "N"
    End Sub
    Private Sub rbTableTitle_Click(sender As Object, e As EventArgs) Handles rbTableTitle.Click
        colorSet()
    End Sub
    Private Sub rbTableItem_Click(sender As Object, e As EventArgs) Handles rbTableItem.Click
        colorSet()
    End Sub
    Private Sub rbOutlook_Click(sender As Object, e As EventArgs) Handles rbOutlook.Click
        colorSet()
    End Sub
    Private Sub rbTentang_Click(sender As Object, e As EventArgs) Handles rbTentang.Click
        colorSet()
    End Sub

#End Region

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Close()
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

                    r = file.GetInteger(reportSection, iniType & " NAB/Unit R", 0)
                    g = file.GetInteger(reportSection, iniType & " NAB/Unit G", 0)
                    b = file.GetInteger(reportSection, iniType & " NAB/Unit B", 0)
                    NAB_R.Text = RGBWrite("R", r)
                    NAB_G.Text = RGBWrite("G", g)
                    NAB_B.Text = RGBWrite("B", b)
                    txtColorNAB.BackColor = Color.FromArgb(r, g, b)
                    txtNAB.Text = file.GetString(reportSection, iniType & " NAB/Unit", "")

                    r = file.GetInteger(reportSection, iniType & " Bank Kustodian R", 0)
                    g = file.GetInteger(reportSection, iniType & " Bank Kustodian G", 0)
                    b = file.GetInteger(reportSection, iniType & " Bank Kustodian B", 0)
                    Bank_R.Text = RGBWrite("R", r)
                    Bank_G.Text = RGBWrite("G", g)
                    Bank_B.Text = RGBWrite("B", b)
                    txtColorBank.BackColor = Color.FromArgb(r, g, b)
                    txtBankKustodian.Text = file.GetString(reportSection, iniType & " Bank Kustodian", "")

                    r = file.GetInteger(reportSection, iniType & " Tanggal R", 0)
                    g = file.GetInteger(reportSection, iniType & " Tanggal G", 0)
                    b = file.GetInteger(reportSection, iniType & " Tanggal B", 0)
                    Tanggal_R.Text = RGBWrite("R", r)
                    Tanggal_G.Text = RGBWrite("G", g)
                    Tanggal_B.Text = RGBWrite("B", b)
                    txtColorTanggal.BackColor = Color.FromArgb(r, g, b)
                    txtTanggalPeluncuran.Text = file.GetString(reportSection, iniType & " Tanggal Peluncuran", "")

                    r = file.GetInteger(reportSection, iniType & " Total AUM R", 0)
                    g = file.GetInteger(reportSection, iniType & " Total AUM G", 0)
                    b = file.GetInteger(reportSection, iniType & " Total AUM B", 0)
                    Total_R.Text = RGBWrite("R", r)
                    Total_G.Text = RGBWrite("G", g)
                    Total_B.Text = RGBWrite("B", b)
                    txtColorTotal.BackColor = Color.FromArgb(r, g, b)
                    txtTotalAUM.Text = file.GetString(reportSection, iniType & " Total AUM", "")

                    r = file.GetInteger(reportSection, iniType & " Mata Uang R", 0)
                    g = file.GetInteger(reportSection, iniType & " Mata Uang G", 0)
                    b = file.GetInteger(reportSection, iniType & " Mata Uang B", 0)
                    MataUang_R.Text = RGBWrite("R", r)
                    MataUang_G.Text = RGBWrite("G", g)
                    MataUang_B.Text = RGBWrite("B", b)
                    txtColorMatUang.BackColor = Color.FromArgb(r, g, b)
                    txtMataUang.Text = file.GetString(reportSection, iniType & " Mata Uang", "")

                    r = file.GetInteger(reportSection, iniType & " Imbal Jasa Manajer R", 0)
                    g = file.GetInteger(reportSection, iniType & " Imbal Jasa Manajer G", 0)
                    b = file.GetInteger(reportSection, iniType & " Imbal Jasa Manajer B", 0)
                    ImbalManajer_R.Text = RGBWrite("R", r)
                    ImbalManajer_G.Text = RGBWrite("G", g)
                    ImbalManajer_B.Text = RGBWrite("B", b)
                    txtColorImbalManajer.BackColor = Color.FromArgb(r, g, b)
                    txtImbalJasaManajer.Text = file.GetString(reportSection, iniType & " Imbal Jasa Manajer Investasi", "")

                    r = file.GetInteger(reportSection, iniType & " Imbal Jasa Bank R", 0)
                    g = file.GetInteger(reportSection, iniType & " Imbal Jasa Bank G", 0)
                    b = file.GetInteger(reportSection, iniType & " Imbal Jasa Bank B", 0)
                    ImbalBank_R.Text = RGBWrite("R", r)
                    ImbalBank_G.Text = RGBWrite("G", g)
                    ImbalBank_B.Text = RGBWrite("B", b)
                    txtColorImbalBank.BackColor = Color.FromArgb(r, g, b)
                    txtImbalJasaBank.Text = file.GetString(reportSection, iniType & " Imbal Jasa Bank Kustodian", "")

                    r = file.GetInteger(reportSection, iniType & " Biaya Pembelian R", 0)
                    g = file.GetInteger(reportSection, iniType & " Biaya Pembelian G", 0)
                    b = file.GetInteger(reportSection, iniType & " Biaya Pembelian B", 0)
                    BiayaBeli_R.Text = RGBWrite("R", r)
                    BiayaBeli_G.Text = RGBWrite("G", g)
                    BiayaBeli_B.Text = RGBWrite("B", b)
                    txtColorBiayaBeli.BackColor = Color.FromArgb(r, g, b)
                    txtBiayaPembelian.Text = file.GetString(reportSection, iniType & " Biaya Pembelian", "")

                    r = file.GetInteger(reportSection, iniType & " Biaya Penjualan R", 0)
                    g = file.GetInteger(reportSection, iniType & " Biaya Penjualan G", 0)
                    b = file.GetInteger(reportSection, iniType & " Biaya Penjualan B", 0)
                    BiayaJual_R.Text = RGBWrite("R", r)
                    BiayaJual_G.Text = RGBWrite("G", g)
                    BiayaJual_B.Text = RGBWrite("B", b)
                    txtColorBiayaJual.BackColor = Color.FromArgb(r, g, b)
                    txtBiayaPenjualaKembali.Text = file.GetString(reportSection, iniType & " Biaya Penjualan Kembali", "")

                    r = file.GetInteger(reportSection, iniType & " Investasi Awal R", 0)
                    g = file.GetInteger(reportSection, iniType & " Investasi Awal G", 0)
                    b = file.GetInteger(reportSection, iniType & " Investasi Awal B", 0)
                    InvestAwal_R.Text = RGBWrite("R", r)
                    InvestAwal_G.Text = RGBWrite("G", g)
                    InvestAwal_B.Text = RGBWrite("B", b)
                    txtColorMinInvest.BackColor = Color.FromArgb(r, g, b)
                    txtMinimalInvestasAwal.Text = file.GetString(reportSection, iniType & " Minimal Investasi Awal", "")

                    r = file.GetInteger(reportSection, iniType & " Investasi Selanjutnya R", 0)
                    g = file.GetInteger(reportSection, iniType & " Investasi Selanjutnya G", 0)
                    b = file.GetInteger(reportSection, iniType & " Investasi Selanjutnya B", 0)
                    InvestSelanjutnya_R.Text = RGBWrite("R", r)
                    InvestSelanjutnya_G.Text = RGBWrite("G", g)
                    InvestSelanjutnya_B.Text = RGBWrite("B", b)
                    txtColorMinInvestSelanjutnya.BackColor = Color.FromArgb(r, g, b)
                    txtMinimalInvestasiSelanjutnya.Text = file.GetString(reportSection, iniType & " Minimal Investasi Selanjutnya", "")

                    r = file.GetInteger(reportSection, iniType & " Kode R", 0)
                    g = file.GetInteger(reportSection, iniType & " Kode G", 0)
                    b = file.GetInteger(reportSection, iniType & " Kode B", 0)
                    Kode_R.Text = RGBWrite("R", r)
                    Kode_G.Text = RGBWrite("G", g)
                    Kode_B.Text = RGBWrite("B", b)
                    txtColorKode.BackColor = Color.FromArgb(r, g, b)
                    txtKodeISINBloomberg.Text = file.GetString(reportSection, iniType & " Kode ISIN/Bloomberg", "")

                    r = file.GetInteger(reportSection, iniType & " Kinerja Sejak Diluncurkan R", 0)
                    g = file.GetInteger(reportSection, iniType & " Kinerja Sejak Diluncurkan G", 0)
                    b = file.GetInteger(reportSection, iniType & " Kinerja Sejak Diluncurkan B", 0)
                    KinerjaAwal_R.Text = RGBWrite("R", r)
                    KinerjaAwal_G.Text = RGBWrite("G", g)
                    KinerjaAwal_B.Text = RGBWrite("B", b)
                    txtColorKinerja.BackColor = Color.FromArgb(r, g, b)
                    txtKinerjaSejakDiluncurkan.Text = file.GetString(reportSection, iniType & " Kinerja Sejak Diluncurkan", "")

                    r = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terbaik R", 0)
                    g = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terbaik G", 0)
                    b = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terbaik B", 0)
                    KinerjaTerbaik_R.Text = RGBWrite("R", r)
                    KinerjaTerbaik_G.Text = RGBWrite("G", g)
                    KinerjaTerbaik_B.Text = RGBWrite("B", b)
                    txtColorKinerjaTerbaik.BackColor = Color.FromArgb(r, g, b)
                    txtKinerjaBulanTerbaik.Text = file.GetString(reportSection, iniType & " Kinerja Bulan Terbaik", "")

                    r = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terburuk R", 0)
                    g = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terburuk G", 0)
                    b = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terburuk B", 0)
                    KinerjaTerburuk_R.Text = RGBWrite("R", r)
                    KinerjaTerburuk_G.Text = RGBWrite("G", g)
                    KinerjaTerburuk_B.Text = RGBWrite("B", b)
                    txtColorKinerjaTerburuk.BackColor = Color.FromArgb(r, g, b)
                    txtKinerjaBulanTerburuk.Text = file.GetString(reportSection, iniType & " Kinerja Bulan Terburuk", "")

                    r = file.GetInteger(reportSection, iniType & " Faktor Risiko Utama R", 0)
                    g = file.GetInteger(reportSection, iniType & " Faktor Risiko Utama G", 0)
                    b = file.GetInteger(reportSection, iniType & " Faktor Risiko Utama B", 0)
                    FaktorRisiko_R.Text = RGBWrite("R", r)
                    FaktorRisiko_G.Text = RGBWrite("G", g)
                    FaktorRisiko_B.Text = RGBWrite("B", b)
                    txtColorFaktorRisiko.BackColor = Color.FromArgb(r, g, b)
                    txtFaktorRisikoUtama.Text = file.GetString(reportSection, iniType & " Faktor Risiko Utama", "")

                    r = file.GetInteger(reportSection, iniType & " Periode Investasi R", 0)
                    g = file.GetInteger(reportSection, iniType & " Periode Investasi G", 0)
                    b = file.GetInteger(reportSection, iniType & " Periode Investasi B", 0)
                    PeriodeInvestasi_R.Text = RGBWrite("R", r)
                    PeriodeInvestasi_G.Text = RGBWrite("G", g)
                    PeriodeInvestasi_B.Text = RGBWrite("B", b)
                    txtColorPeriode.BackColor = Color.FromArgb(r, g, b)
                    txtPerodeInvestasi.Text = file.GetString(reportSection, iniType & " Periode Investasi", "")

                    r = file.GetInteger(reportSection, iniType & " Tingkat Risiko R", 0)
                    g = file.GetInteger(reportSection, iniType & " Tingkat Risiko G", 0)
                    b = file.GetInteger(reportSection, iniType & " Tingkat Risiko B", 0)
                    TingkatRisiko_R.Text = RGBWrite("R", r)
                    TIngkatRisiko_G.Text = RGBWrite("G", g)
                    TingkatRisiko_B.Text = RGBWrite("B", b)
                    txtColorTingkatRisiko.BackColor = Color.FromArgb(r, g, b)
                    txtTingkatRisiko.Text = file.GetString(reportSection, iniType & " Tingkat Risiko", "")

                    r = file.GetInteger(reportSection, iniType & " Tujuan Investasi R", 0)
                    g = file.GetInteger(reportSection, iniType & " Tujuan Investasi G", 0)
                    b = file.GetInteger(reportSection, iniType & " Tujuan Investasi B", 0)
                    TujuanInvestasi_R.Text = RGBWrite("R", r)
                    TujuanInvestasi_G.Text = RGBWrite("G", g)
                    TujuanInvestasi_B.Text = RGBWrite("B", b)
                    txtColorTujuanInvestasi.BackColor = Color.FromArgb(r, g, b)
                    txtTujuanInvestasi.Text = file.GetString(reportSection, iniType & " Tujuan Investasi", "")

                    r = file.GetInteger(reportSection, iniType & " Investasi R", 0)
                    g = file.GetInteger(reportSection, iniType & " Investasi G", 0)
                    b = file.GetInteger(reportSection, iniType & " Investasi B", 0)
                    Investasi_R.Text = RGBWrite("R", r)
                    Investasi_G.Text = RGBWrite("G", g)
                    Investasi_B.Text = RGBWrite("B", b)
                    txtColorInvestasi.BackColor = Color.FromArgb(r, g, b)
                    txtKebijakanInvestasi.Text = file.GetString(reportSection, iniType & " Kebijakan Investasi", "")
                    txtInvestasiSaham.Text = file.GetString(reportSection, iniType & " Saham", "")
                    txtInvestasiObligasi.Text = file.GetString(reportSection, iniType & " Obligasi", "")
                    txtInvesatiPasarUang.Text = file.GetString(reportSection, iniType & " Pasar Uang", "")

                    r = file.GetInteger(reportSection, iniType & " Portofolio R", 0)
                    g = file.GetInteger(reportSection, iniType & " Portofolio G", 0)
                    b = file.GetInteger(reportSection, iniType & " Portofolio B", 0)
                    Portofolio_R.Text = RGBWrite("R", r)
                    Portofolio_G.Text = RGBWrite("G", g)
                    Portofolio_B.Text = RGBWrite("B", b)
                    txtColorKomposisPortofolio.BackColor = Color.FromArgb(r, g, b)
                    txtKomposisiPortofolio.Text = file.GetString(reportSection, iniType & " Komposisi Portofolio", "")
                    txtPortofolioSaham.Text = file.GetString(reportSection, iniType & " Saham", "")
                    txtPortofoliObligasi.Text = file.GetString(reportSection, iniType & " Obligasi", "")
                    txtPortofolioPasarUang.Text = file.GetString(reportSection, iniType & " Pasar Uang", "")

                    r = file.GetInteger(reportSection, iniType & " Line Chart Title R", 0)
                    g = file.GetInteger(reportSection, iniType & " Line Chart Title G", 0)
                    b = file.GetInteger(reportSection, iniType & " Line Chart Title B", 0)
                    KinerjaReksaDana_R.Text = RGBWrite("R", r)
                    KinerjaReksaDana_G.Text = RGBWrite("G", g)
                    KinerjaReksaDana_B.Text = RGBWrite("B", b)
                    txtColorChartTitle.BackColor = Color.FromArgb(r, g, b)
                    txtKinerjaReksaDanaChart.Text = file.GetString(reportSection, iniType & " Line Chart Title", "")

                    r = file.GetInteger(reportSection, iniType & " Line Chart Border R", 0)
                    g = file.GetInteger(reportSection, iniType & " Line Chart Border G", 0)
                    b = file.GetInteger(reportSection, iniType & " Line Chart Border B", 0)
                    ChartBorder_R.Text = RGBWrite("R", r)
                    ChartBorder_G.Text = RGBWrite("G", g)
                    ChartBorder_B.Text = RGBWrite("B", b)
                    txtColorChartBorder.BackColor = Color.FromArgb(r, g, b)
                    If file.GetBoolean(reportSection, iniType & " Line Chart Border", False) Then chkChartBorder.Checked = True Else chkChartBorder.Checked = False

                    r = file.GetInteger(reportSection, iniType & " Title Kepemilikan R", 0)
                    g = file.GetInteger(reportSection, iniType & " Title Kepemilikan G", 0)
                    b = file.GetInteger(reportSection, iniType & " Title Kepemilikan B", 0)
                    TitleKepemilikan_R.Text = RGBWrite("R", r)
                    TitleKepemilikan_G.Text = RGBWrite("G", g)
                    TitleKepemilikan_B.Text = RGBWrite("B", b)
                    txtColorTitleKepemilikan.BackColor = Color.FromArgb(r, g, b)
                    txtTitleKepemilikan.Text = file.GetString(reportSection, iniType & " Title Kepemilikan", "")

                    r = file.GetInteger(reportSection, iniType & " Filter R", 0)
                    g = file.GetInteger(reportSection, iniType & " Filter G", 0)
                    b = file.GetInteger(reportSection, iniType & " Filter B", 0)
                    Filter_R.Text = RGBWrite("R", r)
                    Filter_G.Text = RGBWrite("G", g)
                    Filter_B.Text = RGBWrite("B", b)
                    txtColorFilter.BackColor = Color.FromArgb(r, g, b)
                    txtFilter.Text = file.GetString(reportSection, iniType & " Filter", "")

                    r = file.GetInteger(reportSection, iniType & " Pie Chart Title R", 0)
                    g = file.GetInteger(reportSection, iniType & " Pie Chart Title G", 0)
                    b = file.GetInteger(reportSection, iniType & " Pie Chart Title B", 0)
                    ChartTitle_R.Text = RGBWrite("R", r)
                    ChartTitle_G.Text = RGBWrite("G", g)
                    ChartTitle_B.Text = RGBWrite("B", b)
                    txtColorChartTitlePie.BackColor = Color.FromArgb(r, g, b)
                    txtAlokasiSektor.Text = file.GetString(reportSection, iniType & " Pie Chart Title", "")

                    r = file.GetInteger(reportSection, iniType & " Pie Chart Border R", 0)
                    g = file.GetInteger(reportSection, iniType & " Pie Chart Border G", 0)
                    b = file.GetInteger(reportSection, iniType & " Pie Chart Border B", 0)
                    ChartBorderPie_R.Text = RGBWrite("R", r)
                    ChartBorderPie_G.Text = RGBWrite("G", g)
                    ChartBorderPie_B.Text = RGBWrite("B", b)
                    txtColorChartBorderPie.BackColor = Color.FromArgb(r, g, b)
                    If file.GetBoolean(reportSection, iniType & " Pie Chart Border", False) Then chkChartBorderPie.Checked = True Else chkChartBorderPie.Checked = False

                    r = file.GetInteger(reportSection, iniType & " Table Title R", 0)
                    g = file.GetInteger(reportSection, iniType & " Table Title G", 0)
                    b = file.GetInteger(reportSection, iniType & " Table Title B", 0)
                    TableTitle_R.Text = RGBWrite("R", r)
                    TableTitle_G.Text = RGBWrite("G", g)
                    TableTitle_B.Text = RGBWrite("B", b)
                    txtColorTableTitle.BackColor = Color.FromArgb(r, g, b)
                    txtTableTitle.Text = file.GetString(reportSection, iniType & " Table Title", "")

                    r = file.GetInteger(reportSection, iniType & " Table Item R", 0)
                    g = file.GetInteger(reportSection, iniType & " Table Item G", 0)
                    b = file.GetInteger(reportSection, iniType & " Table Item B", 0)
                    TableItem_R.Text = RGBWrite("R", r)
                    TableItem_G.Text = RGBWrite("G", g)
                    TableItem_B.Text = RGBWrite("B", b)
                    txtColorTableItem.BackColor = Color.FromArgb(r, g, b)
                    txtTableItemReturn.Text = file.GetString(reportSection, iniType & " Return", "")
                    txtTableItemBenchmark.Text = file.GetString(reportSection, iniType & " Benchmark", "")
                    txtTableItem1Bulan.Text = file.GetString(reportSection, iniType & " 1 Bulan", "")
                    txtTableItem3Bulan.Text = file.GetString(reportSection, iniType & " 3 Bulan", "")
                    txtTableItem6Bulan.Text = file.GetString(reportSection, iniType & " 6 Bulan", "")
                    txtTableItem1Tahun.Text = file.GetString(reportSection, iniType & " 1 Tahun", "")
                    txtTableItemDariAwalTahun.Text = file.GetString(reportSection, iniType & " Dari Awal", "")
                    txtTableItemSejakPembentukan.Text = file.GetString(reportSection, iniType & " Sejak Pembentukan", "")

                    r = file.GetInteger(reportSection, iniType & " Title Ulasan R", 0)
                    g = file.GetInteger(reportSection, iniType & " Title Ulasan G", 0)
                    b = file.GetInteger(reportSection, iniType & " Title Ulasan B", 0)
                    OutlookHolding_R.Text = RGBWrite("R", r)
                    OutlookHolding_G.Text = RGBWrite("G", g)
                    OutlookHolding_B.Text = RGBWrite("B", b)
                    txtColorOutlook.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " Title MI R", 0)
                    g = file.GetInteger(reportSection, iniType & " Title MI G", 0)
                    b = file.GetInteger(reportSection, iniType & " Title MI B", 0)
                    TentangHolding_R.Text = RGBWrite("R", r)
                    TentangHolding_G.Text = RGBWrite("G", g)
                    TentangHolding_B.Text = RGBWrite("B", b)
                    txtColorTentang.BackColor = Color.FromArgb(r, g, b)
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

            NAB_R.Text = "R: " & .NAB_R
            NAB_G.Text = "G: " & .NAB_G
            NAB_B.Text = "B: " & .NAB_B
            txtColorNAB.BackColor = Color.FromArgb(.NAB_R, .NAB_G, .NAB_B)
            txtNAB.Text = .NAB

            Bank_R.Text = "R: " & .Bank_R
            Bank_G.Text = "G: " & .Bank_G
            Bank_B.Text = "B: " & .Bank_B
            txtColorBank.BackColor = Color.FromArgb(.Bank_R, .Bank_G, .Bank_B)
            txtBankKustodian.Text = .Bank

            Tanggal_R.Text = "R: " & .Tanggal_R
            Tanggal_G.Text = "G: " & .Tanggal_G
            Tanggal_B.Text = "B: " & .Tanggal_B
            txtColorTanggal.BackColor = Color.FromArgb(.Tanggal_R, .Tanggal_G, .Tanggal_B)
            txtTanggalPeluncuran.Text = .Tanggal

            Total_R.Text = "R: " & .Total_R
            Total_G.Text = "G: " & .Total_G
            Total_B.Text = "B: " & .Total_B
            txtColorTotal.BackColor = Color.FromArgb(.Total_R, .Total_G, .Total_B)
            txtTotalAUM.Text = .Total

            MataUang_R.Text = "R: " & .MataUang_R
            MataUang_G.Text = "G: " & .MataUang_G
            MataUang_B.Text = "B: " & .MataUang_B
            txtColorMatUang.BackColor = Color.FromArgb(.MataUang_R, .MataUang_G, .MataUang_B)
            txtMataUang.Text = .MataUang

            ImbalManajer_R.Text = "R: " & .ImbalManajer_R
            ImbalManajer_G.Text = "G: " & .ImbalManajer_G
            ImbalManajer_B.Text = "B: " & .ImbalManajer_B
            txtColorImbalManajer.BackColor = Color.FromArgb(.ImbalManajer_R, .ImbalManajer_G, .ImbalManajer_B)
            txtImbalJasaManajer.Text = .ImbalManajer

            ImbalBank_R.Text = "R: " & .ImbalBank_R
            ImbalBank_G.Text = "G: " & .ImbalBank_G
            ImbalBank_B.Text = "B: " & .ImbalBank_B
            txtColorImbalBank.BackColor = Color.FromArgb(.ImbalBank_R, .ImbalBank_G, .ImbalBank_B)
            txtImbalJasaBank.Text = .ImbalBank

            BiayaBeli_R.Text = "R: " & .BiayaBeli_R
            BiayaBeli_G.Text = "G: " & .BiayaBeli_G
            BiayaBeli_B.Text = "B: " & .BiayaBeli_B
            txtColorBiayaBeli.BackColor = Color.FromArgb(.BiayaBeli_R, .BiayaBeli_G, .BiayaBeli_B)
            txtBiayaPembelian.Text = .BiayaBeli

            BiayaJual_R.Text = "R: " & .BiayaJual_R
            BiayaJual_G.Text = "G: " & .BiayaJual_G
            BiayaJual_B.Text = "B: " & .BiayaJual_B
            txtColorBiayaJual.BackColor = Color.FromArgb(.BiayaJual_R, .BiayaJual_G, .BiayaJual_B)
            txtBiayaPenjualaKembali.Text = .BiayaJual

            InvestAwal_R.Text = "R: " & .InvestsiAwal_R
            InvestAwal_G.Text = "G: " & .InvestsiAwal_G
            InvestAwal_B.Text = "B: " & .InvestsiAwal_B
            txtColorInvestasi.BackColor = Color.FromArgb(.InvestsiAwal_R, .InvestsiAwal_G, .InvestsiAwal_B)
            txtMinimalInvestasAwal.Text = .InvestsiAwal

            InvestSelanjutnya_R.Text = "R: " & .InvestasiSelanjutnya_R
            InvestSelanjutnya_G.Text = "G: " & .InvestasiSelanjutnya_G
            InvestSelanjutnya_B.Text = "B: " & .InvestasiSelanjutnya_B
            txtColorMinInvestSelanjutnya.BackColor = Color.FromArgb(.InvestasiSelanjutnya_R, .InvestasiSelanjutnya_G, .InvestasiSelanjutnya_B)
            txtMinimalInvestasiSelanjutnya.Text = .InvestasiSelanjutnya

            Kode_R.Text = "R: " & .Kode_R
            Kode_G.Text = "G: " & .Kode_G
            Kode_B.Text = "B: " & .Kode_B
            txtColorKode.BackColor = Color.FromArgb(.Kode_R, .Kode_G, .Kode_B)
            txtKodeISINBloomberg.Text = .Kode

            KinerjaAwal_R.Text = "R: " & .KinerjaAwal_R
            KinerjaAwal_G.Text = "G: " & .KinerjaAwal_G
            KinerjaAwal_B.Text = "B: " & .KinerjaAwal_B
            txtColorKinerja.BackColor = Color.FromArgb(.KinerjaAwal_R, .KinerjaAwal_G, .KinerjaAwal_B)
            txtKinerjaSejakDiluncurkan.Text = .KinerjaAwal

            KinerjaTerbaik_R.Text = "R: " & .KinerjaTerbaik_R
            KinerjaTerbaik_G.Text = "G: " & .KinerjaTerbaik_G
            KinerjaTerbaik_B.Text = "B: " & .KinerjaTerbaik_B
            txtColorKinerjaTerbaik.BackColor = Color.FromArgb(.KinerjaTerbaik_R, .KinerjaTerbaik_G, .KinerjaTerbaik_B)
            txtKinerjaBulanTerbaik.Text = .KinerjaTerbaik

            KinerjaTerburuk_R.Text = "R: " & .KinerjaTerburuk_R
            KinerjaTerburuk_G.Text = "G: " & .KinerjaTerburuk_G
            KinerjaTerburuk_B.Text = "B: " & .KinerjaTerburuk_B
            txtColorKinerjaTerburuk.BackColor = Color.FromArgb(.KinerjaTerburuk_R, .KinerjaTerburuk_G, .KinerjaTerburuk_B)
            txtKinerjaBulanTerburuk.Text = .KinerjaTerburuk

            FaktorRisiko_R.Text = "R: " & .FaktorRisiko_R
            FaktorRisiko_G.Text = "G: " & .FaktorRisiko_G
            FaktorRisiko_B.Text = "B: " & .FaktorRisiko_B
            txtColorFaktorRisiko.BackColor = Color.FromArgb(.FaktorRisiko_R, .FaktorRisiko_G, .FaktorRisiko_B)
            txtFaktorRisikoUtama.Text = .FaktorRisiko

            PeriodeInvestasi_R.Text = "R: " & .PeriodeInvestasi_R
            PeriodeInvestasi_G.Text = "G: " & .PeriodeInvestasi_G
            PeriodeInvestasi_B.Text = "B: " & .PeriodeInvestasi_B
            txtColorPeriode.BackColor = Color.FromArgb(.PeriodeInvestasi_R, .PeriodeInvestasi_G, .PeriodeInvestasi_B)
            txtPerodeInvestasi.Text = .PeriodeInvestasi

            TingkatRisiko_R.Text = "R: " & .TingkatRisiko_R
            TIngkatRisiko_G.Text = "G: " & .TingkatRisiko_G
            TingkatRisiko_B.Text = "B: " & .TingkatRisiko_B
            txtColorTingkatRisiko.BackColor = Color.FromArgb(.TingkatRisiko_R, .TingkatRisiko_G, .TingkatRisiko_B)
            txtTingkatRisiko.Text = .TingkatRisiko

            TujuanInvestasi_R.Text = "R: " & .TujuanInvestasi_R
            TujuanInvestasi_G.Text = "G: " & .TujuanInvestasi_G
            TujuanInvestasi_B.Text = "B: " & .TujuanInvestasi_B
            txtColorTujuanInvestasi.BackColor = Color.FromArgb(.TujuanInvestasi_R, .TujuanInvestasi_G, .TujuanInvestasi_B)
            txtTujuanInvestasi.Text = .TujuanInvestasi

            Investasi_R.Text = "R: " & .Investasi_R
            Investasi_G.Text = "G: " & .Investasi_G
            Investasi_B.Text = "B: " & .Investasi_B
            txtColorInvestasi.BackColor = Color.FromArgb(.Investasi_R, .Investasi_G, .Investasi_B)
            txtKebijakanInvestasi.Text = .Investasi
            txtInvestasiSaham.Text = .InvestasiSaham
            txtInvestasiObligasi.Text = .investasiObligasi
            txtInvesatiPasarUang.Text = .InvestasiPasarUang

            txtInvestasiSaham.Text = .InvestasiSaham
            txtInvestasiObligasi.Text = .investasiObligasi
            txtInvesatiPasarUang.Text = .InvestasiPasarUang
            Portofolio_R.Text = "R: " & .Portofolio_R
            Portofolio_G.Text = "G: " & .Portofolio_G
            Portofolio_B.Text = "B: " & .Portofolio_B
            txtColorKomposisPortofolio.BackColor = Color.FromArgb(.Portofolio_R, .Portofolio_G, .Portofolio_B)
            txtKomposisiPortofolio.Text = .Portofolio
            txtPortofolioSaham.Text = .PortofolioSaham
            txtPortofoliObligasi.Text = .PortifolioObligasi
            txtPortofolioPasarUang.Text = .PortofolioPasarUang

            KinerjaReksaDana_R.Text = "R: " & .KinerjaReksaDana_R
            KinerjaReksaDana_G.Text = "G: " & .KinerjaReksaDana_G
            KinerjaReksaDana_B.Text = "B: " & .KinerjaReksaDana_B
            txtColorChartTitle.BackColor = Color.FromArgb(.KinerjaReksaDana_R, .KinerjaReksaDana_G, .KinerjaReksaDana_B)
            txtKinerjaReksaDanaChart.Text = .KinerjaReksaDana

            ChartBorder_R.Text = "R: " & .ChartBorder_R
            ChartBorder_G.Text = "G: " & .ChartBorder_G
            ChartBorder_B.Text = "B: " & .ChartBorder_B
            txtColorChartBorder.BackColor = Color.FromArgb(.ChartBorder_R, .ChartBorder_G, .ChartBorder_B)

            TitleKepemilikan_R.Text = "R: " & .TitleKepemilikan_R
            TitleKepemilikan_G.Text = "G: " & .TitleKepemilikan_G
            TitleKepemilikan_R.Text = "B: " & .TitleKepemilikan_B
            txtColorTitleKepemilikan.BackColor = Color.FromArgb(.TitleKepemilikan_R, .TitleKepemilikan_G, .TitleKepemilikan_B)
            txtTitleKepemilikan.Text = .TitleKepemilikan

            Filter_R.Text = "R: " & .Filter_R
            Filter_G.Text = "G: " & .Filter_G
            Filter_B.Text = "B: " & .Filter_B
            txtColorFilter.BackColor = Color.FromArgb(.Filter_R, .Filter_G, .Filter_B)
            txtFilter.Text = .Filter

            ChartTitle_R.Text = "R: " & .ChartTitle_R
            ChartTitle_G.Text = "G: " & .ChartTitle_G
            ChartTitle_B.Text = "B: " & .ChartTitle_B
            txtColorChartTitlePie.BackColor = Color.FromArgb(.ChartTitle_R, .ChartTitle_G, .ChartTitle_B)
            txtAlokasiSektor.Text = .ChartTitle

            ChartBorderPie_R.Text = "R: " & .ChartBorderPie_R
            ChartBorderPie_G.Text = "G: " & .ChartBorderPie_G
            ChartBorderPie_B.Text = "B: " & .ChartBorderPie_B
            txtColorChartBorderPie.BackColor = Color.FromArgb(.ChartBorderPie_R, .ChartBorderPie_G, .ChartBorderPie_B)

            TableTitle_R.Text = "R: " & .TableTitle_R
            TableTitle_G.Text = "G: " & .TableTitle_G
            TableTitle_B.Text = "B: " & .TableTitle_B
            txtColorTableTitle.BackColor = Color.FromArgb(.TableTitle_R, .TableTitle_G, .TableTitle_B)
            txtTableTitle.Text = .TableTitle

            TableItem_R.Text = "R: " & .TableItem_R
            TableItem_G.Text = "G: " & .TableItem_G
            TableItem_B.Text = "B: " & .TableItem_B
            txtColorTableItem.BackColor = Color.FromArgb(.TableItem_R, .TableItem_G, .TableItem_B)
            txtTableItemReturn.Text = .TableItemReturn
            txtTableItemBenchmark.Text = .TableItemBenchmark
            txtTableItem1Bulan.Text = .TableItem1Bulan
            txtTableItem3Bulan.Text = .TableItem3Bulan
            txtTableItem6Bulan.Text = .TableItem6Bulan
            txtTableItem1Tahun.Text = .TableItem1Tahun
            txtTableItemDariAwalTahun.Text = .TableItemDariAwal
            txtTableItemSejakPembentukan.Text = .TableItemSejakPembentukan

            OutlookHolding_R.Text = "R: " & .OutlookHolding_R
            OutlookHolding_G.Text = "G: " & .OutlookHolding_G
            OutlookHolding_R.Text = "B: " & .OutlookHolding_B
            txtColorOutlook.BackColor = Color.FromArgb(.OutlookHolding_R, .OutlookHolding_G, .OutlookHolding_B)

            TentangHolding_R.Text = "R: " & .TentangHolding_R
            TentangHolding_G.Text = "G: " & .TentangHolding_G
            TentangHolding_B.Text = "B: " & .TentangHolding_B
            txtColorTentang.BackColor = Color.FromArgb(.TentangHolding_R, .TentangHolding_G, .TentangHolding_B)
        End With
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

                file.WriteInteger(reportSection, iniType & " NAB/Unit R", RGBRead(NAB_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " NAB/Unit G", RGBRead(NAB_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " NAB/Unit B", RGBRead(NAB_B.Text.Trim))
                file.WriteString(reportSection, iniType & " NAB/Unit", txtNAB.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Bank Kustodian R", RGBRead(Bank_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Bank Kustodian G", RGBRead(Bank_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Bank Kustodian B", RGBRead(Bank_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Bank Kustodian", txtBankKustodian.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Tanggal R", RGBRead(Tanggal_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tanggal G", RGBRead(Tanggal_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tanggal B", RGBRead(Tanggal_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Tanggal Peluncuran", txtTanggalPeluncuran.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Total AUM R", RGBRead(Total_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Total AUM G", RGBRead(Total_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Total AUM B", RGBRead(Total_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Total AUM", txtTotalAUM.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Mata Uang R", RGBRead(MataUang_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Mata Uang G", RGBRead(MataUang_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Mata Uang B", RGBRead(MataUang_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Mata Uang", txtMataUang.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Imbal Jasa Manajer R", RGBRead(ImbalManajer_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Imbal Jasa Manajer G", RGBRead(ImbalManajer_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Imbal Jasa Manajer B", RGBRead(ImbalManajer_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Imbal Jasa Manajer Investasi", txtImbalJasaManajer.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Imbal Jasa Bank R", RGBRead(ImbalBank_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Imbal Jasa Bank G", RGBRead(ImbalBank_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Imbal Jasa Bank B", RGBRead(ImbalBank_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Imbal Jasa Bank Kustodian", txtImbalJasaBank.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Biaya Pembelian R", RGBRead(BiayaBeli_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Biaya Pembelian G", RGBRead(BiayaBeli_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Biaya Pembelian B", RGBRead(BiayaBeli_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Biaya Pembelian", txtBiayaPembelian.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Biaya Penjualan R", RGBRead(BiayaJual_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Biaya Penjualan G", RGBRead(BiayaJual_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Biaya Penjualan B", RGBRead(BiayaJual_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Biaya Penjualan Kembali", txtBiayaPenjualaKembali.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Investasi Awal R", RGBRead(InvestAwal_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Investasi Awal G", RGBRead(InvestAwal_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Investasi Awal B", RGBRead(InvestAwal_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Minimal Investasi Awal", txtMinimalInvestasAwal.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Investasi Selanjutnya R", RGBRead(InvestSelanjutnya_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Investasi Selanjutnya G", RGBRead(InvestSelanjutnya_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Investasi Selanjutnya B", RGBRead(InvestSelanjutnya_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Minimal Investasi Selanjutnya", txtMinimalInvestasiSelanjutnya.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Kode R", RGBRead(Kode_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kode G", RGBRead(Kode_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kode B", RGBRead(Kode_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Kode ISIN/Bloomberg", txtKodeISINBloomberg.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Kinerja Sejak Diluncurkan R", RGBRead(KinerjaAwal_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kinerja Sejak Diluncurkan G", RGBRead(KinerjaAwal_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kinerja Sejak Diluncurkan B", RGBRead(KinerjaAwal_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Kinerja Sejak Diluncurkan", txtKinerjaSejakDiluncurkan.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Kinerja Bulan Terbaik R", RGBRead(KinerjaTerbaik_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kinerja Bulan Terbaik G", RGBRead(KinerjaTerbaik_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kinerja Bulan Terbaik B", RGBRead(KinerjaTerbaik_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Kinerja Bulan Terbaik", txtKinerjaBulanTerbaik.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Kinerja Bulan Terburuk R", RGBRead(KinerjaTerburuk_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kinerja Bulan Terburuk G", RGBRead(KinerjaTerburuk_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Kinerja Bulan Terburuk B", RGBRead(KinerjaTerburuk_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Kinerja Bulan Terburuk", txtKinerjaBulanTerburuk.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Faktor Risiko Utama R", RGBRead(FaktorRisiko_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Faktor Risiko Utama G", RGBRead(FaktorRisiko_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Faktor Risiko Utama B", RGBRead(FaktorRisiko_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Faktor Risiko Utama", txtFaktorRisikoUtama.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Periode Investasi R", RGBRead(PeriodeInvestasi_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Periode Investasi G", RGBRead(PeriodeInvestasi_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Periode Investasi B", RGBRead(PeriodeInvestasi_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Periode Investasi", txtPerodeInvestasi.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Tingkat Risiko R", RGBRead(TingkatRisiko_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tingkat Risiko G", RGBRead(TIngkatRisiko_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tingkat Risiko B", RGBRead(TingkatRisiko_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Tingkat Risiko", txtTingkatRisiko.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Tujuan Investasi R", RGBRead(TujuanInvestasi_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tujuan Investasi G", RGBRead(TujuanInvestasi_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tujuan Investasi B", RGBRead(TujuanInvestasi_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Tujuan Investasi", txtTujuanInvestasi.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Investasi R", RGBRead(Investasi_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Investasi G", RGBRead(Investasi_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Investasi B", RGBRead(Investasi_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Kebijakan Investasi", txtKebijakanInvestasi.Text.Trim)
                file.WriteString(reportSection, iniType & " Saham", txtInvestasiSaham.Text.Trim)
                file.WriteString(reportSection, iniType & " Obligasi", txtInvestasiObligasi.Text.Trim)
                file.WriteString(reportSection, iniType & " Pasar Uang", txtInvesatiPasarUang.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Portofolio R", RGBRead(Portofolio_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Portofolio G", RGBRead(Portofolio_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Portofolio B", RGBRead(Portofolio_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Komposisi Portofolio", txtKomposisiPortofolio.Text.Trim)
                file.WriteString(reportSection, iniType & " Saham", txtPortofolioSaham.Text.Trim)
                file.WriteString(reportSection, iniType & " Obligasi", txtPortofoliObligasi.Text.Trim)
                file.WriteString(reportSection, iniType & " Pasar Uang", txtPortofolioPasarUang.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Line Chart Title R", RGBRead(KinerjaReksaDana_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Line Chart Title G", RGBRead(KinerjaReksaDana_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Line Chart Title B", RGBRead(KinerjaReksaDana_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Line Chart Title", txtKinerjaReksaDanaChart.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Line Chart Border R", RGBRead(ChartBorder_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Line Chart Border G", RGBRead(ChartBorder_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Line Chart Border B", RGBRead(ChartBorder_B.Text.Trim))
                file.WriteBoolean(reportSection, iniType & " Line Chart Border", chkChartBorder.Checked)

                file.WriteInteger(reportSection, iniType & " Title kepemilikan R", RGBRead(TitleKepemilikan_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Title Kepemilikan G", RGBRead(TitleKepemilikan_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Title Kepemilikan B", RGBRead(TitleKepemilikan_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Title Kepemilikan", txtTitleKepemilikan.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Filter R", RGBRead(Filter_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Filter G", RGBRead(Filter_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Filter B", RGBRead(Filter_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Filter", txtFilter.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Pie Chart Title R", RGBRead(ChartTitle_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Pie Chart Title G", RGBRead(ChartTitle_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Pie Chart Title B", RGBRead(ChartTitle_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Pie Chart Title", txtAlokasiSektor.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Pie Chart Border R", RGBRead(ChartBorderPie_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Pie Chart Border G", RGBRead(ChartBorderPie_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Pie Chart Border B", RGBRead(ChartBorderPie_B.Text.Trim))
                file.WriteBoolean(reportSection, iniType & " Pie Chart Border", chkChartBorderPie.Checked)

                file.WriteInteger(reportSection, iniType & " Table Title R", RGBRead(TableTitle_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Title G", RGBRead(TableTitle_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Title B", RGBRead(TableTitle_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Table Title", txtTableTitle.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Table Item R", RGBRead(TableItem_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Item G", RGBRead(TableItem_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Item B", RGBRead(TableItem_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Return", txtTableItemReturn.Text.Trim)
                file.WriteString(reportSection, iniType & " Benchmark", txtTableItemBenchmark.Text.Trim)
                file.WriteString(reportSection, iniType & " 1 Bulan", txtTableItem1Bulan.Text.Trim)
                file.WriteString(reportSection, iniType & " 3 Bulan", txtTableItem3Bulan.Text.Trim)
                file.WriteString(reportSection, iniType & " 6 Bulan", txtTableItem6Bulan.Text.Trim)
                file.WriteString(reportSection, iniType & " 1 Tahun", txtTableItem1Tahun.Text.Trim)
                file.WriteString(reportSection, iniType & " Dari Awal", txtTableItemDariAwalTahun.Text.Trim)
                file.WriteString(reportSection, iniType & " Sejak Pembentukan", txtTableItemSejakPembentukan.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Title Ulasan R", RGBRead(OutlookHolding_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Title Ulasan G", RGBRead(OutlookHolding_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Title Ulasan B", RGBRead(OutlookHolding_B.Text.Trim))

                file.WriteInteger(reportSection, iniType & " Title MI R", RGBRead(TentangHolding_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Title MI G", RGBRead(TentangHolding_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Title MI B", RGBRead(TentangHolding_B.Text.Trim))

            End If
            frm.pdfSetting()
            Close()
        Catch ex As Exception
            ExceptionMessage.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


End Class