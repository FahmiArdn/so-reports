Imports simpi.GlobalConnection
Imports simpi.GlobalUtilities

Public Class ReportFundSheetCPFSetting
    Public frm As ReportFundSheetCPF
    Dim reportSection As String = "Report Fund Sheet CPF"

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
            If rbHeader.Checked Then
                txtColorHeader.BackColor = cd.Color
                Header_R.Text = RGBWrite("R", cd.Color.R)
                Header_G.Text = RGBWrite("G", cd.Color.G)
                Header_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbTitle.Checked Then
                txtColorTitle.BackColor = cd.Color
                Title_R.Text = RGBWrite("R", cd.Color.R)
                Title_G.Text = RGBWrite("G", cd.Color.G)
                Title_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbHeaderLine.Checked Then
                txtHeaderLine.BackColor = cd.Color
                HeaderLine_R.Text = RGBWrite("R", cd.Color.R)
                HeaderLine_G.Text = RGBWrite("G", cd.Color.G)
                HeaderLine_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbNAV.Checked Then
                txtColorNAV.BackColor = cd.Color
                NAV_R.Text = RGBWrite("R", cd.Color.R)
                NAV_G.Text = RGBWrite("G", cd.Color.G)
                NAV_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbSummaryLine.Checked Then
                txtColorSummaryLine.BackColor = cd.Color
                SummaryLine_R.Text = RGBWrite("R", cd.Color.R)
                SummaryLine_G.Text = RGBWrite("G", cd.Color.G)
                SummaryLine_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbSummaryItems.Checked Then
                txtColorSummaryItems.BackColor = cd.Color
                SummaryItems_R.Text = RGBWrite("R", cd.Color.R)
                SummaryItems_G.Text = RGBWrite("G", cd.Color.G)
                SummaryItems_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbSubTitle.Checked Then
                txtColorSubTitle.BackColor = cd.Color
                Subtitle_R.Text = RGBWrite("R", cd.Color.R)
                SubTitle_G.Text = RGBWrite("G", cd.Color.G)
                SubTitle_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbTanggalPenting.Checked Then
                txtColorTanggalPenting.BackColor = cd.Color
                TanggalPenting_R.Text = RGBWrite("R", cd.Color.R)
                TanggalPenting_G.Text = RGBWrite("G", cd.Color.G)
                TanggalPenting_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbTanggalPentingLine.Checked Then
                txtColordateLine.BackColor = cd.Color
                TanggalPentingLine_R.Text = RGBWrite("R", cd.Color.R)
                TanggalPentingLine_G.Text = RGBWrite("G", cd.Color.G)
                TanggalPentingLine_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbTableHeader.Checked Then
                txtColorTableHeader.BackColor = cd.Color
                TableHeader_R.Text = RGBWrite("R", cd.Color.R)
                TableHeader_G.Text = RGBWrite("G", cd.Color.G)
                TableHeader_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbTableTitle.Checked Then
                txtColorTableTitle.BackColor = cd.Color
                TableTitle_R.Text = RGBWrite("R", cd.Color.R)
                TableTitle_G.Text = RGBWrite("G", cd.Color.G)
                TableTitle_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbTableLine.Checked Then
                txtColorTableLine.BackColor = cd.Color
                TableLine_R.Text = RGBWrite("R", cd.Color.R)
                TableLine_R.Text = RGBWrite("G", cd.Color.G)
                TableLine_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbFooterTitle1.Checked Then
                txtColorFooterTitle1.BackColor = cd.Color
                FooterTitle1_R.Text = RGBWrite("R", cd.Color.R)
                FooterTitle1_R.Text = RGBWrite("G", cd.Color.G)
                FooterTitle1_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbFooterTitle2.Checked Then
                txtColorFooterTitle2.BackColor = cd.Color
                FooterTitle2_R.Text = RGBWrite("R", cd.Color.R)
                FooterTitle2_R.Text = RGBWrite("G", cd.Color.G)
                FooterTitle2_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbChartTitle.Checked Then
                txtColorChartTitle.BackColor = cd.Color
                ChartTitle_R.Text = RGBWrite("R", cd.Color.R)
                ChartTitle_G.Text = RGBWrite("G", cd.Color.G)
                ChartTitle_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbChartBorder.Checked Then
                txtColorChartBorder.BackColor = cd.Color
                ChartBorder_R.Text = RGBWrite("R", cd.Color.R)
                ChartBorder_G.Text = RGBWrite("G", cd.Color.G)
                ChartBorder_B.Text = RGBWrite("B", cd.Color.B)
            End If
        End If
    End Sub
    Private Sub rbHeader_Click(sender As Object, e As EventArgs) Handles rbHeader.Click
        colorSet()
    End Sub


    Private Sub rbTitle_Click(sender As Object, e As EventArgs) Handles rbTitle.Click
        colorSet()
    End Sub

    Private Sub rbHeaderLine_Click(sender As Object, e As EventArgs) Handles rbHeaderLine.Click
        colorSet()
    End Sub

    Private Sub rbNAV_Click(sender As Object, e As EventArgs) Handles rbNAV.Click
        colorSet()
    End Sub

    Private Sub rbSummaryLine_Click(sender As Object, e As EventArgs) Handles rbSummaryLine.Click
        colorSet()
    End Sub

    Private Sub rbSummaryItems_Click(sender As Object, e As EventArgs) Handles rbSummaryItems.Click
        colorSet()
    End Sub

    Private Sub rbSubTitle_Click(sender As Object, e As EventArgs) Handles rbSubTitle.Click
        colorSet()
    End Sub

    Private Sub rbTanggalPenting_Click(sender As Object, e As EventArgs) Handles rbTanggalPenting.Click
        colorSet()
    End Sub

    Private Sub rbTanggalPentingLine_Click(sender As Object, e As EventArgs) Handles rbTanggalPentingLine.Click
        colorSet()
    End Sub

    Private Sub rbTableHeader_Click(sender As Object, e As EventArgs) Handles rbTableHeader.Click
        colorSet()
    End Sub

    Private Sub rbTableTitle_Click(sender As Object, e As EventArgs) Handles rbTableTitle.Click
        colorSet()
    End Sub
    Private Sub rbTableLine_Click(sender As Object, e As EventArgs) Handles rbTableLine.Click
        colorSet()
    End Sub

    Private Sub rbFooterTitle1_Click(sender As Object, e As EventArgs) Handles rbFooterTitle1.Click
        colorSet()
    End Sub

    Private Sub rbFooterTitle2_Click(sender As Object, e As EventArgs) Handles rbFooterTitle2.Click
        colorSet()
    End Sub

    Private Sub rbChartTitle_Click(sender As Object, e As EventArgs) Handles rbChartTitle.Click
        colorSet()
    End Sub

    Private Sub rbChartBorder_Click(sender As Object, e As EventArgs) Handles rbChartBorder.Click
        colorSet()
    End Sub
#End Region
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
                    r = file.GetInteger(reportSection, iniType & " Header R", 0)
                    g = file.GetInteger(reportSection, iniType & " Header G", 0)
                    b = file.GetInteger(reportSection, iniType & " Header B", 0)
                    Header_R.Text = RGBWrite("R", r)
                    Header_G.Text = RGBWrite("G", g)
                    Header_B.Text = RGBWrite("B", b)
                    txtColorHeader.BackColor = Color.FromArgb(r, g, b)
                    txtHeader.Text = file.GetString(reportSection, iniType & " Header ", "")

                    r = file.GetInteger(reportSection, iniType & " Title R", 0)
                    g = file.GetInteger(reportSection, iniType & " Title G", 0)
                    b = file.GetInteger(reportSection, iniType & " Title B", 0)
                    Title_R.Text = RGBWrite("R", r)
                    Title_G.Text = RGBWrite("G", g)
                    Title_B.Text = RGBWrite("B", b)
                    txtColorTitle.BackColor = Color.FromArgb(r, g, b)
                    txtTitle.Text = file.GetString(reportSection, iniType & " Title", "")

                    r = file.GetInteger(reportSection, iniType & " Header Line R", 0)
                    g = file.GetInteger(reportSection, iniType & " Header Line G", 0)
                    b = file.GetInteger(reportSection, iniType & " Header Line B", 0)
                    HeaderLine_R.Text = RGBWrite("R", r)
                    HeaderLine_G.Text = RGBWrite("G", g)
                    HeaderLine_B.Text = RGBWrite("B", b)
                    txtHeaderLine.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " NAV R", 0)
                    g = file.GetInteger(reportSection, iniType & " NAV G", 0)
                    b = file.GetInteger(reportSection, iniType & " NAV B", 0)
                    NAV_R.Text = RGBWrite("R", r)
                    NAV_G.Text = RGBWrite("G", g)
                    NAV_B.Text = RGBWrite("B", b)
                    txtColorNAV.BackColor = Color.FromArgb(r, g, b)
                    txtNAV.Text = file.GetString(reportSection, iniType & " NAV", "")

                    r = file.GetInteger(reportSection, iniType & " Summary Line R", 0)
                    g = file.GetInteger(reportSection, iniType & " Summary Line G", 0)
                    b = file.GetInteger(reportSection, iniType & " Summary Line B", 0)
                    SummaryLine_R.Text = RGBWrite("R", r)
                    SummaryLine_G.Text = RGBWrite("G", g)
                    SummaryLine_B.Text = RGBWrite("B", b)
                    txtColorSummaryLine.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " Summary Items R", 0)
                    g = file.GetInteger(reportSection, iniType & " Summary Items G", 0)
                    r = file.GetInteger(reportSection, iniType & " Summary Items B", 0)
                    SummaryItems_R.Text = RGBWrite("R", r)
                    SummaryItems_G.Text = RGBWrite("G", g)
                    SummaryItems_B.Text = RGBWrite("B", b)
                    txtColorSummaryItems.BackColor = Color.FromArgb(r, g, b)
                    txtSummaryItemsTanggalLaporan.Text = file.GetString(reportSection, iniType & " Summary Items Tanggal Laporan", "")
                    txtSummaryItemsBankKustodian.Text = file.GetString(reportSection, iniType & " Summary Items Bank Kustodian", "")
                    txtSummaryItemsTanggalPeluncuran.Text = file.GetString(reportSection, iniType & " Summary Items Tanggal Peluncuran", "")
                    txtSummaryItemstanggalJatuhTempo.Text = file.GetString(reportSection, iniType & " Summary Items Tanggal Jatuh Tempo", "")
                    txtSummaryItemsTotalAUM.Text = file.GetString(reportSection, iniType & " Summary Items Total AUM", "")
                    txtSummaryItemsMataUang.Text = file.GetString(reportSection, iniType & " Summary Items Mata Uang", "")
                    txtSummaryItemsImbalJasaManajerInvestasi.Text = file.GetString(reportSection, iniType & " Summary Items Imbal Jasa Manajer Investasi", "")
                    txtSummaryItemsImbalJasaBankKustodian.Text = file.GetString(reportSection, iniType & " Summary Items Imbal Jasa Bank Kustodian", "")
                    txtSummaryItemsBiayaPembelian.Text = file.GetString(reportSection, iniType & " Summary Items Biaya Pembelian", "")
                    txtSummaryItemsBiayaPenjualanKembali.Text = file.GetString(reportSection, iniType & " Summary Items Biaya Penjualan Kembali", "")
                    txtSummaryItemsBiayaPengalihan.Text = file.GetString(reportSection, iniType & " Summary Items Biaya Pengalihan", "")
                    txtSummaryItemsNote.Text = file.GetString(reportSection, iniType & " Summary Items Note", "")

                    r = file.GetInteger(reportSection, iniType & " Sub Title R", 0)
                    b = file.GetInteger(reportSection, iniType & " Sub Title G", 0)
                    g = file.GetInteger(reportSection, iniType & " Sub Title B", 0)
                    Subtitle_R.Text = RGBWrite("R", r)
                    SubTitle_G.Text = RGBWrite("B", b)
                    SubTitle_B.Text = RGBWrite("G", g)
                    txtColorSubTitle.BackColor = Color.FromArgb(r, b, g)
                    txtSubTitle.Text = file.GetString(reportSection, iniType & " Sub Title", "")

                    r = file.GetInteger(reportSection, iniType & " Tanggal Penting R", 0)
                    g = file.GetInteger(reportSection, iniType & " Tanggal Penting G", 0)
                    b = file.GetInteger(reportSection, iniType & " Tanggal Penting B", 0)
                    TanggalPenting_R.Text = RGBWrite("R", r)
                    TanggalPenting_G.Text = RGBWrite("G", g)
                    TanggalPenting_B.Text = RGBWrite("B", b)
                    txtColorTanggalPenting.BackColor = Color.FromArgb(r, g, b)
                    txtTanggalPenting.Text = file.GetString(reportSection, iniType & " Tanggal Penting", "")
                    txtPembagianDividendTerakhir.Text = file.GetString(reportSection, iniType & " Pembagian Dividend Terakhir", "")
                    txtPembagianDividendBerikutnya.Text = file.GetString(reportSection, iniType & " Pembagian Dividend Berikutnya", "")
                    txtUnderlyingAsset.Text = file.GetString(reportSection, iniType & " Underlying Asset", "")

                    r = file.GetInteger(reportSection, iniType & " Tanggal Penting Line R", 0)
                    g = file.GetInteger(reportSection, iniType & " Tanggal Penting Line G", 0)
                    b = file.GetInteger(reportSection, iniType & " Tanggal Penting Line B", 0)
                    TanggalPentingLine_R.Text = RGBWrite("R", r)
                    TanggalPentingLine_G.Text = RGBWrite("G", g)
                    TanggalPentingLine_B.Text = RGBWrite("B", b)
                    txtColordateLine.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " Table Header R", 0)
                    g = file.GetInteger(reportSection, iniType & " Table Header G", 0)
                    b = file.GetInteger(reportSection, iniType & " Table Header B", 0)
                    TableHeader_R.Text = RGBWrite("R", r)
                    TableHeader_G.Text = RGBWrite("G", g)
                    TableHeader_B.Text = RGBWrite("B", b)
                    txtColorTableHeader.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " Table Title R", 0)
                    g = file.GetInteger(reportSection, iniType & " Table Title G", 0)
                    b = file.GetInteger(reportSection, iniType & " Table Title B", 0)
                    TableTitle_R.Text = RGBWrite("R", r)
                    TableTitle_G.Text = RGBWrite("G", g)
                    TableTitle_B.Text = RGBWrite("B", b)
                    txtColorTableTitle.BackColor = Color.FromArgb(r, g, b)
                    txtTable1Bln.Text = file.GetString(reportSection, iniType & " Table Title 1 Bln", "")
                    txtTable3Bln.Text = file.GetString(reportSection, iniType & " Table Title 3 Bln", "")
                    txtTable6Bln.Text = file.GetString(reportSection, iniType & " Table Title 6 Bln", "")
                    txtTable1Thn.Text = file.GetString(reportSection, iniType & " Table Title 1 Tahun", "")
                    txtTableDariAwalTahun.Text = file.GetString(reportSection, iniType & " Table Title Dari Awal Tahun", "")
                    txtTableSejakPembentukan.Text = file.GetString(reportSection, iniType & " Table Title Sejak Pembukuan", "")
                    txtTableIndikasiRateDividenTetap.Text = file.GetString(reportSection, iniType & " Table Title Indikasi Rate Dividen Tetap", "")

                    r = file.GetInteger(reportSection, iniType & " Table Line R", 0)
                    g = file.GetInteger(reportSection, iniType & " Table Line G", 0)
                    b = file.GetInteger(reportSection, iniType & " Table Line B", 0)
                    TableLine_R.Text = RGBWrite("R", r)
                    TableLine_G.Text = RGBWrite("G", g)
                    TableLine_B.Text = RGBWrite("B", b)
                    txtColorTableLine.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " Footer Title 1 R", 0)
                    g = file.GetInteger(reportSection, iniType & " Footer Title 1 R", 0)
                    b = file.GetInteger(reportSection, iniType & " Footer Title 1 R", 0)
                    FooterTitle1_R.Text = RGBWrite("R", r)
                    FooterTitle1_G.Text = RGBWrite("G", g)
                    FooterTitle1_B.Text = RGBWrite("B", b)
                    txtColorFooterTitle1.BackColor = Color.FromArgb(r, g, b)
                    txtFooterTitle1.Text = file.GetString(reportSection, iniType & " Footer Title 1", "")

                    r = file.GetInteger(reportSection, iniType & " Footer Title 2 R", 0)
                    g = file.GetInteger(reportSection, iniType & " Footer Title 2 R", 0)
                    b = file.GetInteger(reportSection, iniType & " Footer Title 2 R", 0)
                    FooterTitle2_R.Text = RGBWrite("R", r)
                    FooterTitle2_G.Text = RGBWrite("G", g)
                    FooterTitle2_B.Text = RGBWrite("B", b)
                    txtColorFooterTitle2.BackColor = Color.FromArgb(r, g, b)
                    txtFooterTitle2.Text = file.GetString(reportSection, iniType & " Footer Title 2", "")

                    r = file.GetInteger(reportSection, iniType & " Chart Title R", 0)
                    g = file.GetInteger(reportSection, iniType & " Chart Title G", 0)
                    b = file.GetInteger(reportSection, iniType & " Chart Title B", 0)
                    ChartTitle_R.Text = RGBWrite("R", r)
                    ChartTitle_G.Text = RGBWrite("G", g)
                    ChartTitle_B.Text = RGBWrite("B", b)
                    txtColorChartTitle.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " Line Chart Border R", 0)
                    g = file.GetInteger(reportSection, iniType & " Line Chart Border G", 0)
                    b = file.GetInteger(reportSection, iniType & " Line Chart Border B", 0)
                    ChartBorder_R.Text = RGBWrite("R", r)
                    ChartBorder_G.Text = RGBWrite("G", g)
                    ChartBorder_B.Text = RGBWrite("B", b)
                    txtColorChartBorder.BackColor = Color.FromArgb(r, g, b)
                    If file.GetBoolean(reportSection, iniType & " Line Chart Border", False) Then chkChartBorder.Checked = True Else chkChartBorder.Checked = False

                End If
            End If
        Catch ex As Exception
            _default()
        End Try
    End Sub
    Private Sub _default()
        With frm.pdfLayout
            frm.pdfColorDefault()
            Header_R.Text = "R: " & .Header_R
            Header_G.Text = "G: " & .Header_G
            Header_B.Text = "B: " & .Header_B
            txtColorHeader.BackColor = Color.FromArgb(.Header_R, .Header_G, .Header_B)
            txtHeader.Text = .Header

            Title_R.Text = "R: " & .Title_R
            Title_G.Text = "G: " & .Title_G
            Title_B.Text = "B: " & .Title_B
            txtColorTitle.BackColor = Color.FromArgb(.Title_R, .Title_G, .Title_B)
            txtTitle.Text = .Title

            ChartTitle_R.Text = "R: " & .ChartTitle_R
            ChartTitle_G.Text = "G: " & .ChartTitle_G
            ChartTitle_B.Text = "B: " & .ChartTitle_B
            txtColorChartTitle.BackColor = Color.FromArgb(.ChartTitle_R, .ChartTitle_G, .ChartTitle_B)
            txtChartTitle.Text = .ChartTitle

            HeaderLine_R.Text = "R: " & .HeaderLine_R
            HeaderLine_G.Text = "G: " & .HeaderLine_G
            HeaderLine_B.Text = "B: " & .HeaderLine_B
            txtHeaderLine.BackColor = Color.FromArgb(.HeaderLine_R, .HeaderLine_G, .HeaderLine_B)

            NAV_R.Text = "R: " & .NAV_R
            NAV_G.Text = "G: " & .NAV_G
            NAV_B.Text = "B: " & .NAV_B
            txtColorNAV.BackColor = Color.FromArgb(.NAV_R, .NAV_G, .NAV_B)
            txtNAV.Text = .NAV

            FooterTitle1_R.Text = "R: " & .FooterTitle1_R
            FooterTitle1_B.Text = "G: " & .FooterTitle1_G
            FooterTitle1_G.Text = "B: " & .FooterTitle1_B
            txtColorFooterTitle1.BackColor = Color.FromArgb(.FooterTitle1_R, .FooterTitle1_G, .FooterTitle1_B)
            txtFooterTitle1.Text = .Footer1

            FooterTitle2_R.Text = "R: " & .FooterTitle2_R
            FooterTitle2_B.Text = "G: " & .FooterTitle2_G
            FooterTitle2_G.Text = "B: " & .FooterTitle2_B
            txtColorFooterTitle2.BackColor = Color.FromArgb(.FooterTitle2_R, .FooterTitle2_G, .FooterTitle2_B)
            txtFooterTitle2.Text = .Footer2

            TableLine_R.Text = "R: " & .TableLine_R
            TableLine_G.Text = "G: " & .TableLine_G
            TableLine_B.Text = "B: " & .TableLine_B
            txtColorTableLine.BackColor = Color.FromArgb(.TableLine_R, .TableLine_G, .TableLine_B)

            SummaryLine_R.Text = "R: " & .SummaryLine_R
            SummaryLine_G.Text = "G: " & .SummaryLine_G
            SummaryLine_B.Text = "B: " & .SummaryLine_B
            txtColorSummaryLine.BackColor = Color.FromArgb(.SummaryLine_R, .SummaryLine_G, .SummaryLine_B)

            SummaryItems_R.Text = "R: " & .SummaryItems_R
            SummaryItems_G.Text = "G: " & .SummaryItems_G
            SummaryItems_B.Text = "B: " & .SummaryItems_B
            txtColorSummaryItems.BackColor = Color.FromArgb(.SummaryItems_R, .SummaryItems_G, .SummaryItems_B)
            txtSummaryItemsTanggalLaporan.Text = .SummaryItemsTanggalLaporan
            txtSummaryItemsBankKustodian.Text = .SummaryItemsBankKustodian
            txtSummaryItemsTanggalPeluncuran.Text = .SummaryItemsTanggalPeluncuran
            txtSummaryItemstanggalJatuhTempo.Text = .SummaryItemstanggalJatuhTempo
            txtSummaryItemsTotalAUM.Text = .SummaryItemsTotalAUM
            txtSummaryItemsMataUang.Text = .SummaryItemsMataUang
            txtSummaryItemsImbalJasaManajerInvestasi.Text = .SummaryItemsImbalJasaManajerInvestasi
            txtSummaryItemsImbalJasaBankKustodian.Text = .SummaryItemsImbalJasaBankKustodian
            txtSummaryItemsBiayaPembelian.Text = .SummaryItemsBiayaPembelian
            txtSummaryItemsBiayaPenjualanKembali.Text = .SummaryItemsBiayaPenjualanKembali
            txtSummaryItemsBiayaPengalihan.Text = .SummaryItemsBiayaPengalihan
            txtSummaryItemsNote.Text = .SummaryItemsNote

            Subtitle_R.Text = "R: " & .Subtitle_R
            SubTitle_G.Text = "G: " & .Subtitle_G
            SubTitle_B.Text = "B: " & .Subtitle_B
            txtColorSubTitle.BackColor = Color.FromArgb(.Subtitle_R, .Subtitle_G, .Subtitle_B)
            txtSubTitle.Text = .Subtitle

            TanggalPenting_R.Text = "R: " & .TanggalPenting_R
            TanggalPenting_G.Text = "G: " & .TanggalPenting_G
            TanggalPenting_B.Text = "B: " & .TanggalPenting_B
            txtColorTanggalPenting.BackColor = Color.FromArgb(.TanggalPenting_R, .TanggalPenting_G, .TanggalPenting_B)
            txtTanggalPenting.Text = .TanggalPenting
            txtPembagianDividendTerakhir.Text = .PembagianDividendTerakhir
            txtPembagianDividendBerikutnya.Text = .PembagianDividendBerikutnya
            txtUnderlyingAsset.Text = .UnderlyingAsset

            TanggalPentingLine_R.Text = "R: " & .TanggalPentingLine_R
            TanggalPentingLine_G.Text = "G: " & .TanggalPentingLine_G
            TanggalPentingLine_B.Text = "B: " & .TanggalPentingLine_B
            txtColordateLine.BackColor = Color.FromArgb(.TanggalPentingLine_R, .TanggalPentingLine_G, .TanggalPentingLine_B)

            TableHeader_R.Text = "R: " & .TableHeader_R
            TableHeader_G.Text = "G: " & .TableHeader_G
            TableHeader_B.Text = "B: " & .TableHeader_B
            txtColorTableHeader.BackColor = Color.FromArgb(.TableHeader_R, .TableHeader_G, .TableHeader_B)
            txtTableHeader.Text = .TableHeader

            TableTitle_R.Text = "R: " & .TableTitle_R
            TableTitle_G.Text = "G: " & .TableTitle_G
            TableTitle_B.Text = "B: " & .TableTitle_B
            txtColorTableTitle.BackColor = Color.FromArgb(.TableTitle_R, .TableTitle_G, .TableTitle_B)
            txtTable1Bln.Text = .Table1Bln
            txtTable3Bln.Text = .Table3Bln
            txtTable6Bln.Text = .Table6Bln
            txtTable1Thn.Text = .Table1Thn
            txtTableDariAwalTahun.Text = .TableDariAwalTahun
            txtTableSejakPembentukan.Text = .TableSejakPembentukan
            txtTableIndikasiRateDividenTetap.Text = .TableIndikasiRateDividenTetap

            ChartBorder_R.Text = "R: " & .ChartBorder_R
            ChartBorder_G.Text = "G: " & .ChartBorder_G
            ChartBorder_B.Text = "B: " & .ChartBorder_B
            txtColorChartBorder.BackColor = Color.FromArgb(.ChartBorder_R, .ChartBorder_G, .ChartBorder_B)
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
                file.WriteInteger(reportSection, iniType & " Header R", RGBRead(Header_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Header G", RGBRead(Header_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Header B", RGBRead(Header_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Header", txtHeader.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Title R", RGBRead(Title_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Title G", RGBRead(Title_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Title B", RGBRead(Title_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Title", txtTitle.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Chart Title R", RGBRead(ChartTitle_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Title  G", RGBRead(ChartTitle_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Title  B", RGBRead(ChartTitle_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Chart Title ", txtChartTitle.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Header Line R", RGBRead(HeaderLine_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Header Line G", RGBRead(HeaderLine_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Header Line B", RGBRead(HeaderLine_B.Text.Trim))

                file.WriteInteger(reportSection, iniType & " Footer Title 1 R", RGBRead(FooterTitle1_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Footer Title 1 G", RGBRead(FooterTitle1_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Footer Title 1 B", RGBRead(FooterTitle1_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Footer Title 1", txtFooterTitle1.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Footer Title 2 R", RGBRead(FooterTitle2_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Footer Title 2 G", RGBRead(FooterTitle2_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Footer Title 2 B", RGBRead(FooterTitle2_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Footer Title 2", txtFooterTitle2.Text.Trim)

                file.WriteInteger(reportSection, iniType & " NAV R", RGBRead(NAV_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " NAV G", RGBRead(NAV_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " NAV B", RGBRead(NAV_B.Text.Trim))
                file.WriteString(reportSection, iniType & " NAV", txtNAV.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Summary Line R", RGBRead(SummaryLine_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Summary Line G", RGBRead(SummaryLine_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Summary Line B", RGBRead(SummaryLine_B.Text.Trim))

                file.WriteInteger(reportSection, iniType & " Summary Items R", RGBRead(SummaryItems_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Summary Items G", RGBRead(SummaryItems_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Summary Items B", RGBRead(SummaryItems_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Summary Items Tanggal Laporan", txtSummaryItemsTanggalLaporan.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Bank Kustodian", txtSummaryItemsBankKustodian.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Tanggal Peluncuran ", txtSummaryItemsTanggalPeluncuran.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Tanggal Jatuh Tempo", txtSummaryItemstanggalJatuhTempo.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Total AUM", txtSummaryItemsTotalAUM.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Mata Uang", txtSummaryItemsMataUang.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Imbal Jasa Manajer Investasi", txtSummaryItemsImbalJasaManajerInvestasi.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Imbal Jasa Bank Kustodian", txtSummaryItemsImbalJasaBankKustodian.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Biaya Pembelian", txtSummaryItemsBiayaPembelian.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Biaya Penjualan Kembali", txtSummaryItemsBiayaPenjualanKembali.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Biaya Pengalihan", txtSummaryItemsBiayaPengalihan.Text.Trim)
                file.WriteString(reportSection, iniType & " Summary Items Note", txtSummaryItemsNote.Text.Trim)
                ''

                file.WriteInteger(reportSection, iniType & " Sub Title R", RGBRead(Subtitle_R.Text.Trim()))
                file.WriteInteger(reportSection, iniType & " Sub Title G", RGBRead(SubTitle_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Sub Title B", RGBRead(SubTitle_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Sub Title", txtSubTitle.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Tanggal Penting R", RGBRead(TanggalPenting_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tanggal Penting G", RGBRead(TanggalPenting_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tanggal Penting B", RGBRead(TanggalPenting_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Tanggal Penting", txtTanggalPenting.Text.Trim)
                file.WriteString(reportSection, iniType & " Pembagian Dividend Terakhir", txtPembagianDividendTerakhir.Text.Trim)
                file.WriteString(reportSection, iniType & " Pembagian Dividend Berikutnya", txtPembagianDividendBerikutnya.Text.Trim)
                file.WriteString(reportSection, iniType & " Underlying Asset", txtUnderlyingAsset.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Tanggal Penting Line R", RGBRead(TanggalPentingLine_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tanggal Penting Line G", RGBRead(TanggalPentingLine_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tanggal Penting Line B", RGBRead(TanggalPentingLine_B.Text.Trim))

                file.WriteInteger(reportSection, iniType & " Table header R", RGBRead(TableHeader_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table header G", RGBRead(TableHeader_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table header B", RGBRead(TableHeader_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Table Header", txtTableHeader.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Table Line R", RGBRead(TableLine_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Line G", RGBRead(TableLine_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Line B", RGBRead(TableLine_B.Text.Trim))

                file.WriteInteger(reportSection, iniType & " Table Title R", RGBRead(TableTitle_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Title G", RGBRead(TableTitle_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Title B", RGBRead(TableTitle_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Table Title 1 Bln", txtTable1Bln.Text.Trim)
                file.WriteString(reportSection, iniType & " Table Title 3 Bln", txtTable3Bln.Text.Trim)
                file.WriteString(reportSection, iniType & " Table Title 6 Bln", txtTable6Bln.Text.Trim)
                file.WriteString(reportSection, iniType & " Table Title 1 Tahun", txtTable1Thn.Text.Trim)
                file.WriteString(reportSection, iniType & " Table Title Dari Awal Tahun", txtTableDariAwalTahun.Text.Trim)
                file.WriteString(reportSection, iniType & " Table Title Sejak Pembukuan", txtTableSejakPembentukan.Text.Trim)
                file.WriteString(reportSection, iniType & " Table Title Indikasi Rate Dividen Tetap", txtTableIndikasiRateDividenTetap.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Line Chart Border R", RGBRead(ChartBorder_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Line Chart Border G", RGBRead(ChartBorder_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Line Chart Border B", RGBRead(ChartBorder_B.Text.Trim))
                file.WriteBoolean(reportSection, iniType & " Line Chart Border", chkChartBorder.Checked)
            End If
            frm.pdfSetting()
            Close()
        Catch ex As Exception
            ExceptionMessage.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Close()
    End Sub

    Private Sub rbDefault_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rbDefault.SelectedIndexChanged
        iniCheck()
    End Sub

    Private Sub rbOption1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rbOption1.SelectedIndexChanged
        iniCheck()
    End Sub

    Private Sub rbOption2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rbOption2.SelectedIndexChanged
        iniCheck()
    End Sub
End Class