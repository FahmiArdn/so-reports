Imports simpi.GlobalUtilities
Imports simpi.GlobalConnection
Imports simpi.CoreData
Imports simpi.MasterPortfolio
Imports System.Drawing.Imaging
Imports C1.Win.C1Chart

Public Class ReportFundSheetDividend
    Dim objPortfolio As New MasterPortfolio
    Dim objNAV As New PortfolioNAV
    Dim objReturn As New PortfolioReturn
    Dim objBenchmark As New simpi.CoreData.PortfolioBenchmark
    Dim objSecurities As New PositionSecurities

    Dim dtReturn As New DataTable
    Dim dtDividend As New DataTable
    Dim reportSection As String = "REPORT FUND SHEET DIVIDEND"
    Public pdfLayout As New pdfColor
#Region "pdf"
    Structure pdfColor
        Public layoutType As String
        'Left Column
        Public ReportTitle_R As Integer
        Public ReportTitle_G As Integer
        Public ReportTitle_B As Integer
        Public ReportTitle As String
        Public TujuanInvestasi_R As Integer
        Public TujuanInvestasi_G As Integer
        Public TujuanInvestasi_B As Integer
        Public TujuanInvestasi As String
        Public InformasiReksaDana_R As Integer
        Public InformasiReksaDana_G As Integer
        Public InformasiReksaDana_B As Integer
        Public InformasiReksaDana As String
        Public JenisReksa As String
        Public TanggalPeluncuran As String
        Public DanaKelolaan As String
        Public MataUang As String
        Public FrekuensiValuasi As String
        Public BankKustodian As String
        Public TolakUkur As String
        Public NabUnit As String
        Public InvestasiDanBiaya_R As Integer
        Public InvestasiDanBiaya_G As Integer
        Public InvestasiDanBiaya_B As Integer
        Public InvestasiDanBiaya As String
        Public MinInvestasiawal As String
        Public MinInvestasiSelanjutnya As String
        Public BiayaPembelian As String
        Public BiayaPenjualan As String
        Public BiayaPengalihan As String
        Public BiayaJasaPengelola As String
        Public BiayaJasaBank As String
        Public StatistikReksadana_R As Integer
        Public StatistikReksadana_G As Integer
        Public StatistikReksadana_B As Integer
        Public StatistikReksadana As String
        Public KinerjaSejakDiluncurkan As String
        Public StandarDeviasiDisetahunkan As String
        Public KinerjaBulanTerbaik As String
        Public KinerjaBulanTerburuk As String
        Public KinerjaTerbaikSetahunTerakhir As String
        Public RisikoInvestasi_R As Integer
        Public RisikoInvestasi_G As Integer
        Public RisikoInvestasi_B As Integer
        Public RisikoInvestasi As String
        Public KlasifikasiRisiko_R As Integer
        Public KlasifikasiRisiko_G As Integer
        Public KlasifikasiRisiko_B As Integer
        Public KlasifikasiRisiko As String
        Public ChartTitle_R As Integer
        Public ChartTitle_G As Integer
        Public ChartTitle_B As Integer
        Public ChartTitle As String
        Public KinerjaKumulatif_R As Integer
        Public KinerjaKumulatif_G As Integer
        Public KinerjaKumulatif_B As Integer
        Public KinerjaKumulatif As String
        Public KebijakanInvestasi_R As Integer
        Public KebijakanInvestasi_G As Integer
        Public KebijakanInvestasi_B As Integer
        Public KebijakanInvestasi As String
        Public EfekPortfolio_R As Integer
        Public EfekPortfolio_G As Integer
        Public EfekPortfolio_B As Integer
        Public EfekPortfolio As String
        Public InformasiDividend_R As Integer
        Public InformasiDividend_G As Integer
        Public InformasiDividend_B As Integer
        Public InformasiDividend As String
        Public ReportLine_R As Integer
        Public ReportLine_G As Integer
        Public ReportLine_B As Integer
        Public ItemWarna_R As Integer
        Public ItemWarna_G As Integer
        Public ItemWarna_B As Integer
        Public ChartBorder_R As Integer
        Public ChartBorder_G As Integer
        Public ChartBorder_B As Integer
        Public ChartBorder As Boolean
    End Structure

    Public Sub pdfColorDefault()
        pdfLayout.layoutType = "DEFAULT"
        pdfLayout.ReportTitle_R = 128
        pdfLayout.ReportTitle_G = 0
        pdfLayout.ReportTitle_B = 121
        pdfLayout.ReportTitle = "FUND FACT SHEET"
        pdfLayout.TujuanInvestasi_R = 128
        pdfLayout.TujuanInvestasi_G = 0
        pdfLayout.TujuanInvestasi_B = 121
        pdfLayout.TujuanInvestasi = "Tujuan Investasi"
        pdfLayout.InformasiReksaDana_R = 128
        pdfLayout.InformasiReksaDana_G = 0
        pdfLayout.InformasiReksaDana_B = 121
        pdfLayout.InformasiReksaDana = "Informasi Reksa Dana"
        pdfLayout.JenisReksa = "Jenis Reksa Dana"
        pdfLayout.TanggalPeluncuran = "Tanggal Peluncuran"
        pdfLayout.DanaKelolaan = "Dana Kelolaan (Rp Mil)"
        pdfLayout.MataUang = "Mata Uang"
        pdfLayout.FrekuensiValuasi = "Frekuensi Valuasi"
        pdfLayout.BankKustodian = "Bank Kustodian"
        pdfLayout.TolakUkur = "Tolak Ukur"
        pdfLayout.NabUnit = "Nab/Unit (Rp/Unit)"
        pdfLayout.InvestasiDanBiaya_R = 128
        pdfLayout.InvestasiDanBiaya_G = 0
        pdfLayout.InvestasiDanBiaya_B = 121
        pdfLayout.InvestasiDanBiaya = "Investasi dan Biaya-Biaya"
        pdfLayout.MinInvestasiawal = "Minimal Investasi Awal (Rp)"
        pdfLayout.MinInvestasiSelanjutnya = "Minimal Investasi Selanjutnya (Rp)"
        pdfLayout.BiayaPembelian = "Biaya Pembelian (%)"
        pdfLayout.BiayaPenjualan = "Biaya Penjualan (%)"
        pdfLayout.BiayaPengalihan = "Biaya Pengalihan (%)"
        pdfLayout.BiayaJasaPengelola = "Biaya Jasa Pengelolaan MI (%)"
        pdfLayout.BiayaJasaBank = "Biaya Jasa Bank Kustodian (%)"
        pdfLayout.StatistikReksadana_R = 128
        pdfLayout.StatistikReksadana_G = 0
        pdfLayout.StatistikReksadana_B = 121
        pdfLayout.StatistikReksadana = "Statistik Reksadana"
        pdfLayout.RisikoInvestasi_R = 128
        pdfLayout.RisikoInvestasi_G = 0
        pdfLayout.RisikoInvestasi_B = 121
        pdfLayout.RisikoInvestasi = "Risiko Investasi"
        pdfLayout.KlasifikasiRisiko_R = 128
        pdfLayout.KlasifikasiRisiko_G = 0
        pdfLayout.KlasifikasiRisiko_B = 121
        pdfLayout.KlasifikasiRisiko = "Klasifikasi Risiko"
        pdfLayout.ChartTitle_R = 128
        pdfLayout.ChartTitle_G = 0
        pdfLayout.ChartTitle_B = 121
        pdfLayout.ChartTitle = "Grafik Kinerja Reksa Dana Satu Tahun Terakhir"
        pdfLayout.KinerjaKumulatif_R = 128
        pdfLayout.KinerjaKumulatif_G = 0
        pdfLayout.KinerjaKumulatif_B = 121
        pdfLayout.KinerjaKumulatif = "Kinerja Kumulatif (%)"
        pdfLayout.KebijakanInvestasi_R = 128
        pdfLayout.KebijakanInvestasi_G = 0
        pdfLayout.KebijakanInvestasi_B = 121
        pdfLayout.KebijakanInvestasi = "Kebijakan Investasi"
        pdfLayout.EfekPortfolio_R = 128
        pdfLayout.EfekPortfolio_G = 0
        pdfLayout.EfekPortfolio_B = 121
        pdfLayout.EfekPortfolio = "5 Besar Efek Dalam Portfolio"
        pdfLayout.InformasiDividend_R = 128
        pdfLayout.InformasiDividend_G = 0
        pdfLayout.InformasiDividend_B = 121
        pdfLayout.InformasiDividend = "INFORMASI DIVIDEND"
        pdfLayout.ReportLine_R = 0
        pdfLayout.ReportLine_G = 0
        pdfLayout.ReportLine_B = 0
        pdfLayout.ItemWarna_R = 0
        pdfLayout.ItemWarna_G = 0
        pdfLayout.ItemWarna_B = 0
        pdfLayout.ChartBorder_R = 0
        pdfLayout.ChartBorder_G = 0
        pdfLayout.ChartBorder_B = 0
        pdfLayout.ChartBorder = False
    End Sub

    Public Sub pdfSetting()
        Try
            Dim strFile As String = simpiFile("simpi.ini")
            If GlobalFileWindows.FileExists(strFile) Then
                Dim file As New GlobalINI(strFile)
                Dim iniType As String = file.GetString(reportSection, "LAYOUT", "")
                If iniType = "" Or iniType = "DEFAULT" Then
                    pdfColorDefault()
                Else
                    pdfLayout.layoutType = iniType
                    pdfLayout.ReportTitle_R = file.GetInteger(reportSection, iniType & " Report Title R", 0)
                    pdfLayout.ReportTitle_G = file.GetInteger(reportSection, iniType & " Report Title G", 0)
                    pdfLayout.ReportTitle_B = file.GetInteger(reportSection, iniType & " Report Title B", 0)
                    pdfLayout.ReportTitle = file.GetString(reportSection, iniType & " Report Title", "")

                    pdfLayout.TujuanInvestasi_R = file.GetInteger(reportSection, iniType & " Tujuan Investasi R", 0)
                    pdfLayout.TujuanInvestasi_G = file.GetInteger(reportSection, iniType & " Tujuan Investasi G", 0)
                    pdfLayout.TujuanInvestasi_B = file.GetInteger(reportSection, iniType & " Tujuan Investasi B", 0)
                    pdfLayout.TujuanInvestasi = file.GetString(reportSection, iniType & " Tujuan Investasi", "")

                    pdfLayout.InformasiReksaDana_R = file.GetInteger(reportSection, iniType & " Informasi Reksa Dana R", 0)
                    pdfLayout.InformasiReksaDana_G = file.GetInteger(reportSection, iniType & " Informasi Reksa Dana G", 0)
                    pdfLayout.InformasiReksaDana_B = file.GetInteger(reportSection, iniType & " Informasi Reksa Dana B", 0)
                    pdfLayout.InformasiReksaDana = file.GetString(reportSection, iniType & " Informasi Reksa Dana", "")
                    pdfLayout.JenisReksa = file.GetString(reportSection, iniType & " Jenis Reksa", "")
                    pdfLayout.TanggalPeluncuran = file.GetString(reportSection, iniType & " Tanggal Peluncuran", "")
                    pdfLayout.DanaKelolaan = file.GetString(reportSection, iniType & " Dana Kelolaan", "")
                    pdfLayout.MataUang = file.GetString(reportSection, iniType & " Mata Uang", "")
                    pdfLayout.FrekuensiValuasi = file.GetString(reportSection, iniType & " Frekuensi Valuasi", "")
                    pdfLayout.BankKustodian = file.GetString(reportSection, iniType & " Bank Kustodian", "")
                    pdfLayout.TolakUkur = file.GetString(reportSection, iniType & " Tolak Ukur", "")
                    pdfLayout.NabUnit = file.GetString(reportSection, iniType & " Nab Unit", "")

                    pdfLayout.InvestasiDanBiaya_R = file.GetInteger(reportSection, iniType & " Investasi dan Biaya R", 0)
                    pdfLayout.InvestasiDanBiaya_G = file.GetInteger(reportSection, iniType & " Investasi dan Biaya G", 0)
                    pdfLayout.InvestasiDanBiaya_B = file.GetInteger(reportSection, iniType & " Investasi dan Biaya B", 0)
                    pdfLayout.InvestasiDanBiaya = file.GetString(reportSection, iniType & " Investasi dan Biaya", "")
                    pdfLayout.MinInvestasiawal = file.GetString(reportSection, iniType & " Minimal Investasi Awal (Rp)", "")
                    pdfLayout.MinInvestasiSelanjutnya = file.GetString(reportSection, iniType & " Minimal Investasi Selanjutnya (Rp)", "")
                    pdfLayout.BiayaPembelian = file.GetString(reportSection, iniType & " Biaya Pembelian (%)", "")
                    pdfLayout.BiayaPenjualan = file.GetString(reportSection, iniType & " Biaya Penjualan (%)", "")
                    pdfLayout.BiayaPengalihan = file.GetString(reportSection, iniType & " Biaya Pengalihan (%)", "")
                    pdfLayout.BiayaJasaPengelola = file.GetString(reportSection, iniType & " Biaya Jasa Pengelolaan MI (%)", "")
                    pdfLayout.BiayaJasaBank = file.GetString(reportSection, iniType & " Biaya Jasa Bank Kustodian (%)", "")

                    pdfLayout.StatistikReksadana_R = file.GetInteger(reportSection, iniType & " Statistik Reksadana R", 0)
                    pdfLayout.StatistikReksadana_G = file.GetInteger(reportSection, iniType & " Statistik Reksadana G", 0)
                    pdfLayout.StatistikReksadana_B = file.GetInteger(reportSection, iniType & " Statistik Reksadana B", 0)
                    pdfLayout.StatistikReksadana = file.GetString(reportSection, iniType & " Statistik Reksadana", "")

                    pdfLayout.RisikoInvestasi_R = file.GetInteger(reportSection, iniType & " Risiko Investasi R", 0)
                    pdfLayout.RisikoInvestasi_G = file.GetInteger(reportSection, iniType & " Risiko Investasi G", 0)
                    pdfLayout.RisikoInvestasi_B = file.GetInteger(reportSection, iniType & " Risiko Investasi B", 0)
                    pdfLayout.RisikoInvestasi = file.GetString(reportSection, iniType & " Risiko Investasi ", "")

                    pdfLayout.KlasifikasiRisiko_R = file.GetInteger(reportSection, iniType & " Klasifikasi Risiko R", 0)
                    pdfLayout.KlasifikasiRisiko_G = file.GetInteger(reportSection, iniType & " Klasifikasi Risiko G", 0)
                    pdfLayout.KlasifikasiRisiko_B = file.GetInteger(reportSection, iniType & " Klasifikasi Risiko B", 0)
                    pdfLayout.KlasifikasiRisiko = file.GetString(reportSection, iniType & " Klasifikasi Risiko", "")

                    pdfLayout.ChartTitle_R = file.GetInteger(reportSection, iniType & " Chart Title R", 0)
                    pdfLayout.ChartTitle_G = file.GetInteger(reportSection, iniType & " Chart Title G", 0)
                    pdfLayout.ChartTitle_B = file.GetInteger(reportSection, iniType & " Chart Title B", 0)
                    pdfLayout.ChartTitle = file.GetString(reportSection, iniType & " Chart Title", "")

                    pdfLayout.KinerjaKumulatif_R = file.GetInteger(reportSection, iniType & " Kinerja Kumulatif R", 0)
                    pdfLayout.KinerjaKumulatif_G = file.GetInteger(reportSection, iniType & " Kinerja Kumulatif G", 0)
                    pdfLayout.KinerjaKumulatif_B = file.GetInteger(reportSection, iniType & " Kinerja Kumulatif B", 0)
                    pdfLayout.KinerjaKumulatif = file.GetString(reportSection, iniType & " Kinerja Kumulatif", "")

                    pdfLayout.KebijakanInvestasi_R = file.GetInteger(reportSection, iniType & " Kebijakan Investasi R", 0)
                    pdfLayout.KebijakanInvestasi_G = file.GetInteger(reportSection, iniType & " Kebijakan Investasi G", 0)
                    pdfLayout.KebijakanInvestasi_B = file.GetInteger(reportSection, iniType & " Kebijakan Investasi B", 0)
                    pdfLayout.KebijakanInvestasi = file.GetString(reportSection, iniType & " Kebijakan Investasi", "")

                    pdfLayout.EfekPortfolio_R = file.GetInteger(reportSection, iniType & " Efek Portfolio R", 0)
                    pdfLayout.EfekPortfolio_G = file.GetInteger(reportSection, iniType & " Efek Portfolio G", 0)
                    pdfLayout.EfekPortfolio_B = file.GetInteger(reportSection, iniType & " Efek Portfolio B", 0)
                    pdfLayout.EfekPortfolio = file.GetString(reportSection, iniType & " Efek Portfolio", "")

                    pdfLayout.InformasiDividend_R = file.GetInteger(reportSection, iniType & " Informasi Dividend R", 0)
                    pdfLayout.InformasiDividend_G = file.GetInteger(reportSection, iniType & " Informasi Dividend G", 0)
                    pdfLayout.InformasiDividend_B = file.GetInteger(reportSection, iniType & " Informasi Dividend B", 0)
                    pdfLayout.InformasiDividend = file.GetString(reportSection, iniType & " Informasi Dividend", "")

                    pdfLayout.ReportLine_R = file.GetInteger(reportSection, iniType & " Report Line R", 0)
                    pdfLayout.ReportLine_G = file.GetInteger(reportSection, iniType & " Report Line G", 0)
                    pdfLayout.ReportLine_B = file.GetInteger(reportSection, iniType & " Report Line B", 0)

                    pdfLayout.ItemWarna_R = file.GetInteger(reportSection, iniType & " Item Warna R", 0)
                    pdfLayout.ItemWarna_G = file.GetInteger(reportSection, iniType & " Item Warna G", 0)
                    pdfLayout.ItemWarna_B = file.GetInteger(reportSection, iniType & " Item Warna B", 0)

                    pdfLayout.ChartBorder_R = file.GetInteger(reportSection, iniType & " Chart Border R", 0)
                    pdfLayout.ChartBorder_G = file.GetInteger(reportSection, iniType & " Chart Border G", 0)
                    pdfLayout.ChartBorder_B = file.GetInteger(reportSection, iniType & " Chart Border B", 0)
                    If file.GetBoolean(reportSection, iniType & " Chart Border", False) Then pdfLayout.ChartBorder = True Else pdfLayout.ChartBorder = False
                End If
            Else
                pdfColorDefault()
            End If
        Catch ex As Exception
            pdfColorDefault()
        End Try
    End Sub

#End Region

    Private Sub ReportProductFocusEQ_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetMasterSimpi()
        GetInstrumentUser()
        GetParameterInstrumentType()
        dtAs.Value = Now

        objPortfolio.UserAccess = objAccess
        objNAV.UserAccess = objAccess
        objReturn.UserAccess = objAccess
        objBenchmark.UserAccess = objAccess
        objSecurities.UserAccess = objAccess

        DBGPerformance1.FetchRowStyles = True
        DBGHolding.FetchRowStyles = True
        DBGDividend.FetchRowStyles = True
    End Sub

    Private Sub btnSearchPortfolio_Click(sender As Object, e As EventArgs) Handles btnSearchPortfolio.Click
        PortfolioSearch()
    End Sub

    Private Sub PortfolioSearch()
        Dim form As New SelectMasterPortfolio
        form.lblCode = lblPortfolioCode
        form.lblName = lblPortfolioName
        form.lblSimpiEmail = lblSimpiEmail
        form.lblSimpiName = lblSimpiName
        form.Show()
        form.MdiParent = MDIMENU
        lblPortfolioCode.Text = ""
        lblPortfolioName.Text = ""
        lblSimpiEmail.Text = ""
        lblSimpiName.Text = ""
        objPortfolio.Clear()
    End Sub

    Private Sub lblSimpiEmail_TextChanged(sender As Object, e As EventArgs) Handles lblSimpiEmail.TextChanged
        PortfolioLoad()
    End Sub

    Private Sub PortfolioLoad()
        If lblPortfolioCode.Text.Trim <> "" Then
            objPortfolio.Clear()
            objPortfolio.LoadCode(objMasterSimpi, lblPortfolioCode.Text)
            If objPortfolio.ErrID = 0 Then


            Else
                ExceptionMessage.Show(objPortfolio.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        DataLoad()
        DataDisplay()
    End Sub

    Private Sub DataLoad()
        If objPortfolio.GetPortfolioID > 0 Then
            objNAV.Clear()
            objNAV.LoadAt(objPortfolio, dtAs.Value)
            'objPortfolio.LoadCode(objMasterSimpi, dtAs.Value)
            'objReturn.LoadAt(objPortfolio, dtAs.Value)
            'objPortfolio.LoadCode(objPortfolio, )
            If objNAV.ErrID = 0 Then
                txtCcy.Text = objPortfolio.GetPortfolioCcy.GetCcyDescription & " (" & objPortfolio.GetPortfolioCcy.GetCcy & ")"
                txtInception.Text = objPortfolio.GetInceptionDate.ToString
                txtBenchmark.Text = "GetPortfolioBenchmark.GetBenchmarkName not found"
                txtNAVUnit.Text = objNAV.GetNAVPerUnit.ToString("n2")
                txtAUM.Text = (objNAV.GetNAV / 1000000).ToString("n4")
            End If
            dtDividend = objNAV.SearchHistoryLast(objPortfolio, objNAV.GetPositionDate, 0)
            dtReturn = objReturn.SearchHistoryLast(objPortfolio, objReturn.GetPositionDate, 0)
            With chartPerformance
                .Style.Border.BorderStyle = C1.Win.C1Chart.BorderStyleEnum.None
                Dim ds As C1.Win.C1Chart.ChartDataSeriesCollection = .ChartGroups(0).ChartData.SeriesList
                ds.Clear()
                Dim series As C1.Win.C1Chart.ChartDataSeries = ds.AddNewSeries()
                series.Label = "Series 1"
                series.LineStyle.Color = Color.LightSteelBlue

                series.LineStyle.Thickness = 2
                series.SymbolStyle.Shape = C1.Win.C1Chart.SymbolShapeEnum.None
                series.FitType = C1.Win.C1Chart.FitTypeEnum.Line

                series.X.CopyDataIn((From u In dtDividend.AsEnumerable Select u.Field(Of Date)("PositionDate")).ToArray)
                series.Y.CopyDataIn((From u In dtDividend.AsEnumerable Select u.Field(Of Decimal)("GeometricIndex") - 1).ToArray)
                series.PointData.Length = dtDividend.Rows.Count
                .ChartArea.AxisX.AnnoFormat = FormatEnum.DateYear

                'Dim ds2 As C1.Win.C1Chart.ChartDataSeriesCollection = .ChartGroups(1).ChartData.SeriesList
                'ds2.Clear()
                'Dim series2 As C1.Win.C1Chart.ChartDataSeries = ds.AddNewSeries()
                'series2.Label = "Series 2"
                'series2.LineStyle.Color = Color.Blue
                'series2.LineStyle.Thickness = 2
                'series2.SymbolStyle.Shape = C1.Win.C1Chart.SymbolShapeEnum.None
                'series2.FitType = C1.Win.C1Chart.FitTypeEnum.Line

                'series2.X.CopyDataIn((From u In dtDividend.AsEnumerable Select u.Field(Of Date)("PositionDate")).ToArray)
                'series2.Y.CopyDataIn((From u In dtDividend.AsEnumerable Select u.Field(Of Decimal)("GeometricIndex") - 1).ToArray)
                'series2.PointData.Length = dtDividend.Rows.Count
                '.ChartArea.AxisX.AnnoFormat = FormatEnum.DateYear
                'series.PointData.Length =
            End With
        End If
    End Sub

    Private Sub DataDisplay()
        If objNAV.GetNAV > 0 Then

        End If
    End Sub

    Private Sub btnEmail_Click(sender As Object, e As EventArgs) Handles btnEmail.Click
        ReportEmail()
    End Sub

    Private Sub ReportEmail()
        'If DBGDividend.RowCount > 0 Then
        '    Dim frm As New ReportFundSheetDividendEmail
        '    frm.Show()
        '    frm.frm = Me
        '    frm.MdiParent = MDISO
        'End If
    End Sub

    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles btnPDF.Click
        ExportPDF(False)
    End Sub

    Public Function ExportPDF(ByVal isAttachment As Boolean) As String
        Return PrintPDF(isAttachment)
    End Function
    Private Function PrintPDF(ByVal isAttachment As Boolean) As String
        'Stuff
        Dim strFile As String = ""
        Dim strLayout As String = ""
        Dim layout = Image.FromFile("..\..\Template\Report Fund Sheet Dividend - Portrait.jpg")
        Dim myBrush As New SolidBrush(Color.FromArgb(0, 61, 121))
        Dim detailBrush As New SolidBrush(Color.Black)
        Dim headerBrush As New SolidBrush(Color.White)
        Dim koordX As Single = 15, koordY As Single = 35
        Dim fontType = "Calibri", fontSize = 8
        With c1pdf
            .Clear()
            .PaperKind = Printing.PaperKind.A4
            'Another Stuff
            Dim rc As RectangleF = .PageRectangle
            .DrawImage(layout, rc)
            Dim pdf_height As Integer = c1pdf.PageRectangle.Height
            Dim pdf_width As Integer = c1pdf.PageRectangle.Width
            'Dim str As String
            Dim strRight, strCenter, strLeft, sfCenter, sfLCenter, sfRCenter As New StringFormat()
            strLeft.Alignment = StringAlignment.Near
            strRight.Alignment = StringAlignment.Far
            strCenter.Alignment = StringAlignment.Center
            sfCenter.Alignment = StringAlignment.Center
            sfCenter.LineAlignment = StringAlignment.Near
            sfLCenter.Alignment = StringAlignment.Near
            sfLCenter.LineAlignment = StringAlignment.Near
            sfRCenter.Alignment = StringAlignment.Far
            sfRCenter.LineAlignment = StringAlignment.Near

            .DrawString(Date.Now.ToString("dd MMMM yyyy"), New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 538, koordY + 40), strRight)
            'Column separator
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 250, koordY + 45), New PointF(koordX + 250, koordY + 580))
#Region "Column"
            'Left Column
            .DrawString(pdfLayout.TujuanInvestasi, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 26, koordY + 45))
            'Right Column
            .DrawString(pdfLayout.ChartTitle, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 260, koordY + 45))
            koordY += 60
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 26, koordY), New PointF(koordX + 240, koordY))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 260, koordY), New PointF(koordX + 535, koordY))
            koordY += 5
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                "Avrist Prime Income Fund (APIF) bertujuan untuk memberikan pertumbuhan nilai investasi yang relatif stabil melalui investasi pada efek bersifat utang melalui pemilihan penerbit surat utang secara hati-hati",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc)
            'Chart
            'chartPerformance.ChartGroups(0).ChartData.SeriesList(0).LineStyle.Color = Color.FromArgb(pdfLayout.ChartLine_R, pdfLayout.ChartLine_G, pdfLayout.ChartLine_B)
            chartPerformance.ChartArea.AxisX.GridMajor.Visible = False
            Dim imgChart = chartPerformance.GetImage(ImageFormat.Emf)
            rc = New RectangleF(koordX + 260, koordY, 280, 80)
            .DrawImage(imgChart, rc, ContentAlignment.MiddleCenter, C1.C1Pdf.ImageSizeModeEnum.Stretch)
            If pdfLayout.ChartBorder Then
                Dim pnNAV As New Pen(New SolidBrush(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)), 0.5)
                .DrawRectangle(pnNAV, rc)
            Else
                .DrawRectangle(Pens.White, rc)
            End If
            'Informasi Reksa Dana
            koordY += 45
            .DrawString(pdfLayout.InformasiReksaDana, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 26, koordY))
            koordY += 10
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 26, koordY), New PointF(koordX + 240, koordY))
            koordY += 5
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                pdfLayout.JenisReksa,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtAssetType.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                pdfLayout.TanggalPeluncuran,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtInceptionDate.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                pdfLayout.DanaKelolaan,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtAUM.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                pdfLayout.MataUang,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtCcy.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                pdfLayout.FrekuensiValuasi,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtValuation.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                pdfLayout.BankKustodian,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtCustodian.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                pdfLayout.TolakUkur,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtBenchmark.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                pdfLayout.NabUnit,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtNAVUnit.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            'Investasi dan Biaya-biaya | Kinerja Kumulatif
            koordY += 15
            .DrawString(pdfLayout.InvestasiDanBiaya, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 26, koordY))
            .DrawString(pdfLayout.KinerjaKumulatif, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 260, koordY))
            koordY += 10
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 26, koordY), New PointF(koordX + 240, koordY))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 260, koordY), New PointF(koordX + 535, koordY))
            '.DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
            '          New PointF(koordX + 260, koordY + 15), New PointF(koordX + 535, koordY + 15))
            '.DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
            '          New PointF(koordX + 260, koordY + 35), New PointF(koordX + 535, koordY + 35))
            koordY += 5
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                InputLabel10.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtMinimumInitialSubscription.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                InputLabel11.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtMinimumAdditionalSubscription.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                "Biaya Penjualan (%)",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "Maks 1.00%",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                InputLabel12.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtSellingFee.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                "0% untuk kepemilikan diatas 1 tahun",
                New Font(fontType, 8, FontStyle.Bold),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                InputLabel13.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtRedemptionFee.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                InputLabel14.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtManagementFee.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                InputLabel15.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                txtCustodianFee.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)

            'Statistik Reksadana | Kebijakan Investasi
            koordY += 15
            .DrawString(pdfLayout.StatistikReksadana, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 26, koordY))
            .DrawString(pdfLayout.KebijakanInvestasi, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 260, koordY))
            koordY += 10
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 26, koordY), New PointF(koordX + 240, koordY))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 260, koordY), New PointF(koordX + 535, koordY))
            '.DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
            '          New PointF(koordX + 260, koordY + 10), New PointF(koordX + 535, koordY + 10))
            '.DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
            '          New PointF(koordX + 260, koordY + 35), New PointF(koordX + 535, koordY + 35))
            koordY += 5
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                "1 Statistik Reksadana",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "1 Statistik Reksadana",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                "2 Statistik Reksadana",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "2 Statistik Reksadana",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                "3 Statistik Reksadana",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "3 Statistik Reksadana",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                "4 Statistik Reksadana          3.61",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "4 Statistik Reksadana",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                "5 Statistik Reksadana         (1.74)",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "5 Statistik Reksadana",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                "6 Statistik Reksadana",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "6 Statistik Reksadana",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strRight)
            'Risiko Investasi | Efek Portfolio
            koordY += 15
            .DrawString(pdfLayout.RisikoInvestasi, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 26, koordY))
            .DrawString(pdfLayout.EfekPortfolio, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 260, koordY))
            koordY += 10
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 26, koordY), New PointF(koordX + 240, koordY))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 260, koordY), New PointF(koordX + 535, koordY))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 260, koordY + 15), New PointF(koordX + 535, koordY + 15))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 260, koordY + 65), New PointF(koordX + 535, koordY + 65))
            koordY += 5
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                txtRisk1.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "Efek",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), New PointF(koordX + 260, koordY), strLeft)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                txtRisk2.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "Efek 1",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), New PointF(koordX + 260, koordY), strLeft)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                txtRisk3.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "Efek 2",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), New PointF(koordX + 260, koordY), strLeft)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                txtRisk4.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "Efek 3",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), New PointF(koordX + 260, koordY), strLeft)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                txtRisk5.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "Efek 4",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), New PointF(koordX + 260, koordY), strLeft)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                txtRisk6.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            .DrawString(
                "Efek 5",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), New PointF(koordX + 260, koordY), strLeft)
            koordY += 10
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                txtRisk7.Text,
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            'Klasifikasi Risiko | Informasi Dividend
            koordY += 15
            .DrawString(pdfLayout.KlasifikasiRisiko, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 26, koordY))
            .DrawString(pdfLayout.InformasiDividend, New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 260, koordY))
            koordY += 10
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 26, koordY), New PointF(koordX + 240, koordY))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 260, koordY), New PointF(koordX + 535, koordY))

            koordY += 5
            .FillRectangle(Brushes.AliceBlue, New RectangleF(46, 530, 35, 10))
            .DrawString("1", New Font("calibri", 8), Brushes.Black, New RectangleF(46, 530, 35, 8), strCenter)
            .FillRectangle(Brushes.AliceBlue, New RectangleF(76, 530, 35, 10))
            .DrawString("2", New Font("calibri", 8), Brushes.Black, New RectangleF(76, 530, 35, 8), strCenter)
            .FillRectangle(Brushes.AliceBlue, New RectangleF(111, 530, 35, 10))
            .DrawString("3", New Font("calibri", 8), Brushes.Black, New RectangleF(111, 530, 35, 8), strCenter)
            .FillRectangle(Brushes.AliceBlue, New RectangleF(146, 530, 35, 10))
            .DrawString("4", New Font("calibri", 8), Brushes.Black, New RectangleF(146, 530, 35, 8), strCenter)
            .FillRectangle(Brushes.Purple, New RectangleF(182, 530, 35, 10))
            .DrawString("5", New Font("calibri", 8), Brushes.White, New RectangleF(182, 530, 35, 8), strCenter)
            koordY += 5
            'Mengenai Manajer Investasi
            koordY += 15
            .DrawString("Mengenal Manajer Investasi", New Font(fontType, 8, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)),
                        New PointF(koordX + 26, koordY))
            koordY += 10
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
                      New PointF(koordX + 26, koordY), New PointF(koordX + 240, koordY))
            koordY += 5
#End Region

#Region "Footer"
            rc = New RectangleF(koordX + 26, koordY, 214, 50)
            .DrawString(
                "PT Avrist Asset Management merupakan anak perusahaan dari PT Avrist Assurance ('Avrist'). PT. Avrist Asset Management didukung oleh professional yang berpengalaman di bidang investasi dan menawarkan beragam solusi investasi yang disesuaikan dengan kondisi pasar dan tujuan investasi pemodal.",
                New Font(fontType, 8),
                New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc, strLeft)
            koordY += 60
            rc = New RectangleF(koordX + 26, koordY, 509, 70)
            .DrawRectangle(New Pen(New SolidBrush(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B))), rc)
            rc = New RectangleF(koordX + 30, koordY, 509, 70)
            .DrawString("Akhir Maret, pasar obligasi domestik ditutup naik (INDOBEX Composite Index +0.6% mom) dengan yield obligasi pemerintah bertenor 10 tahun bergerak naik (4.3 bps mom) menjadi 6.68%. Investor asing melakukan aksi beli yang tercatat sebesar Rp.10.56tn selama Maret walaupun porsi persentase kepimilikan asing menjadi turun menjadi 39.31% dari bulan sebelumnya sebesar 39.83%. " & ControlChars.Cr & ControlChars.Cr & ControlChars.Lf & "Rupiah menutup perdagangan bulan Maret dengan ditutup pada level Rp13.768 atau melemah -0.17% mom. Tekanan dan volatilitas Rupiah cenderung mereda setelah trader cenderung melepas USD yang terpapar sentiment deficit perdagangan US serta intervensi BI untuk menstabilkan rupiah.",
                        New Font(fontType, 8),
                        New SolidBrush(Color.FromArgb(pdfLayout.ItemWarna_R, pdfLayout.ItemWarna_G, pdfLayout.ItemWarna_B)), rc)
            '.DrawLine(New Pen(Color.FromArgb(pdfLayout.ReportLine_R, pdfLayout.ReportLine_G, pdfLayout.ReportLine_B)),
            '          New PointF(koordX + 26, koordY), New PointF(koordX + 536, koordY))
            'koordY += 5
            '.DrawString("Mengenal Manajer Investasi", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 445, 175, 8))
            '.DrawLine(Pens.Black, 25, 455, 265, 455)
            'rc = New RectangleF(koordX + 21, koordY, 209, 50)
            '.DrawString(
            '    "PT Avrist Asset Management merupakan anak perusahaan dari PT Avrist Assurance ('Avrist'). PT. Avrist Asset Management didukung oleh professional yang berpengalaman di bidang investasi dan menawarkan beragam solusi investasi yang disesuaikan dengan kondisi pasar dan tujuan investasi pemodal.",
            '    New Font(fontType, 8),
            '    New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)), rc)
            '.DrawLine(Pens.Black, 225, 70, 225, 506)

            ''Grafik Kinerja Reksa Dana
            '.DrawString("Grafik Kinerja Reksa Dana", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(230, 60, 175, 8))
            '.DrawLine(Pens.Black, 230, 70, pdf_width - 73, 70)

            ''Informasi Reksa Dana
            '.DrawString("Informasi Reksa Dana", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 125, 175, 8))
            '.DrawLine(Pens.Black, 25, 135, 215, 134)
            '.DrawString("Jenis Reksa Dana", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 136, 150, 8))
            '.DrawString("Tanggal Peluncuran", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 144, 150, 8))
            '.DrawString("Dana Kelolaan (Rp. Mil)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 152, 150, 8))
            '.DrawString("Mata Uang", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 160, 150, 8))
            '.DrawString("Frekuensi Valuasi", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 168, 150, 8))
            '.DrawString("Bank Kustodian", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 176, 150, 8))
            '.DrawString("Tolok Ukur", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 184, 150, 8))
            '.DrawString("NAB/Unit (Rp/Unit)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 192, 150, 8))
            ''Value Informasi Reksa Dana
            '.DrawString("Ekultas", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 135, 150, 8), strRight)
            '.DrawString("18-Dec-2017", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 143, 150, 8), strRight)
            '.DrawString("135.92", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 151, 150, 8), strRight)
            '.DrawString("Rupiah", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 159, 150, 8), strRight)
            '.DrawString("Harian", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 167, 150, 8), strRight)
            '.DrawString("Standard Chartered Bank", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 175, 150, 8), strRight)
            '.DrawString("IDX30 Indeks", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 183, 150, 8), strRight)
            '.DrawString("960.97", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 191, 150, 8), strRight)

            ''Investasi dan Biaya-Biaya
            '.DrawString("Investasi dan Biaya-Biaya", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 207, 175, 8))
            '.DrawLine(Pens.Black, 25, 217, 215, 217)
            '.DrawString("Minimal Investasi Awal (Rp)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 218, 150, 8))
            '.DrawString("Minimal Investasi Selanjutnya (Rp)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 226, 150, 8))
            '.DrawString("Biaya Pembelian (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 234, 150, 8))
            '.DrawString("Biaya Penjualan (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 242, 150, 8))
            '.DrawString("Biaya Jasa Pengelolaan Ml (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 250, 150, 8))
            '.DrawString("Biaya Jasa Bank Kustodian (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 258, 150, 8))

            ''Value Investasi dan Biaya-Biaya
            '.DrawString("100.000", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 218, 150, 8), strRight)
            '.DrawString("100.000", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 226, 150, 8), strRight)
            '.DrawString("Maks 1.00", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 234, 150, 8), strRight)
            '.DrawString("Maks 1.00", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 242, 150, 8), strRight)
            '.DrawString("Maks 1.00", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 250, 150, 8), strRight)
            '.DrawString("Maks 0.25", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 258, 150, 8), strRight)

            ''Kinerja Kumulatif
            '.DrawString("Kinerja Kumulatif (%)", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(230, 207, 175, 8))
            '.DrawLine(Pens.Black, 230, 217, pdf_width - 73, 217)
            '.DrawString("1 Bln", New Font("calibri", 8), Brushes.Black, New RectangleF(347, 219, 25, 8), strRight)
            '.DrawString("3 Bln", New Font("calibri", 8), Brushes.Black, New RectangleF(372, 219, 25, 8), strRight)
            '.DrawString("6 Bln", New Font("calibri", 8), Brushes.Black, New RectangleF(397, 219, 25, 8), strRight)
            '.DrawString("YTD", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 219, 25, 8), strRight)
            '.DrawString("1 Thn", New Font("calibri", 8), Brushes.Black, New RectangleF(447, 219, 25, 8), strRight)
            '.DrawString("3 Thn", New Font("calibri", 8), Brushes.Black, New RectangleF(472, 219, 25, 8), strRight)
            '.DrawString("5 Thn", New Font("calibri", 8), Brushes.Black, New RectangleF(497, 219, 25, 8), strRight)
            '.DrawString("SP", New Font("calibri", 8), Brushes.Black, New RectangleF(522, 219, 25, 8), strRight)
            '.DrawLine(Pens.Black, 230, 229, pdf_width - 73, 229)
            '.DrawString("Avrist IDX 30", New Font("calibri", 8), Brushes.Black, New RectangleF(232, 229, 175, 8))
            '.DrawString("(8.51)", New Font("calibri", 8), Brushes.Black, New RectangleF(347, 229, 25, 8), strRight)
            '.DrawString("(7.73)", New Font("calibri", 8), Brushes.Black, New RectangleF(372, 229, 25, 8), strRight)
            '.DrawString("n/a", New Font("calibri", 8), Brushes.Black, New RectangleF(397, 229, 25, 8), strRight)
            '.DrawString("(7.73)", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 229, 25, 8), strRight)
            '.DrawString("n/a", New Font("calibri", 8), Brushes.Black, New RectangleF(447, 229, 25, 8), strRight)
            '.DrawString("n/a", New Font("calibri", 8), Brushes.Black, New RectangleF(472, 229, 25, 8), strRight)
            '.DrawString("n/a", New Font("calibri", 8), Brushes.Black, New RectangleF(497, 229, 25, 8), strRight)
            '.DrawString("(3.90)", New Font("calibri", 8), Brushes.Black, New RectangleF(522, 229, 25, 8), strRight)
            '.DrawString("Tolak Ukur", New Font("calibri", 8), Brushes.Black, New RectangleF(232, 237, 175, 8))
            '.DrawString("(8.51)", New Font("calibri", 8), Brushes.Black, New RectangleF(347, 237, 25, 8), strRight)
            '.DrawString("(7.73)", New Font("calibri", 8), Brushes.Black, New RectangleF(372, 237, 25, 8), strRight)
            '.DrawString("n/a", New Font("calibri", 8), Brushes.Black, New RectangleF(397, 237, 25, 8), strRight)
            '.DrawString("(7.73)", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 237, 25, 8), strRight)
            '.DrawString("n/a", New Font("calibri", 8), Brushes.Black, New RectangleF(447, 237, 25, 8), strRight)
            '.DrawString("n/a", New Font("calibri", 8), Brushes.Black, New RectangleF(472, 237, 25, 8), strRight)
            '.DrawString("n/a", New Font("calibri", 8), Brushes.Black, New RectangleF(497, 237, 25, 8), strRight)
            '.DrawString("(3.90)", New Font("calibri", 8), Brushes.Black, New RectangleF(522, 237, 25, 8), strRight)
            '.DrawLine(Pens.Black, 230, 248, pdf_width - 73, 248)
            '.DrawString("*SP : Sejak Peluncuran", New Font("calibri", 8), Brushes.Black, New RectangleF(256, 259, 175, 8))

            ''Kebijakan Investasi
            '.DrawString("Kebijakan Investasi", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(230, 274, 175, 8))
            '.DrawLine(Pens.Black, 230, 274, pdf_width - 73, 274)
            '.DrawString("%", New Font("calibri", 8), Brushes.Black, New RectangleF(372, 275, 15, 8), strRight)
            '.DrawString("%", New Font("calibri", 8), Brushes.Black, New RectangleF(pdf_width - 90, 275, 15, 8), strRight)
            '.DrawLine(Pens.Black, 230, 285, pdf_width - 73, 285)

            ''Statistik Reksa Dana
            '.DrawString("Statistik Reksadana", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 274, 175, 8))
            '.DrawLine(Pens.Black, 25, 284, 215, 284)
            '.DrawString("Kinerja Sejak Diluncurkan (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 285, 150, 8))
            '.DrawString("Standar Deviasi Disetahunkan (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 293, 150, 8))
            '.DrawString("Beta", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 301, 150, 8))
            '.DrawString("Kinerja Bulanan Terbaik (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 309, 150, 8))
            '.DrawString("Kinerja Bulanan Terburuk (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 317, 150, 8))
            '.DrawString("Kinerja terbaik setahun terakhir (%)", New Font("calibri", 8, FontStyle.Bold), Brushes.Black, New RectangleF(25, 325, 150, 8))

            ''Value Statistik Reksadana
            '.DrawString("(3.90)", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 285, 150, 8), strRight)
            '.DrawString("18.86", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 293, 150, 8), strRight)
            '.DrawString("0.97", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 301, 150, 8), strRight)
            '.DrawString("4.15", New Font("calibri", 8), Brushes.Black, New RectangleF(100, 309, 125, 8), strRight)
            '.DrawString("(8.51)", New Font("calibri", 8), Brushes.Black, New RectangleF(100, 317, 125, 8), strRight)
            '.DrawString("Dec-17", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 309, 150, 8), strRight)
            '.DrawString("Mar-18", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 317, 150, 8), strRight)
            '.DrawString("4.15", New Font("calibri", 8, FontStyle.Bold), Brushes.Black, New RectangleF(65, 325, 150, 8), strRight)

            ''Risiko Investasi
            '.DrawString("Risiko Investasi", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 341, 175, 8))
            '.DrawLine(Pens.Black, 25, 351, 215, 351)
            '.DrawString("1. Risiko Perubahan Kondisi Ekonomi dan Politik", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 352, 150, 8))
            '.DrawString("2. Risiko Pasar", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 360, 150, 8))
            '.DrawString("3. Risiko Kredit/Wanprestasi", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 368, 150, 8))
            '.DrawString("4. Risiko Likuiditas", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 376, 150, 8))
            '.DrawString("5. Risiko Perubahan Peraturan", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 384, 150, 8))
            '.DrawString("6. Risiko Berkurangnya Nilai Aktiva Bersih", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 392, 150, 8))
            '.DrawString("7. Risiko Pembubaran dan Likuidasi", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 400, 150, 8))

            ''Klasifikasi Risiko
            '.DrawString("Klasifikasi Risiko", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 416, 175, 8))
            '.DrawLine(Pens.Black, 25, 426, 215, 426)
            '.FillRectangle(Brushes.AliceBlue, New RectangleF(25, 427, 35, 10))
            '.DrawString("1", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 427, 35, 8), strRight)
            '.FillRectangle(Brushes.AliceBlue, New RectangleF(61, 427, 35, 10))
            '.DrawString("2", New Font("calibri", 8), Brushes.Black, New RectangleF(61, 427, 35, 8), strRight)
            '.FillRectangle(Brushes.AliceBlue, New RectangleF(97, 427, 35, 10))
            '.DrawString("3", New Font("calibri", 8), Brushes.Black, New RectangleF(97, 427, 35, 8), strRight)
            '.FillRectangle(Brushes.AliceBlue, New RectangleF(133, 427, 35, 10))
            '.DrawString("4", New Font("calibri", 8), Brushes.Black, New RectangleF(133, 427, 35, 8), strRight)
            '.FillRectangle(Brushes.Purple, New RectangleF(169, 427, 35, 10))
            '.DrawString("5", New Font("calibri", 8), Brushes.White, New RectangleF(169, 427, 35, 8), strRight)

            '.DrawString("Mengenal Manajer Investasi", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 445, 175, 8))
            '.DrawLine(Pens.Black, 25, 455, 215, 455)
            '.DrawString(
            '"PT Avrist Asset Management merupakan anak perusahaan dari PT Avrist Assurance ('Avrist'). PT. Avrist Asset Management didukung oleh professional yang berpengalaman di bidang investasi dan menawarkan beragam solusi investasi yang disesuaikan dengan kondisi pasar dan tujuan investasi pemodal.", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 456, 190, 54))
            '.DrawLine(Pens.Black, 225, 70, 225, 506)

            '.DrawRectangle(Pens.Black, New RectangleF(25, 510, pdf_width - 98, 78))
            '.DrawString("Akhir Maret, IHSG ditutup turun -6.2% (mom) ke level 6.188 sedangkan indeks LQ45 ditutup turun -8.6% (mom) ke level 1.005. Selama Maret, investor asing melakukan penjualan bersih sebesar Rp. 14.9 tn. Volatilitas rupiah yang disebabkan ketidakpastian kebijakan moneter ditambah dengan sentiment kekhawatiran peranf dagang menyeret indeks ditutup", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 514, pdf_width - 102, 7))
            '.DrawString("bersih sebesar Rp. 14.9 tn. Volatilitas rupiah yang disebabkan ketidakpastian kebijakan moneter ditambah dengan sentiment kekhawatiran perang dagang menyeret indeks ditutup pada teritori negatif.", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 526, pdf_width - 100, 7))
            '.DrawString("pada teritori negatif.", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 538, pdf_width - 100, 7))
            '.DrawString("Sementara itu pasar saham global mayoritas ditutup turun (S&P 500-2.7%, FTSE 100-2.4%, Nikkel 225-4.1% mom). Diawali ketidakpastian landscape kebijakan moneter US ", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 550, pdf_width - 100, 7))
            '.DrawString("dan memanasnya suhu politik di US, bursa global kembali tertekan setelah US mengesahkan trade protectionism program yang menyeret mayoritas bursa global pada teritori", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 562, pdf_width - 100, 7))
            '.DrawString("negatif.", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 574, pdf_width - 100, 7))
#End Region
            strFile = reportFileExists("Report Fund Sheet Dividend " & lblPortfolioCode.Text.Trim & dtAs.Value.ToString("yyyymmdd") & ".pdf")
            .Save(strFile)

        End With
        If Not isAttachment Then Process.Start(strFile)
        Return strFile
    End Function
    Private Sub ReportSetting()
        Dim frm As New ReportFundSheetDividendSetting
        frm.frm = Me
        frm.Show()
        frm.FormLoad()
        frm.MdiParent = MDIMENU
    End Sub

    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click
        ReportSetting
    End Sub

End Class