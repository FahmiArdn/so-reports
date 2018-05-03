Imports simpi.GlobalUtilities
Imports simpi.CoreData
Imports simpi.MasterPortfolio
Imports simpi.GlobalConnection
Imports C1.Win.C1TrueDBGrid

Public Class ReportFundSheetSectorEQ
    Dim objPortfolio As New MasterPortfolio
    Dim objSimpi As New simpi.MasterSimpi.MasterSimpi
    Dim objCodeset As New PortfolioCodeset
    Dim objNAV As New PortfolioNAV
    Dim objReturn As New PortfolioReturn
    Dim objBenchmark As New simpi.CoreData.PortfolioBenchmark
    Dim objSecurities As New simpi.CoreData.PositionSecurities
    Dim dtPerformance As New DataTable
    Dim dtBenchmark As New DataTable
    Dim dtNAV As New DataTable
    Dim dtReturn As New DataTable
    Dim dtSecurities As New DataTable
    Dim reportSection As String = "Report Fund Sheet Sector EQ"


    Public pdfLayout As New pdfColor


    Private Sub ReportProductFocusEQ_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetInstrumentUser()
        GetParameterInstrumentType()
        dtAs.Value = Now

        objPortfolio.UserAccess = objAccess
        objSimpi.UserAccess = objAccess
        objCodeset.UserAccess = objAccess
        objNAV.UserAccess = objAccess
        objReturn.UserAccess = objAccess
        objBenchmark.UserAccess = objAccess
        objSecurities.UserAccess = objAccess
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
            objSimpi.Clear()
            objSimpi.Load(lblSimpiEmail.Text)
            If objSimpi.ErrID = 0 Then
                objPortfolio.Clear()
                objPortfolio.LoadCode(objSimpi, lblPortfolioCode.Text)
                If objPortfolio.ErrID = 0 Then

                Else
                    ExceptionMessage.Show(objPortfolio.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                ExceptionMessage.Show(objSimpi.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        DataLoad()
    End Sub

    Private Sub DataLoad()
        If objPortfolio.GetPortfolioID > 0 Then
            objNAV.Clear()
            objNAV.LoadAt(objPortfolio, dtAs.Value)
            objPortfolio.LoadCode(objMasterSimpi, dtAs.Value)
            objBenchmark.LoadAt(objPortfolio, dtAs.Value)
            objSecurities.Search(objPortfolio, dtAs.Value)
            DBGPerformance1.Columns.Clear()
            dtNAV = objNAV.SearchHistoryLast(objPortfolio, objNAV.GetPositionDate)
            dtBenchmark = objBenchmark.SearchHistoryLast(objPortfolio, objBenchmark.GetPositionDate)
            dtReturn = objReturn.SearchHistoryLast(objPortfolio, objReturn.GetPositionDate)
            dtSecurities = objSecurities.Search(objPortfolio, objSecurities.GetPositionDate)
            If objNAV.ErrID = 0 Then
                'data load

                'data display
                txtAssetType.Text = objPortfolio.GetAssetType.GetAssetTypeDescription.ToString
                txtInceptionDate.Text = objPortfolio.GetInceptionDate.ToString
                txtAUM.Text = (objNAV.GetNAV / 1000000).ToString("n2")
                txtCcy.Text = objPortfolio.GetPortfolioCcy.GetCcy.ToString
                txtValuation.Text = "Not Found"
                txtCustodian.Text = "Not Found"
                txtBenchmark.Text = objPortfolio.GetPortfolioBenchmarkClass.GetClassName.ToString
                txtNAVUnit.Text = objNAV.GetNAVPerUnit.ToString
                txtInception.Text = "Not Found"
                'txtBestMonth.Text = (From u In dtNAV.AsEnumerable Select u.Field(Of Integer)("")).ToString

            Else
                ExceptionMessage.Show(objNAV.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If


            'If dtNAV.Rows.Count > 0 AndAlso dtBenchmark.Rows.Count > 0 Then
            With chartPerformance
                Dim firstdate, lastdate As Date

                firstdate = CDate(dtNAV.Rows(0)("PositionDate"))
                lastdate = CDate(dtNAV.Rows(dtNAV.Rows.Count - 1)("PositionDate"))
                .Style.Border.BorderStyle = C1.Win.C1Chart.BorderStyleEnum.None
                .BackColor = Color.Transparent
                .ChartArea.AxisX.AutoMax = True
                .ChartArea.AxisX.AutoMin = True
                .ChartArea.AxisX.AutoMajor = True
                .ChartArea.AxisX.AutoMinor = True
                .ChartArea.AxisY.AutoMax = True
                .ChartArea.AxisY.AutoMin = True
                .ChartArea.AxisY.AutoMajor = True
                .ChartArea.AxisY.AutoMinor = True
                .ChartArea.AxisX.AnnoFormat = C1.Win.C1Chart.FormatEnum.DateManual
                .ChartArea.AxisX.Origin = .ChartArea.AxisX.Min
                .ChartArea.Style.Border.BorderStyle = C1.Win.C1Chart.BorderStyleEnum.None
                .ChartArea.AxisX.Min = lastdate.ToOADate
                .ChartArea.AxisX.Max = firstdate.ToOADate
                Dim ds As C1.Win.C1Chart.ChartDataSeriesCollection = .ChartGroups(0).ChartData.SeriesList
                Dim series As C1.Win.C1Chart.ChartDataSeries = ds.AddNewSeries()
                series.Label = "NAV"
                'series.LineStyle.Color = Color.FromArgb(pdfLayout.ChartLine_R, pdfLayout.ChartLine_G, pdfLayout.ChartLine_B)
                series.LineStyle.Color = Color.Green
                series.LineStyle.Thickness = 2
                series.SymbolStyle.Shape = C1.Win.C1Chart.SymbolShapeEnum.None
                series.FitType = C1.Win.C1Chart.FitTypeEnum.Line
                series.X.CopyDataIn((From u In dtNAV.AsEnumerable Select u.Field(Of Date)("PositionDate")).ToArray)
                series.Y.CopyDataIn((From u In dtNAV.AsEnumerable Select u.Field(Of Decimal)("GeometricIndex") - 1).ToArray)
                series.Y1.CopyDataIn((From u In dtNAV.AsEnumerable Select u.Field(Of Decimal)("GeometricIndex") - 1).ToArray)
                series.PointData.Length = dtNAV.Rows.Count


            End With
            'Pie Chart
            With chartSector
                .ChartArea.Inverted = True
                .ChartGroups(0).ChartType = C1.Win.C1Chart.Chart2DTypeEnum.Pie
                ' Clear previous data
                .ChartGroups(0).ChartData.SeriesList.Clear()
                .BorderStyle = C1.Win.C1Chart.BorderStyleEnum.None
                ' Add Data
                Dim SectorName As String() = {"Pasar Uang 0.4", "Keuangan 38.5", "Konsumer Kebutuhan Pokok", "Industrial 2.0", "Konsumer Kebutuhan Sekunder 8.3", "Utilitas 1.4", "Jasa Telekomunikasi", "Energi 4.6", "Layanan Kesehatan", "Real Estate"}
                Dim PriceX As Integer() = {80, 400, 20, 60, 150, 300,
                130, 500, 100, 100}
                'get series collection
                Dim dscoll As C1.Win.C1Chart.ChartDataSeriesCollection = .ChartGroups(0).ChartData.SeriesList
                For i As Integer = 0 To PriceX.Length - 1
                    'populate the series
                    Dim series As C1.Win.C1Chart.ChartDataSeries = dscoll.AddNewSeries()
                    'Add one point to show one pie
                    series.PointData.Length = 1
                    'Assign the prices to the Y Data series
                    series.Y(0) = PriceX(i)
                    'format the product name and product price on the legend
                    series.Label = String.Format("{0} ({1:c})",
                    SectorName(i), PriceX(i))
                Next
                ' show pie Legend
                .Legend.Visible = True
                'add a title to the chart legend
                .Legend.Text = "Sumber : Bloomberg, PT. Avrist Asset Management"
            End With
        End If
        DisplayPerformance()
    End Sub

    Private Sub DisplayPerformance()
        Dim dtPerformance1 As New DataTable
        If dtPerformance1.Rows.Count = 0 Then
            dtPerformance1.Columns.AddRange(New DataColumn() {
                    New DataColumn("Items", GetType(String)),
                    New DataColumn("1D", GetType(String)),
                    New DataColumn("MTD", GetType(String)),
                    New DataColumn("30D", GetType(String)),
                    New DataColumn("1Mo", GetType(String)),
                    New DataColumn("3Mo", GetType(String)),
                    New DataColumn("6Mo", GetType(String)),
                    New DataColumn("YTD", GetType(String)),
                    New DataColumn("1Y", GetType(String)),
                    New DataColumn("2Y", GetType(String)),
                    New DataColumn("3Y", GetType(String)),
                    New DataColumn("5Y", GetType(String)),
                    New DataColumn("10Y", GetType(String)),
                    New DataColumn("Inception", GetType(String))})
        End If

        dtPerformance1.Clear()
        dtPerformance1.Rows.Add(objPortfolio.GetPortfolioCode, (objReturn.Getr1D * 100).ToString("n2"),
                       (objReturn.GetrMTD * 100).ToString("n2"), (objReturn.Getr30D * 100).ToString("n2"),
                       (objReturn.Getr1Mo * 100).ToString("n2"), (objReturn.Getr3Mo * 100).ToString("n2"),
                       (objReturn.Getr6Mo * 100).ToString("n2"), (objReturn.GetrYTD * 100).ToString("n2"),
                       (objReturn.Getr1Y * 100).ToString("n2"), (objReturn.Getr2Y * 100).ToString("n2"),
                       (objReturn.Getr3Y * 100).ToString("n2"), (objReturn.Getr5Y * 100).ToString("n2"),
                       (objReturn.Getr10Y * 100).ToString("n2"), (objReturn.GetrInception * 100).ToString("n2"))


        dtPerformance1.Rows.Add("Benchmark", (objBenchmark.Getr1D * 100).ToString("n2"),
                (objBenchmark.GetrMTD * 100).ToString("n2"), (objBenchmark.Getr30D * 100).ToString("n2"),
                (objBenchmark.Getr1Mo * 100).ToString("n2"), (objBenchmark.Getr3Mo * 100).ToString("n2"),
                (objBenchmark.Getr6Mo * 100).ToString("n2"), (objBenchmark.GetrYTD * 100).ToString("n2"),
                (objBenchmark.Getr1Y * 100).ToString("n2"), (objBenchmark.Getr2Y * 100).ToString("n2"),
                (objBenchmark.Getr3Y * 100).ToString("n2"), (objBenchmark.Getr5Y * 100).ToString("n2"),
                (objBenchmark.Getr10Y * 100).ToString("n2"), (objBenchmark.GetrInception * 100).ToString("n2"))


        '1: portfolio vs benchmark: 1d, mtd, 30d, 1Mo, 3Mo, 6Mo, YTD, 1Y, 2Y, 3Y, 5Y, 10Y, Inception
        '2: this year vs last year: JAN - DEC, Q1 - Q4
        'report list of fund: MTD, YTD, 2016, 2015, 2014 - 10 tahun, Average 1Y
        With DBGPerformance1
            .AllowAddNew = False
            .AllowDelete = False
            .AllowUpdate = False
            .Style.WrapText = False
            .Columns.Clear()
            .DataSource = dtPerformance1

            .Splits(0).DisplayColumns("1D").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("MTD").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("30D").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("1Mo").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("3Mo").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("6Mo").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("YTD").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("1Y").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("2Y").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("3Y").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("5Y").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("10Y").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            .Splits(0).DisplayColumns("Inception").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far

            For Each column As C1DisplayColumn In .Splits(0).DisplayColumns
                column.AutoSize()
                .Splits(0).DisplayColumns(column.Name).HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
            Next

        End With

    End Sub

    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles btnPDF.Click
        objReturn.LoadAt(objPortfolio, dtAs.Value)
        objBenchmark.LoadAt(objPortfolio, dtAs.Value)

        Dim pdf As New C1.C1Pdf.C1PdfDocument
        Dim pdf_width = pdf.PageRectangle.Width
        Dim pdf_height = pdf.PageRectangle.Height
        Dim sf As New StringFormat()
        sf.Alignment = StringAlignment.Far
        sf.LineAlignment = StringAlignment.Center
        Dim cc As New StringFormat()
        cc.Alignment = StringAlignment.Center
        cc.LineAlignment = StringAlignment.Center
        Dim cf As New StringFormat
        cf.Alignment = StringAlignment.Center
        cf.LineAlignment = StringAlignment.Near
        Dim urlImg = Image.FromFile("..\..\Template\Fund Sheet Sector EQ - Portrait.jpg")
        pdf.DrawString("Title", New Font("calibri", 40, FontStyle.Bold), Brushes.Black, New RectangleF(25, 10, 200, 25))
        pdf.DrawImage(urlImg, New RectangleF(pdf_width - 170, 2, 100, 60))
        pdf.DrawString("Tujuan Investasi", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 60, 175, 8))
        pdf.DrawLine(Pens.Black, 25, 70, 215, 70)
        pdf.DrawString("AVRIST IDX30 bertujuan untuk berinvestasi pada perusahaan dengan kapitalisasi saham besar, likuiditas tinggi, dan kondisi keuangan yang baik, yang masuk ke dalam Indeks IDX30 serta memberikan hasil investasi yang setara dengan kinerja Indeks IDX30", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 71, 190, 48))

        'Grafik Kinerja Reksa Dana
        Dim imageChart = chartPerformance.GetImage
        pdf.DrawString("Grafik Kinerja Reksa Dana", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(230, 60, 175, 8))
        pdf.DrawLine(Pens.Black, 230, 70, pdf_width - 73, 70)
        pdf.DrawString("29 Maret 2008", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(490, 50, 165, 8))
        pdf.DrawImage(imageChart, New RectangleF(230, 72, 325, 122))

        'Informasi Reksa Dana
        pdf.DrawString("Informasi Reksa Dana", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 125, 175, 8))
        pdf.DrawLine(Pens.Black, 25, 135, 215, 134)
        pdf.DrawString("Jenis Reksa Dana", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 136, 150, 8))
        pdf.DrawString("Tanggal Peluncuran", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 144, 150, 8))
        pdf.DrawString("Dana Kelolaan (Rp. Mil)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 152, 150, 8))
        pdf.DrawString("Mata Uang", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 160, 150, 8))
        pdf.DrawString("Frekuensi Valuasi", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 168, 150, 8))
        pdf.DrawString("Bank Kustodian", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 176, 150, 8))
        pdf.DrawString("Tolok Ukur", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 184, 150, 8))
        pdf.DrawString("NAB/Unit (Rp/Unit)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 192, 150, 8))
        'Value Informasi Reksa Dana
        pdf.DrawString("Ekultas", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 135, 150, 8), sf)
        pdf.DrawString("18-Dec-2017", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 143, 150, 8), sf)
        pdf.DrawString("135.92", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 151, 150, 8), sf)
        pdf.DrawString("Rupiah", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 159, 150, 8), sf)
        pdf.DrawString("Harian", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 167, 150, 8), sf)
        pdf.DrawString("Standard Chartered Bank", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 175, 150, 8), sf)
        pdf.DrawString("IDX30 Indeks", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 183, 150, 8), sf)
        pdf.DrawString("960.97", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 191, 150, 8), sf)

        'Investasi dan Biaya-Biaya
        pdf.DrawString("Investasi dan Biaya-Biaya", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 207, 175, 8))
        pdf.DrawLine(Pens.Black, 25, 217, 215, 217)
        pdf.DrawString("Minimal Investasi Awal (Rp)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 218, 150, 8))
        pdf.DrawString("Minimal Investasi Selanjutnya (Rp)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 226, 150, 8))
        pdf.DrawString("Biaya Pembelian (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 234, 150, 8))
        pdf.DrawString("Biaya Penjualan (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 242, 150, 8))
        pdf.DrawString("Biaya Jasa Pengelolaan Ml (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 250, 150, 8))
        pdf.DrawString("Biaya Jasa Bank Kustodian (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 258, 150, 8))

        'Value Investasi dan Biaya-Biaya
        pdf.DrawString("100.000", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 218, 150, 8), sf)
        pdf.DrawString("100.000", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 226, 150, 8), sf)
        pdf.DrawString("Maks 1.00", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 234, 150, 8), sf)
        pdf.DrawString("Maks 1.00", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 242, 150, 8), sf)
        pdf.DrawString("Maks 1.00", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 250, 150, 8), sf)
        pdf.DrawString("Maks 0.25", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 258, 150, 8), sf)

        'Kinerja Kumulatif
        pdf.DrawString("Kinerja Kumulatif (%)", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(230, 207, 175, 8))
        pdf.DrawLine(Pens.Black, 230, 217, pdf_width - 73, 217)
        pdf.DrawString("1 Bln", New Font("calibri", 8), Brushes.Black, New RectangleF(347, 219, 25, 8), cf)
        pdf.DrawString("3 Bln", New Font("calibri", 8), Brushes.Black, New RectangleF(372, 219, 25, 8), cf)
        pdf.DrawString("6 Bln", New Font("calibri", 8), Brushes.Black, New RectangleF(397, 219, 25, 8), cf)
        pdf.DrawString("YTD", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 219, 25, 8), cf)
        pdf.DrawString("1 Thn", New Font("calibri", 8), Brushes.Black, New RectangleF(447, 219, 25, 8), cf)
        pdf.DrawString("3 Thn", New Font("calibri", 8), Brushes.Black, New RectangleF(472, 219, 25, 8), cf)
        pdf.DrawString("5 Thn", New Font("calibri", 8), Brushes.Black, New RectangleF(497, 219, 25, 8), cf)
        pdf.DrawString("SP", New Font("calibri", 8), Brushes.Black, New RectangleF(522, 219, 25, 8), cf)
        pdf.DrawLine(Pens.Black, 230, 229, pdf_width - 73, 229)
        pdf.DrawString("Avrist IDX 30", New Font("calibri", 8), Brushes.Black, New RectangleF(232, 229, 175, 8))
        pdf.DrawString(objReturn.Getr1Mo.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(347, 229, 25, 8), cf)
        pdf.DrawString(objReturn.Getr3Mo.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(372, 229, 25, 8), cf)
        pdf.DrawString(objReturn.Getr6Mo.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(397, 229, 25, 8), cf)
        pdf.DrawString(objReturn.GetrYTD.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(422, 229, 25, 8), cf)
        pdf.DrawString(objReturn.Getr1Y.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(447, 229, 25, 8), cf)
        pdf.DrawString(objReturn.Getr3Y.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(472, 229, 25, 8), cf)
        pdf.DrawString(objReturn.Getr5Y.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(497, 229, 25, 8), cf)
        pdf.DrawString(objReturn.GetrInception, New Font("calibri", 8), Brushes.Black, New RectangleF(522, 229, 25, 8), cf)
        pdf.DrawString("Tolak Ukur", New Font("calibri", 8), Brushes.Black, New RectangleF(232, 237, 175, 8))
        pdf.DrawString(objBenchmark.Getr1Mo.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(347, 237, 25, 8), cf)
        pdf.DrawString(objBenchmark.Getr3Mo.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(372, 237, 25, 8), cf)
        pdf.DrawString(objBenchmark.Getr6Mo.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(397, 237, 25, 8), cf)
        pdf.DrawString(objBenchmark.GetrYTD.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(422, 237, 25, 8), cf)
        pdf.DrawString(objBenchmark.GetDate1Y.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(447, 237, 25, 8), cf)
        pdf.DrawString(objBenchmark.GetDate3Y.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(472, 237, 25, 8), cf)
        pdf.DrawString(objBenchmark.GetDate5Y.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(497, 237, 25, 8), cf)
        pdf.DrawString(objBenchmark.GetrInception.ToString, New Font("calibri", 8), Brushes.Black, New RectangleF(522, 237, 25, 8), cf)
        pdf.DrawLine(Pens.Black, 230, 248, pdf_width - 73, 248)
        pdf.DrawString("*SP : Sejak Peluncuran", New Font("calibri", 8), Brushes.Black, New RectangleF(256, 259, 175, 8))

        'Kebijakan Investasi
        pdf.DrawString("Kebijakan Investasi", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(230, 274, 175, 8))
        pdf.DrawString("Komposisi Portofolio", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(422, 274, 175, 8))
        pdf.DrawLine(Pens.Black, 230, 274, pdf_width - 73, 274)
        pdf.DrawString("%", New Font("calibri", 8), Brushes.Black, New RectangleF(372, 275, 15, 8), sf)
        pdf.DrawString("%", New Font("calibri", 8), Brushes.Black, New RectangleF(pdf_width - 90, 275, 15, 8), sf)
        pdf.DrawLine(Pens.Black, 230, 285, pdf_width - 73, 285)
        pdf.DrawString("Ekuitas", New Font("calibri", 8), Brushes.Black, New RectangleF(230, 286, 175, 8))
        pdf.DrawString("Ekuitas", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 286, 175, 8))
        pdf.DrawString("80 - 100", New Font("calibri", 8), Brushes.Black, New RectangleF(287, 286, 100, 8), sf)
        pdf.DrawString("99.56", New Font("calibri", 8), Brushes.Black, New RectangleF(pdf_width - 175, 286, 100, 8), sf)
        pdf.DrawString("Pasar Ulang", New Font("calibri", 8), Brushes.Black, New RectangleF(230, 294, 175, 8))
        pdf.DrawString("Pasar Ulang", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 294, 175, 8))
        pdf.DrawString("0 - 20", New Font("calibri", 8), Brushes.Black, New RectangleF(287, 294, 100, 8), sf)
        pdf.DrawString("0.44", New Font("calibri", 8), Brushes.Black, New RectangleF(pdf_width - 175, 294, 100, 8), sf)
        pdf.DrawLine(Pens.Black, 230, 306, pdf_width - 73, 306)

        'Statistik Reksa Dana
        pdf.DrawString("Statistik Reksadana", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 274, 175, 8))
        pdf.DrawLine(Pens.Black, 25, 284, 215, 284)
        pdf.DrawString("Kinerja Sejak Diluncurkan (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 285, 150, 8))
        pdf.DrawString("Standar Deviasi Disetahunkan (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 293, 150, 8))
        pdf.DrawString("Beta", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 301, 150, 8))
        pdf.DrawString("Kinerja Bulanan Terbaik (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 309, 150, 8))
        pdf.DrawString("Kinerja Bulanan Terburuk (%)", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 317, 150, 8))
        pdf.DrawString("Kinerja terbaik setahun terakhir (%)", New Font("calibri", 8, FontStyle.Bold), Brushes.Black, New RectangleF(25, 325, 150, 8))

        '5 Besar Efek Dalam Portofolio
        pdf.DrawString("5 Besar Efek Dalam Portofolio", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(230, 325, 150, 8))
        pdf.DrawLine(Pens.Black, 230, 335, pdf_width - 73, 335)
        'Efek
        pdf.DrawString("Efek", New Font("calibri", 8), Brushes.Black, New RectangleF(230, 341, 175, 8))
        pdf.DrawLine(Pens.Black, 230, 351, pdf_width - 73, 351)
        pdf.DrawString("Bank Central Asia Tbk.", New Font("calibri", 8), Brushes.Black, New RectangleF(230, 352, 150, 8))
        pdf.DrawString("Bank Rakyat Indonesia (Persero Tbk.)", New Font("calibri", 8), Brushes.Black, New RectangleF(230, 360, 150, 8))
        pdf.DrawString("H.M. Sampoerna Tbk.", New Font("calibri", 8), Brushes.Black, New RectangleF(230, 368, 150, 8))
        pdf.DrawString("Telekomunikasi Indonesia (Persero) Tbk.", New Font("calibri", 8), Brushes.Black, New RectangleF(230, 376, 150, 8))
        pdf.DrawString("Unilever Indonesia Tbk.", New Font("calibri", 8), Brushes.Black, New RectangleF(230, 384, 150, 8))
        'Sektor
        pdf.DrawString("Sektor", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 341, 175, 8))
        pdf.DrawString("Keuangan", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 352, 150, 8))
        pdf.DrawString("Keuangan", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 360, 150, 8))
        pdf.DrawString("Konsumer Kebutuhan Pokok", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 368, 150, 8))
        pdf.DrawString("Jasa Telekomunikasi", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 376, 150, 8))
        pdf.DrawString("Konsumer Kebutuhan Pokok", New Font("calibri", 8), Brushes.Black, New RectangleF(422, 384, 150, 8))

        'Value Statistik Reksadana
        pdf.DrawString("(3.90)", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 285, 150, 8), sf)
        pdf.DrawString("18.86", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 293, 150, 8), sf)
        pdf.DrawString("0.97", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 301, 150, 8), sf)
        pdf.DrawString("4.15", New Font("calibri", 8), Brushes.Black, New RectangleF(100, 309, 125, 8), cf)
        pdf.DrawString("(8.51)", New Font("calibri", 8), Brushes.Black, New RectangleF(100, 317, 125, 8), cf)
        pdf.DrawString("Dec-17", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 309, 150, 8), sf)
        pdf.DrawString("Mar-18", New Font("calibri", 8), Brushes.Black, New RectangleF(65, 317, 150, 8), sf)
        pdf.DrawString("4.15", New Font("calibri", 8, FontStyle.Bold), Brushes.Black, New RectangleF(65, 325, 150, 8), sf)

        'Risiko Investasi
        pdf.DrawString("Risiko Investasi", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 341, 175, 8))
        pdf.DrawLine(Pens.Black, 25, 351, 215, 351)
        pdf.DrawString("1. Risiko Perubahan Kondisi Ekonomi dan Politik", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 352, 150, 8))
        pdf.DrawString("2. Risiko Pasar", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 360, 150, 8))
        pdf.DrawString("3. Risiko Kredit/Wanprestasi", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 368, 150, 8))
        pdf.DrawString("4. Risiko Likuiditas", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 376, 150, 8))
        pdf.DrawString("5. Risiko Perubahan Peraturan", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 384, 150, 8))
        pdf.DrawString("6. Risiko Berkurangnya Nilai Aktiva Bersih", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 392, 150, 8))
        pdf.DrawString("7. Risiko Pembubaran dan Likuidasi", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 400, 150, 8))

        'Alokasi Sektoral(%)
        Dim imgPie = chartSector.GetImage
        pdf.DrawString("Alokasi Sektoral (%)", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(230, 400, 150, 8))
        pdf.DrawLine(Pens.Black, 230, 410, pdf_width - 73, 410)
        pdf.DrawImage(imgPie, New RectangleF(230, 412, 300, 95))
        'Klasifikasi Risiko
        pdf.DrawString("Klasifikasi Risiko", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 416, 175, 8))
        pdf.DrawLine(Pens.Black, 25, 426, 215, 426)
        pdf.FillRectangle(Brushes.AliceBlue, New RectangleF(25, 427, 35, 10))
        pdf.DrawString("1", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 427, 35, 8), cf)
        pdf.FillRectangle(Brushes.AliceBlue, New RectangleF(61, 427, 35, 10))
        pdf.DrawString("2", New Font("calibri", 8), Brushes.Black, New RectangleF(61, 427, 35, 8), cf)
        pdf.FillRectangle(Brushes.AliceBlue, New RectangleF(97, 427, 35, 10))
        pdf.DrawString("3", New Font("calibri", 8), Brushes.Black, New RectangleF(97, 427, 35, 8), cf)
        pdf.FillRectangle(Brushes.AliceBlue, New RectangleF(133, 427, 35, 10))
        pdf.DrawString("4", New Font("calibri", 8), Brushes.Black, New RectangleF(133, 427, 35, 8), cf)
        pdf.FillRectangle(Brushes.Purple, New RectangleF(169, 427, 35, 10))
        pdf.DrawString("5", New Font("calibri", 8), Brushes.White, New RectangleF(169, 427, 35, 8), cf)

        pdf.DrawString("Mengenal Manajer Investasi", New Font("calibri", 8, FontStyle.Bold), Brushes.Purple, New RectangleF(25, 445, 175, 8))
        pdf.DrawLine(Pens.Black, 25, 455, 215, 455)
        pdf.DrawString(
            "PT Avrist Asset Management merupakan anak perusahaan dari PT Avrist Assurance ('Avrist'). PT. Avrist Asset Management didukung oleh professional yang berpengalaman di bidang investasi dan menawarkan beragam solusi investasi yang disesuaikan dengan kondisi pasar dan tujuan investasi pemodal.", New Font("calibri", 8), Brushes.Black, New RectangleF(25, 456, 190, 54))
        pdf.DrawLine(Pens.Black, 225, 70, 225, 506)

        pdf.DrawRectangle(Pens.Black, New RectangleF(25, 510, pdf_width - 98, 78))
        pdf.DrawString("Akhir Maret, IHSG ditutup turun -6.2% (mom) ke level 6.188 sedangkan indeks LQ45 ditutup turun -8.6% (mom) ke level 1.005. Selama Maret, investor asing melakukan penjualan bersih sebesar Rp. 14.9 tn. Volatilitas rupiah yang disebabkan ketidakpastian kebijakan moneter ditambah dengan sentiment kekhawatiran peranf dagang menyeret indeks ditutup", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 514, pdf_width - 102, 7))
        pdf.DrawString("bersih sebesar Rp. 14.9 tn. Volatilitas rupiah yang disebabkan ketidakpastian kebijakan moneter ditambah dengan sentiment kekhawatiran perang dagang menyeret indeks ditutup pada teritori negatif.", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 526, pdf_width - 100, 7))
        pdf.DrawString("pada teritori negatif.", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 538, pdf_width - 100, 7))
        pdf.DrawString("Sementara itu pasar saham global mayoritas ditutup turun (S&P 500-2.7%, FTSE 100-2.4%, Nikkel 225-4.1% mom). Diawali ketidakpastian landscape kebijakan moneter US ", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 550, pdf_width - 100, 7))
        pdf.DrawString("dan memanasnya suhu politik di US, bursa global kembali tertekan setelah US mengesahkan trade protectionism program yang menyeret mayoritas bursa global pada teritori", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 562, pdf_width - 100, 7))
        pdf.DrawString("negatif.", New Font("calibri", 7), Brushes.Black, New RectangleF(27, 574, pdf_width - 100, 7))

        pdf.DrawRectangle(Pens.Black, New RectangleF(25, 592, pdf_width - 98, 29))
        pdf.DrawString("INVESTASI MELALUI REKSA DANA MENGANDUNG RISIKO. CALON INVESTOR WAJIB DAN MEMAHAMI PROSPEKTUS SEBELUM MEMUTUSKAN UNTUK BERINVESTASI MELALUI REKSA DANA. KINERJA MASA LALU TIDAK MENCERMINKAN KINERJA MASA DATANG. PT. AVRIST ASSET MANAGEMENT TELAH MEMILIKI IZIN USAHA, TERDAFTAR DAN DIAWASI OLEH OTORITAS JASA KEUANGAN", New Font("calibri", 7, FontStyle.Bold), Brushes.Black, New RectangleF(25, 592, pdf_width - 98, 50), cf)

        pdf.FillRectangle(Brushes.Purple, New RectangleF(25, 626, pdf_width - 98, 67))
        pdf.DrawStringRtf("{\b \i Disclaimer :}", New Font("calibri", 8), Brushes.White, New RectangleF((pdf_width - 98) / 2, 628, pdf_width - 98, 8))
        pdf.DrawString("Laporan ini disajikan oleh PT. Avrist Asset Management hanya untuk tujuan informasi dan tidak dapat digunakan atau dijadikan dasar sebagai penawaran atau rekomendasi untuk menjual atau membeli. Laporan ini dibuat berdasarkan keadaan yang telah terjadi dan telah disusun secara seksama oleh PT. Avrist Asset Management meskipun demikian PT. Avrist Asset Management tidak menjamin keakuratan atau kelengkapan dari laporan tersebut. PT. Avrist Asset Management maupun officer atau karyawannya tidak bertanggung jawab apapun terhadap setiap kerugian yang timbul baik langsung maupun tidak langsung sebagai akibat dari setiap penggunaan laporan ini. Setiap keputusan investasi haruslah merupakan keputusan individu, sehingga tanggung jawabnya ada pada masing-masing individu yang membuat keputusan investasi tersebut. Kinerja masa lalu tidak mencerminkan kinerja masa mendatang. Calon pemodal wajib memahami risiko berinvestasi di Pasar Modal oleh sebab itu calon pemodal wajib membaca dan memahami isi Prospektus sebelum memutsukan untuk berinvestasi.", New Font("calibri", 6), Brushes.White, New RectangleF(25, 640, pdf_width - 98, 60), cf)

        pdf.DrawStringRtf("{\b PT AVRIST ASSET MANAGEMENT} : Gedung WTC 5 Lt.9", New Font("calibiri", 6), Brushes.Black, New RectangleF(25, 706, pdf_width - 98, 8))
        pdf.DrawString("Jl. Jend. Sudirman Kav. 29, Jakarta 12920 | t +62 21 252 1662, f +62 21 252 2106 |", New Font("calibiri", 6), Brushes.Black, New RectangleF(25, 714, pdf_width - 98, 8))
        pdf.DrawString("CS.AVRAM@Avrist.com", New Font("calibiri", 6, FontStyle.Bold), Brushes.Black, New RectangleF(25, 722, pdf_width - 98, 8))

        pdf.DrawLine(Pens.Black, (pdf_width / 2) - 25, 706, (pdf_width / 2) - 25, 746)

        pdf.DrawStringRtf("NOMOR REKENING:", New Font("calibiri", 6, FontStyle.Bold), Brushes.Black, New RectangleF(pdf_width / 2, 706, pdf_width - 98, 8))
        pdf.DrawString("REKSA DANA INDEKS AVRIST IDX30", New Font("calibiri", 6, FontStyle.Bold), Brushes.Black, New RectangleF(pdf_width / 2, 722, pdf_width - 98, 8))
        pdf.DrawString("Standard Chartered Bank", New Font("calibiri", 6), Brushes.Black, New RectangleF(pdf_width / 2, 730, pdf_width - 98, 8))
        pdf.DrawString("A/C # 306-8112029-8", New Font("calibiri", 6, FontStyle.Bold), Brushes.Black, New RectangleF(pdf_width / 2, 738, pdf_width - 98, 8))


        pdf.Save("D:\test.pdf")
        'Process.Start("D:\test.pdf")
    End Sub

    Structure pdfColor

        Public LayoutType As String

        Public ReportLine_R As Integer
        Public ReportLine_G As Integer
        Public ReportLine_B As Integer

        Public Tanggal_R As Integer
        Public Tanggal_G As Integer
        Public Tanggal_B As Integer
        Public Tanggal As String

        Public TujuanInvestasi_R As Integer
        Public TujuanInvestasi_G As Integer
        Public TujuanInvestasi_B As Integer
        Public TujuanInvestasi As String

        Public ValueTujuanInvestasi_R As Integer
        Public ValueTujuanInvestasi_G As Integer
        Public ValueTujuanInvestasi_B As Integer

        Public InformasiReksaDana_R As Integer
        Public InformasiReksaDana_G As Integer
        Public InformasiReksaDana_B As Integer
        Public InformasiReksaDana As String

        Public ValueInformasiReksaDana_R As Integer
        Public ValueInformasiReksaDana_G As Integer
        Public ValueInformasiReksaDana_B As Integer

        Public IIRD1 As String
        Public IIRD2 As String
        Public IIRD3 As String
        Public IIRD4 As String
        Public IIRD5 As String
        Public IIRD6 As String
        Public IIRD7 As String
        Public IIRD8 As String

        Public InvestasiDanaBiayaBiaya_R As Integer
        Public InvestasiDanaBiayaBiaya_G As Integer
        Public InvestasiDanaBiayaBiaya_B As Integer
        Public InvestasiDanaBiayaBiaya As String

        Public ValueInvestasiDanaBiayaBiaya_R As Integer
        Public ValueInvestasiDanaBiayaBiaya_G As Integer
        Public ValueInvestasiDanaBiayaBiaya_B As Integer

        Public IIBB1 As String
        Public IIBB2 As String
        Public IIBB3 As String
        Public IIBB4 As String
        Public IIBB5 As String
        Public IIBB6 As String

        Public StatistikReksadana_R As Integer
        Public StatistikReksadana_G As Integer
        Public StatistikReksadana_B As Integer
        Public StatistikReksadana As String

        Public ValueStatistikReksadana_R As Integer
        Public ValueStatistikReksadana_G As Integer
        Public ValueStatistikReksadana_B As Integer

        Public ISR1 As String
        Public ISR2 As String
        Public ISR3 As String
        Public ISR4 As String
        Public ISR5 As String
        Public ISR6 As String

        Public RisikoInvestasi_R As Integer
        Public RisikoInvestasi_G As Integer
        Public RisikoInvestasi_B As Integer
        Public RisikoInvestasi As String

        Public ValueRisikoInvestasi_R As Integer
        Public ValueRisikoInvestasi_G As Integer
        Public ValueRisikoInvestasi_B As Integer

        Public IRI1 As String
        Public IRI2 As String
        Public IRI3 As String
        Public IRI4 As String
        Public IRI5 As String
        Public IRI6 As String
        Public IRI7 As String

        Public KlasifikasiRisiko_R As Integer
        Public KlasifikasiRisiko_G As Integer
        Public KlasifikasiRisiko_B As Integer
        Public KlasifikasiRisiko As String

        Public MengenaiManajerInvestasi_R As Integer
        Public MengenaiManajerInvestasi_G As Integer
        Public MengenaiManajerInvestasi_B As Integer
        Public MengenaiManajerInvestasi As String

        Public ValueMengenaiManajerInvestasi_R As Integer
        Public ValueMengenaiManajerInvestasi_G As Integer
        Public ValueMengenaiManajerInvestasi_B As Integer
        Public ValueMengenaiManajerInvestasi As String

        Public GrafikKinerja_R As Integer
        Public GrafikKinerja_G As Integer
        Public GrafikKinerja_B As Integer
        Public GrafikKinerja As String

        Public KinerjaKumulatif_R As Integer
        Public KinerjaKumulatif_G As Integer
        Public KinerjaKumulatif_B As Integer
        Public KinerjaKumulatif As String

        Public ValueKinerjaKumulatif_R As Integer
        Public ValueKinerjaKumulatif_G As Integer
        Public ValueKinerjaKumulatif_B As Integer

        Public KebijakanInvestasi_R As Integer
        Public KebijakanInvestasi_G As Integer
        Public KebijakanInvestasi_B As Integer
        Public KebijakanInvestasi As String

        Public KomposisiPortofolio_R As Integer
        Public KomposisiPortofolio_G As Integer
        Public KomposisiPortofolio_B As Integer
        Public KomposisiPortofolio As String

        Public AlokasiSektoral_R As Integer
        Public AlokasiSektoral_G As Integer
        Public AlokasiSektoral_B As Integer
        Public AlokasiSektoral As String

        Public BEDP_R As Integer
        Public BEDP_G As Integer
        Public BEDP_B As Integer
        Public BEDP As String

        Public IBEDP_R As Integer
        Public IBEDP_G As Integer
        Public IBEDP_B As Integer

        Public BorderSatu_R As Integer
        Public BorderSatu_G As Integer
        Public BorderSatu_B As Integer

        Public ValueBorderSatu_R As Integer
        Public ValueBorderSatu_G As Integer
        Public ValueBorderSatu_B As Integer

        Public BorderDua_R As Integer
        Public BorderDua_G As Integer
        Public BorderDua_B As Integer

        Public ValueBorderDua_R As Integer
        Public ValueBorderDua_G As Integer
        Public ValueBorderDua_B As Integer

        Public FillSatu_R As Integer
        Public FillSatu_G As Integer
        Public FillSatu_B As Integer

        Public ValueFillSatu_R As Integer
        Public ValueFillSatu_G As Integer
        Public ValueFillSatu_B As Integer

        Public ChartTitle_R As Integer
        Public ChartTitle_G As Integer
        Public ChartTitle_B As Integer
        Public ChartTitle As String
        Public ChartAxisX As String
        Public ChartAxisY As String

        Public ChartLine_R As Integer
        Public ChartLine_G As Integer
        Public ChartLine_B As Integer

        Public ChartBorder_R As Integer
        Public ChartBorder_G As Integer
        Public ChartBorder_B As Integer
        Public ChartBorder As Boolean

        Public TableHeader_R As Integer
        Public TableHeader_G As Integer
        Public TableHeader_B As Integer

        Public TableItem_R As Integer
        Public TableItem_G As Integer
        Public TableItem_B As Integer

        Public UlasanPasar_R As Integer
        Public UlasanPasar_G As Integer
        Public UlasanPasar_B As Integer
        Public UlasanPasar As String
    End Structure

    Public Sub pdfColorDefault()
        pdfLayout.LayoutType = "DEFAULT"

        pdfLayout.ReportLine_R = 0
        pdfLayout.ReportLine_G = 0
        pdfLayout.ReportLine_B = 0

        pdfLayout.Tanggal_R = 0
        pdfLayout.Tanggal_G = 61
        pdfLayout.Tanggal_B = 121

        pdfLayout.TujuanInvestasi_R = 0
        pdfLayout.TujuanInvestasi_G = 61
        pdfLayout.TujuanInvestasi_B = 121
        pdfLayout.TujuanInvestasi = "Tujuan Investasi"

        pdfLayout.ValueTujuanInvestasi_R = 0
        pdfLayout.ValueTujuanInvestasi_G = 61
        pdfLayout.ValueTujuanInvestasi_B = 121

        pdfLayout.InformasiReksaDana_R = 0
        pdfLayout.InformasiReksaDana_G = 61
        pdfLayout.InformasiReksaDana_B = 121
        pdfLayout.InformasiReksaDana = "Informasi Reksa Dana"

        pdfLayout.IIRD1 = "Jenis Reksa Dana"
        pdfLayout.IIRD2 = "Tanggal Peluncuran"
        pdfLayout.IIRD3 = "Dana Kelolaan (Rp Mil)"
        pdfLayout.IIRD4 = "Mata Uang"
        pdfLayout.IIRD5 = "Frekuensi Valuasi"
        pdfLayout.IIRD6 = "Bank Kustodian"
        pdfLayout.IIRD7 = "Tolok Ukur"
        pdfLayout.IIRD8 = "NAB/Unit (Rp/Unit)"

        pdfLayout.ValueInformasiReksaDana_R = 0
        pdfLayout.ValueInformasiReksaDana_G = 61
        pdfLayout.ValueInformasiReksaDana_B = 121

        pdfLayout.InvestasiDanaBiayaBiaya_R = 0
        pdfLayout.InvestasiDanaBiayaBiaya_G = 61
        pdfLayout.InvestasiDanaBiayaBiaya_B = 121
        pdfLayout.InvestasiDanaBiayaBiaya = "Investasi dan Biaya-Biaya"

        pdfLayout.ValueInvestasiDanaBiayaBiaya_R = 0
        pdfLayout.ValueInvestasiDanaBiayaBiaya_G = 61
        pdfLayout.ValueInvestasiDanaBiayaBiaya_B = 121
        pdfLayout.IIBB1 = "Minimal Investasi Awal (Rp)"
        pdfLayout.IIBB2 = "Standar Deviasi Disetahunkan (%)"
        pdfLayout.IIBB3 = "Biaya Pembelian (%)"
        pdfLayout.IIBB4 = "Biaya Penjualan (%)"
        pdfLayout.IIBB5 = "Biaya Jasa Pengelolaan MI (Rp)"
        pdfLayout.IIBB6 = "Biaya Jasa Bank Kustodian (Rp)"

        pdfLayout.StatistikReksadana_R = 0
        pdfLayout.StatistikReksadana_G = 61
        pdfLayout.StatistikReksadana_B = 121
        pdfLayout.StatistikReksadana = "Statistik Reksadana"

        pdfLayout.ValueStatistikReksadana_R = 0
        pdfLayout.ValueStatistikReksadana_G = 61
        pdfLayout.ValueStatistikReksadana_B = 121
        pdfLayout.ISR1 = "Kinerja Sejak Diluncurkan (%)"
        pdfLayout.ISR2 = "Standar Deviasi  Disetahunkan (%)"
        pdfLayout.ISR3 = "Beta (%)"
        pdfLayout.ISR4 = "Kinerja Bulanan Terbaik (%)"
        pdfLayout.ISR5 = "Kinerja Bulanan Terburuk (%)"
        pdfLayout.ISR6 = "Biaya Jasa Bank Kustodian (%)"

        pdfLayout.RisikoInvestasi_R = 0
        pdfLayout.RisikoInvestasi_G = 61
        pdfLayout.RisikoInvestasi_B = 121
        pdfLayout.RisikoInvestasi = "Risiko Investasi"

        pdfLayout.IRI1 = "Risiko Perubahan Kondisi Ekonomi dan Politik"
        pdfLayout.IRI2 = "Risiko Pasar"
        pdfLayout.IRI3 = "Risiko Kredit/wanprestasi"
        pdfLayout.IRI4 = "Risiko Likuiditas"
        pdfLayout.IRI5 = "Risiko Perubahan Peraturan"
        pdfLayout.IRI6 = "Risiko Berkurangnya Nilai Aktiva Bersih"
        pdfLayout.IRI7 = "Risiko Pembubaran dan Likuidasi"

        pdfLayout.ValueRisikoInvestasi_R = 0
        pdfLayout.ValueRisikoInvestasi_G = 61
        pdfLayout.ValueRisikoInvestasi_B = 121

        pdfLayout.KlasifikasiRisiko_R = 0
        pdfLayout.KlasifikasiRisiko_G = 61
        pdfLayout.KlasifikasiRisiko_B = 121

        pdfLayout.MengenaiManajerInvestasi_R = 0
        pdfLayout.MengenaiManajerInvestasi_G = 61
        pdfLayout.MengenaiManajerInvestasi_B = 121
        pdfLayout.MengenaiManajerInvestasi = "Mengenal Manajer Investasi"

        pdfLayout.ValueMengenaiManajerInvestasi_R = 0
        pdfLayout.ValueMengenaiManajerInvestasi_G = 61
        pdfLayout.ValueMengenaiManajerInvestasi_B = 121
        pdfLayout.ValueMengenaiManajerInvestasi = "PT Avrist Asset Management merupakan anak perusahaan dari PT Avrist Assurance ('Avrist'). PT. Avrist Asset Management didukung oleh professional yang berpengalaman di bidang investasi dan menawarkan beragam solusi investasi yang disesuaikan dengan kondisi pasar dan tujuan investasi pemodal."

        pdfLayout.GrafikKinerja_R = 0
        pdfLayout.GrafikKinerja_G = 61
        pdfLayout.GrafikKinerja_B = 121
        pdfLayout.GrafikKinerja = "Grafik Kinerja"

        pdfLayout.KinerjaKumulatif_R = 0
        pdfLayout.KinerjaKumulatif_G = 61
        pdfLayout.KinerjaKumulatif_B = 121
        pdfLayout.KinerjaKumulatif = "Kinerja Kumulatif"

        pdfLayout.KebijakanInvestasi_R = 0
        pdfLayout.KebijakanInvestasi_G = 61
        pdfLayout.KebijakanInvestasi_B = 121
        pdfLayout.KebijakanInvestasi = "Kebijakan Investasi"

        pdfLayout.KomposisiPortofolio_R = 0
        pdfLayout.KomposisiPortofolio_G = 61
        pdfLayout.KomposisiPortofolio_B = 121
        pdfLayout.KomposisiPortofolio = "Komposisi Portofolio"

        pdfLayout.AlokasiSektoral_R = 0
        pdfLayout.AlokasiSektoral_G = 61
        pdfLayout.AlokasiSektoral_B = 121
        pdfLayout.AlokasiSektoral = "Alokasi Sektoral"

        pdfLayout.KlasifikasiRisiko_R = 0
        pdfLayout.KlasifikasiRisiko_G = 61
        pdfLayout.KlasifikasiRisiko_B = 121
        pdfLayout.KlasifikasiRisiko = "Klasifikasi Risiko"

        pdfLayout.BEDP_R = 0
        pdfLayout.BEDP_G = 61
        pdfLayout.BEDP_B = 121
        pdfLayout.BEDP = "5 Besar Efek Dalam Portofolio"

        pdfLayout.ChartTitle_R = 0
        pdfLayout.ChartTitle_G = 61
        pdfLayout.ChartTitle_B = 121
        pdfLayout.ChartTitle = "Grafik Kinerja Reksa Dana"

        pdfLayout.ChartBorder_R = 0
        pdfLayout.ChartBorder_G = 61
        pdfLayout.ChartBorder_B = 121

        pdfLayout.ChartLine_R = 0
        pdfLayout.ChartLine_G = 61
        pdfLayout.ChartLine_B = 121

        pdfLayout.BorderSatu_R = 0
        pdfLayout.BorderSatu_G = 61
        pdfLayout.BorderSatu_B = 121

        pdfLayout.ValueBorderSatu_R = 0
        pdfLayout.ValueBorderSatu_G = 61
        pdfLayout.ValueBorderSatu_B = 121

        pdfLayout.BorderDua_R = 0
        pdfLayout.BorderDua_G = 61
        pdfLayout.BorderDua_B = 121

        pdfLayout.ValueBorderDua_R = 0
        pdfLayout.ValueBorderDua_G = 61
        pdfLayout.ValueBorderDua_B = 121

        pdfLayout.FillSatu_R = 0
        pdfLayout.FillSatu_G = 61
        pdfLayout.FillSatu_B = 121

        pdfLayout.ValueFillSatu_R = 0
        pdfLayout.ValueFillSatu_G = 61
        pdfLayout.ValueFillSatu_B = 121

        pdfLayout.UlasanPasar_R = 0
        pdfLayout.UlasanPasar_G = 0
        pdfLayout.UlasanPasar_B = 0
        pdfLayout.UlasanPasar = "Akhir Maret, IHSG ditutup turun -6.2% (mom) ke level 6.188 sedangkan indeks LQ45 ditutup turun -8.6% (mom) ke level 1.005. Selama Maret, investor asing melakukan penjualan bersih sebesar Rp. 14.9 tn. Volatilitas rupiah yang disebabkan ketidakpastian kebijakan moneter ditambah dengan sentiment kekhawatiran peranf dagang menyeret indeks ditutup bersih sebesar Rp. 14.9 tn. Volatilitas rupiah yang disebabkan ketidakpastian kebijakan moneter ditambah dengan sentiment kekhawatiran perang dagang menyeret indeks ditutup pada teritori negatif.  pada teritori negatif. Sementara itu pasar saham global mayoritas ditutup turun (S&P 500-2.7%, FTSE 100-2.4%, Nikkel 225-4.1% mom). Diawali ketidakpastian landscape kebijakan moneter US dan memanasnya suhu politik di US, bursa global kembali tertekan setelah US mengesahkan trade protectionism program yang menyeret mayoritas bursa global pada teritori negatif."
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
                    pdfLayout.LayoutType = iniType
                    pdfLayout.TujuanInvestasi_R = file.GetInteger(reportSection, iniType & " Tujuan Investasi R", 0)
                    pdfLayout.TujuanInvestasi_G = file.GetInteger(reportSection, iniType & " Tujuan Investasi G", 0)
                    pdfLayout.TujuanInvestasi_B = file.GetInteger(reportSection, iniType & " Tujuan Investasi B", 0)
                    pdfLayout.TujuanInvestasi = file.GetString(reportSection, iniType & " Tujuan Investasi", "")

                    pdfLayout.Tanggal_R = file.GetInteger(reportSection, iniType & " Tanggal R", 0)
                    pdfLayout.Tanggal_G = file.GetInteger(reportSection, iniType & " Tanggal G", 0)
                    pdfLayout.Tanggal_B = file.GetInteger(reportSection, iniType & " Tanggal B", 0)

                    pdfLayout.ReportLine_R = file.GetInteger(reportSection, iniType & " Report Line R", 0)
                    pdfLayout.ReportLine_G = file.GetInteger(reportSection, iniType & " Report Line G", 0)
                    pdfLayout.ReportLine_B = file.GetInteger(reportSection, iniType & " Report Line B", 0)

                    pdfLayout.InformasiReksaDana_R = file.GetInteger(reportSection, iniType & " IRD R", 0)
                    pdfLayout.InformasiReksaDana_G = file.GetInteger(reportSection, iniType & " IRD G", 0)
                    pdfLayout.InformasiReksaDana_B = file.GetInteger(reportSection, iniType & " IRD B", 0)
                    pdfLayout.InformasiReksaDana = file.GetString(reportSection, iniType & " IRD", "")

                    pdfLayout.ValueInformasiReksaDana_R = file.GetInteger(reportSection, iniType & " IIRD R", 0)
                    pdfLayout.ValueInformasiReksaDana_G = file.GetInteger(reportSection, iniType & " IIRD R", 0)
                    pdfLayout.ValueInformasiReksaDana_B = file.GetInteger(reportSection, iniType & " IIRD R", 0)

                    pdfLayout.IIRD1 = file.GetString(reportSection, iniType & " IIRD1", "")
                    pdfLayout.IIRD2 = file.GetString(reportSection, iniType & " IIRD2", "")
                    pdfLayout.IIRD3 = file.GetString(reportSection, iniType & " IIRD3", "")
                    pdfLayout.IIRD4 = file.GetString(reportSection, iniType & " IIRD4", "")
                    pdfLayout.IIRD5 = file.GetString(reportSection, iniType & " IIRD5", "")
                    pdfLayout.IIRD6 = file.GetString(reportSection, iniType & " IIRD6", "")
                    pdfLayout.IIRD7 = file.GetString(reportSection, iniType & " IIRD7", "")
                    pdfLayout.IIRD8 = file.GetString(reportSection, iniType & " IIRD8", "")

                    pdfLayout.KinerjaKumulatif_R = file.GetInteger(reportSection, iniType & " KK R", 0)
                    pdfLayout.KinerjaKumulatif_G = file.GetInteger(reportSection, iniType & " KK G", 0)
                    pdfLayout.KinerjaKumulatif_B = file.GetInteger(reportSection, iniType & " KK B", 0)
                    pdfLayout.KinerjaKumulatif = file.GetString(reportSection, iniType & " KK", "")

                    pdfLayout.KebijakanInvestasi_R = file.GetInteger(reportSection, iniType & " KI R", 0)
                    pdfLayout.KebijakanInvestasi_G = file.GetInteger(reportSection, iniType & " KI G", 0)
                    pdfLayout.KebijakanInvestasi_B = file.GetInteger(reportSection, iniType & " KI B", 0)
                    pdfLayout.KebijakanInvestasi = file.GetString(reportSection, iniType & " KI", "")

                    pdfLayout.AlokasiSektoral_R = file.GetInteger(reportSection, iniType & " AS R", 0)
                    pdfLayout.AlokasiSektoral_G = file.GetInteger(reportSection, iniType & " AS G", 0)
                    pdfLayout.AlokasiSektoral_B = file.GetInteger(reportSection, iniType & " AS B", 0)
                    pdfLayout.AlokasiSektoral = file.GetString(reportSection, iniType & " AS", "")

                    pdfLayout.InvestasiDanaBiayaBiaya_R = file.GetInteger(reportSection, iniType & " IBB R", 0)
                    pdfLayout.InvestasiDanaBiayaBiaya_G = file.GetInteger(reportSection, iniType & " IBB G", 0)
                    pdfLayout.InvestasiDanaBiayaBiaya_B = file.GetInteger(reportSection, iniType & " IBB B", 0)
                    pdfLayout.InvestasiDanaBiayaBiaya = file.GetString(reportSection, iniType & " IBB", "")

                    pdfLayout.IIBB1 = file.GetString(reportSection, iniType & " IIBB1", "")
                    pdfLayout.IIBB2 = file.GetString(reportSection, iniType & " IIBB2", "")
                    pdfLayout.IIBB3 = file.GetString(reportSection, iniType & " IIBB3", "")
                    pdfLayout.IIBB4 = file.GetString(reportSection, iniType & " IIBB4", "")
                    pdfLayout.IIBB5 = file.GetString(reportSection, iniType & " IIBB5", "")
                    pdfLayout.IIBB6 = file.GetString(reportSection, iniType & " IIBB6", "")

                    pdfLayout.StatistikReksadana_R = file.GetInteger(reportSection, iniType & " SR R", 0)
                    pdfLayout.StatistikReksadana_G = file.GetInteger(reportSection, iniType & " SR G", 0)
                    pdfLayout.StatistikReksadana_B = file.GetInteger(reportSection, iniType & " SR B", 0)
                    pdfLayout.StatistikReksadana = file.GetString(reportSection, iniType & " SR", "")

                    pdfLayout.ISR1 = file.GetString(reportSection, iniType & " ISR1", "")
                    pdfLayout.ISR2 = file.GetString(reportSection, iniType & " ISR2", "")
                    pdfLayout.ISR3 = file.GetString(reportSection, iniType & " ISR3", "")
                    pdfLayout.ISR4 = file.GetString(reportSection, iniType & " ISR4", "")
                    pdfLayout.ISR5 = file.GetString(reportSection, iniType & " ISR5", "")
                    pdfLayout.ISR6 = file.GetString(reportSection, iniType & " ISR6", "")

                    pdfLayout.KomposisiPortofolio_R = file.GetInteger(reportSection, iniType & " KP R", 0)
                    pdfLayout.KomposisiPortofolio_G = file.GetInteger(reportSection, iniType & " KP G", 0)
                    pdfLayout.KomposisiPortofolio_B = file.GetInteger(reportSection, iniType & " KP B", 0)
                    pdfLayout.KomposisiPortofolio = file.GetString(reportSection, iniType & " KP", "")

                    pdfLayout.BEDP_R = file.GetInteger(reportSection, iniType & " BEDP R", 0)
                    pdfLayout.BEDP_G = file.GetInteger(reportSection, iniType & " BEDP G", 0)
                    pdfLayout.BEDP_B = file.GetInteger(reportSection, iniType & " BEDP B", 0)
                    pdfLayout.BEDP = file.GetString(reportSection, iniType & " BEDP", "")

                    pdfLayout.IBEDP_R = file.GetInteger(reportSection, iniType & " IBEDP R", 0)
                    pdfLayout.IBEDP_G = file.GetInteger(reportSection, iniType & " IBEDP G", 0)
                    pdfLayout.IBEDP_B = file.GetInteger(reportSection, iniType & " IBEDP B", 0)

                    pdfLayout.RisikoInvestasi_R = file.GetInteger(reportSection, iniType & " RI R", 0)
                    pdfLayout.RisikoInvestasi_G = file.GetInteger(reportSection, iniType & " RI G", 0)
                    pdfLayout.RisikoInvestasi_B = file.GetInteger(reportSection, iniType & " RI B", 0)
                    pdfLayout.RisikoInvestasi = file.GetString(reportSection, iniType & " RI", "")

                    pdfLayout.IRI1 = file.GetString(reportSection, iniType & " IRI1", "")
                    pdfLayout.IRI2 = file.GetString(reportSection, iniType & " IRI2", "")
                    pdfLayout.IRI3 = file.GetString(reportSection, iniType & " IRI3", "")
                    pdfLayout.IRI4 = file.GetString(reportSection, iniType & " IRI4", "")
                    pdfLayout.IRI5 = file.GetString(reportSection, iniType & " IRI5", "")
                    pdfLayout.IRI6 = file.GetString(reportSection, iniType & " IRI6", "")
                    pdfLayout.IRI7 = file.GetString(reportSection, iniType & " IRI7", "")

                    pdfLayout.KlasifikasiRisiko_R = file.GetInteger(reportSection, iniType & " KR R", 0)
                    pdfLayout.KlasifikasiRisiko_G = file.GetInteger(reportSection, iniType & " KR G", 0)
                    pdfLayout.KlasifikasiRisiko_B = file.GetInteger(reportSection, iniType & " KR B", 0)
                    pdfLayout.KlasifikasiRisiko = file.GetString(reportSection, iniType & " KR", "")

                    pdfLayout.MengenaiManajerInvestasi_R = file.GetInteger(reportSection, iniType & " MMI R", 0)
                    pdfLayout.MengenaiManajerInvestasi_G = file.GetInteger(reportSection, iniType & " MMI G", 0)
                    pdfLayout.MengenaiManajerInvestasi_B = file.GetInteger(reportSection, iniType & " MMI B", 0)
                    pdfLayout.MengenaiManajerInvestasi = file.GetString(reportSection, iniType & " MMI", "")

                    pdfLayout.ValueMengenaiManajerInvestasi_R = file.GetInteger(reportSection, iniType & " IMMI R", 0)
                    pdfLayout.ValueMengenaiManajerInvestasi_G = file.GetInteger(reportSection, iniType & " IMMI G", 0)
                    pdfLayout.ValueMengenaiManajerInvestasi_B = file.GetInteger(reportSection, iniType & " IMMI B", 0)
                    pdfLayout.ValueMengenaiManajerInvestasi = file.GetString(reportSection, iniType & " IMMI", "")

                    pdfLayout.UlasanPasar_R = file.GetInteger(reportSection, iniType & " UP R", 0)
                    pdfLayout.UlasanPasar_G = file.GetInteger(reportSection, iniType & " UP G", 0)
                    pdfLayout.UlasanPasar_B = file.GetInteger(reportSection, iniType & " UP B", 0)
                    pdfLayout.UlasanPasar = file.GetString(reportSection, iniType & " UP", "")

                    pdfLayout.ChartTitle_R = file.GetInteger(reportSection, iniType & " Chart Title R", 0)
                    pdfLayout.ChartTitle_G = file.GetInteger(reportSection, iniType & " Chart Title G", 0)
                    pdfLayout.ChartTitle_B = file.GetInteger(reportSection, iniType & " Chart Title B", 0)
                    pdfLayout.ChartTitle = file.GetString(reportSection, iniType & " Grafik Kinerja Reksa Dana", "")

                    pdfLayout.ChartBorder_R = file.GetInteger(reportSection, iniType & " Chart Border R", 0)
                    pdfLayout.ChartBorder_G = file.GetInteger(reportSection, iniType & " Chart Border G", 0)
                    pdfLayout.ChartBorder_B = file.GetInteger(reportSection, iniType & " Chart Border B", 0)
                    If file.GetBoolean(reportSection, iniType & " Chart Border", False) Then pdfLayout.ChartBorder = True Else pdfLayout.ChartBorder = False

                    pdfLayout.ChartLine_R = file.GetInteger(reportSection, iniType & " Chart Line R", 0)
                    pdfLayout.ChartLine_G = file.GetInteger(reportSection, iniType & " Chart Line G", 0)
                    pdfLayout.ChartLine_B = file.GetInteger(reportSection, iniType & " Chart Line B", 0)
                    pdfLayout.ChartAxisX = file.GetString(reportSection, iniType & " Chart Label 1", "")
                    pdfLayout.ChartAxisY = file.GetString(reportSection, iniType & " Chart Label 2", "")

                    pdfLayout.TableHeader_R = file.GetInteger(reportSection, iniType & " Table Header R", 0)
                    pdfLayout.TableHeader_G = file.GetInteger(reportSection, iniType & " Table Header G", 0)
                    pdfLayout.TableHeader_B = file.GetInteger(reportSection, iniType & " Table Header B", 0)

                    pdfLayout.TableItem_B = file.GetInteger(reportSection, iniType & " Table Items R", 0)
                    pdfLayout.TableItem_G = file.GetInteger(reportSection, iniType & " Table Items G", 0)
                    pdfLayout.TableItem_B = file.GetInteger(reportSection, iniType & " Table Items B", 0)
                End If
            Else
                pdfColorDefault()
            End If
        Catch ex As Exception
            pdfColorDefault()
        End Try
    End Sub

    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click
        ReportSetting()
    End Sub

    Private Sub ReportSetting()
        Dim frm As New ReportFundSheetSectorEQSetting
        frm.frm = Me
        frm.Show()
        frm.FormLoad()
        frm.MdiParent = MDIMENU
    End Sub

    Private Sub txtWorstMonth_TextChanged(sender As Object, e As EventArgs) Handles txtWorstMonth.TextChanged

    End Sub
End Class