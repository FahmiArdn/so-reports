Imports simpi.GlobalUtilities
Imports simpi.CoreData
Imports simpi.MasterPortfolio
Imports System.Drawing.Imaging
Imports simpi.GlobalConnection
Imports simpi.MasterSecurities
Imports C1.Win.C1Chart
Imports C1.Win.C1TrueDBGrid
Imports simpi.MarketInstrument.ParameterSecurities

Public Class ReportFundSheetEQSyariah
    Dim objPortfolio As New MasterPortfolio
    Dim objSimpi As New simpi.MasterSimpi.MasterSimpi
    Dim objCodeset As New PortfolioCodeset
    Dim objNAV As New PortfolioNAV

    Dim objReturn As New PortfolioReturn
    Dim objBenchmark As New simpi.CoreData.PortfolioBenchmark
    Dim objSecurities As New PositionSecurities
    Dim objSector As New ParameterSectorClass
    Dim objTerm As New simpi.MasterSimpi.SimpiTerm
    Dim objReview As New PortfolioMarketReview
    Dim dtNAV As New DataTable
    Dim dtBenchmark As New DataTable
    Dim dtSecurities As New DataTable
    Dim dtSector As New DataTable

    Private Sub ReportFundSheetEQ_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetInstrumentUser()
        GetParameterInstrumentType()
        dtAs.Value = Now

        objPortfolio.UserAccess = objAccess
        objSimpi.UserAccess = objAccess
        objCodeset.UserAccess = objAccess
        objNAV.UserAccess = objAccess
        'new
        objReturn.UserAccess = objAccess
        objBenchmark.UserAccess = objAccess
        objSecurities.UserAccess = objAccess
        objSector.UserAccess = objAccess
        objTerm.UserAccess = objAccess
        objReview.UserAccess = objAccess

        DBGHolding.FetchRowStyles = True
        DBGPerformance.FetchRowStyles = True
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
                    txtCcy.Text = objPortfolio.GetPortfolioCcy.GetCcy
                    'txtCurrency.Text = objPortfolio.GetPortfolioCcy.GetCcyDescription
                    txtInception.Text = objPortfolio.GetInceptionDate.ToString("dd-MMMM-yyyy")
                    'txtCustodian.Text = objCodeset.GetCodeset(objPortfolio, 2)

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
            If objNAV.ErrID = 0 Then
                'data load
                objSecurities.Clear()
                dtSecurities = objSecurities.Search(objPortfolio, objNAV.GetPositionDate)
                objReturn.LoadAt(objPortfolio, objNAV.GetPositionDate)
                objBenchmark.LoadAt(objPortfolio, objNAV.GetPositionDate)
                dtNAV = objNAV.SearchHistoryLast(objPortfolio, objNAV.GetPositionDate, 0)
                dtBenchmark = objBenchmark.SearchHistoryLast(objPortfolio, objNAV.GetPositionDate, 0)

                'data display
                txtNAVPerUnit.Text = objNAV.GetNAVPerUnit.ToString("n4")
                txtAUM.Text = (objNAV.GetNAV / 1000000000).ToString("n2")

                DisplayPerformance()
                DisplayHolding()
                DisplaySector()
                DisplayNAV()
            Else
                ExceptionMessage.Show(objNAV.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
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
        With DBGPerformance
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

    End Sub 'Table Performance

    Private Sub DBGPerformance_FetchRowStyle(sender As Object, e As FetchRowStyleEventArgs) Handles DBGPerformance.FetchRowStyle
        If e.Row Mod 2 = 0 Then e.CellStyle.BackColor = Color.LemonChiffon
    End Sub

    Private Sub DisplayHolding()
        If dtSecurities IsNot Nothing AndAlso dtSecurities.Rows.Count > 0 AndAlso
           dtInstrumentUser IsNot Nothing AndAlso dtInstrumentUser.Rows.Count > 0 AndAlso
           dtParameterCountry IsNot Nothing AndAlso dtParameterCountry.Rows.Count > 0 AndAlso
           dtParameterInstrumentType IsNot Nothing AndAlso dtParameterInstrumentType.Rows.Count > 0 AndAlso
           dtSector IsNot Nothing AndAlso dtSector.Rows.Count > 0 Then

            Dim query = From p In dtSecurities.AsEnumerable
                        Join i In dtInstrumentUser On i.Field(Of Long)("SecuritiesID") Equals p.Field(Of Long)("SecuritiesID")
                        Join c In dtParameterCountry On i.Field(Of Integer)("CcyID") Equals c.Field(Of Integer)("CountryID")
                        Join t In dtParameterInstrumentType On i.Field(Of Integer)("TypeID") Equals t.Field(Of Integer)("TypeID")
                        Group Join s In dtSector.AsEnumerable On i.Field(Of Integer)("CountryID") Equals s.Field(Of Integer)("CountryID")
                               Into sp = Group Let s = sp.FirstOrDefault
                        Order By p.Field(Of Decimal)("MarketValue") Descending
                        Select ID = i.Field(Of Long)("SecuritiesID"),
                            Code = i.Field(Of String)("SecuritiesCode"),
                            Name = i.Field(Of String)("SecuritiesNameShort"),
                            TypeID = i.Field(Of Integer)("TypeID"),
                            Type = t.Field(Of String)("TypeCode"),
                            Sector = If(s Is Nothing, "No Sector", s.Field(Of String)("SectorName")),
                            Ccy = c.Field(Of String)("Ccy"),
                            Qty = p.Field(Of Decimal)("Qty"),
                            Price = p.Field(Of Decimal)("MarketPrice"),
                            Value = p.Field(Of Decimal)("MarketValue"),
                            Bobot = (p.Field(Of Decimal)("MarketValue") * 100 / objNAV.GetNAV)

            txtCompositionEQ.Text = (From q In query Where q.TypeID = SetEQ() Select q.Bobot).Sum.ToString("n2")
            txtCompositionFI.Text = (From q In query Where q.TypeID = SetFI() Select q.Bobot).Sum.ToString("n2")
            txtCompositionMM.Text = (100 - CDbl(txtCompositionEQ.Text) - CDbl(txtCompositionFI.Text)).ToString("n2")

            'ID, Code, Name  isprivate, ccy, Qty, price, ccy portfolio + value, %NAV 
            With DBGHolding
                .AllowAddNew = False
                .AllowDelete = False
                .AllowUpdate = False
                .Style.WrapText = False
                .Columns.Clear()
                .DataSource = query.ToList

                .Columns("Qty").NumberFormat = "n0"
                .Columns("Price").NumberFormat = "n2"
                .Columns("Value").NumberFormat = "n2"
                .Columns("Bobot").NumberFormat = "n2"

                .Splits(0).DisplayColumns("ID").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
                .Splits(0).DisplayColumns("Code").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Near
                .Splits(0).DisplayColumns("Name").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Near
                .Splits(0).DisplayColumns("Type").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
                .Splits(0).DisplayColumns("Sector").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Near
                .Splits(0).DisplayColumns("Ccy").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
                .Splits(0).DisplayColumns("Qty").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
                .Splits(0).DisplayColumns("Price").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
                .Splits(0).DisplayColumns("Value").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
                .Splits(0).DisplayColumns("Bobot").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far

                .Columns("Bobot").Caption = "%"

                For Each column As C1DisplayColumn In .Splits(0).DisplayColumns
                    If column.Name.Trim = "TypeID" Then column.Visible = False Else column.AutoSize()
                    .Splits(0).DisplayColumns(column.Name).HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
                Next

            End With
        End If

    End Sub 'Table Holding

    Private Sub DBGHolding_FetchRowStyle(sender As Object, e As FetchRowStyleEventArgs) Handles DBGHolding.FetchRowStyle
        If e.Row Mod 2 = 0 Then e.CellStyle.BackColor = Color.LemonChiffon
    End Sub

    Private Sub DisplaySector()
        If dtSecurities IsNot Nothing AndAlso dtSecurities.Rows.Count > 0 AndAlso
           dtInstrumentUser IsNot Nothing AndAlso dtInstrumentUser.Rows.Count > 0 AndAlso
           dtSector IsNot Nothing AndAlso dtSector.Rows.Count > 0 Then

            Dim query = From i In dtSecurities.AsEnumerable Join m In dtInstrumentUser
                               On i.Field(Of Long)("SecuritiesID") Equals m.Field(Of Long)("SecuritiesID")
                        Where m.Field(Of Integer)("TypeID") = SetEQ()
                        Group Join s In dtSector.AsEnumerable
                            On m.Field(Of Integer)("CompanyID") Equals s.Field(Of Integer)("CompanyID")
                            Into si = Group Let s = si.FirstOrDefault
                        Select
                            Sector = If(s Is Nothing, "No Sector", s.Field(Of String)("SectorName")),
                            MarketValue = i.Field(Of Decimal)("MarketValue")

            Dim query2 = From i In query.AsEnumerable Group value = i.MarketValue
                         By key = i.Sector Into s = Group Order By s.Sum Descending
                         Select Sector = key, Value = s.Sum()

            Dim aumTotal As Double = (From c In query2 Select aum = c.Value).Sum()

            If query2.Count > 0 Then
                chartSector.Style.Border.BorderStyle = BorderStyleEnum.None
                chartSector.ChartLabels.DefaultLabelStyle.BackColor = SystemColors.Info
                chartSector.ChartLabels.DefaultLabelStyle.Border.BorderStyle = BorderStyleEnum.Solid

                Dim grp As ChartGroup = chartSector.ChartGroups(0)
                grp.ChartType = Chart2DTypeEnum.Pie
                grp.Pie.OtherOffset = 0
                grp.Pie.InnerRadius = 65

                Dim dat As ChartData = grp.ChartData
                dat.SeriesList.Clear()

                Dim ColorValue() As Color = {Color.OrangeRed, Color.Tan, Color.LightGreen, Color.MediumTurquoise,
                                             Color.DodgerBlue, Color.Magenta, Color.GreenYellow, Color.MediumBlue}

                Dim slice, max As Integer
                Dim itemTotal As Double
                itemTotal = 1
                If query2.Count < 5 Then max = query2.Count Else max = 3
                slice = 0
                For Each item In query2.Take(max)
                    Dim series As ChartDataSeries = dat.SeriesList.AddNewSeries()
                    series.PointData.Length = 1
                    series.PointData(0) = New PointF(1.0F, item.Value * 100 / aumTotal)
                    series.LineStyle.Color = ColorValue(slice)
                    slice += 1
                    itemTotal -= (item.Value / aumTotal)
                    If item.Sector.Length > 15 Then
                        series.Label = item.Sector.Substring(0, 15).ToLowerInvariant & " "
                    Else
                        series.Label = item.Sector.ToLowerInvariant & " "
                    End If
                    series.Label &= (item.Value * 100 / aumTotal).ToString("n2") & "%"
                Next

                If itemTotal > 0 Then
                    Dim series As ChartDataSeries = dat.SeriesList.AddNewSeries()
                    series.PointData.Length = 1
                    series.PointData(0) = New PointF(1.0F, itemTotal * 100)
                    series.LineStyle.Color = ColorValue(slice)
                    series.Label = "other(s) " & itemTotal.ToString("n2") & "%"
                End If

                chartSector.Legend.Visible = True
                chartSector.Legend.Compass = CompassEnum.East
                chartSector.ChartGroups(0).ShowOutline = True
                chartSector.ChartArea.PlotArea.View3D.Elevation = 45

                'tbarHoleRadius.Maximum = 90
                'tbarHoleRadius.Minimum = 0
                'tbarHoleRadius.Value = chartSector.ChartGroups.Group0.Pie.InnerRadius

            End If
        End If
    End Sub 'Doughnut

    Private Sub DisplayNAV()
        If dtNAV IsNot Nothing AndAlso dtNAV.Rows.Count > 0 Then
            With chartNAV
                .Style.Border.BorderStyle = BorderStyleEnum.None
                Dim ds As ChartDataSeriesCollection = .ChartGroups(0).ChartData.SeriesList
                ds.Clear()
                Dim series As ChartDataSeries = ds.AddNewSeries()
                series.LineStyle.Color = Color.Green
                series.LineStyle.Thickness = 1
                series.SymbolStyle.Shape = SymbolShapeEnum.None
                series.FitType = FitTypeEnum.Line

                series.X.CopyDataIn((From q In dtNAV.AsEnumerable Order By q.Field(Of Date)("PositionDate") Ascending Select q.Field(Of Date)("PositionDate")).ToArray)
                series.Y.CopyDataIn((From q In dtNAV.AsEnumerable Order By q.Field(Of Date)("PositionDate") Ascending Select q.Field(Of Decimal)("GeometricIndex") - 1).ToArray)
                series.PointData.Length = dtNAV.Rows.Count

                .BackColor = Color.Transparent
                .ChartArea.AxisX.Max = CDate(dtNAV.Rows(0)("PositionDate")).ToOADate
                .ChartArea.AxisX.Min = CDate(dtNAV.Rows(dtNAV.Rows.Count - 1)("PositionDate")).ToOADate
                .ChartArea.AxisX.AutoMajor = True
                .ChartArea.AxisX.AutoMinor = True
                .ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual
                .ChartArea.AxisX.AnnoFormatString = "MMM-yy"
                .ChartArea.AxisX.AnnotationRotation = 25
                .ChartArea.AxisX.Origin = .ChartArea.AxisX.Min


                Dim Max As Double = (Decimal.Floor((From q In dtNAV.AsEnumerable Select q.Field(Of Decimal)("GeometricIndex") - 1).Max * 10) + 1) / 10D
                Dim Min As Double = Decimal.Floor((From q In dtNAV.AsEnumerable Select q.Field(Of Decimal)("GeometricIndex") - 1).Min * 10) / 10D
                .ChartArea.AxisY.Max = Max
                .ChartArea.AxisY.Min = Min
                If Max < 0 And Min < 0 Then .ChartArea.AxisY.Origin = (Min + Max) / 2D Else .ChartArea.AxisY.Origin = 0D
                .ChartArea.AxisY.AutoMajor = True
                .ChartArea.AxisY.AutoMinor = True
                .ChartArea.AxisY.AnnoFormat = FormatEnum.NumericManual
                .ChartArea.AxisY.AnnoFormatString = "p0"

                .ChartArea.AxisX.Thickness = 1
                .ChartArea.AxisY.Thickness = 1
            End With

            'pgNAV.SelectedObject = chartNAV
            'pgNAV.Text = chartNAV.Name
        End If
    End Sub 'Line

    Private Sub DisplayMonthlyReturn()

    End Sub 'Column

#Region "export"
    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles btnPDF.Click
        ExportPDF(False)
    End Sub

    Public Function ExportPDF(ByVal isAttachment As Boolean) As String
        Return PrintPDF(isAttachment)
    End Function

    Private Sub btnEmail_Click(sender As Object, e As EventArgs) Handles btnEmail.Click
        ReportEmail()
    End Sub

    Private Sub ReportEmail()
        'Dim frm As New ReportFundSheetEmail
        'frm.Show()
        'frm.frm = Me
        'frm.MdiParent = MDISO
    End Sub

    Private Function PrintPDF(ByVal isAttachment As Boolean) As String
        Dim strFile As String = ""
        Dim strLayout As String = ""
        Dim myBrush As New SolidBrush(Color.FromArgb(0, 6, 121))
        Dim detailBrush As New SolidBrush(Color.Black)
        Dim headerBrush As New SolidBrush(Color.White)
        Dim koordX As Single = 40, koordY As Single = 35
        Dim fontType = "Calibri", fontSize = 7
        Dim str As String = ""
        Dim efek As Integer = 5
        With c1pdf
            .Clear()
            .PaperKind = Printing.PaperKind.A4
            Dim rc As RectangleF = .PageRectangle
            strLayout = reportPDFPortrait("REPORT TEMPLATE PORTRAIT")
            If GlobalFileWindows.FileExists(strLayout) Then
                Dim img As Image = Image.FromFile(strLayout)
                .DrawImage(img, rc)
            End If
            simpiLogo(c1pdf, rc)
            'pdfColorDefault()

            .DrawString("Laporan Kinerja Bulanan", New Font(fontType, 10, FontStyle.Bold), myBrush, New PointF(koordX + 190, koordY - 10))
            .DrawString("Feb - 18", New Font(fontType, 10, FontStyle.Bold), myBrush, New PointF(koordX + 350, koordY - 10))
            .DrawLine(New Pen(Color.Gold, 1.5), New PointF(koordX + 105, koordY + 13), New PointF(koordX + 385, koordY + 13))
            .DrawString("Mandiri Global Sharia Equity Dollar", New Font(fontType, 14, FontStyle.Bold), myBrush, New PointF(koordX + 145, koordY + 17))

            koordY += 45
            Dim Bar1 = New SolidBrush(Color.FromArgb(23, 146, 27)),
                Bar2 = New SolidBrush(Color.FromArgb(173, 202, 80)),
                Bar3 = New SolidBrush(Color.FromArgb(92, 164, 74)),
                Bar4 = New SolidBrush(Color.FromArgb(225, 242, 187))
            .DrawString("Klasifikasi Tingkat Resiko", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 5, koordY))
            .DrawLine(New Pen(Color.Black, 1), New PointF(koordX + 145, koordY), New PointF(koordX + 145, koordY + 10)) '2
            .DrawLine(New Pen(Color.Black, 1), New PointF(koordX + 190, koordY), New PointF(koordX + 190, koordY + 10)) '4
            .FillRectangle(Bar1, New Rectangle(koordX + 115, koordY + 2, 125, 6.5))
            .DrawLine(New Pen(Color.Black, 1), New PointF(koordX + 115, koordY), New PointF(koordX + 115, koordY + 10)) '1
            .DrawLine(New Pen(Color.Black, 1), New PointF(koordX + 170, koordY), New PointF(koordX + 170, koordY + 10)) '3
            .DrawString("Rendah", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 155, koordY + 10))
            .FillRectangle(Bar2, New Rectangle(koordX + 240, koordY + 2, 125, 6.5))
            .DrawLine(New Pen(Color.Black, 1), New PointF(koordX + 240, koordY), New PointF(koordX + 240, koordY + 10)) '5
            .DrawString("Sedang", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 230, koordY + 10))
            .DrawLine(New Pen(Color.Black, 1), New PointF(koordX + 305, koordY), New PointF(koordX + 305, koordY + 10)) '6
            .DrawLine(New Pen(Color.Black, 1), New PointF(koordX + 365, koordY), New PointF(koordX + 365, koordY + 10)) '7
            .DrawString("Tinggi", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 365, koordY + 10))
            .DrawString("Jangka Waktu Investasi", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 5, koordY + 20))
            .FillRectangle(Bar3, New Rectangle(koordX + 115, koordY + 22, 40, 6.5))
            .DrawLine(New Pen(Color.Green, 1), New PointF(koordX + 115, koordY + 20), New PointF(koordX + 115, koordY + 30))    '1
            .DrawString("0", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 113, koordY + 30))
            .FillRectangle(Bar4, New Rectangle(koordX + 155, koordY + 22, 210, 6.5))
            .DrawLine(New Pen(Color.Green, 1), New PointF(koordX + 155, koordY + 20), New PointF(koordX + 155, koordY + 30))    '2
            .DrawString("1", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 153, koordY + 30))
            .DrawLine(New Pen(Color.Green, 1), New PointF(koordX + 175, koordY + 20), New PointF(koordX + 175, koordY + 30))    '3
            .DrawString("3", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 173, koordY + 30))
            .DrawLine(New Pen(Color.Green, 1), New PointF(koordX + 205, koordY + 20), New PointF(koordX + 205, koordY + 30))    '4
            .DrawString("6", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 203, koordY + 30))
            .DrawLine(New Pen(Color.Green, 1), New PointF(koordX + 240, koordY + 20), New PointF(koordX + 240, koordY + 30))    '5
            .DrawString("9", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 238, koordY + 30))
            .DrawLine(New Pen(Color.Green, 1), New PointF(koordX + 295, koordY + 20), New PointF(koordX + 295, koordY + 30))    '6
            .DrawString("12", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 292, koordY + 30))
            .DrawLine(New Pen(Color.Green, 1), New PointF(koordX + 365, koordY + 20), New PointF(koordX + 365, koordY + 30))    '7
            .DrawString("18 (Bulan)", New Font(fontType, 8, FontStyle.Bold), myBrush, New PointF(koordX + 362, koordY + 30))
            .DrawArc(Pens.Red, New RectangleF(koordX + 360, koordY, 10, 10), 0, 360) 'Cirlce
            .DrawArc(Pens.Red, New RectangleF(koordX + 360, koordY + 20, 10, 10), 0, 360) 'Cirlce           

            koordY += 45
            Dim clrColumn = New SolidBrush(Color.FromArgb(249, 249, 165))
            Dim clrColumnHdr = New SolidBrush(Color.FromArgb(68, 161, 65))
            Dim sf As StringFormat = New StringFormat()
            'Left Column
            rc = New RectangleF(koordX, koordY, 225, 600) 'Column
            .FillRectangle(clrColumn, rc)
            rc = New RectangleF(koordX, koordY, 225, 14) 'Header
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Tujuan Investasi", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY))
            rc = New RectangleF(koordX + 5, koordY + 15, 220, 30)
            .DrawString(txtInvestmentGoal.Text,
                        New Font(fontType, fontSize), Brushes.Black, rc)

            rc = New RectangleF(koordX, koordY + 45, 225, 14) 'Header
            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Center
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Kebijakan Investasi", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 45))
            .DrawString("Pasar Uang Syariah", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 60))
            .DrawString("0% -20%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 95, koordY + 60, 80, 10), sf)
            .DrawString("Saham Syariah", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 70))
            .DrawString("80% - 100%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 95, koordY + 70, 80, 10), sf)
            .DrawString("Sukuk", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 80))
            .DrawString("0% - 20%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 95, koordY + 80, 80, 10), sf)
            .DrawString("Efek Luar Negri", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 90))
            .DrawString("Min. 51%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 95, koordY + 90, 80, 10), sf)

            rc = New RectangleF(koordX, koordY + 105, 225, 14)
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Ulasan Singkat Market Outlook", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 105))
            .DrawString(txtUlasan.Text, New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 5, koordY + 115, 210, 120), sf)

            rc = New RectangleF(koordX, koordY + 230, 225, 14)
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Kepemilikan Terbesar", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 230))
            .DrawString("Nama Efek", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New PointF(koordX + 5, koordY + 245))
            .DrawString("Saham - Alphabetic Inc.", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 255))
            .DrawString("Saham - Apple Inc.", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 265))
            .DrawString("Saham - Facebook Inc.", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 275))
            .DrawString("Saham - Microsoft Corp.", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 285))
            .DrawString("Saham - Tencent Holding Ltd.", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 295))

            rc = New RectangleF(koordX, koordY + 305, 225, 14)
            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Center
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Komposisi Portofolio", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 305))
            .DrawString("Pasar Uang", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 320))
            .DrawString("4.64%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 105, koordY + 320, 80, 10), sf)
            .DrawString("Obligasi", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 330))
            .DrawString("0.00%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 105, koordY + 330, 80, 10), sf)
            .DrawString("Saham", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 340))
            .DrawString("95.36%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 105, koordY + 340, 80, 10), sf)

            rc = New RectangleF(koordX, koordY + 350, 225, 14)
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Posisi", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 350))
            .DrawString("Total Nilai Aktiva Bersih", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 365))
            .DrawString("USD", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 80, koordY + 365, 80, 10), sf)
            .DrawString("71.18 Juta", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 115, koordY + 365, 80, 10), sf)
            .DrawString("Nilai Aktiva Bersih per unit", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 375))
            .DrawString("USD", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 80, koordY + 375, 80, 10), sf)
            .DrawString("1.1251", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 115, koordY + 375, 80, 10), sf)
            .DrawString("Jumlah Outstanding unit", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 385))
            .DrawString("63.26 Juta", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 115, koordY + 385, 80, 10), sf)

            rc = New RectangleF(koordX, koordY + 400, 225, 14)
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Menganai Manajer Investasi", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 400))
            .DrawString(txtMIDesc.Text, New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 5, koordY + 415, 220, 80))

            rc = New RectangleF(koordX, koordY + 490, 225, 14)
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Informasi Lainnya", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 490))
            .DrawString("Minimum Investasi", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 505))
            .DrawString(": USD 10,000", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 70, koordY + 505, 190, 10))
            .DrawString("Bank Kustodian", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 515))
            .DrawString(": " & txtCustodian.Text & "", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 70, koordY + 515, 190, 10))
            .DrawString("Biaya Investasi", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 525))
            .DrawString("-Manajemen", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 535))
            .DrawString(": maks 2.50% p.a ", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 70, koordY + 535, 190, 10))
            .DrawString("-Pembelian", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 545))
            .DrawString(": maks 1.00%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 70, koordY + 545, 190, 10))
            .DrawString("-Penjualan Kembali", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 555))
            .DrawString(": maks 1.00% (< 1 tahun) ; 0% > 1 Tahun", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 70, koordY + 555, 190, 10))
            .DrawString("-Pengalihan", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 565))
            .DrawString(": maks 1.00%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 70, koordY + 565, 190, 10))
            .DrawString("Tanggal Efektif OJK", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 575))
            .DrawString(": 6 April 2016", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 70, koordY + 575, 190, 10))
            .DrawString("Cabang Penjualan", New Font(fontType, fontSize), Brushes.Black, New PointF(koordX + 5, koordY + 585))
            .DrawString(": Bank Mandiri", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 70, koordY + 585, 190, 10))

            'Right Column
            koordX += 235
            rc = New RectangleF(koordX, koordY, 280, 605) 'Column
            .FillRectangle(clrColumn, rc)
            rc = New RectangleF(koordX, koordY, 280, 14) 'Header
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Alokasi Aset", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY))  'CHART DOUGHNUT
            chartSector.Style.Border.BorderStyle = BorderStyle.None
            chartSector.Style.BackColor = Color.Transparent
            chartSector.ChartArea.PlotArea.BackColor = Color.Transparent
            chartSector.ChartArea.Style.BackColor = Color.Transparent
            chartSector.Legend.Style.BackColor = clrColumn.Color
            Dim imgSector = chartSector.GetImage(ImageFormat.Jpeg)
            rc = New RectangleF(New PointF(koordX + 60, koordY + 5), New SizeF(0.6 * chartSector.Size.Width, 0.6 * chartSector.Size.Height))
            c1pdf.DrawImage(imgSector, rc, ContentAlignment.TopCenter, C1.C1Pdf.ImageSizeModeEnum.Scale)

            rc = New RectangleF(koordX, koordY + 105, 280, 14)
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Kinerja Sejak Diluncurkan", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 105))   'CHART LINE
            chartNAV.ChartArea.AxisX.GridMajor.Visible = False
            chartNAV.ChartArea.AxisY.GridMajor.Visible = False
            'If pdfLayout.ChartBorder Then
            '    chartNAV.BorderStyle = BorderStyle.FixedSingle
            '    chartNAV.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.Solid
            '    chartNAV.ChartArea.Style.Border.Color = Color.FromArgb(pdfLayout.ChartBorder_R, pdfLayout.ChartBorder_G, pdfLayout.ChartBorder_B)
            'Else
            '    chartNAV.Style.Border.BorderStyle = BorderStyle.None
            '    chartNAV.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.None
            'End If
            chartNAV.Style.Border.BorderStyle = BorderStyle.None
            chartNAV.Style.BackColor = Color.Transparent
            chartNAV.ChartArea.Style.BackColor = Color.Transparent
            Dim imgPortfolio = chartNAV.GetImage(ImageFormat.Emf)
            rc = New RectangleF(New PointF(koordX, koordY + 120), New SizeF(0.47 * chartNAV.Size.Width, 0.47 * chartNAV.Size.Height))
            .DrawImage(imgPortfolio, rc, ContentAlignment.TopLeft, C1.C1Pdf.ImageSizeModeEnum.Scale)

            rc = New RectangleF(koordX, koordY + 230, 280, 14)
            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Center
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Kinerja Mandiri Investa Aktif dan Tolak Ukur", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 230))
            .DrawString("1 Bulan", New Font(fontType, 7, FontStyle.Bold), Brushes.Black, New Rectangle(koordX + 85, koordY + 250, 55, 10), sf)
            .DrawString("3 Bulan", New Font(fontType, 7, FontStyle.Bold), Brushes.Black, New Rectangle(koordX + 120, koordY + 250, 55, 10), sf)
            .DrawString("6 Bulan", New Font(fontType, 7, FontStyle.Bold), Brushes.Black, New Rectangle(koordX + 165, koordY + 250, 55, 10), sf)
            .DrawString("1 Tahun", New Font(fontType, 7, FontStyle.Bold), Brushes.Black, New Rectangle(koordX + 220, koordY + 250, 55, 10), sf)
            'DBGPerformance1.Row = 0
            .DrawString("Fund", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New PointF(koordX + 15, koordY + 260))
            .DrawString(":", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New PointF(koordX + 65, koordY + 260))
            'str = DBGPerformance1.Columns("1Mo").Text
            .DrawString("0.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 85, koordY + 260, 55, 10), sf)
            'str = DBGPerformance1.Columns("3Mo").Text
            .DrawString("0.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 120, koordY + 260, 55, 10), sf)
            'str = DBGPerformance1.Columns("6Mo").Text
            .DrawString("0.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 165, koordY + 260, 55, 10), sf)
            'str = DBGPerformance1.Columns("1Y").Text
            .DrawString("0.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 220, koordY + 260, 55, 10), sf)
            'DBGPerformance1.Row = 1
            .DrawString("Tolak Ukur", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New PointF(koordX + 15, koordY + 275))
            .DrawString(":", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New PointF(koordX + 65, koordY + 275))
            'str = DBGPerformance1.Columns("1Mo").Text
            .DrawString("0.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 85, koordY + 275, 55, 10), sf)
            'str = DBGPerformance1.Columns("3Mo").Text
            .DrawString("0.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 120, koordY + 275, 55, 10), sf)
            'str = DBGPerformance1.Columns("6Mo").Text
            .DrawString("0.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 165, koordY + 275, 55, 10), sf)
            'str = DBGPerformance1.Columns("1Y").Text
            .DrawString("0.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 220, koordY + 275, 55, 10), sf)

            .DrawString("3 Tahun", New Font(fontType, 7, FontStyle.Bold), Brushes.Black, New Rectangle(koordX + 85, koordY + 295, 55, 10), sf)
            .DrawString("YTD", New Font(fontType, 7, FontStyle.Bold), Brushes.Black, New Rectangle(koordX + 120, koordY + 295, 55, 10), sf)
            .DrawString("Sejak Diluncurkan", New Font(fontType, 7, FontStyle.Bold), Brushes.Black, New Rectangle(koordX + 165, koordY + 295, 55, 10), sf)
            .DrawString("SI Anualized", New Font(fontType, 7, FontStyle.Bold), Brushes.Black, New Rectangle(koordX + 220, koordY + 295, 55, 10), sf)
            'DBGPerformance1.Row = 0
            .DrawString("Fund", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New PointF(koordX + 15, koordY + 305))
            .DrawString(":", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New PointF(koordX + 65, koordY + 305))
            'str = DBGPerformance1.Columns("1Y").Text
            .DrawString("99.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 85, koordY + 305, 55, 10), sf)
            'str = DBGPerformance1.Columns("YTD").Text
            .DrawString("99.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 120, koordY + 305, 55, 10), sf)
            'str = DBGPerformance1.Columns("Sejak Diluncurkan").Text
            .DrawString("99.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 165, koordY + 305, 55, 10), sf)
            'str = DBGPerformance1.Columns("1Y").Text
            .DrawString("99.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 220, koordY + 305, 55, 10), sf)
            'DBGPerformance1.Row = 1
            .DrawString("Tolak Ukur", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New PointF(koordX + 15, koordY + 320))
            .DrawString(":", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New PointF(koordX + 65, koordY + 320))
            'str = DBGPerformance1.Columns("1Y").Text
            .DrawString("99.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 85, koordY + 320, 55, 10), sf)
            'str = DBGPerformance1.Columns("YTD").Text
            .DrawString("99.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 120, koordY + 320, 55, 10), sf)
            'str = DBGPerformance1.Columns("Sejak Diluncurkan").Text
            .DrawString("99.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 165, koordY + 320, 55, 10), sf)
            'str = DBGPerformance1.Columns("1Y").Text
            .DrawString("99.99%", New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 220, koordY + 320, 55, 10), sf)
            'str = DBGPerformance1.Columns("YTD").Text

            Dim Br1 = New SolidBrush(Color.FromArgb(60, 107, 178)),
                Br2 = New SolidBrush(Color.FromArgb(96, 126, 189)),
                Br3 = New SolidBrush(Color.FromArgb(130, 149, 203))
            .FillRectangle(Br1, New Rectangle(koordX + 10, koordY + 380, 85, 20))
            .DrawString("Bulan Terbaik" & vbCrLf & "Bulan Terburuk", New Font(fontType, fontSize, FontStyle.Bold), Brushes.White, New Rectangle(koordX + 15, koordY + 380, 75, 20))
            .DrawString("Bulan", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New Rectangle(koordX + 105, koordY + 365, 75, 20), sf)
            .FillRectangle(Br2, New Rectangle(koordX + 100, koordY + 380, 85, 20))
            .DrawString(txtBestMonthDate.Text & vbCrLf & txtWorstMonthDate.Text, New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 105, koordY + 380, 75, 20), sf)
            .DrawString("Kinerja", New Font(fontType, fontSize, FontStyle.Bold), Brushes.Black, New Rectangle(koordX + 195, koordY + 365, 75, 20), sf)
            .FillRectangle(Br3, New Rectangle(koordX + 190, koordY + 380, 85, 20))
            .DrawString(txtBestMonth.Text & vbCrLf & txtWorstMonth.Text, New Font(fontType, fontSize), Brushes.Black, New Rectangle(koordX + 195, koordY + 380, 75, 20), sf)

            rc = New RectangleF(koordX, koordY + 420, 280, 14)
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Tingkat Pengembalian Bulanan", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 420)) 'CHART COLUMN
            chartMonthly.ChartArea.AxisX.GridMajor.Visible = False
            chartMonthly.ChartArea.AxisY.GridMajor.Visible = False
            'If pdfLayout.ChartBorder Then
            '    chartNAV.BorderStyle = BorderStyle.FixedSingle
            '    chartNAV.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.Solid
            '    chartNAV.ChartArea.Style.Border.Color = Color.FromArgb(pdfLayout.ChartBorder_R, pdfLayout.ChartBorder_G, pdfLayout.ChartBorder_B)
            'Else
            '    chartNAV.Style.Border.BorderStyle = BorderStyle.None
            '    chartNAV.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.None
            'End If
            chartMonthly.Style.Border.BorderStyle = BorderStyle.None
            chartMonthly.Style.BackColor = Color.Transparent
            chartMonthly.ChartArea.Style.BackColor = Color.Transparent
            Dim imgColumn = chartMonthly.GetImage(ImageFormat.Emf)
            rc = New RectangleF(New PointF(koordX, koordY + 435), New SizeF(0.45 * chartMonthly.Size.Width, 0.45 * chartMonthly.Size.Height))
            .DrawImage(imgColumn, rc, ContentAlignment.TopLeft, C1.C1Pdf.ImageSizeModeEnum.Scale)

            rc = New RectangleF(koordX, koordY + 525, 280, 14)
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Risiko Investasi", New Font(fontType, 10), Brushes.White, New PointF(koordX + 5, koordY + 525))
            Dim factorList As String() = {txtRisk1.Text, txtRisk2.Text,
                                          txtRisk3.Text, txtRisk4.Text,
                                          txtRisk5.Text, txtRisk6.Text, txtRisk7.Text}
            Dim cnt As Integer = 0
            For i As Integer = 0 To factorList.Length - 1
                rc = New RectangleF(koordX + 5, koordY + 540, 275, 50)
                cnt += 1
                If factorList(i) IsNot "" Then
                    .DrawString("" & cnt & ". " & factorList(i), New Font(fontType, fontSize), Brushes.Black, rc)
                    koordY = koordY + fontSize + 2
                    'koordY = If(factorList(i).Length > 27, koordY + 2 * (fontSize + 1), koordY + fontSize)
                End If
            Next

            koordY = 735
            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Center
            rc = New RectangleF(koordX + 190, koordY, 80, 30)
            .FillRectangle(clrColumnHdr, rc)
            .DrawString("Kinerja Bulan ini: " & vbCrLf & " -3.00% " & vbCrLf &
                        " NAB/Unit : " & txtCcy.Text & " " & txtNAVPerUnit.Text & "", New Font(fontType, 7, FontStyle.Bold), Brushes.White, New PointF(koordX + 230, koordY + 15), sf)

            strFile = reportFileExists("ReportFS" & lblPortfolioCode.Text.Trim & dtAs.Value.ToString("yyyymmdd") & ".pdf")
            .Save(strFile)
            If Not isAttachment Then Process.Start(strFile)
        End With
        Return strFile
    End Function

    Private Sub ReportSetting()
        'Dim frm As New ReportKinerjaBulananSetting
        'frm.frm = Me
        'frm.Show()
        'frm.FormLoad()
        'frm.MdiParent = MDISO
    End Sub

    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click
        ReportSetting()
    End Sub

#End Region

End Class