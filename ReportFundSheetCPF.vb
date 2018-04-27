Imports System.Drawing.Imaging
Imports C1.Win.C1Chart
Imports simpi.CoreData
Imports simpi.GlobalConnection
Imports simpi.GlobalUtilities
Imports simpi.MasterPortfolio

Public Class ReportFundSheetCPF
    Dim objPortfolio As New MasterPortfolio
    Dim objNAV As New PortfolioNAV
    Public pdfLayout As New pdfColor
    Dim reportSection As String = "Report Fund Sheet CPF"
    Dim objReturn As New PortfolioReturn
    Dim dtNAV As New DataTable
    'Dim isChart As Boolean = False
    Dim x As Integer = 0
    Dim y As Integer = 0

    Public Property frm As ReportFundSheetCPF
    Private Sub ReportSetting()
        Dim frm As New ReportFundSheetCPFSetting
        frm.frm = Me
        frm.Show()
        frm.FormLoad()
    End Sub

#Region "pdf"
    Structure pdfColor
        Public layoutType As String
        Public Header_R As Integer
        Public Header_G As Integer
        Public Header_B As Integer
        Public Header As String

        Public Title_R As Integer
        Public Title_G As Integer
        Public Title_B As Integer
        Public Title As String

        Public HeaderLine_R As Integer
        Public HeaderLine_G As Integer
        Public HeaderLine_B As Integer

        Public NAV_R As Integer
        Public NAV_G As Integer
        Public NAV_B As Integer
        Public NAV As String

        Public ChartTitle_R As Integer
        Public ChartTitle_G As Integer
        Public ChartTitle_B As Integer
        Public ChartTitle As String

        Public SummaryLine_R As Integer
        Public SummaryLine_G As Integer
        Public SummaryLine_B As Integer
        Public SummaryItems_R As Integer
        Public SummaryItems_G As Integer
        Public SummaryItems_B As Integer
        Public SummaryItemsTanggalLaporan As String
        Public SummaryItemsBankKustodian As String
        Public SummaryItemsTanggalPeluncuran As String
        Public SummaryItemstanggalJatuhTempo As String
        Public SummaryItemsTotalAUM As String
        Public SummaryItemsMataUang As String
        Public SummaryItemsImbalJasaManajerInvestasi As String
        Public SummaryItemsImbalJasaBankKustodian As String
        Public SummaryItemsBiayaPembelian As String
        Public SummaryItemsBiayaPenjualanKembali As String
        Public SummaryItemsBiayaPengalihan As String
        Public SummaryItemsNote As String

        Public Subtitle_R As Integer
        Public Subtitle_G As Integer
        Public Subtitle_B As Integer
        Public Subtitle As String

        Public TanggalPenting_R As Integer
        Public TanggalPenting_G As Integer
        Public TanggalPenting_B As Integer
        Public TanggalPenting As String
        Public PembagianDividendTerakhir As String
        Public PembagianDividendBerikutnya As String
        Public UnderlyingAsset As String

        Public TanggalPentingLine_R As Integer
        Public TanggalPentingLine_G As Integer
        Public TanggalPentingLine_B As Integer

        Public TableHeader_R As Integer
        Public TableHeader_G As Integer
        Public TableHeader_B As Integer
        Public TableHeader As String

        Public TableTitle_R As Integer
        Public TableTitle_G As Integer
        Public TableTitle_B As Integer
        Public Table1Bln As String
        Public Table3Bln As String
        Public Table6Bln As String
        Public Table1Thn As String
        Public TableDariAwalTahun As String
        Public TableSejakPembentukan As String
        Public TableIndikasiRateDividenTetap As String

        Public FooterTitle1_R As Integer
        Public FooterTitle1_G As Integer
        Public FooterTitle1_B As Integer
        Public Footer1 As String

        Public FooterTitle2_R As Integer
        Public FooterTitle2_G As Integer
        Public FooterTitle2_B As Integer
        Public Footer2 As String

        Public TableLine_R As Integer
        Public TableLine_G As Integer
        Public TableLine_B As Integer

        Public ChartBorder_R As Integer
        Public ChartBorder_G As Integer
        Public ChartBorder_B As Integer
        Public ChartBorder As Boolean

    End Structure

    Public Sub pdfColorDefault()
        pdfLayout.layoutType = "DEFAULT"
        pdfLayout.Header_R = 9
        pdfLayout.Header_G = 62
        pdfLayout.Header_B = 111
        pdfLayout.Header = "Fund Fact Sheet"

        pdfLayout.Title_R = 9
        pdfLayout.Title_G = 62
        pdfLayout.Title_B = 111
        pdfLayout.Title = "Mandiri Protected Growth Dollar 3 "

        pdfLayout.ChartTitle_R = 9
        pdfLayout.ChartTitle_G = 62
        pdfLayout.ChartTitle_B = 111
        pdfLayout.ChartTitle = "Kinerja Reksa Dana "

        pdfLayout.HeaderLine_R = 255
        pdfLayout.HeaderLine_G = 255
        pdfLayout.HeaderLine_B = 0

        pdfLayout.NAV_R = 9
        pdfLayout.NAV_G = 62
        pdfLayout.NAV_B = 111
        pdfLayout.NAV = "NAV/Unit "

        pdfLayout.SummaryLine_R = 255
        pdfLayout.SummaryLine_G = 255
        pdfLayout.SummaryLine_B = 0

        pdfLayout.SummaryItems_R = 9
        pdfLayout.SummaryItems_G = 62
        pdfLayout.SummaryItems_B = 111
        pdfLayout.SummaryItemsTanggalLaporan = "Tanggal Laporan : "
        pdfLayout.SummaryItemsBankKustodian = "Bank Kustodian : "
        pdfLayout.SummaryItemsTanggalPeluncuran = "Tanggal Peluncuran : "
        pdfLayout.SummaryItemstanggalJatuhTempo = "Tanggal Jatuh Tempo : "
        pdfLayout.SummaryItemsTotalAUM = "Total AUM : "
        pdfLayout.SummaryItemsMataUang = "Mata Uang : "
        pdfLayout.SummaryItemsImbalJasaManajerInvestasi = "Imbal Jasa Manajer Investasi : "
        pdfLayout.SummaryItemsImbalJasaBankKustodian = "Imbal Jasa Bank Kustodian: "
        pdfLayout.SummaryItemsBiayaPembelian = "Biaya Pembelian : "
        pdfLayout.SummaryItemsBiayaPenjualanKembali = "Biaya Penjualan Kembali : "
        pdfLayout.SummaryItemsBiayaPengalihan = "Biaya Pengaliah : "
        pdfLayout.SummaryItemsNote = "Note : "

        pdfLayout.Subtitle_R = 9
        pdfLayout.Subtitle_G = 62
        pdfLayout.Subtitle_B = 111
        pdfLayout.Subtitle = "Kebijakan Investasi "

        pdfLayout.TanggalPenting_R = 9
        pdfLayout.TanggalPenting_G = 62
        pdfLayout.TanggalPenting_B = 111
        pdfLayout.TanggalPenting = "Tanggal Penting"
        pdfLayout.PembagianDividendTerakhir = "Pembagian Devidend Terakhir"
        pdfLayout.PembagianDividendBerikutnya = "Pembagian Devidend Berikutnya"
        pdfLayout.UnderlyingAsset = "Underlying Asset"

        pdfLayout.TanggalPentingLine_R = 255
        pdfLayout.TanggalPentingLine_G = 255
        pdfLayout.TanggalPentingLine_B = 0


        pdfLayout.TableHeader_R = 9
        pdfLayout.TableHeader_G = 62
        pdfLayout.TableHeader_B = 111
        pdfLayout.TableHeader = "Kinerja Historis (%)"

        pdfLayout.TableTitle_R = 9
        pdfLayout.TableTitle_G = 62
        pdfLayout.TableTitle_B = 111
        pdfLayout.Table1Bln = "1 Bulan"
        pdfLayout.Table3Bln = "3 Bulan"
        pdfLayout.Table6Bln = "6 Bulan"
        pdfLayout.Table1Thn = "1 Tahun"
        pdfLayout.TableDariAwalTahun = "Dari Awal Tahun"
        pdfLayout.TableSejakPembentukan = "Sejak Pembentukan"
        pdfLayout.TableIndikasiRateDividenTetap = "Indikasi Rate Dividen Tetap"

        pdfLayout.FooterTitle1_R = 9
        pdfLayout.FooterTitle1_G = 62
        pdfLayout.FooterTitle1_B = 111
        pdfLayout.Footer1 = "Tujuan Investasi"

        pdfLayout.FooterTitle2_R = 9
        pdfLayout.FooterTitle2_G = 62
        pdfLayout.FooterTitle2_B = 111
        pdfLayout.Footer2 = "Tentang Mandiri Investasi"

        pdfLayout.TableLine_R = 255
        pdfLayout.TableLine_G = 255
        pdfLayout.TableLine_B = 0

        pdfLayout.ChartBorder_R = 255
        pdfLayout.ChartBorder_G = 255
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
                    pdfLayout.Header_R = file.GetInteger(reportSection, iniType & " Header R", 0)
                    pdfLayout.Header_G = file.GetInteger(reportSection, iniType & " Header G", 0)
                    pdfLayout.Header_B = file.GetInteger(reportSection, iniType & " Header B", 0)
                    pdfLayout.Header = file.GetString(reportSection, iniType & " Header", "")

                    pdfLayout.Title_R = file.GetInteger(reportSection, iniType & " Title R", 0)
                    pdfLayout.Title_G = file.GetInteger(reportSection, iniType & " Title G", 0)
                    pdfLayout.Title_B = file.GetInteger(reportSection, iniType & " Title B", 0)
                    pdfLayout.Title = file.GetString(reportSection, iniType & " Title", "")

                    pdfLayout.HeaderLine_R = file.GetInteger(reportSection, iniType & " Header Line R", 0)
                    pdfLayout.HeaderLine_G = file.GetInteger(reportSection, iniType & " Header Line G", 0)
                    pdfLayout.HeaderLine_B = file.GetInteger(reportSection, iniType & " Header Line B", 0)

                    pdfLayout.NAV_R = file.GetInteger(reportSection, iniType & " NAV R", 0)
                    pdfLayout.NAV_G = file.GetInteger(reportSection, iniType & " NAV G", 0)
                    pdfLayout.NAV_B = file.GetInteger(reportSection, iniType & " NAV B", 0)
                    pdfLayout.NAV = file.GetString(reportSection, iniType & " NAV", "")

                    pdfLayout.SummaryLine_R = file.GetInteger(reportSection, iniType & " Summary Line R", 0)
                    pdfLayout.SummaryLine_G = file.GetInteger(reportSection, iniType & " Summary Line G", 0)
                    pdfLayout.SummaryLine_B = file.GetInteger(reportSection, iniType & " Summary Line B", 0)

                    pdfLayout.SummaryItems_R = file.GetInteger(reportSection, iniType & " Summary Items R", 0)
                    pdfLayout.SummaryItems_G = file.GetInteger(reportSection, iniType & " Summary Items G", 0)
                    pdfLayout.SummaryItems_B = file.GetInteger(reportSection, iniType & " Summary Items B", 0)
                    pdfLayout.SummaryItemsTanggalLaporan = file.GetString(reportSection, iniType & " Summary Items Tanggal Laporan", "")
                    pdfLayout.SummaryItemsBankKustodian = file.GetString(reportSection, iniType & " Summary Items Bank Kustodian", "")
                    pdfLayout.SummaryItemsTanggalPeluncuran = file.GetString(reportSection, iniType & " Summary Items Tanggal Peluncuran", "")
                    pdfLayout.SummaryItemstanggalJatuhTempo = file.GetString(reportSection, iniType & " Summary Items Tanggal Jatuh Tempo", "")
                    pdfLayout.SummaryItemsTotalAUM = file.GetString(reportSection, iniType & " Summary Items Total AUM", "")
                    pdfLayout.SummaryItemsMataUang = file.GetString(reportSection, iniType & " Summary Items Mata Uang", "")
                    pdfLayout.SummaryItemsImbalJasaManajerInvestasi = file.GetString(reportSection, iniType & " Summary Items Imbal Jasa Manajer Investasi", "")
                    pdfLayout.SummaryItemsImbalJasaBankKustodian = file.GetString(reportSection, iniType & " Summary Items Imbal Jasa Bank Kustodian", "")
                    pdfLayout.SummaryItemsBiayaPembelian = file.GetString(reportSection, iniType & " Summary Items Biaya Pembelian", "")
                    pdfLayout.SummaryItemsBiayaPenjualanKembali = file.GetString(reportSection, iniType & " Summary Items Biaya Penjualan Kembali", "")
                    pdfLayout.SummaryItemsBiayaPengalihan = file.GetString(reportSection, iniType & "Summary Items Biaya Pengalihan", "")
                    pdfLayout.SummaryItemsNote = file.GetString(reportSection, iniType & " Summary Items Note", "")

                    pdfLayout.Subtitle_R = file.GetInteger(reportSection, iniType & " Sub Title R", 0)
                    pdfLayout.Subtitle_G = file.GetInteger(reportSection, iniType & " Sub Title G", 0)
                    pdfLayout.Subtitle_B = file.GetInteger(reportSection, iniType & " Sub Title B", 0)
                    pdfLayout.Subtitle = file.GetString(reportSection, iniType & " Sub Title", "")

                    pdfLayout.TanggalPenting_R = file.GetInteger(reportSection, iniType & " Tanggal Penting R", 0)
                    pdfLayout.TanggalPenting_G = file.GetInteger(reportSection, iniType & " Tanggal Penting G", 0)
                    pdfLayout.TanggalPenting_B = file.GetInteger(reportSection, iniType & " Tanggal Penting B", 0)
                    pdfLayout.TanggalPenting = file.GetString(reportSection, iniType & " Tanggal Penting", "")
                    pdfLayout.PembagianDividendTerakhir = file.GetString(reportSection, iniType & " Pembagian Dividend Terakhir", "")
                    pdfLayout.PembagianDividendBerikutnya = file.GetString(reportSection, iniType & " Pembagian Dividend Berikutnya", "")
                    pdfLayout.UnderlyingAsset = file.GetString(reportSection, iniType & " Underlying Asset", "")

                    pdfLayout.TanggalPentingLine_R = file.GetInteger(reportSection, iniType & " Tanggal Penting Line R", 0)
                    pdfLayout.TanggalPentingLine_G = file.GetInteger(reportSection, iniType & " Tanggal Penting Line G", 0)
                    pdfLayout.TanggalPentingLine_B = file.GetInteger(reportSection, iniType & " Tanggal Penting Line B", 0)

                    pdfLayout.TableHeader_R = file.GetInteger(reportSection, iniType & " Pie Title 1 R", 0)
                    pdfLayout.TableHeader_G = file.GetInteger(reportSection, iniType & " Pie Title 1 G", 0)
                    pdfLayout.TableHeader_B = file.GetInteger(reportSection, iniType & " Pie Title 1 B", 0)
                    pdfLayout.TableHeader = file.GetString(reportSection, iniType & " Table Header", "")

                    pdfLayout.TableTitle_R = file.GetInteger(reportSection, iniType & " Pie Border 1 R", 0)
                    pdfLayout.TableTitle_G = file.GetInteger(reportSection, iniType & " Pie Border 1 G", 0)
                    pdfLayout.TableTitle_B = file.GetInteger(reportSection, iniType & " Pie Border 1 B", 0)
                    pdfLayout.Table1Bln = file.GetString(reportSection, iniType & " Table Title 1 Bln", "")
                    pdfLayout.Table3Bln = file.GetString(reportSection, iniType & " Table Title 3 Bln", "")
                    pdfLayout.Table6Bln = file.GetString(reportSection, iniType & " Table Title 6 Bln", "")
                    pdfLayout.Table1Thn = file.GetString(reportSection, iniType & " Table Title 1 Tahun", "")
                    pdfLayout.TableDariAwalTahun = file.GetString(reportSection, iniType & " Table Title Dari Awal Tahun", "")
                    pdfLayout.TableSejakPembentukan = file.GetString(reportSection, iniType & " Table Title Sejak Pembukuan", "")
                    pdfLayout.TableIndikasiRateDividenTetap = file.GetString(reportSection, iniType & " Table Title Indikasi Rate Dividen Tetap", "")

                    pdfLayout.FooterTitle1_R = file.GetInteger(reportSection, iniType & " Footer Title 1 R", 0)
                    pdfLayout.FooterTitle1_G = file.GetInteger(reportSection, iniType & " Footer Title 1 G", 0)
                    pdfLayout.FooterTitle1_B = file.GetInteger(reportSection, iniType & " Footer Title 1 B", 0)
                    pdfLayout.Footer1 = file.GetString(reportSection, iniType & " Footer Title 1", "")

                    pdfLayout.FooterTitle2_R = file.GetInteger(reportSection, iniType & " Footer Title 2 R", 0)
                    pdfLayout.FooterTitle2_G = file.GetInteger(reportSection, iniType & " Footer Title 2 G", 0)
                    pdfLayout.FooterTitle2_B = file.GetInteger(reportSection, iniType & " Footer Title 2 B", 0)
                    pdfLayout.Footer2 = file.GetString(reportSection, iniType & " Footer Title 2", "")

                    pdfLayout.TableLine_R = file.GetInteger(reportSection, iniType & " Table Line R", 0)
                    pdfLayout.TableLine_G = file.GetInteger(reportSection, iniType & " Table Line G", 0)
                    pdfLayout.TableLine_B = file.GetInteger(reportSection, iniType & " Table Line B", 0)

                    pdfLayout.ChartTitle_R = file.GetInteger(reportSection, iniType & " Chart Title R", 0)
                    pdfLayout.ChartTitle_G = file.GetInteger(reportSection, iniType & " Chart Title G", 0)
                    pdfLayout.ChartTitle_B = file.GetInteger(reportSection, iniType & " Chart Title B", 0)
                    pdfLayout.ChartTitle = file.GetString(reportSection, iniType & " Chart Title", "")

                    pdfLayout.ChartBorder_R = file.GetInteger(reportSection, iniType & " Line Chart Border R", 0)
                    pdfLayout.ChartBorder_G = file.GetInteger(reportSection, iniType & " Line Chart Border G", 0)
                    pdfLayout.ChartBorder_B = file.GetInteger(reportSection, iniType & " Line Chart Border B", 0)
                    If file.GetBoolean(reportSection, iniType & " Line Chart Border", False) Then pdfLayout.ChartBorder = True Else pdfLayout.ChartBorder = False
                End If
            Else
                pdfColorDefault()
            End If
        Catch ex As Exception
            pdfColorDefault()
        End Try
    End Sub
#End Region

    Private Sub ReportFundSheetCPF_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetMasterSimpi()
        GetInstrumentUser()
        GetParameterInstrumentType()
        dtAs.Value = Now

        objPortfolio.UserAccess = objAccess
        objNAV.UserAccess = objAccess
        objReturn.UserAccess = objAccess
        pdfSetting()

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
                txtCcy.Text = objPortfolio.GetPortfolioCcy.GetCcy.ToString
                txtCurrency.Text = objPortfolio.GetPortfolioCcy.GetCcyDescription.ToString
                txtInception.Text = objPortfolio.GetInceptionDate.ToString("dd-MMMM-yyyy")

            Else
                ExceptionMessage.Show(objPortfolio.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        DataLoad()
        DataDisplay()
        DisplayNAV()

    End Sub

    Private Sub DataLoad()
        If objPortfolio.GetPortfolioID > 0 Then
            objNAV.Clear()
            objNAV.LoadAt(objPortfolio, dtAs.Value)
            objReturn.LoadAt(objPortfolio, dtAs.Value)

        End If
    End Sub

    Private Sub DataDisplay()
        If objNAV.GetNAV > 0 Then
            txtNAVPerUnit.Text = objNAV.GetNAVPerUnit.ToString("n4")
            txtAUM.Text = (objNAV.GetNAV / 100000).ToString("n2")
            InputTextBox7.Text = objReturn.Getr1Mo.ToString("n2")
            InputTextBox8.Text = objReturn.Getr3Mo.ToString("n2")
            InputTextBox9.Text = objReturn.Getr6Mo.ToString("n2")
            InputTextBox10.Text = objReturn.GetrYTD.ToString("n2")
            InputTextBox11.Text = objReturn.Getr1Y.ToString("n2")
            InputTextBox12.Text = objReturn.GetrInception.ToString("n2")

        End If


    End Sub

    Private Sub DisplayNAV()
        Dim tbl As New DataTable

        'tbl = objReturn.SearchHistoryLast(objPortfolio, dtAs.Value, 0)
        tbl = objNAV.SearchHistoryLast(objPortfolio, objNAV.GetPositionDate)
        'dtNAV = tbl.DefaultView.ToTable("nav", True, "PositionDate", "GeometricIndex")

        With chartNAV
            .Style.Border.BorderStyle = BorderStyleEnum.None
            Dim ds As ChartDataSeriesCollection = .ChartGroups(0).ChartData.SeriesList
            ds.Clear()
            Dim series As ChartDataSeries = ds.AddNewSeries()
            'series.Label = "Price"
            series.LineStyle.Color = Color.Green
            series.LineStyle.Thickness = 2
            series.SymbolStyle.Shape = SymbolShapeEnum.None
            series.FitType = FitTypeEnum.Line

            series.X.CopyDataIn((From q In tbl.AsEnumerable Select q.Field(Of Date)("PositionDate")).ToArray)
            series.Y.CopyDataIn((From q In tbl.AsEnumerable Select q.Field(Of Decimal)("GeometricIndex") - 1).ToArray)
            series.PointData.Length = tbl.Rows.Count 

            'Dim fundData1 As New DataView(dtNAV)
            'series.PointData.Length = fundData1.Count
            'For i As Integer = 0 To fundData1.Count - 1
            '    series.X(i) = fundData1(i)("PositionDate")
            '    series.Y(i) = fundData1(i)("GeometricIndex") - 1
            'Next i

            .BackColor = Color.Transparent
            'If dtNAV.Rows.Count > 0 Then
            Dim firstDate, lastDate As Date
            firstDate = CDate(tbl.Rows(0)("PositionDate"))
            lastDate = CDate(tbl.Rows(tbl.Rows.Count - 1)("PositionDate"))
            ''    If CalculateDays(firstDate, lastDate, "A") <= 7 Then
            '.firstDate = firstDate.ToOADate
            .ChartArea.AxisX.Min = lastDate.ToOADate
            .ChartArea.AxisX.Max = firstDate.ToOADate

            ''        .ChartArea.AxisX.UnitMajor = 1
            ''        .ChartArea.AxisX.UnitMinor = 1
            ''        .ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual
            ''        .ChartArea.AxisX.AnnoFormatString = "dd-MMM-yy"
            ''    ElseIf CalculateDays(firstDate, lastDate, "A") <= 30 Then
            ''        .ChartArea.AxisX.Min = firstDate.ToOADate
            ''        .ChartArea.AxisX.Max = lastDate.ToOADate
            ''        .ChartArea.AxisX.UnitMajor = 3
            ''        .ChartArea.AxisX.UnitMinor = 1
            ''        .ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual
            ''        .ChartArea.AxisX.AnnoFormatString = "dd-MMM-yy"
            ''    Else
            ''        .ChartArea.AxisX.AutoMax = True
            ''        .ChartArea.AxisX.AutoMin = True
            ''        .ChartArea.AxisX.AutoMajor = True
            ''        .ChartArea.AxisX.AutoMinor = True
            ''        .ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual
            ''        .ChartArea.AxisX.AnnoFormatString = "MMM-yy"
            ''    End If
            ''Else

            .ChartArea.AxisX.AutoMajor = True
            .ChartArea.AxisX.AutoMinor = True
            .ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual
            .ChartArea.AxisX.AnnoFormatString = "MMM-yy"
            'End If
            .ChartArea.AxisX.AnnotationRotation = 25

            .ChartArea.AxisY.AutoMax = True
            .ChartArea.AxisY.AutoMin = True
            .ChartArea.AxisY.AutoMajor = True
            .ChartArea.AxisY.AutoMinor = True
            .ChartArea.AxisY.AnnoFormat = FormatEnum.NumericManual
            .ChartArea.AxisY.AnnoFormatString = "p2"

            .ChartArea.AxisX.Origin = lastDate.ToOADate
            .ChartArea.AxisY.Origin = 100
        End With
    End Sub

    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles btnPDF.Click
        ExportPDF(False)
    End Sub
    Public Function ExportPDF(ByVal isAttachment As Boolean) As String
        Return PrintPDF(isAttachment)
    End Function
    Private Function PrintPDF(ByVal isAttachment As Boolean) As String
        pdfSetting()
        Dim strFile As String = ""
        Dim strLayout As String = ""
        Dim myBrush As New SolidBrush(Color.FromArgb(0, 61, 121))
        Dim detailBrush As New SolidBrush(Color.Black)
        Dim headerBrush As New SolidBrush(Color.White)
        Dim koordX As Single = 40, koordY As Single = 35
        Dim fontType = "Calibri", fontSize = 8
        Dim layout = Image.FromFile("..\..\Template\Fund Sheet CPF - Portrait.jpg")
        With C1PdfDocument1
            .Clear()
            .PaperKind = Printing.PaperKind.A4
            Dim rc As RectangleF = .PageRectangle
            Dim sf As New StringFormat()
            sf.Alignment = StringAlignment.Center
            .DrawImage(layout, rc)
            rc = New RectangleF(koordX, koordY, 150, fontSize)
            .DrawStringRtf("{\b " & pdfLayout.Header & ", }" & Date.Today.ToString("MMMM yyyy"), New Font(fontType, 10), New SolidBrush(Color.FromArgb(pdfLayout.Header_R, pdfLayout.Header_G, pdfLayout.Header_B)), rc)
            koordY += 12
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.HeaderLine_R, pdfLayout.HeaderLine_G, pdfLayout.HeaderLine_B)), New PointF(koordX, koordY), New PointF(koordX + 300, koordY))
            koordY += 6
            rc = New RectangleF(koordX, koordY, 300, fontSize)
            .DrawStringRtf("{\b " + pdfLayout.Title + "}", New Font(fontType, 20), New SolidBrush(Color.FromArgb(pdfLayout.Title_R, pdfLayout.Title_G, pdfLayout.Title_B)), rc)
            .DrawString("Reksa Dana Terproteksi", New Font(fontType, 10), New SolidBrush(Color.FromArgb(pdfLayout.Title_R, pdfLayout.Title_G, pdfLayout.Title_B)), New PointF(koordX, koordY + 20))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.SummaryLine_R, pdfLayout.SummaryLine_G, pdfLayout.SummaryLine_B)), New PointF(koordX + 145, koordY + 50), New PointF(koordX + 145, koordY + 500))

            .DrawString(pdfLayout.NAV + " " + txtCcy.Text + " " + txtNAVPerUnit.Text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.NAV_R, pdfLayout.NAV_G, pdfLayout.NAV_B)), New PointF(koordX, koordY + 50))
            .DrawString(pdfLayout.SummaryItemsTanggalLaporan, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 65))
            .DrawString("28-Februari-2018", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 75))
            .DrawString(pdfLayout.SummaryItemsBankKustodian, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 90))
            .DrawString(txtCustodian.Text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 100))
            .DrawString(pdfLayout.SummaryItemsTanggalPeluncuran, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 115))
            .DrawString(txtInception.Text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 125))
            .DrawString(pdfLayout.SummaryItemstanggalJatuhTempo, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 140))
            .DrawString(txtISIN.Text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 150))
            .DrawString(pdfLayout.SummaryItemsTotalAUM, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 165))
            .DrawString(txtAUM.Text + " Juta", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 175))
            .DrawString(pdfLayout.SummaryItemsMataUang, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 190))
            .DrawString(txtCurrency.Text + " (" + txtCcy.Text + ")", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 200))
            .DrawString(pdfLayout.SummaryItemsImbalJasaManajerInvestasi, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 215))
            .DrawString(txtManagementFee.Text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 225))
            .DrawString(pdfLayout.SummaryItemsImbalJasaBankKustodian, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 240))
            .DrawString("Maks." + txtCustodianFee.Text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 250))
            .DrawString(pdfLayout.SummaryItemsBiayaPembelian, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 265))
            .DrawString("-", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 280))
            .DrawString(pdfLayout.SummaryItemsBiayaPenjualanKembali, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 295))
            .DrawString("Maks. 2.00%", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 305))
            .DrawString(pdfLayout.SummaryItemsBiayaPengalihan, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 320))
            .DrawString(txtSwitchingFee.Text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 330))
            .DrawString(pdfLayout.SummaryItemsNote, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), New PointF(koordX, koordY + 345))
            Dim rc2 As New Rectangle(koordX, koordY + 355, 145, koordY + 375)
            Dim text = txtInvestmentPeriode.Text
            .DrawString(text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.SummaryItems_R, pdfLayout.SummaryItems_G, pdfLayout.SummaryItems_B)), rc2)
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.SummaryLine_R, pdfLayout.SummaryLine_G, pdfLayout.SummaryLine_B)), New PointF(koordX, koordY + 510), New PointF(koordX + 520, koordY + 510))
            .DrawString(pdfLayout.Footer1, New Font(fontType, 12, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.FooterTitle1_R, pdfLayout.FooterTitle1_G, pdfLayout.FooterTitle1_B)), New PointF(koordX, koordY + 520))
            text = txtInvestmentGoal.Text
            rc = New Rectangle(koordX, koordY + 540, 500, koordY + 20)
            .DrawString(text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.FooterTitle1_R, pdfLayout.FooterTitle1_G, pdfLayout.FooterTitle1_B)), rc)
            .DrawString(pdfLayout.Footer2, New Font(fontType, 12, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.FooterTitle2_R, pdfLayout.FooterTitle2_G, pdfLayout.FooterTitle2_B)), New PointF(koordX, koordY + 600))
            text = txtAboutUs.Text
            rc = New Rectangle(koordX, koordY + 620, 500, koordY + 20)
            .DrawString(text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.FooterTitle2_R, pdfLayout.FooterTitle2_G, pdfLayout.FooterTitle2_B)), rc)


            koordX += 150
            .DrawString(pdfLayout.Subtitle, New Font(fontType, 12, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.Subtitle_R, pdfLayout.Subtitle_G, pdfLayout.Subtitle_B)), New PointF(koordX + 20, koordY + 50))
            .DrawString("Warrant", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.Subtitle_R, pdfLayout.Subtitle_G, pdfLayout.Subtitle_B)), New PointF(koordX + 20, koordY + 70))
            .DrawString(": 0%-30%", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.Subtitle_R, pdfLayout.Subtitle_G, pdfLayout.Subtitle_B)), New PointF(koordX + 200, koordY + 70))
            .DrawString("Surat utang Pemerintah Indonesia", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.Subtitle_R, pdfLayout.Subtitle_G, pdfLayout.Subtitle_B)), New PointF(koordX + 20, koordY + 80))
            .DrawString(": 70%-100%", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.Subtitle_R, pdfLayout.Subtitle_G, pdfLayout.Subtitle_B)), New PointF(koordX + 200, koordY + 80))
            .DrawString(pdfLayout.ChartTitle, New Font(fontType, 12, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.ChartTitle_R, pdfLayout.ChartTitle_G, pdfLayout.ChartTitle_B)), New PointF(koordX + 20, koordY + 100))
            .DrawString(pdfLayout.TanggalPenting, New Font(fontType, 12, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TanggalPenting_R, pdfLayout.TanggalPenting_G, pdfLayout.TanggalPenting_B)), New PointF(koordX, koordY + 345))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.TanggalPentingLine_R, pdfLayout.TanggalPentingLine_G, pdfLayout.TanggalPentingLine_B)), New PointF(koordX, koordY + 360), New PointF(koordX + 170, koordY + 360))
            .DrawString(pdfLayout.PembagianDividendTerakhir, New Font(fontType, 10, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TanggalPenting_R, pdfLayout.TanggalPenting_G, pdfLayout.TanggalPenting_B)), New PointF(koordX + 20, koordY + 370))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.TanggalPentingLine_R, pdfLayout.TanggalPentingLine_G, pdfLayout.TanggalPentingLine_B)), New PointF(koordX + 20, koordY + 385), New PointF(koordX + 170, koordY + 385))
            .DrawString(InputTextBox4.Text, New Font(fontType, 10), New SolidBrush(Color.FromArgb(pdfLayout.TanggalPenting_R, pdfLayout.TanggalPenting_G, pdfLayout.TanggalPenting_B)), New PointF(koordX + 50, koordY + 385))
            .DrawString(pdfLayout.PembagianDividendBerikutnya, New Font(fontType, 10, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TanggalPenting_R, pdfLayout.TanggalPenting_G, pdfLayout.TanggalPenting_B)), New PointF(koordX + 20, koordY + 400))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.TanggalPentingLine_R, pdfLayout.TanggalPentingLine_G, pdfLayout.TanggalPentingLine_B)), New PointF(koordX + 20, koordY + 415), New PointF(koordX + 170, koordY + 415))
            .DrawString(InputTextBox5.Text, New Font(fontType, 10), New SolidBrush(Color.FromArgb(pdfLayout.TanggalPenting_R, pdfLayout.TanggalPenting_G, pdfLayout.TanggalPenting_B)), New PointF(koordX + 60, koordY + 415))
            .DrawString(pdfLayout.UnderlyingAsset, New Font(fontType, 10, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TanggalPenting_R, pdfLayout.TanggalPenting_G, pdfLayout.TanggalPenting_B)), New PointF(koordX + 60, koordY + 430))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.TanggalPentingLine_R, pdfLayout.TanggalPentingLine_G, pdfLayout.TanggalPentingLine_B)), New PointF(koordX + 20, koordY + 445), New PointF(koordX + 170, koordY + 445))
            text = InputTextBox6.Text
            rc = New Rectangle(koordX + 20, koordY + 450, 145, koordY + 430)
            .DrawString(text, New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.TanggalPenting_R, pdfLayout.TanggalPenting_G, pdfLayout.TanggalPenting_B)), rc, sf)

            koordX += 230
            .DrawString(pdfLayout.TableHeader, New Font(fontType, 12, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableHeader_R, pdfLayout.TableHeader_G, pdfLayout.TableHeader_B)), New PointF(koordX, koordY + 345))
            .DrawLine(New Pen(Color.FromArgb(pdfLayout.TableLine_R, pdfLayout.TableLine_G, pdfLayout.TableLine_B)), New PointF(koordX, koordY + 360), New PointF(koordX + 140, koordY + 360))
            .DrawString(pdfLayout.Table1Bln, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 40, koordY + 370))
            .DrawString(pdfLayout.Table3Bln, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 90, koordY + 370))
            .DrawString("MPGD 3 :", New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX, koordY + 380))
            .DrawString(InputTextBox7.Text + "%", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 40, koordY + 380))
            .DrawString(InputTextBox8.Text + "%", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 90, koordY + 380))
            .DrawString(pdfLayout.Table6Bln, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 40, koordY + 400))
            .DrawString(pdfLayout.Table1Thn, New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 90, koordY + 400))
            .DrawString("MPGD 3 :", New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX, koordY + 410))
            .DrawString(InputTextBox9.Text + "%", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 40, koordY + 410))
            .DrawString(InputTextBox11.Text + "%", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 90, koordY + 410))
            rc = New Rectangle(koordX + 35, koordY + 435, 30, koordY + 20)
            .DrawString(pdfLayout.TableDariAwalTahun, New Font(fontType, 6, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), rc, sf)
            rc = New Rectangle(koordX + 80, koordY + 435, 40, koordY + 20)
            .DrawString(pdfLayout.TableSejakPembentukan, New Font(fontType, 6, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), rc, sf)
            .DrawString("MPGD 3 :", New Font(fontType, 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX, koordY + 450))
            .DrawString(InputTextBox10.Text + "%", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 40, koordY + 450))
            .DrawString(InputTextBox12.Text + "%", New Font(fontType, 8), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 90, koordY + 450))
            .DrawString(pdfLayout.TableIndikasiRateDividenTetap, New Font(fontType, 6, FontStyle.Bold), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX, koordY + 470))
            .DrawString(InputTextBox13.Text + "%", New Font(fontType, 6), New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX + 90, koordY + 470))

            'Dim imgPortfolio = chartNAV.GetImage(ImageFormat.Emf)
            'chartNAV.ChartArea.AxisX.GridMajor.Visible = False
            'chartNAV.ChartArea.AxisY.GridMajor.Visible = False
            'chartNAV.BorderStyle = BorderStyle.None
            'chartNAV.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.None
            'rc = New RectangleF(New PointF(koordX - 200, koordY + 120), New SizeF(chartNAV.Size.Width - 300, chartNAV.Size.Height + 30))
            '.DrawImage(imgPortfolio, rc, ContentAlignment.TopLeft, C1.C1Pdf.ImageSizeModeEnum.Stretch)

            chartNAV.ChartArea.AxisX.GridMajor.Visible = False
            chartNAV.ChartArea.AxisY.GridMajor.Visible = False
            If pdfLayout.ChartBorder Then
                chartNAV.BorderStyle = BorderStyle.FixedSingle
                chartNAV.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.Solid
                chartNAV.ChartArea.Style.Border.Color = Color.FromArgb(pdfLayout.ChartBorder_R, pdfLayout.ChartBorder_G, pdfLayout.ChartBorder_B)
            Else
                chartNAV.BorderStyle = BorderStyle.None
                chartNAV.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.None
            End If
            Dim imgPortfolio = chartNAV.GetImage(ImageFormat.Emf)
            rc = New RectangleF(New PointF(koordX - 230, koordY + 140), New SizeF(chartNAV.Size.Width * 0.9, chartNAV.Size.Height))
            .DrawImage(imgPortfolio, rc, ContentAlignment.TopLeft, C1.C1Pdf.ImageSizeModeEnum.Scale)

            strFile = reportFileExists("Report CPF " & lblPortfolioCode.Text.Trim & dtAs.Value.ToString("yyyymmdd") & ".pdf")
            .Save(strFile)

        End With
        If Not isAttachment Then Process.Start(strFile)
        Return strFile
    End Function

    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click
        ReportSetting()
    End Sub
End Class