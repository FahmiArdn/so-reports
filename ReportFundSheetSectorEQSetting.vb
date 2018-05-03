Imports simpi.GlobalConnection
Imports simpi.GlobalException
Imports simpi.GlobalGateway
Imports simpi.GlobalString
Imports simpi.GlobalUtilities

Public Class ReportFundSheetSectorEQSetting
    Public frm As ReportFundSheetSectorEQ
    Dim reportSection As String = "Report Fund SheetSector"
    Public Sub FormLoad()
        If frm.pdfLayout.LayoutType = "DEFAULT" Then
            rbDefault.Checked = True
        ElseIf frm.pdfLayout.LayoutType = "OPTION1" Then
            rbOption1.Checked = True
        Else
            rbOption2.Checked = True
        End If
    End Sub

    Private Sub rbDefault_CheckedChanged(sender As Object, e As EventArgs) Handles rbDefault.CheckedChanged
        iniCheck()
    End Sub

    Private Sub rbOption1_CheckedChanged(sender As Object, e As EventArgs) Handles rbOption1.CheckedChanged
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
                    r = file.GetInteger(reportSection, iniType & " Tujuan Investasi R", 0)
                    g = file.GetInteger(reportSection, iniType & " Tujuan Investasi G", 0)
                    b = file.GetInteger(reportSection, iniType & " Tujuan Investasi B", 0)
                    TujuanInvestasi_R.Text = RGBWrite("R", r)
                    TujuanInvestasi_G.Text = RGBWrite("G", g)
                    TujuanInvestasi_B.Text = RGBWrite("B", b)
                    txtColorTujuanInvestasi.BackColor = Color.FromArgb(r, g, b)
                    txtTujuanInvestasi.Text = file.GetString(reportSection, iniType & " Tujuan Investasi", "")

                    r = file.GetInteger(reportSection, iniType & " Tanggal R", 0)
                    g = file.GetInteger(reportSection, iniType & " Tanggal G", 0)
                    b = file.GetInteger(reportSection, iniType & " Tanggal B", 0)
                    Tanggal_R.Text = RGBWrite("R", r)
                    Tanggal_G.Text = RGBWrite("G", g)
                    Tanggal_B.Text = RGBWrite("B", b)
                    txtColorTanggal.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " IRD R", 0)
                    g = file.GetInteger(reportSection, iniType & " IRD G", 0)
                    b = file.GetInteger(reportSection, iniType & " IRD B", 0)
                    IRD_R.Text = RGBWrite("R", r)
                    IRD_G.Text = RGBWrite("G", g)
                    IRD_B.Text = RGBWrite("B", b)
                    txtColorIRD.BackColor = Color.FromArgb(r, g, b)
                    txtInformasiReksaDana.Text = file.GetString(reportSection, iniType & " IRD", "")

                    r = file.GetInteger(reportSection, iniType & " IIRD R", 0)
                    g = file.GetInteger(reportSection, iniType & " IIRD G", 0)
                    b = file.GetInteger(reportSection, iniType & " IIRD B", 0)
                    IIRD_R.Text = RGBWrite("R", r)
                    IIRD_G.Text = RGBWrite("G", g)
                    IIRD_B.Text = RGBWrite("B", b)
                    txtColorItemIRD.BackColor = Color.FromArgb(r, g, b)
                    txtItemIRD1.Text = file.GetString(reportSection, iniType & " IIRD1", "")
                    txtItemIRD2.Text = file.GetString(reportSection, iniType & " IIRD2", "")
                    txtItemIRD3.Text = file.GetString(reportSection, iniType & " IIRD3", "")
                    txtItemIRD4.Text = file.GetString(reportSection, iniType & " IIRD4", "")
                    txtItemIRD5.Text = file.GetString(reportSection, iniType & " IIRD5", "")
                    txtItemIRD6.Text = file.GetString(reportSection, iniType & " IIRD6", "")
                    txtItemIRD7.Text = file.GetString(reportSection, iniType & " IIRD7", "")
                    txtItemIRD8.Text = file.GetString(reportSection, iniType & " IIRD8", "")

                    r = file.GetInteger(reportSection, iniType & " KI R", 0)
                    g = file.GetInteger(reportSection, iniType & " KI G", 0)
                    b = file.GetInteger(reportSection, iniType & " KI B", 0)
                    KI_R.Text = RGBWrite("R", r)
                    KI_G.Text = RGBWrite("G", g)
                    KI_B.Text = RGBWrite("B", b)
                    txtColorRI.BackColor = Color.FromArgb(r, g, b)
                    txtKI.Text = file.GetString(reportSection, iniType & " KI", "")

                    r = file.GetInteger(reportSection, iniType & " KK R", 0)
                    g = file.GetInteger(reportSection, iniType & " KK G", 0)
                    b = file.GetInteger(reportSection, iniType & " KK B", 0)
                    KK_R.Text = RGBWrite("R", r)
                    KK_G.Text = RGBWrite("G", g)
                    KK_B.Text = RGBWrite("B", b)
                    txtColorRI.BackColor = Color.FromArgb(r, g, b)
                    txtKK.Text = file.GetString(reportSection, iniType & " KK", "")

                    r = file.GetInteger(reportSection, iniType & " KP R", 0)
                    g = file.GetInteger(reportSection, iniType & " KP G", 0)
                    b = file.GetInteger(reportSection, iniType & " KP B", 0)
                    KP_R.Text = RGBWrite("R", r)
                    KP_G.Text = RGBWrite("G", g)
                    KP_B.Text = RGBWrite("B", b)
                    txtColorRI.BackColor = Color.FromArgb(r, g, b)
                    txtKP.Text = file.GetString(reportSection, iniType & " KP", "")

                    r = file.GetInteger(reportSection, iniType & " AS R", 0)
                    g = file.GetInteger(reportSection, iniType & " AS G", 0)
                    b = file.GetInteger(reportSection, iniType & " AS B", 0)
                    RI_R.Text = RGBWrite("R", r)
                    RI_G.Text = RGBWrite("G", g)
                    RI_B.Text = RGBWrite("B", b)
                    txtColorRI.BackColor = Color.FromArgb(r, g, b)
                    txtAS.Text = file.GetString(reportSection, iniType & " AS", "")

                    r = file.GetInteger(reportSection, iniType & " Report Line R", 0)
                    g = file.GetInteger(reportSection, iniType & " Report Line G", 0)
                    b = file.GetInteger(reportSection, iniType & " Report Line B", 0)
                    ReportLine_R.Text = RGBWrite("R", r)
                    ReportLine_G.Text = RGBWrite("G", g)
                    ReportLine_B.Text = RGBWrite("B", b)
                    txtColoeReportLine.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " IBB R", 0)
                    g = file.GetInteger(reportSection, iniType & " IBB G", 0)
                    b = file.GetInteger(reportSection, iniType & " IBB B", 0)
                    IBB_R.Text = RGBWrite("R", r)
                    IBB_G.Text = RGBWrite("G", g)
                    IBB_B.Text = RGBWrite("B", b)
                    txtColorIBB.BackColor = Color.FromArgb(r, g, b)
                    txtIBB.Text = file.GetString(reportSection, iniType & " IBB", "")

                    r = file.GetInteger(reportSection, iniType & " IIBB R", 0)
                    g = file.GetInteger(reportSection, iniType & " IIBB G", 0)
                    b = file.GetInteger(reportSection, iniType & " IIBB B", 0)
                    IIBB_R.Text = RGBWrite("R", r)
                    IIBB_G.Text = RGBWrite("G", g)
                    IIBB_B.Text = RGBWrite("B", b)
                    txtColorItemIBB.BackColor = Color.FromArgb(r, g, b)
                    txtIBB1.Text = file.GetString(reportSection, iniType & " IIBB1", "")
                    txtIBB2.Text = file.GetString(reportSection, iniType & " IIBB2", "")
                    txtIBB3.Text = file.GetString(reportSection, iniType & " IIBB3", "")
                    txtIBB4.Text = file.GetString(reportSection, iniType & " IIBB4", "")
                    txtIBB5.Text = file.GetString(reportSection, iniType & " IIBB5", "")
                    txtIBB6.Text = file.GetString(reportSection, iniType & " IIBB6", "")

                    r = file.GetInteger(reportSection, iniType & " SR R", 0)
                    g = file.GetInteger(reportSection, iniType & " SR G", 0)
                    b = file.GetInteger(reportSection, iniType & " SR B", 0)
                    SR_R.Text = RGBWrite("R", r)
                    SR_G.Text = RGBWrite("G", g)
                    SR_B.Text = RGBWrite("B", b)
                    txtColorSR.BackColor = Color.FromArgb(r, g, b)
                    txtSR.Text = file.GetString(reportSection, iniType & " SR", "")

                    r = file.GetInteger(reportSection, iniType & " ISR R", 0)
                    g = file.GetInteger(reportSection, iniType & " ISR G", 0)
                    b = file.GetInteger(reportSection, iniType & " ISR B", 0)
                    ISR_R.Text = RGBWrite("R", r)
                    ISR_G.Text = RGBWrite("G", g)
                    ISR_B.Text = RGBWrite("B", b)
                    txtColorItemIBB.BackColor = Color.FromArgb(r, g, b)
                    txtItemSR1.Text = file.GetString(reportSection, iniType & " ISR1", "")
                    txtItemSR2.Text = file.GetString(reportSection, iniType & " ISR2", "")
                    txtItemSR3.Text = file.GetString(reportSection, iniType & " ISR3", "")
                    txtItemSR4.Text = file.GetString(reportSection, iniType & " ISR4", "")
                    txtItemSR5.Text = file.GetString(reportSection, iniType & " ISR5", "")
                    txtItemSR6.Text = file.GetString(reportSection, iniType & " ISR6", "")

                    r = file.GetInteger(reportSection, iniType & " RI R", 0)
                    g = file.GetInteger(reportSection, iniType & " RI G", 0)
                    b = file.GetInteger(reportSection, iniType & " RI B", 0)
                    RI_R.Text = RGBWrite("R", r)
                    RI_G.Text = RGBWrite("G", g)
                    RI_B.Text = RGBWrite("B", b)
                    txtColorRI.BackColor = Color.FromArgb(r, g, b)
                    txtRI.Text = file.GetString(reportSection, iniType & " RI", "")

                    r = file.GetInteger(reportSection, iniType & " IRI R", 0)
                    g = file.GetInteger(reportSection, iniType & " IRI G", 0)
                    b = file.GetInteger(reportSection, iniType & " IRI B", 0)
                    IRI_R.Text = RGBWrite("R", r)
                    IRI_G.Text = RGBWrite("G", g)
                    IRI_B.Text = RGBWrite("B", b)
                    txtColorItemIBB.BackColor = Color.FromArgb(r, g, b)
                    txtItemRI1.Text = file.GetString(reportSection, iniType & " IRI1", "")
                    txtItemRI2.Text = file.GetString(reportSection, iniType & " IRI2", "")
                    txtItemRI3.Text = file.GetString(reportSection, iniType & " IRI3", "")
                    txtItemRI4.Text = file.GetString(reportSection, iniType & " IRI4", "")
                    txtItemRI5.Text = file.GetString(reportSection, iniType & " IRI5", "")
                    txtItemRI6.Text = file.GetString(reportSection, iniType & " IRI6", "")
                    txtItemRI7.Text = file.GetString(reportSection, iniType & " IRI7", "")

                    r = file.GetInteger(reportSection, iniType & " KR R", 0)
                    g = file.GetInteger(reportSection, iniType & " KR G", 0)
                    r = file.GetInteger(reportSection, iniType & " KR B", 0)
                    KR_R.Text = RGBWrite("R", r)
                    KR_G.Text = RGBWrite("G", g)
                    KR_B.Text = RGBWrite("B", b)
                    txtColorKR.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " MMI R", 0)
                    g = file.GetInteger(reportSection, iniType & " MMI G", 0)
                    r = file.GetInteger(reportSection, iniType & " MMI B", 0)
                    MMI_R.Text = RGBWrite("R", r)
                    MMI_G.Text = RGBWrite("G", g)
                    MMI_B.Text = RGBWrite("B", b)
                    txtColorMMI.BackColor = Color.FromArgb(r, g, b)
                    txtMMI.Text = file.GetString(reportSection, iniType & " RI", "")

                    r = file.GetInteger(reportSection, iniType & " IMMI R", 0)
                    g = file.GetInteger(reportSection, iniType & " IMMI G", 0)
                    b = file.GetInteger(reportSection, iniType & " IMMI B", 0)
                    IMMI_R.Text = RGBWrite("R", r)
                    IMMI_G.Text = RGBWrite("G", g)
                    IMMI_B.Text = RGBWrite("B", b)
                    txtColorItemIBB.BackColor = Color.FromArgb(r, g, b)
                    'txtIMMI1.Text = file.GetString(reportSection, iniType & " IMMI1", "")

                    r = file.GetInteger(reportSection, iniType & " BEDP R", 0)
                    g = file.GetInteger(reportSection, iniType & " BEDP G", 0)
                    r = file.GetInteger(reportSection, iniType & " BEDP B", 0)
                    BEDP_R.Text = RGBWrite("R", r)
                    BEDP_G.Text = RGBWrite("G", g)
                    BEDP_B.Text = RGBWrite("B", b)
                    txtColorMMI.BackColor = Color.FromArgb(r, g, b)
                    txtBEDP.Text = file.GetString(reportSection, iniType & " BEDP", "")

                    r = file.GetInteger(reportSection, iniType & " Chart Title R", 0)
                    b = file.GetInteger(reportSection, iniType & " Chart Title G", 0)
                    g = file.GetInteger(reportSection, iniType & " Chart Title B", 0)
                    ChartTitle_R.Text = RGBWrite("R", r)
                    ChartTitle_G.Text = RGBWrite("B", b)
                    ChartTitle_B.Text = RGBWrite("G", g)
                    txtColorGKRD.BackColor = Color.FromArgb(r, b, g)
                    txtGKRD.Text = file.GetString(reportSection, iniType & " Investment Growth", "")

                    r = file.GetInteger(reportSection, iniType & " Chart Border R", 0)
                    g = file.GetInteger(reportSection, iniType & " Chart Border G", 0)
                    b = file.GetInteger(reportSection, iniType & " Chart Border B", 0)
                    ChartBorder_R.Text = RGBWrite("R", r)
                    ChartBorder_G.Text = RGBWrite("G", g)
                    ChartBorder_B.Text = RGBWrite("B", b)
                    txtColorChartBorder.BackColor = Color.FromArgb(r, g, b)
                    If file.GetBoolean(reportSection, iniType & " Chart Border", False) Then chkChartBorder.Checked = True Else chkChartBorder.Checked = False

                    r = file.GetInteger(reportSection, iniType & " Chart Line R", 0)
                    g = file.GetInteger(reportSection, iniType & " Chart Line G", 0)
                    b = file.GetInteger(reportSection, iniType & " Chart Line B", 0)
                    ChartLine_R.Text = RGBWrite("R", r)
                    ChartLine_G.Text = RGBWrite("G", g)
                    ChartLine_B.Text = RGBWrite("B", b)
                    txtColorChartLine.BackColor = Color.FromArgb(r, g, b)
                    txtAxisX.Text = file.GetString(reportSection, iniType & " Chart Label 1", "")
                    txtAxisY.Text = file.GetString(reportSection, iniType & " Chart Label 2", "")

                    r = file.GetInteger(reportSection, iniType & " Table Header R", 0)
                    g = file.GetInteger(reportSection, iniType & " Table Header G", 0)
                    b = file.GetInteger(reportSection, iniType & " Table Header B", 0)
                    TableHeader_R.Text = RGBWrite("R", r)
                    TableHeader_G.Text = RGBWrite("G", g)
                    TableHeader_B.Text = RGBWrite("B", b)
                    txtColorTableHeader.BackColor = Color.FromArgb(r, g, b)
                    txtHeader.Text = file.GetString(reportSection, iniType & " Table Header", "")

                    r = file.GetInteger(reportSection, iniType & " Table Items R", 0)
                    g = file.GetInteger(reportSection, iniType & " Table Items G", 0)
                    b = file.GetInteger(reportSection, iniType & " Table Items B", 0)
                    TableItems_R.Text = RGBWrite("R", r)
                    TableItems_G.Text = RGBWrite("G", g)
                    TableItems_B.Text = RGBWrite("B", b)
                    txtColorTableItems.BackColor = Color.FromArgb(r, g, b)

                    r = file.GetInteger(reportSection, iniType & " UP R", 0)
                    g = file.GetInteger(reportSection, iniType & " UP G", 0)
                    b = file.GetInteger(reportSection, iniType & " UP B", 0)
                    UP_R.Text = RGBWrite("R", r)
                    UP_G.Text = RGBWrite("G", g)
                    UP_B.Text = RGBWrite("B", b)
                    txtColorUP.BackColor = Color.FromArgb(r, g, b)

                End If
            End If
        Catch ex As Exception
            _default()
        End Try
    End Sub

    Private Sub _default()
        With frm.pdfLayout
            frm.pdfColorDefault()

            Tanggal_R.Text = "R: " & .Tanggal_R
            Tanggal_G.Text = "G: " & .Tanggal_G
            Tanggal_B.Text = "B: " & .Tanggal_B
            txtColorTanggal.BackColor = Color.FromArgb(.Tanggal_R, .Tanggal_G, .Tanggal_B)

            ReportLine_R.Text = "R: " & .ReportLine_R
            ReportLine_G.Text = "G: " & .ReportLine_G
            ReportLine_B.Text = "B: " & .ReportLine_B
            txtColoeReportLine.BackColor = Color.FromArgb(.ReportLine_R, .ReportLine_G, .ReportLine_B)

            TujuanInvestasi_R.Text = "R: " & .TujuanInvestasi_R
            TujuanInvestasi_G.Text = "G: " & .TujuanInvestasi_G
            TujuanInvestasi_B.Text = "B: " & .TujuanInvestasi_B
            txtColorTujuanInvestasi.BackColor = Color.FromArgb(.TujuanInvestasi_R, .TujuanInvestasi_G, .TujuanInvestasi_B)
            txtTujuanInvestasi.Text = .TujuanInvestasi

            IRD_R.Text = "R: " & .InformasiReksaDana_R
            IRD_R.Text = "G: " & .InformasiReksaDana_G
            IRD_R.Text = "B: " & .InformasiReksaDana_B
            txtColorIRD.BackColor = Color.FromArgb(.InformasiReksaDana_R, .InformasiReksaDana_G, .InformasiReksaDana_B)
            txtInformasiReksaDana.Text = .InformasiReksaDana

            IIRD_R.Text = "R: " & .ValueInformasiReksaDana_R
            IIRD_R.Text = "G: " & .ValueInformasiReksaDana_G
            IIRD_R.Text = "B: " & .ValueInformasiReksaDana_B
            txtColorItemIRD.BackColor = Color.FromArgb(.InformasiReksaDana_R, .InformasiReksaDana_G, .InformasiReksaDana_B)

            txtItemIRD1.Text = .IIRD1
            txtItemIRD2.Text = .IIRD2
            txtItemIRD3.Text = .IIRD3
            txtItemIRD4.Text = .IIRD4
            txtItemIRD5.Text = .IIRD5
            txtItemIRD6.Text = .IIRD6
            txtItemIRD7.Text = .IIRD7
            txtItemIRD8.Text = .IIRD8

            KK_R.Text = "R: " & .KinerjaKumulatif_R
            KK_G.Text = "G: " & .KinerjaKumulatif_G
            KK_B.Text = "B: " & .KinerjaKumulatif_B
            txtColorKK.BackColor = Color.FromArgb(.InformasiReksaDana_R, .InformasiReksaDana_G, .InformasiReksaDana_B)
            txtKK.Text = .KinerjaKumulatif

            KI_R.Text = "R: " & .KebijakanInvestasi_R
            KI_G.Text = "R: " & .KebijakanInvestasi_G
            KI_B.Text = "R: " & .KebijakanInvestasi_B
            txtColorKI.BackColor = Color.FromArgb(.KebijakanInvestasi_R, .KebijakanInvestasi_G, .KebijakanInvestasi_B)
            txtKI.Text = .KebijakanInvestasi

            AS_R.Text = "R: " & .AlokasiSektoral_R
            AS_G.Text = "R: " & .AlokasiSektoral_G
            AS_B.Text = "R: " & .AlokasiSektoral_B
            txtColorAS.BackColor = Color.FromArgb(.AlokasiSektoral_R, .AlokasiSektoral_G, .AlokasiSektoral_B)
            txtAS.Text = .AlokasiSektoral

            IBB_R.Text = "R: " & .InvestasiDanaBiayaBiaya_R
            IBB_G.Text = "G: " & .InvestasiDanaBiayaBiaya_G
            IBB_B.Text = "B: " & .InvestasiDanaBiayaBiaya_B

            IIBB_R.Text = "R: " & .ValueInvestasiDanaBiayaBiaya_R
            IIBB_G.Text = "R: " & .ValueInvestasiDanaBiayaBiaya_G
            IIBB_B.Text = "R: " & .ValueInvestasiDanaBiayaBiaya_B

            txtIBB1.Text = .IIBB1
            txtIBB2.Text = .IIBB2
            txtIBB3.Text = .IIBB3
            txtIBB4.Text = .IIBB4
            txtIBB5.Text = .IIBB5
            txtIBB6.Text = .IIBB6

            SR_R.Text = "R: " & .StatistikReksadana_R
            SR_G.Text = "G: " & .StatistikReksadana_G
            SR_B.Text = "B: " & .StatistikReksadana_B
            txtSR.Text = .StatistikReksadana

            ISR_R.Text = "R: " & .ValueStatistikReksadana_R
            ISR_G.Text = "G: " & .ValueStatistikReksadana_G
            ISR_B.Text = "B: " & .ValueStatistikReksadana_B

            txtItemSR1.Text = .ISR1
            txtItemSR2.Text = .ISR2
            txtItemSR3.Text = .ISR3
            txtItemSR4.Text = .ISR4
            txtItemSR5.Text = .ISR5
            txtItemSR6.Text = .ISR6

            KP_R.Text = "R: " & .KomposisiPortofolio_R
            KP_G.Text = "G: " & .KomposisiPortofolio_G
            KP_B.Text = "B: " & .KomposisiPortofolio_B
            txtKP.Text = .KomposisiPortofolio

            BEDP_R.Text = "R: " & .BEDP_R
            BEDP_G.Text = "G: " & .BEDP_G
            BEDP_B.Text = "B: " & .BEDP_B
            txtBEDP.Text = .BEDP

            IBEDP_R.Text = "R: " & .IBEDP_R
            IBEDP_G.Text = "G: " & .IBEDP_G
            IBEDP_B.Text = "B: " & .IBEDP_B

            RI_R.Text = "R: " & .RisikoInvestasi_R
            RI_R.Text = "G: " & .RisikoInvestasi_G
            RI_R.Text = "B: " & .RisikoInvestasi_B
            txtRI.Text = "R: " & .RisikoInvestasi

            IRI_R.Text = "R: " & .ValueRisikoInvestasi_R
            IRI_G.Text = "G: " & .ValueRisikoInvestasi_G
            IRI_B.Text = "B: " & .ValueRisikoInvestasi_B
            txtItemRI1.Text = .IRI1
            txtItemRI2.Text = .IRI2
            txtItemRI3.Text = .IRI3
            txtItemRI4.Text = .IRI4
            txtItemRI5.Text = .IRI5
            txtItemRI6.Text = .IRI6
            txtItemRI7.Text = .IRI7

            KR_R.Text = "R: " & .KlasifikasiRisiko_R
            KR_G.Text = "G: " & .KlasifikasiRisiko_G
            KR_B.Text = "B: " & .KlasifikasiRisiko_B
            txtKR.Text = .KlasifikasiRisiko

            MMI_R.Text = "R: " & .MengenaiManajerInvestasi_R
            MMI_G.Text = "G: " & .MengenaiManajerInvestasi_G
            MMI_B.Text = "B: " & .MengenaiManajerInvestasi_B
            txtMMI.Text = .MengenaiManajerInvestasi

            IMMI_R.Text = "R: " & .ValueMengenaiManajerInvestasi_R
            IMMI_G.Text = "G: " & .ValueMengenaiManajerInvestasi_G
            IMMI_B.Text = "B: " & .ValueMengenaiManajerInvestasi_B
            'txtIMMI1.Text = .ValueMengenaiManajerInvestasi

            ChartTitle_R.Text = "R: " & .ChartTitle_R
            ChartTitle_G.Text = "G: " & .ChartTitle_G
            ChartTitle_B.Text = "B: " & .ChartTitle_B
            txtColorGKRD.BackColor = Color.FromArgb(.ChartTitle_R, .ChartTitle_G, .ChartTitle_B)
            txtGKRD.Text = .ChartTitle

            ChartBorder_R.Text = "R: " & .ChartBorder_R
            ChartBorder_G.Text = "G: " & .ChartBorder_G
            ChartBorder_B.Text = "B: " & .ChartBorder_B
            txtColorChartBorder.BackColor = Color.FromArgb(.ChartBorder_R, .ChartBorder_G, .ChartBorder_B)
            chkChartBorder.Checked = .ChartBorder

            ChartLine_R.Text = "R: " & .ChartLine_R
            ChartLine_G.Text = "G: " & .ChartLine_G
            ChartLine_B.Text = "B: " & .ChartLine_B
            txtColorChartLine.BackColor = Color.FromArgb(.ChartLine_R, .ChartLine_G, .ChartLine_B)
            txtAxisX.Text = .ChartAxisX
            txtAxisY.Text = .ChartAxisY

            TableHeader_R.Text = "R: " & .TableHeader_R
            TableHeader_G.Text = "G: " & .TableHeader_G
            TableHeader_B.Text = "B: " & .TableHeader_B
            txtColorTableHeader.BackColor = Color.FromArgb(.TableHeader_R, .TableHeader_G, .TableHeader_B)

            TableItems_R.Text = "R: " & .TableItem_R
            TableItems_G.Text = "G: " & .TableItem_G
            TableItems_B.Text = "B: " & .TableItem_B
            txtColorTableItems.BackColor = Color.FromArgb(.TableItem_R, .TableItem_G, .TableItem_B)

            UP_R.Text = "R: " & .UlasanPasar_R
            UP_G.Text = "G: " & .UlasanPasar_G
            UP_B.Text = "B: " & .UlasanPasar_B
            'txtUP.Text = .UlasanPasar
        End With
    End Sub

    Private Sub rbOption2_CheckedChanged(sender As Object, e As EventArgs) Handles rbOption2.CheckedChanged
        iniCheck()
    End Sub
#Region "setting"
    Private Sub colorSet()
        If cd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            If rbTujuanInvestasi.Checked Then
                txtColorTujuanInvestasi.BackColor = cd.Color
                TujuanInvestasi_R.Text = RGBWrite("R", cd.Color.R)
                TujuanInvestasi_G.Text = RGBWrite("G", cd.Color.G)
                TujuanInvestasi_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbTanggal.Checked Then
                txtColorTanggal.BackColor = cd.Color
                Tanggal_R.Text = RGBWrite("R", cd.Color.R)
                Tanggal_G.Text = RGBWrite("G", cd.Color.G)
                Tanggal_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbReportLine.Checked Then
                txtColoeReportLine.BackColor = cd.Color
                ReportLine_R.Text = RGBWrite("R", cd.Color.R)
                ReportLine_G.Text = RGBWrite("G", cd.Color.G)
                ReportLine_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbIRD.Checked Then
                txtColorIRD.BackColor = cd.Color
                IRD_R.Text = RGBWrite("R", cd.Color.R)
                IRD_G.Text = RGBWrite("G", cd.Color.G)
                IRD_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbItemIRD.Checked Then
                txtColorItemIRD.BackColor = cd.Color
                IIRD_R.Text = RGBWrite("R", cd.Color.R)
                IIRD_G.Text = RGBWrite("G", cd.Color.G)
                IIRD_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbKK.Checked Then
                txtColorKK.BackColor = cd.Color
                KK_R.Text = RGBWrite("R", cd.Color.R)
                KK_G.Text = RGBWrite("G", cd.Color.G)
                KK_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbKI.Checked Then
                txtColorKI.BackColor = cd.Color
                KI_R.Text = RGBWrite("R", cd.Color.R)
                KI_G.Text = RGBWrite("G", cd.Color.G)
                KI_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbAS.Checked Then
                txtColorAS.BackColor = cd.Color
                AS_R.Text = RGBWrite("R", cd.Color.R)
                AS_G.Text = RGBWrite("G", cd.Color.G)
                AS_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbIBB.Checked Then
                txtColorIBB.BackColor = cd.Color
                IBB_R.Text = RGBWrite("R", cd.Color.R)
                IBB_G.Text = RGBWrite("G", cd.Color.G)
                IBB_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbIIBB.Checked Then
                txtColorItemIBB.BackColor = cd.Color
                IIBB_R.Text = RGBWrite("R", cd.Color.R)
                IIBB_G.Text = RGBWrite("G", cd.Color.G)
                IIBB_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbSR.Checked Then
                txtColorSR.BackColor = cd.Color
                SR_R.Text = RGBWrite("R", cd.Color.R)
                SR_G.Text = RGBWrite("G", cd.Color.G)
                SR_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbItemSR.Checked Then
                txtColorItemSR.BackColor = cd.Color
                ISR_R.Text = RGBWrite("R", cd.Color.R)
                ISR_G.Text = RGBWrite("G", cd.Color.G)
                ISR_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbKP.Checked Then
                txtColorKP.BackColor = cd.Color
                KP_R.Text = RGBWrite("R", cd.Color.R)
                KP_G.Text = RGBWrite("G", cd.Color.G)
                KP_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbRI.Checked Then
                txtColorRI.BackColor = cd.Color
                RI_R.Text = RGBWrite("R", cd.Color.R)
                RI_G.Text = RGBWrite("G", cd.Color.G)
                RI_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbIRI.Checked Then
                TxtColorItemRI.BackColor = cd.Color
                IRI_R.Text = RGBWrite("R", cd.Color.R)
                IRI_G.Text = RGBWrite("G", cd.Color.G)
                IRI_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbKR.Checked Then
                txtColorKR.BackColor = cd.Color
                KR_R.Text = RGBWrite("R", cd.Color.R)
                KR_G.Text = RGBWrite("G", cd.Color.G)
                KR_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbMMI.Checked Then
                txtColorMMI.BackColor = cd.Color
                MMI_R.Text = RGBWrite("R", cd.Color.R)
                MMI_G.Text = RGBWrite("G", cd.Color.G)
                MMI_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbIMMI.Checked Then
                txtColorIMMI.BackColor = cd.Color
                IMMI_R.Text = RGBWrite("R", cd.Color.R)
                IMMI_G.Text = RGBWrite("G", cd.Color.G)
                IMMI_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbChartTitle.Checked Then
                txtColorGKRD.BackColor = cd.Color
                ChartTitle_R.Text = RGBWrite("R", cd.Color.R)
                ChartTitle_G.Text = RGBWrite("G", cd.Color.G)
                ChartTitle_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbChartBorder.Checked Then
                txtColorChartBorder.BackColor = cd.Color
                ChartBorder_R.Text = RGBWrite("R", cd.Color.R)
                ChartBorder_G.Text = RGBWrite("G", cd.Color.G)
                ChartBorder_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbChartLine.Checked Then
                txtColorChartLine.BackColor = cd.Color
                ChartLine_R.Text = RGBWrite("R", cd.Color.R)
                ChartLine_G.Text = RGBWrite("G", cd.Color.G)
                ChartLine_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbChartTitle.Checked Then
                txtColorGKRD.BackColor = cd.Color
                ChartTitle_R.Text = RGBWrite("R", cd.Color.R)
                ChartTitle_G.Text = RGBWrite("G", cd.Color.G)
                ChartTitle_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf chkChartBorder.Checked Then
                If rbChartBorder.Checked Then
                    txtColorChartBorder.BackColor = cd.Color
                    ChartBorder_R.Text = RGBWrite("R", cd.Color.R)
                    ChartBorder_G.Text = RGBWrite("G", cd.Color.G)
                    ChartBorder_B.Text = RGBWrite("B", cd.Color.B)
                End If
            ElseIf rbChartLine.Checked Then
                txtColorChartLine.BackColor = cd.Color
                ChartLine_R.Text = RGBWrite("R", cd.Color.R)
                ChartLine_G.Text = RGBWrite("G", cd.Color.G)
                ChartLine_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbTableHeader.Checked Then
                txtColorTableHeader.BackColor = cd.Color
                TableHeader_R.Text = RGBWrite("R", cd.Color.R)
                TableHeader_G.Text = RGBWrite("G", cd.Color.G)
                TableHeader_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbBEDP.Checked Then
                txtColorBEDP.BackColor = cd.Color
                BEDP_R.Text = RGBWrite("R", cd.Color.R)
                BEDP_G.Text = RGBWrite("G", cd.Color.G)
                BEDP_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbItemBEDP.Checked Then
                txtColorIBEDP.BackColor = cd.Color
                IBEDP_R.Text = RGBWrite("R", cd.Color.R)
                IBEDP_G.Text = RGBWrite("G", cd.Color.G)
                IBEDP_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbTableItems.Checked Then
                txtColorTableItems.BackColor = cd.Color
                TableItems_R.Text = RGBWrite("R", cd.Color.R)
                TableItems_G.Text = RGBWrite("G", cd.Color.G)
                TableItems_B.Text = RGBWrite("B", cd.Color.B)
            ElseIf rbUP.Checked Then
                txtColorUP.BackColor = cd.Color
                UP_R.Text = RGBWrite("R", cd.Color.R)
                UP_G.Text = RGBWrite("G", cd.Color.G)
                UP_B.Text = RGBWrite("B", cd.Color.B)
            End If
        End If
    End Sub
#End Region

   
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
                file.WriteInteger(reportSection, iniType & " Tujuan Investasi R", RGBRead(TujuanInvestasi_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tujuan Investasi G", RGBRead(TujuanInvestasi_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tujuan Investasi B", RGBRead(TujuanInvestasi_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Tujuan Investasi", txtTujuanInvestasi.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Tanggal R", RGBRead(Tanggal_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tanggal G", RGBRead(Tanggal_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Tanggal B", RGBRead(Tanggal_B.Text.Trim))

                file.WriteInteger(reportSection, iniType & " Report Line R", RGBRead(ReportLine_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Report Line G", RGBRead(ReportLine_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Report Line B", RGBRead(ReportLine_B.Text.Trim))

                file.WriteInteger(reportSection, iniType & " IRD R", RGBRead(IRD_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " IRD G", RGBRead(IRD_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " IRD B", RGBRead(IRD_B.Text.Trim))
                file.WriteString(reportSection, iniType & " IRD", txtInformasiReksaDana.Text.Trim)

                file.WriteInteger(reportSection, iniType & " IIRD R", RGBRead(IIRD_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " IIRD G", RGBRead(IIRD_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " IIRD B", RGBRead(IIRD_B.Text.Trim))

                file.WriteString(reportSection, iniType & " IIRD1", txtItemIRD1.Text.Trim)
                file.WriteString(reportSection, iniType & " IIRD2", txtItemIRD2.Text.Trim)
                file.WriteString(reportSection, iniType & " IIRD3", txtItemIRD3.Text.Trim)
                file.WriteString(reportSection, iniType & " IIRD4", txtItemIRD4.Text.Trim)
                file.WriteString(reportSection, iniType & " IIRD5", txtItemIRD5.Text.Trim)
                file.WriteString(reportSection, iniType & " IIRD6", txtItemIRD6.Text.Trim)
                file.WriteString(reportSection, iniType & " IIRD7", txtItemIRD7.Text.Trim)
                file.WriteString(reportSection, iniType & " IIRD8", txtItemIRD8.Text.Trim)

                file.WriteInteger(reportSection, iniType & " KK R", RGBRead(KK_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " KK G", RGBRead(KK_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " KK B", RGBRead(KK_B.Text.Trim))
                file.WriteString(reportSection, iniType & " KK", txtInformasiReksaDana.Text.Trim)

                file.WriteInteger(reportSection, iniType & " KI R", RGBRead(KI_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " KI G", RGBRead(KI_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " KI B", RGBRead(KI_B.Text.Trim))
                file.WriteString(reportSection, iniType & " KI", "")

                file.WriteInteger(reportSection, iniType & " AS R", RGBRead(AS_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " AS G", RGBRead(AS_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " AS B", RGBRead(AS_B.Text.Trim))
                file.WriteString(reportSection, iniType & " AS", txtAS.Text.Trim)

                file.WriteInteger(reportSection, iniType & " IBB R", RGBRead(IBB_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " IBB G", RGBRead(IBB_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " IBB B", RGBRead(IBB_B.Text.Trim))
                file.WriteString(reportSection, iniType & " IBB", txtIBB.Text.Trim)

                file.WriteString(reportSection, iniType & " IIBB1", txtIBB1.Text.Trim)
                file.WriteString(reportSection, iniType & " IIBB2", txtIBB2.Text.Trim)
                file.WriteString(reportSection, iniType & " IIBB3", txtIBB3.Text.Trim)
                file.WriteString(reportSection, iniType & " IIBB4", txtIBB4.Text.Trim)
                file.WriteString(reportSection, iniType & " IIBB5", txtIBB5.Text.Trim)
                file.WriteString(reportSection, iniType & " IIBB6", txtIBB6.Text.Trim)

                file.WriteInteger(reportSection, iniType & " SR R", RGBRead(SR_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " SR G", RGBRead(SR_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " SR B", RGBRead(SR_B.Text.Trim))
                file.WriteString(reportSection, iniType & " SR", txtSR.Text.Trim)

                file.WriteString(reportSection, iniType & " ISR1", txtItemSR1.Text.Trim)
                file.WriteString(reportSection, iniType & " ISR2", txtItemSR2.Text.Trim)
                file.WriteString(reportSection, iniType & " ISR3", txtItemSR3.Text.Trim)
                file.WriteString(reportSection, iniType & " ISR4", txtItemSR4.Text.Trim)
                file.WriteString(reportSection, iniType & " ISR5", txtItemSR5.Text.Trim)
                file.WriteString(reportSection, iniType & " ISR6", txtItemSR6.Text.Trim)

                file.WriteInteger(reportSection, iniType & " KP R", RGBRead(KP_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " KP G", RGBRead(KP_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " KP B", RGBRead(KP_B.Text.Trim))
                file.WriteString(reportSection, iniType & " KP", "")

                file.WriteInteger(reportSection, iniType & " BEDP R", RGBRead(BEDP_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " BEDP G", RGBRead(BEDP_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " BEDP B", RGBRead(BEDP_B.Text.Trim))
                file.WriteString(reportSection, iniType & " BEDP", txtBEDP.Text.Trim)

                file.WriteInteger(reportSection, iniType & " IBEDP R", RGBRead(IBEDP_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " IBEDP G", RGBRead(IBEDP_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " IBEDP B", RGBRead(IBEDP_B.Text.Trim))

                file.WriteInteger(reportSection, iniType & " RI R", RGBRead(RI_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " RI G", RGBRead(RI_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " RI B", RGBRead(RI_B.Text.Trim))
                file.WriteString(reportSection, iniType & " RI", txtRI.Text.Trim)

                file.WriteString(reportSection, iniType & " IRI1", txtItemRI1.Text.Trim)
                file.WriteString(reportSection, iniType & " IRI2", txtItemRI2.Text.Trim)
                file.WriteString(reportSection, iniType & " IRI3", txtItemRI3.Text.Trim)
                file.WriteString(reportSection, iniType & " IRI4", txtItemRI4.Text.Trim)
                file.WriteString(reportSection, iniType & " IRI5", txtItemRI5.Text.Trim)
                file.WriteString(reportSection, iniType & " IRI6", txtItemRI6.Text.Trim)
                file.WriteString(reportSection, iniType & " IRI7", txtItemRI7.Text.Trim)

                file.WriteInteger(reportSection, iniType & " KR R", RGBRead(KR_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " KR G", RGBRead(KR_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " KR B", RGBRead(KR_B.Text.Trim))
                file.WriteString(reportSection, iniType & " KR", txtKR.Text.Trim)

                file.WriteInteger(reportSection, iniType & " MMI R", RGBRead(MMI_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " MMI G", RGBRead(MMI_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " MMI B", RGBRead(MMI_B.Text.Trim))
                file.WriteString(reportSection, iniType & " MMI", txtMMI.Text.Trim)

                file.WriteInteger(reportSection, iniType & " IMMI R", RGBRead(IMMI_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " IMMI G", RGBRead(IMMI_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " IMMI B", RGBRead(IMMI_B.Text.Trim))
                'file.WriteString(reportSection, iniType & " IMMI", txtIMMI1.Text.Trim)

                file.WriteInteger(reportSection, iniType & " UP R", RGBRead(UP_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " UP G", RGBRead(UP_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " UP B", RGBRead(UP_B.Text.Trim))
                'file.WriteString(reportSection, iniType & " UP", txtUP.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Chart Title R", RGBRead(ChartTitle_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Title G", RGBRead(ChartTitle_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Title B", RGBRead(ChartTitle_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Investment Growth", txtGKRD.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Chart Border R", RGBRead(ChartBorder_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Border G", RGBRead(ChartBorder_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Border B", RGBRead(ChartBorder_B.Text.Trim))
                file.WriteBoolean(reportSection, iniType & " Chart Border", chkChartBorder.Checked)

                file.WriteInteger(reportSection, iniType & " Chart Line R", RGBRead(ChartLine_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Line G", RGBRead(ChartLine_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Chart Line B", RGBRead(ChartLine_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Chart Label 1", txtAxisX.Text.Trim)
                file.WriteString(reportSection, iniType & " Chart Label 2", txtAxisY.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Table Header R", RGBRead(TableHeader_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Header G", RGBRead(TableHeader_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Header B", RGBRead(TableHeader_B.Text.Trim))
                file.WriteString(reportSection, iniType & " Table Header", txtHeader.Text.Trim)

                file.WriteInteger(reportSection, iniType & " Table Items R", RGBRead(TableItems_R.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Items G", RGBRead(TableItems_G.Text.Trim))
                file.WriteInteger(reportSection, iniType & " Table Items B", RGBRead(TableItems_B.Text.Trim))

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

    Private Sub rbTujuanInvestasi_Click(sender As Object, e As EventArgs) Handles rbTujuanInvestasi.Click
        colorSet()
    End Sub

    Private Sub rbTanggal_Click(sender As Object, e As EventArgs) Handles rbTanggal.Click
        colorSet()
    End Sub

    Private Sub rbIRD_Click(sender As Object, e As EventArgs) Handles rbIRD.Click
        colorSet()
    End Sub

    Private Sub rbItemIRD_Click(sender As Object, e As EventArgs) Handles rbItemIRD.Click
        colorSet()
    End Sub

    Private Sub rbReportLine_Click(sender As Object, e As EventArgs) Handles rbReportLine.Click
        colorSet()
    End Sub

    Private Sub rbAS_Click(sender As Object, e As EventArgs) Handles rbAS.Click
        colorSet()
    End Sub

    Private Sub rbIBB_Click(sender As Object, e As EventArgs) Handles rbIBB.Click
        colorSet()
    End Sub

    Private Sub rbIIBB_Click(sender As Object, e As EventArgs) Handles rbIIBB.Click
        colorSet()
    End Sub

    Private Sub rbSR_Click(sender As Object, e As EventArgs) Handles rbSR.Click
        colorSet()
    End Sub

    Private Sub rbItemSR_Click(sender As Object, e As EventArgs) Handles rbItemSR.Click
        colorSet()
    End Sub

    Private Sub rbKP_Click(sender As Object, e As EventArgs) Handles rbKP.Click
        colorSet()
    End Sub

    Private Sub rbBEDP_Click(sender As Object, e As EventArgs) Handles rbBEDP.Click
        colorSet()
    End Sub

    Private Sub rbItemBEDP_Click(sender As Object, e As EventArgs) Handles rbItemBEDP.Click
        colorSet()
    End Sub

    Private Sub rbRI_Click(sender As Object, e As EventArgs) Handles rbRI.Click
        colorSet()
    End Sub

    Private Sub rbIRI_Click(sender As Object, e As EventArgs) Handles rbIRI.Click
        colorSet()
    End Sub

    Private Sub rbKR_Click(sender As Object, e As EventArgs) Handles rbKR.Click
        colorSet()
    End Sub

    Private Sub rbMMI_Click(sender As Object, e As EventArgs) Handles rbMMI.Click
        colorSet()
    End Sub

    Private Sub rbIMMI_Click(sender As Object, e As EventArgs) Handles rbIMMI.Click
        colorSet()
    End Sub

    Private Sub rbChartTitle_Click(sender As Object, e As EventArgs) Handles rbChartTitle.Click
        colorSet()
    End Sub

    Private Sub rbChartBorder_Click(sender As Object, e As EventArgs) Handles rbChartBorder.Click
        colorSet()
    End Sub

    Private Sub rbChartLine_Click(sender As Object, e As EventArgs) Handles rbChartLine.Click
        colorSet()
    End Sub

    Private Sub rbTableHeader_Click(sender As Object, e As EventArgs) Handles rbTableHeader.Click
        colorSet()
    End Sub

    Private Sub rbTableItems_Click(sender As Object, e As EventArgs) Handles rbTableItems.Click
        colorSet()
    End Sub

    Private Sub rbUP_Click(sender As Object, e As EventArgs) Handles rbUP.Click
        colorSet()
    End Sub

    Private Sub rbKI_Click(sender As Object, e As EventArgs) Handles rbKI.Click
        colorSet()
    End Sub

    Private Sub rbKK_Click(sender As Object, e As EventArgs) Handles rbKK.Click
        colorSet()
    End Sub
End Class