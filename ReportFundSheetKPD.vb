Imports simpi.GlobalUtilities
Imports simpi.CoreData
Imports simpi.MasterPortfolio
Imports simpi.GlobalConnection
Imports C1.Win.C1Chart
Imports System.Drawing.Imaging
Imports simpi.MasterSecurities
Imports C1.Win.C1TrueDBGrid
Imports simpi.MarketInstrument.ParameterSecurities

Public Class ReportFundSheetKPD
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
    Dim reportSection As String = "REPORT FUND SHEET KPD"
    Public pdfLayout As New pdfColor

#Region "pdf"
    Structure pdfColor
        Public layoutType As String
        'Left Column
        Public ReportTitle_R As Integer
        Public ReportTitle_G As Integer
        Public ReportTitle_B As Integer
        Public ReportTitle As String
        Public NAB_R As Integer
        Public NAB_G As Integer
        Public NAB_B As Integer
        Public NAB As String
        Public TglLaporan_R As Integer
        Public TglLaporan_G As Integer
        Public TglLaporan_B As Integer
        Public TglLaporan As String
        Public Bank_R As Integer
        Public Bank_G As Integer
        Public Bank_B As Integer
        Public Bank As String
        Public TglPeluncuran_R As Integer
        Public TglPeluncuran_G As Integer
        Public TglPeluncuran_B As Integer
        Public TglPeluncuran As String
        Public Total_R As Integer
        Public Total_G As Integer
        Public Total_B As Integer
        Public Total As String
        Public MataUang_R As Integer
        Public MataUang_G As Integer
        Public MataUang_B As Integer
        Public MataUang As String
        Public ImbalManajer_R As Integer
        Public ImbalManajer_G As Integer
        Public ImbalManajer_B As Integer
        Public ImbalManajer As String
        Public ImbalBank_R As Integer
        Public ImbalBank_G As Integer
        Public ImbalBank_B As Integer
        Public ImbalBank As String
        Public BiayaBeli_R As Integer
        Public BiayaBeli_G As Integer
        Public BiayaBeli_B As Integer
        Public BiayaBeli As String
        Public BiayaJual_R As Integer
        Public BiayaJual_G As Integer
        Public BiayaJual_B As Integer
        Public BiayaJual As String
        Public SwitchingFee_R As Integer
        Public SwitchingFee_G As Integer
        Public SwitchingFee_B As Integer
        Public SwitchingFee As String
        Public InvestasiSelanjutnya_R As Integer
        Public InvestasiSelanjutnya_G As Integer
        Public InvestasiSelanjutnya_B As Integer
        Public InvestasiSelanjutnya As String
        Public Kode_R As Integer
        Public Kode_G As Integer
        Public Kode_B As Integer
        Public Kode As String
        Public KinerjaTerbaik_R As Integer
        Public KinerjaTerbaik_G As Integer
        Public KinerjaTerbaik_B As Integer
        Public KinerjaTerbaik As String
        Public KinerjaTerburuk_R As Integer
        Public KinerjaTerburuk_G As Integer
        Public KinerjaTerburuk_B As Integer
        Public KinerjaTerburuk As String
        Public FaktorRisiko_R As Integer
        Public FaktorRisiko_G As Integer
        Public FaktorRisiko_B As Integer
        Public FaktorRisiko As String
        Public TotalReturn_R As Integer
        Public TotalReturn_G As Integer
        Public TotalReturn_B As Integer
        Public TotalReturn As String
        Public Fund_R As Integer
        Public Fund_G As Integer
        Public Fund_B As Integer
        Public Fund As String
        'Right Column
        Public TujuanInvestasi_R As Integer
        Public TujuanInvestasi_G As Integer
        Public TujuanInvestasi_B As Integer
        Public TujuanInvestasi As String
        Public Investasi_R As Integer
        Public Investasi_G As Integer
        Public Investasi_B As Integer
        Public Investasi As String
        Public InvestasiSaham As String
        Public investasiObligasi As String
        Public InvestasiPasarUang As String
        Public Portofolio_R As Integer
        Public Portofolio_G As Integer
        Public Portofolio_B As Integer
        Public Portofolio As String
        Public PortofolioSaham As String
        Public PortifolioObligasi As String
        Public PortofolioPasarUang As String
        Public KinerjaReksaDana_R As Integer
        Public KinerjaReksaDana_G As Integer
        Public KinerjaReksaDana_B As Integer
        Public KinerjaReksaDana As String
        Public ChartBorder_R As Integer
        Public ChartBorder_G As Integer
        Public ChartBorder_B As Integer
        Public ChartBorder As Boolean
        Public TitleKepemilikan_R As Integer
        Public TitleKepemilikan_G As Integer
        Public TitleKepemilikan_B As Integer
        Public TitleKepemilikan As String
        Public Filter_R As Integer
        Public Filter_G As Integer
        Public Filter_B As Integer
        Public Filter As String
        Public ChartTitle_R As Integer
        Public ChartTitle_G As Integer
        Public ChartTitle_B As Integer
        Public ChartTitle As String
        Public ChartBorderPie_R As Integer
        Public ChartBorderPie_G As Integer
        Public ChartBorderPie_B As Integer
        Public ChartBorderPie As Boolean
        'Table
        Public TableTitle_R As Integer
        Public TableTitle_G As Integer
        Public TableTitle_B As Integer
        Public TableTitle As String
        Public TableItem_R As Integer
        Public TableItem_G As Integer
        Public TableItem_B As Integer
        Public TableItemReturn As String
        Public TableItemBenchmark As String
        Public TableItem1Bulan As String
        Public TableItem3Bulan As String
        Public TableItem6Bulan As String
        Public TableItem1Tahun As String
        Public TableItem3Tahun As String
        Public TableItem5Tahun As String
        Public TableItemDariAwal As String
        Public TableItemSejakPembentukan As String
        Public OutlookHolding_R As Integer
        Public OutlookHolding_G As Integer
        Public OutlookHolding_B As Integer
        Public TentangHolding_R As Integer
        Public TentangHolding_G As Integer
        Public TentangHolding_B As Integer
    End Structure

    Public Sub pdfColorDefault()
        pdfLayout.layoutType = "DEFAULT"
        pdfLayout.ReportTitle_R = 0
        pdfLayout.ReportTitle_G = 61
        pdfLayout.ReportTitle_B = 121
        pdfLayout.ReportTitle = "FUND FACT SHEET"
        pdfLayout.NAB_R = 0
        pdfLayout.NAB_G = 61
        pdfLayout.NAB_B = 121
        pdfLayout.NAB = "Price IDR"
        pdfLayout.TglLaporan_R = 0
        pdfLayout.TglLaporan_G = 61
        pdfLayout.TglLaporan_B = 121
        pdfLayout.TglLaporan = "Report Date"
        pdfLayout.Bank_R = 0
        pdfLayout.Bank_G = 61
        pdfLayout.Bank_B = 121
        pdfLayout.Bank = "Custodian Bank"
        pdfLayout.TglPeluncuran_R = 0
        pdfLayout.TglPeluncuran_G = 61
        pdfLayout.TglPeluncuran_B = 121
        pdfLayout.TglPeluncuran = "Inception Date"
        pdfLayout.Total_R = 0
        pdfLayout.Total_G = 61
        pdfLayout.Total_B = 121
        pdfLayout.Total = "Asset Under Management"
        pdfLayout.MataUang_R = 0
        pdfLayout.MataUang_G = 61
        pdfLayout.MataUang_B = 121
        pdfLayout.MataUang = "Currency"
        pdfLayout.ImbalManajer_R = 0
        pdfLayout.ImbalManajer_G = 61
        pdfLayout.ImbalManajer_B = 121
        pdfLayout.ImbalManajer = "Manegement Fee"
        pdfLayout.ImbalBank_R = 0
        pdfLayout.ImbalBank_G = 61
        pdfLayout.ImbalBank_B = 121
        pdfLayout.ImbalBank = "Custodian Fee"
        pdfLayout.BiayaBeli_R = 0
        pdfLayout.BiayaBeli_G = 61
        pdfLayout.BiayaBeli_B = 121
        pdfLayout.BiayaBeli = "Subscription Fee"
        pdfLayout.BiayaJual_R = 0
        pdfLayout.BiayaJual_G = 61
        pdfLayout.BiayaJual_B = 121
        pdfLayout.BiayaJual = "Redemption Fee"
        pdfLayout.SwitchingFee_R = 0
        pdfLayout.SwitchingFee_G = 61
        pdfLayout.SwitchingFee_B = 121
        pdfLayout.SwitchingFee = "Switching Fee"
        pdfLayout.Kode_R = 0
        pdfLayout.Kode_G = 61
        pdfLayout.Kode_B = 121
        pdfLayout.Kode = "ISIN / Bloomberg Code"
        pdfLayout.KinerjaTerbaik_R = 0
        pdfLayout.KinerjaTerbaik_G = 61
        pdfLayout.KinerjaTerbaik_B = 121
        pdfLayout.KinerjaTerbaik = "Best Month"
        pdfLayout.KinerjaTerburuk_R = 0
        pdfLayout.KinerjaTerburuk_G = 61
        pdfLayout.KinerjaTerburuk_B = 121
        pdfLayout.KinerjaTerburuk = "Worst Month"
        pdfLayout.FaktorRisiko_R = 0
        pdfLayout.FaktorRisiko_G = 61
        pdfLayout.FaktorRisiko_B = 121
        pdfLayout.FaktorRisiko = "Main Risk Factor"
        pdfLayout.TotalReturn_R = 0
        pdfLayout.TotalReturn_G = 61
        pdfLayout.TotalReturn_B = 121
        pdfLayout.TotalReturn = "Total Return Since Inception"
        pdfLayout.Fund_R = 0
        pdfLayout.Fund_G = 61
        pdfLayout.Fund_B = 121
        pdfLayout.Fund = "Fund"
        'Right Column
        pdfLayout.TujuanInvestasi_R = 0
        pdfLayout.TujuanInvestasi_G = 61
        pdfLayout.TujuanInvestasi_B = 121
        pdfLayout.TujuanInvestasi = "Investment Objective"
        pdfLayout.Investasi_R = 0
        pdfLayout.Investasi_G = 61
        pdfLayout.Investasi_B = 121
        pdfLayout.Investasi = "Investment Policy"
        pdfLayout.InvestasiSaham = "Saham"
        pdfLayout.investasiObligasi = "Obligasi"
        pdfLayout.InvestasiPasarUang = "Pasar Uang"
        pdfLayout.Portofolio_R = 0
        pdfLayout.Portofolio_G = 61
        pdfLayout.Portofolio_B = 121
        pdfLayout.Portofolio = "Portofolio Allocation"
        pdfLayout.PortofolioSaham = "Saham"
        pdfLayout.PortifolioObligasi = "Obligasi"
        pdfLayout.PortofolioPasarUang = "Pasar Uang"
        pdfLayout.KinerjaReksaDana_R = 0
        pdfLayout.KinerjaReksaDana_G = 61
        pdfLayout.KinerjaReksaDana_B = 121
        pdfLayout.KinerjaReksaDana = "Fund Performance"
        pdfLayout.ChartBorder_R = 0
        pdfLayout.ChartBorder_G = 61
        pdfLayout.ChartBorder_B = 121
        pdfLayout.ChartBorder = False
        pdfLayout.TitleKepemilikan_R = 0
        pdfLayout.TitleKepemilikan_G = 61
        pdfLayout.TitleKepemilikan_B = 121
        pdfLayout.TitleKepemilikan = "Top Holdings"
        pdfLayout.Filter_R = 0
        pdfLayout.Filter_G = 61
        pdfLayout.Filter_B = 121
        pdfLayout.Filter = "In Alphabetical Order"
        pdfLayout.ChartTitle_R = 0
        pdfLayout.ChartTitle_G = 61
        pdfLayout.ChartTitle_B = 121
        pdfLayout.ChartTitle = "Monthly Return"
        pdfLayout.ChartBorderPie_R = 0
        pdfLayout.ChartBorderPie_G = 61
        pdfLayout.ChartBorderPie_B = 121
        pdfLayout.ChartBorderPie = False
        'Table
        pdfLayout.TableTitle_R = 0
        pdfLayout.TableTitle_G = 61
        pdfLayout.TableTitle_B = 121
        pdfLayout.TableTitle = "Performace - dd/mm/yyyy"
        pdfLayout.TableItem_R = 0
        pdfLayout.TableItem_G = 61
        pdfLayout.TableItem_B = 111
        pdfLayout.TableItemReturn = "Return"
        pdfLayout.TableItemBenchmark = "Benchmark"
        pdfLayout.TableItem1Bulan = "1 Bulan"
        pdfLayout.TableItem3Bulan = "3 Bulan"
        pdfLayout.TableItem6Bulan = "6 Bulan"
        pdfLayout.TableItem1Tahun = "1 Tahun"
        pdfLayout.TableItem3Tahun = "3 Tahun"
        pdfLayout.TableItem5Tahun = "5 Tahun"
        pdfLayout.TableItemDariAwal = "YTD"
        pdfLayout.TableItemSejakPembentukan = "Since Inc."
        pdfLayout.OutlookHolding_R = 0
        pdfLayout.OutlookHolding_G = 61
        pdfLayout.OutlookHolding_B = 121
        pdfLayout.TentangHolding_R = 0
        pdfLayout.TentangHolding_G = 61
        pdfLayout.TentangHolding_B = 121
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

                    pdfLayout.NAB_R = file.GetInteger(reportSection, iniType & " NAB/Unit R", 0)
                    pdfLayout.NAB_G = file.GetInteger(reportSection, iniType & " NAB/Unit G", 0)
                    pdfLayout.NAB_B = file.GetInteger(reportSection, iniType & " NAB/Unit B", 0)
                    pdfLayout.NAB = file.GetString(reportSection, iniType & " NAB/Unit", "")

                    pdfLayout.TglLaporan_R = file.GetInteger(reportSection, iniType & " Tanggal Laporan R", 0)
                    pdfLayout.TglLaporan_G = file.GetInteger(reportSection, iniType & " Tanggal Laporan G", 0)
                    pdfLayout.TglLaporan_B = file.GetInteger(reportSection, iniType & " Tanggal Laporan B", 0)
                    pdfLayout.TglLaporan = file.GetString(reportSection, iniType & " Tanggal Laporan", "")

                    pdfLayout.Bank_R = file.GetInteger(reportSection, iniType & " Bank Kustodian R", 0)
                    pdfLayout.Bank_G = file.GetInteger(reportSection, iniType & " Bank Kustodian G", 0)
                    pdfLayout.Bank_B = file.GetInteger(reportSection, iniType & " Bank Kustodian B", 0)
                    pdfLayout.Bank = file.GetString(reportSection, iniType & " Bank Kustodian", "")

                    pdfLayout.TglPeluncuran_R = file.GetInteger(reportSection, iniType & " Tanggal Peluncuran R", 0)
                    pdfLayout.TglPeluncuran_G = file.GetInteger(reportSection, iniType & " Tanggal Peluncuran G", 0)
                    pdfLayout.TglPeluncuran_B = file.GetInteger(reportSection, iniType & " Tanggal Peluncuran B", 0)
                    pdfLayout.TglPeluncuran = file.GetString(reportSection, iniType & " Tanggal Peluncuran", "")

                    pdfLayout.Total_R = file.GetInteger(reportSection, iniType & " Total AUM R", 0)
                    pdfLayout.Total_G = file.GetInteger(reportSection, iniType & " Total AUM G", 0)
                    pdfLayout.Total_B = file.GetInteger(reportSection, iniType & " Total AUM B", 0)
                    pdfLayout.Total = file.GetString(reportSection, iniType & " Total AUM", "")

                    pdfLayout.MataUang_R = file.GetInteger(reportSection, iniType & " Mata Uang R", 0)
                    pdfLayout.MataUang_G = file.GetInteger(reportSection, iniType & " Mata Uang G", 0)
                    pdfLayout.MataUang_B = file.GetInteger(reportSection, iniType & " Mata Uang B", 0)
                    pdfLayout.MataUang = file.GetString(reportSection, iniType & " Mata Uang", "")

                    pdfLayout.ImbalManajer_R = file.GetInteger(reportSection, iniType & " Imbal Jasa Manajer R", 0)
                    pdfLayout.ImbalManajer_G = file.GetInteger(reportSection, iniType & " Imbal Jasa Manajer G", 0)
                    pdfLayout.ImbalManajer_B = file.GetInteger(reportSection, iniType & " Imbal Jasa Manajer B", 0)
                    pdfLayout.ImbalManajer = file.GetString(reportSection, iniType & " Imbal Jasa Manajer Investasi", "")

                    pdfLayout.ImbalBank_R = file.GetInteger(reportSection, iniType & " Imbal Jasa Bank R", 0)
                    pdfLayout.ImbalBank_G = file.GetInteger(reportSection, iniType & " Imbal Jasa Bank G", 0)
                    pdfLayout.ImbalBank_B = file.GetInteger(reportSection, iniType & " Imbal Jasa Bank B", 0)
                    pdfLayout.ImbalBank = file.GetString(reportSection, iniType & " Imbal Jasa Bank Kustodian", "")

                    pdfLayout.BiayaBeli_R = file.GetInteger(reportSection, iniType & " Biaya Pembelian R", 0)
                    pdfLayout.BiayaBeli_G = file.GetInteger(reportSection, iniType & " Biaya Pembelian G", 0)
                    pdfLayout.BiayaBeli_B = file.GetInteger(reportSection, iniType & " Biaya Pembelian B", 0)
                    pdfLayout.BiayaBeli = file.GetString(reportSection, iniType & " Biaya Pembelian", "")

                    pdfLayout.BiayaJual_R = file.GetInteger(reportSection, iniType & " Biaya Penjualan R", 0)
                    pdfLayout.BiayaJual_G = file.GetInteger(reportSection, iniType & " Biaya Penjualan G", 0)
                    pdfLayout.BiayaJual_B = file.GetInteger(reportSection, iniType & " Biaya Penjualan B", 0)
                    pdfLayout.BiayaJual = file.GetString(reportSection, iniType & " Biaya Penjualan Kembali", "")

                    pdfLayout.SwitchingFee_R = file.GetInteger(reportSection, iniType & " Switching Fee R", 0)
                    pdfLayout.SwitchingFee_G = file.GetInteger(reportSection, iniType & " Switching Fee G", 0)
                    pdfLayout.SwitchingFee_B = file.GetInteger(reportSection, iniType & " Switching Fee B", 0)
                    pdfLayout.SwitchingFee = file.GetString(reportSection, iniType & " Switching Fee", "")

                    pdfLayout.InvestasiSelanjutnya_R = file.GetInteger(reportSection, iniType & " Investasi Selanjutnya R", 0)
                    pdfLayout.InvestasiSelanjutnya_G = file.GetInteger(reportSection, iniType & " Investasi Selanjutnya G", 0)
                    pdfLayout.InvestasiSelanjutnya_B = file.GetInteger(reportSection, iniType & " Investasi Selanjutnya B", 0)
                    pdfLayout.InvestasiSelanjutnya = file.GetString(reportSection, iniType & " Minimal Investasi Selanjutnya", "")

                    pdfLayout.Kode_R = file.GetInteger(reportSection, iniType & " Kode R", 0)
                    pdfLayout.Kode_G = file.GetInteger(reportSection, iniType & " Kode G", 0)
                    pdfLayout.Kode_B = file.GetInteger(reportSection, iniType & " Kode B", 0)
                    pdfLayout.Kode = file.GetString(reportSection, iniType & " Kode ISIN/Bloomberg", "")

                    pdfLayout.KinerjaTerbaik_R = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terbaik R", 0)
                    pdfLayout.KinerjaTerbaik_G = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terbaik G", 0)
                    pdfLayout.KinerjaTerbaik_B = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terbaik B", 0)
                    pdfLayout.KinerjaTerbaik = file.GetString(reportSection, iniType & " Kinerja Bulan Terbaik", "")

                    pdfLayout.KinerjaTerburuk_R = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terburuk R", 0)
                    pdfLayout.KinerjaTerburuk_G = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terburuk G", 0)
                    pdfLayout.KinerjaTerburuk_B = file.GetInteger(reportSection, iniType & " Kinerja Bulan Terburuk B", 0)
                    pdfLayout.KinerjaTerburuk = file.GetString(reportSection, iniType & " Kinerja Bulan Terburuk", "")

                    pdfLayout.FaktorRisiko_R = file.GetInteger(reportSection, iniType & " Faktor Risiko Utama R", 0)
                    pdfLayout.FaktorRisiko_G = file.GetInteger(reportSection, iniType & " Faktor Risiko Utama G", 0)
                    pdfLayout.FaktorRisiko_B = file.GetInteger(reportSection, iniType & " Faktor Risiko Utama B", 0)
                    pdfLayout.FaktorRisiko = file.GetString(reportSection, iniType & " Faktor Risiko Utama", "")

                    pdfLayout.TotalReturn_R = file.GetInteger(reportSection, iniType & " Periode Investasi R", 0)
                    pdfLayout.TotalReturn_G = file.GetInteger(reportSection, iniType & " Periode Investasi G", 0)
                    pdfLayout.TotalReturn_B = file.GetInteger(reportSection, iniType & " Periode Investasi B", 0)
                    pdfLayout.TotalReturn = file.GetString(reportSection, iniType & " Periode Investasi", "")

                    pdfLayout.Fund_R = file.GetInteger(reportSection, iniType & " Fund R", 0)
                    pdfLayout.Fund_G = file.GetInteger(reportSection, iniType & " Fund G", 0)
                    pdfLayout.Fund_B = file.GetInteger(reportSection, iniType & " Fund B", 0)
                    pdfLayout.Fund = file.GetString(reportSection, iniType & " Fund", "")
                    'Right Column
                    pdfLayout.TujuanInvestasi_R = file.GetInteger(reportSection, iniType & " Tujuan Investasi R", 0)
                    pdfLayout.TujuanInvestasi_G = file.GetInteger(reportSection, iniType & " Tujuan Investasi G", 0)
                    pdfLayout.TujuanInvestasi_B = file.GetInteger(reportSection, iniType & " Tujuan Investasi B", 0)
                    pdfLayout.TujuanInvestasi = file.GetString(reportSection, iniType & " Tujuan Investasi", "")

                    pdfLayout.Investasi_R = file.GetInteger(reportSection, iniType & " Investasi R", 0)
                    pdfLayout.Investasi_G = file.GetInteger(reportSection, iniType & " Investasi G", 0)
                    pdfLayout.Investasi_B = file.GetInteger(reportSection, iniType & " Investasi B", 0)
                    pdfLayout.Investasi = file.GetString(reportSection, iniType & " Kebijakan Investasi", "")
                    pdfLayout.InvestasiSaham = file.GetString(reportSection, iniType & " Saham", "")
                    pdfLayout.investasiObligasi = file.GetString(reportSection, iniType & " Obligasi", "")
                    pdfLayout.InvestasiPasarUang = file.GetString(reportSection, iniType & " Pasar Uang", "")

                    pdfLayout.Portofolio_R = file.GetInteger(reportSection, iniType & " Portofolio R", 0)
                    pdfLayout.Portofolio_G = file.GetInteger(reportSection, iniType & " Portofolio G", 0)
                    pdfLayout.Portofolio_B = file.GetInteger(reportSection, iniType & " Portofolio B", 0)
                    pdfLayout.Portofolio = file.GetString(reportSection, iniType & " Komposisi Portofolio", "")
                    pdfLayout.PortofolioSaham = file.GetString(reportSection, iniType & " Saham", "")
                    pdfLayout.PortifolioObligasi = file.GetString(reportSection, iniType & " Obligasi", "")
                    pdfLayout.PortofolioPasarUang = file.GetString(reportSection, iniType & " Pasar Uang", "")

                    pdfLayout.KinerjaReksaDana_R = file.GetInteger(reportSection, iniType & " Line Chart Title R", 0)
                    pdfLayout.KinerjaReksaDana_G = file.GetInteger(reportSection, iniType & " Line Chart Title G", 0)
                    pdfLayout.KinerjaReksaDana_B = file.GetInteger(reportSection, iniType & " Line Chart Title B", 0)
                    pdfLayout.KinerjaReksaDana = file.GetString(reportSection, iniType & " Line Chart Title", "")

                    pdfLayout.ChartBorder_R = file.GetInteger(reportSection, iniType & " Line Chart Border R", 0)
                    pdfLayout.ChartBorder_G = file.GetInteger(reportSection, iniType & " Line Chart Border G", 0)
                    pdfLayout.ChartBorder_B = file.GetInteger(reportSection, iniType & " Line Chart Border B", 0)
                    If file.GetBoolean(reportSection, iniType & " Line Chart Border", False) Then pdfLayout.ChartBorder = True Else pdfLayout.ChartBorder = False

                    pdfLayout.TitleKepemilikan_R = file.GetInteger(reportSection, iniType & " Title Kepemilikan R", 0)
                    pdfLayout.TitleKepemilikan_G = file.GetInteger(reportSection, iniType & " Title Kepemilikan G", 0)
                    pdfLayout.TitleKepemilikan_B = file.GetInteger(reportSection, iniType & " Title Kepemilikan B", 0)
                    pdfLayout.TitleKepemilikan = file.GetString(reportSection, iniType & " Title Kepemilikan", "")

                    pdfLayout.Filter_R = file.GetInteger(reportSection, iniType & " Filter R", 0)
                    pdfLayout.Filter_G = file.GetInteger(reportSection, iniType & " Filter G", 0)
                    pdfLayout.Filter_B = file.GetInteger(reportSection, iniType & " Filter B", 0)
                    pdfLayout.Filter = file.GetString(reportSection, iniType & " Filter", "")

                    pdfLayout.ChartTitle_R = file.GetInteger(reportSection, iniType & " Pie Chart Title R", 0)
                    pdfLayout.ChartTitle_G = file.GetInteger(reportSection, iniType & " Pie Chart Title G", 0)
                    pdfLayout.ChartTitle_B = file.GetInteger(reportSection, iniType & " Pie Chart Title B", 0)
                    pdfLayout.ChartTitle = file.GetString(reportSection, iniType & " Pie Chart Title", "")

                    pdfLayout.ChartBorderPie_R = file.GetInteger(reportSection, iniType & " Pie Chart Border R", 0)
                    pdfLayout.ChartBorderPie_G = file.GetInteger(reportSection, iniType & " Pie Chart Border G", 0)
                    pdfLayout.ChartBorderPie_B = file.GetInteger(reportSection, iniType & " Pie Chart Border B", 0)
                    If file.GetBoolean(reportSection, iniType & " Pie Chart Border", False) Then pdfLayout.ChartBorder = True Else pdfLayout.ChartBorder = False
                    'Table
                    pdfLayout.TableTitle_R = file.GetInteger(reportSection, iniType & " Table Title R", 0)
                    pdfLayout.TableTitle_G = file.GetInteger(reportSection, iniType & " Table Title G", 0)
                    pdfLayout.TableTitle_B = file.GetInteger(reportSection, iniType & " Table Title B", 0)
                    pdfLayout.TableTitle = file.GetString(reportSection, iniType & " Table Title", "")

                    pdfLayout.TableItem_R = file.GetInteger(reportSection, iniType & " Table Item R", 0)
                    pdfLayout.TableItem_G = file.GetInteger(reportSection, iniType & " Table Item G", 0)
                    pdfLayout.TableItem_B = file.GetInteger(reportSection, iniType & " Table Item B", 0)
                    pdfLayout.TableItemReturn = file.GetString(reportSection, iniType & " Return", "")
                    pdfLayout.TableItemBenchmark = file.GetString(reportSection, iniType & " Benchmark", "")
                    pdfLayout.TableItem1Bulan = file.GetString(reportSection, iniType & " 1 Bulan", "")
                    pdfLayout.TableItem3Bulan = file.GetString(reportSection, iniType & " 3 Bulan", "")
                    pdfLayout.TableItem6Bulan = file.GetString(reportSection, iniType & " 6 Bulan", "")
                    pdfLayout.TableItem1Tahun = file.GetString(reportSection, iniType & " 1 Tahun", "")
                    pdfLayout.TableItem3Tahun = file.GetString(reportSection, iniType & " 3 Tahun", "")
                    pdfLayout.TableItem5Tahun = file.GetString(reportSection, iniType & " 5 Tahun", "")
                    pdfLayout.TableItemDariAwal = file.GetString(reportSection, iniType & " Dari Awal", "")
                    pdfLayout.TableItemSejakPembentukan = file.GetString(reportSection, iniType & " Sejak Pembentukan", "")

                    pdfLayout.OutlookHolding_R = file.GetInteger(reportSection, iniType & " Title Ulasan R", 0)
                    pdfLayout.OutlookHolding_G = file.GetInteger(reportSection, iniType & " Title Ulasan G", 0)
                    pdfLayout.OutlookHolding_B = file.GetInteger(reportSection, iniType & " Title Ulasan B", 0)

                    pdfLayout.TentangHolding_R = file.GetInteger(reportSection, iniType & " Title MI R", 0)
                    pdfLayout.TentangHolding_G = file.GetInteger(reportSection, iniType & " Title MI G", 0)
                    pdfLayout.TentangHolding_B = file.GetInteger(reportSection, iniType & " Title MI B", 0)
                End If
            Else
                pdfColorDefault()
            End If
        Catch ex As Exception
            pdfColorDefault()
        End Try
    End Sub
#End Region

    Private Sub ReportFundSheetEQ_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetInstrumentUser()
        GetParameterInstrumentType()
        dtAs.Value = Now
        pdfSetting()
        objPortfolio.UserAccess = objAccess
        objSimpi.UserAccess = objAccess
        objCodeset.UserAccess = objAccess
        objNAV.UserAccess = objAccess

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
                    txtCurrency.Text = objPortfolio.GetPortfolioCcy.GetCcyDescription
                    txtInception.Text = objPortfolio.GetInceptionDate.ToString("dd-MMMM-yyyy")
                    txtCustodian.Text = objCodeset.GetCodeset(objPortfolio, 2)
                    Dim tmp As String = objCodeset.GetCodeset(objPortfolio, 3)
                    tmp = objPortfolio.ExternalID_Get(2)
                    txtISIN.Text = IIf(tmp.Trim = "", "-", tmp.Trim)
                    tmp = objPortfolio.ExternalID_Get(5)
                    txtBloomberg.Text = IIf(tmp.Trim = "", "-", tmp.Trim)
                    tmp = objCodeset.GetCodeset(objPortfolio, 5)
                    tmp = objCodeset.GetCodeset(objPortfolio, 6)
                    txtInvestmentGoal.Text = IIf(tmp.Trim = "", "-", tmp.Trim)
                    txtPolicyEQ.Text = objCodeset.GetCodeset(objPortfolio, 7)
                    txtPolicyFI.Text = objCodeset.GetCodeset(objPortfolio, 8)
                    txtPolicyMM.Text = objCodeset.GetCodeset(objPortfolio, 9)
                    dtSector.Clear()
                    objSector.Clear()
                    dtSector = objSector.Company_Member(objPortfolio.GetPortfolioSector.GetClassID, 0, 0)
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

                objReview.Clear()
                objReview.Load(objPortfolio, objNAV.GetPositionDate)
                'If objReview.ErrID = 0 Then
                '    lblReviewID.Text = objReview.GetReviewID
                '    lblLastReview.Text = objReview.GetReviewDate.ToString("dd-MMM-yyy")
                '    txtReviewText.Text = objReview.GetReviewText
                'Else
                '    reviewClear()
                'End If

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
        With chartSector
            .Style.Border.BorderStyle = BorderStyleEnum.None
            Dim ds As ChartDataSeriesCollection = .ChartGroups(0).ChartData.SeriesList
            ds.Clear()
            Dim series As ChartDataSeries = ds.AddNewSeries()
            'series.LineStyle.Color = Color.Green
            'series.LineStyle.Thickness = 1
            'series.SymbolStyle.Shape = SymbolShapeEnum.Box
            'series.FitType = FitTypeEnum.Line

            series.X.CopyDataIn((From q In dtNAV.AsEnumerable Order By q.Field(Of Date)("PositionDate") Ascending Select q.Field(Of Date)("PositionDate")).ToArray)
            series.Y.CopyDataIn((From q In dtNAV.AsEnumerable Order By q.Field(Of Date)("PositionDate") Ascending Select q.Field(Of Decimal)("GeometricIndex") - 1).ToArray)
            series.Y1.CopyDataIn((From q In dtBenchmark.AsEnumerable Order By q.Field(Of Date)("PositionDate") Ascending Select q.Field(Of Decimal)("GeometricIndex") - 1).ToArray)
            series.PointData.Length = dtNAV.Rows.Count
            'series.PointData.Length = dtBenchmark.Rows.Count

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
    End Sub 'Bar

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
        Dim myBrush As New SolidBrush(Color.FromArgb(0, 61, 121))
        Dim detailBrush As New SolidBrush(Color.Black)
        Dim headerBrush As New SolidBrush(Color.White)
        Dim koordX As Single = 40, koordY As Single = 35
        Dim fontType = "Calibri", fontSize = 8
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
            pdfSetting()

            rc = New RectangleF(koordX, koordY, 150, fontSize)
            .DrawStringRtf("{\b " & pdfLayout.ReportTitle & ",} " & dtAs.Value.Day & "-" & MonthName(dtAs.Value.Month, False) & " " & dtAs.Value.Year, New Font(fontType, 10),
                           New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)), rc)

            koordY += 23
            .DrawLine(New Pen(Color.FromArgb(251, 188, 47), 1), New PointF(koordX, koordY), New PointF(koordX + 300, koordY))
            koordY += fontSize
            Dim point = New PointF(koordX, koordY)
            point.Y += fontSize + 30
            .DrawString(lblPortfolioName.Text, New Font(fontType, 22, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)), New PointF(koordX, koordY))

            If objPortfolio.GetPortfolioAccount.GetAccountID = 2 Then
                str = "Reksa Dana"
            ElseIf objPortfolio.GetPortfolioAccount.GetAccountID = 1 Then
                str = "Kontrak Pengelolalan Dana"
            End If
            If objPortfolio.GetAssetType.GetAssetTypeCode = "EQ" Then
                str &= " Saham"
            ElseIf objPortfolio.GetAssetType.GetAssetTypeCode = "FI" Then
                str &= " Pendapatan Tetap"
            ElseIf objPortfolio.GetAssetType.GetAssetTypeCode = "MIX" Then
                str &= " Campuran"
            ElseIf objPortfolio.GetAssetType.GetAssetTypeCode = "MM" Then
                str &= " Pasar Uang"
            ElseIf objPortfolio.GetAssetType.GetAssetTypeCode = "CPF" Then
                str &= " Terproteksi"
            End If
            koordY += 30
            .DrawString(str, New Font(fontType, 11),
                        New SolidBrush(Color.FromArgb(pdfLayout.ReportTitle_R, pdfLayout.ReportTitle_G, pdfLayout.ReportTitle_B)), New PointF(koordX, koordY))

            koordY += 20
            .DrawLine(New Pen(Color.FromArgb(224, 225, 226), 2), New PointF(koordX + 140, koordY + 10), New PointF(koordX + 140, koordY + 490))

            koordX += 160
            .DrawString(pdfLayout.TujuanInvestasi, New Font(fontType, 16),
                        New SolidBrush(Color.FromArgb(pdfLayout.TujuanInvestasi_R, pdfLayout.TujuanInvestasi_G, pdfLayout.TujuanInvestasi_B)), New PointF(koordX, koordY + 5))
            rc = New RectangleF(koordX, koordY + 25, 350, fontSize * 3)
            .DrawString(txtInvestmentGoal.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TujuanInvestasi_R, pdfLayout.TujuanInvestasi_G, pdfLayout.TujuanInvestasi_B)), rc)

            .DrawString(pdfLayout.Investasi, New Font(fontType, 16),
                        New SolidBrush(Color.FromArgb(pdfLayout.Investasi_R, pdfLayout.Investasi_G, pdfLayout.Investasi_B)), New PointF(koordX, koordY + 50))
            .DrawString(pdfLayout.Portofolio, New Font(fontType, 16),
                        New SolidBrush(Color.FromArgb(pdfLayout.Portofolio_R, pdfLayout.Portofolio_G, pdfLayout.Portofolio_B)), New PointF(koordX + 215, koordY + 50))

            .DrawString(pdfLayout.InvestasiSaham, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Investasi_R, pdfLayout.Investasi_G, pdfLayout.Investasi_B)), New PointF(koordX, koordY + 70))
            .DrawString(": " & txtPolicyEQ.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Investasi_R, pdfLayout.Investasi_G, pdfLayout.Investasi_B)), New PointF(koordX + 70, koordY + 70))
            .DrawString(pdfLayout.PortofolioSaham, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Portofolio_R, pdfLayout.Portofolio_G, pdfLayout.Portofolio_B)), New PointF(koordX + 215, koordY + 70))
            .DrawString(": " & txtCompositionEQ.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Portofolio_R, pdfLayout.Portofolio_G, pdfLayout.Portofolio_B)), New PointF(koordX + 285, koordY + 70))
            .DrawString(pdfLayout.investasiObligasi, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Investasi_R, pdfLayout.Investasi_G, pdfLayout.Investasi_B)), New PointF(koordX, koordY + 80))
            .DrawString(": " & txtPolicyFI.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Investasi_R, pdfLayout.Investasi_G, pdfLayout.Investasi_B)), New PointF(koordX + 70, koordY + 80))
            .DrawString(pdfLayout.PortifolioObligasi, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Portofolio_R, pdfLayout.Portofolio_G, pdfLayout.Portofolio_B)), New PointF(koordX + 215, koordY + 80))
            .DrawString(": " & txtCompositionFI.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Portofolio_R, pdfLayout.Portofolio_G, pdfLayout.Portofolio_B)), New PointF(koordX + 285, koordY + 80))
            .DrawString(pdfLayout.InvestasiPasarUang, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Investasi_R, pdfLayout.Investasi_G, pdfLayout.Investasi_B)), New PointF(koordX, koordY + 90))
            .DrawString(": " & txtPolicyMM.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Investasi_R, pdfLayout.Investasi_G, pdfLayout.Investasi_B)), New PointF(koordX + 70, koordY + 90))
            .DrawString(pdfLayout.PortofolioPasarUang, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Portofolio_R, pdfLayout.Portofolio_G, pdfLayout.Portofolio_B)), New PointF(koordX + 215, koordY + 90))
            .DrawString(": " & txtCompositionMM.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Portofolio_R, pdfLayout.Portofolio_G, pdfLayout.Portofolio_B)), New PointF(koordX + 285, koordY + 90))

            .DrawString(pdfLayout.KinerjaReksaDana, New Font(fontType, 16),
                        New SolidBrush(Color.FromArgb(pdfLayout.KinerjaReksaDana_R, pdfLayout.KinerjaReksaDana_G, pdfLayout.KinerjaReksaDana_B)), New PointF(koordX, koordY + 100))

            chartNAV.ChartArea.AxisX.GridMajor.Visible = False
            chartNAV.ChartArea.AxisY.GridMajor.Visible = False
            If pdfLayout.ChartBorder Then
                chartNAV.BorderStyle = BorderStyle.FixedSingle
                chartNAV.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.Solid
                chartNAV.ChartArea.Style.Border.Color = Color.FromArgb(pdfLayout.ChartBorder_R, pdfLayout.ChartBorder_G, pdfLayout.ChartBorder_B)
            Else
                chartNAV.Style.Border.BorderStyle = BorderStyle.None
                chartNAV.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.None
            End If
            Dim imgPortfolio = chartNAV.GetImage(ImageFormat.Emf)
            rc = New RectangleF(New PointF(koordX, koordY + 120), New SizeF(0.6 * chartNAV.Size.Width, 0.6 * chartNAV.Size.Height))
            .DrawImage(imgPortfolio, rc, ContentAlignment.TopLeft, C1.C1Pdf.ImageSizeModeEnum.Scale)
            'If pdfLayout.ChartBorder Then
            '    Dim pnNAV As New Pen(New SolidBrush(Color.FromArgb(pdfLayout.ChartBorder_R, pdfLayout.ChartBorder_G, pdfLayout.ChartBorder_B)), 0.5)
            '    .DrawRectangle(pnNAV, rc)
            'Else
            '    .DrawRectangle(Pens.White, rc)
            'End If

            .DrawString(pdfLayout.TitleKepemilikan, New Font(fontType, 16),
                        New SolidBrush(Color.FromArgb(pdfLayout.TitleKepemilikan_R, pdfLayout.TitleKepemilikan_G, pdfLayout.TitleKepemilikan_B)), New PointF(koordX, koordY + 270))
            .DrawString(pdfLayout.ChartTitle, New Font(fontType, 16),
                        New SolidBrush(Color.FromArgb(pdfLayout.ChartTitle_R, pdfLayout.ChartTitle_G, pdfLayout.ChartTitle_B)), New PointF(koordX + 215, koordY + 270))
            .DrawString(pdfLayout.Filter, New Font(fontType, 7),
                        New SolidBrush(Color.FromArgb(pdfLayout.Filter_R, pdfLayout.Filter_G, pdfLayout.Filter_B)), New PointF(koordX, koordY + 290))

            If DBGHolding.RowCount > 5 Then efek = 5 Else efek = DBGHolding.RowCount - 1
            Dim pos = koordY + 300
            For i = 0 To efek
                DBGHolding.Row = i

                rc = New RectangleF(koordX, pos, 150, 10)
                str = DBGHolding.Columns("Name").Text
                .DrawString(str, New Font(fontType, fontSize),
                            New SolidBrush(Color.FromArgb(pdfLayout.TitleKepemilikan_R, pdfLayout.TitleKepemilikan_G, pdfLayout.TitleKepemilikan_B)), rc)

                rc = New RectangleF(koordX + 150, pos, 150, 10)
                str = DBGHolding.Columns("Type").Text
                .DrawString(str, New Font(fontType, fontSize),
                            New SolidBrush(Color.FromArgb(pdfLayout.TitleKepemilikan_R, pdfLayout.TitleKepemilikan_G, pdfLayout.TitleKepemilikan_B)), rc)

                pos += fontSize + 1
            Next

            'If pdfLayout.ChartBorderPie Then
            '    chartSector.BorderStyle = BorderStyle.FixedSingle
            '    chartSector.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.Solid
            '    chartSector.ChartArea.Style.Border.Color = Color.FromArgb(pdfLayout.ChartBorderPie_R, pdfLayout.ChartBorderPie_G, pdfLayout.ChartBorderPie_B)
            'Else
            '    chartSector.BorderStyle = BorderStyle.None
            '    chartSector.ChartArea.Style.Border.BorderStyle = BorderStyleEnum.None
            'End If
            Dim imgSector = chartSector.GetImage(ImageFormat.Png)
            rc = New RectangleF(New PointF(koordX + 220, koordY + 290), New SizeF(0.4 * chartSector.Size.Width, 0.4 * chartSector.Size.Height))
            c1pdf.DrawImage(imgSector, rc, ContentAlignment.TopCenter, C1.C1Pdf.ImageSizeModeEnum.Scale)
            If pdfLayout.ChartBorder Then
                Dim pnNAV As New Pen(New SolidBrush(Color.FromArgb(pdfLayout.ChartBorderPie_R, pdfLayout.ChartBorderPie_G, pdfLayout.ChartBorderPie_B)), 0.5)
                .DrawRectangle(pnNAV, rc)
            Else
                .DrawRectangle(Pens.White, rc)
            End If

            Dim sf As StringFormat = New StringFormat()
            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Center
            .DrawString(pdfLayout.TableTitle, New Font(fontType, 16),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableTitle_R, pdfLayout.TableTitle_G, pdfLayout.TableTitle_B)), New PointF(koordX, koordY + 390))
            .DrawLine(New Pen(Color.FromArgb(0, 61, 121), 0.5), New PointF(koordX, koordY + 410), New PointF(koordX + 375, koordY + 410))
            .DrawString(pdfLayout.TableItem1Bulan, New Font(fontType, 7, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 45, koordY + 420, 65, 10), sf)
            .DrawString(pdfLayout.TableItem3Bulan, New Font(fontType, 7, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 85, koordY + 420, 65, 10), sf)
            .DrawString(pdfLayout.TableItem6Bulan, New Font(fontType, 7, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 125, koordY + 420, 65, 10), sf)
            .DrawString(pdfLayout.TableItem1Tahun, New Font(fontType, 7, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 165, koordY + 420, 65, 10), sf)
            .DrawString(pdfLayout.TableItem3Tahun, New Font(fontType, 7, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 205, koordY + 420, 65, 10), sf)
            .DrawString(pdfLayout.TableItem5Tahun, New Font(fontType, 7, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 245, koordY + 420, 65, 10), sf)
            .DrawString(pdfLayout.TableItemDariAwal, New Font(fontType, 7, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 285, koordY + 420, 65, 10), sf)
            .DrawString(pdfLayout.TableItemSejakPembentukan, New Font(fontType, 7, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 325, koordY + 420, 65, 10), sf)

            DBGPerformance.Row = 0
            .DrawString(pdfLayout.TableItemReturn, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New PointF(koordX, koordY + 430))
            str = DBGPerformance.Columns("1Mo").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 45, koordY + 430, 65, 10), sf)
            str = DBGPerformance.Columns("3Mo").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 85, koordY + 430, 65, 10), sf)
            str = DBGPerformance.Columns("6Mo").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 125, koordY + 430, 65, 10), sf)
            str = DBGPerformance.Columns("1Y").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 165, koordY + 430, 65, 10), sf)
            str = DBGPerformance.Columns("3Y").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 205, koordY + 430, 65, 10), sf)
            str = DBGPerformance.Columns("5Y").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 245, koordY + 430, 65, 10), sf)
            str = DBGPerformance.Columns("YTD").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 285, koordY + 430, 65, 10), sf)
            str = DBGPerformance.Columns("Inception").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 325, koordY + 430, 65, 10), sf)

            DBGPerformance.Row = 1
            .DrawString(pdfLayout.TableItemSejakPembentukan, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New PointF(koordX, koordY + 445))
            str = DBGPerformance.Columns("1Mo").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 45, koordY + 445, 65, 10), sf)
            str = DBGPerformance.Columns("3Mo").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 85, koordY + 445, 65, 10), sf)
            str = DBGPerformance.Columns("6Mo").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 125, koordY + 445, 65, 10), sf)
            str = DBGPerformance.Columns("1Y").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 165, koordY + 445, 65, 10), sf)
            str = DBGPerformance.Columns("3Y").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 205, koordY + 445, 65, 10), sf)
            str = DBGPerformance.Columns("5Y").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 245, koordY + 445, 65, 10), sf)
            str = DBGPerformance.Columns("YTD").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 285, koordY + 445, 65, 10), sf)
            str = DBGPerformance.Columns("Inception").Text
            .DrawString(str, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TableItem_R, pdfLayout.TableItem_G, pdfLayout.TableItem_B)), New Rectangle(koordX + 325, koordY + 445, 65, 10), sf)
            koordX -= 160

            koordY += 10
            .DrawString("" & pdfLayout.NAB & " " & txtNAVPerUnit.Text, New Font(fontType, 10),
                        New SolidBrush(Color.FromArgb(pdfLayout.NAB_R, pdfLayout.NAB_G, pdfLayout.NAB_B)), New PointF(koordX, koordY - 10))

            koordY += fontSize + 6 'Tanggal Laporan
            .DrawString(pdfLayout.TglLaporan, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TglLaporan_R, pdfLayout.TglLaporan_G, pdfLayout.TglLaporan_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(dtAs.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TglLaporan_R, pdfLayout.TglLaporan_G, pdfLayout.TglLaporan_B)), New PointF(koordX, koordY))

            koordY += fontSize + 6
            .DrawString(pdfLayout.Bank, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.Bank_R, pdfLayout.Bank_G, pdfLayout.Bank_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(txtCustodian.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Bank_R, pdfLayout.Bank_G, pdfLayout.Bank_B)), New PointF(koordX, koordY))

            koordY += fontSize + 6
            .DrawString(pdfLayout.TglPeluncuran, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TglPeluncuran_R, pdfLayout.TglPeluncuran_G, pdfLayout.TglPeluncuran_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(txtInception.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.TglPeluncuran_R, pdfLayout.TglPeluncuran_G, pdfLayout.TglPeluncuran_B)), New PointF(koordX, koordY))

            koordY += fontSize + 6
            .DrawString(pdfLayout.Total, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.Total_R, pdfLayout.Total_G, pdfLayout.Total_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(txtAUM.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Total_R, pdfLayout.Total_G, pdfLayout.Total_B)), New PointF(koordX, koordY))

            koordY += fontSize + 6
            .DrawString(pdfLayout.MataUang, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.MataUang_R, pdfLayout.MataUang_G, pdfLayout.MataUang_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(txtCurrency.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.MataUang_R, pdfLayout.MataUang_G, pdfLayout.MataUang_B)), New PointF(koordX, koordY))

            koordY += fontSize + 6
            .DrawString(pdfLayout.ImbalManajer, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ImbalManajer_R, pdfLayout.ImbalManajer_G, pdfLayout.ImbalBank_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(txtManagementFee.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.ImbalManajer_R, pdfLayout.ImbalBank_G, pdfLayout.ImbalBank_B)), New PointF(koordX, koordY))

            koordY += fontSize + 6
            .DrawString(pdfLayout.ImbalBank, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.ImbalBank_R, pdfLayout.ImbalBank_G, pdfLayout.ImbalBank_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(txtCustodianFee.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.ImbalBank_R, pdfLayout.ImbalBank_G, pdfLayout.ImbalBank_B)), New PointF(koordX, koordY))

            koordY += fontSize + 6
            .DrawString(pdfLayout.BiayaBeli, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.BiayaBeli_R, pdfLayout.BiayaBeli_G, pdfLayout.BiayaBeli_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(txtSellingFee.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.BiayaBeli_R, pdfLayout.BiayaBeli_G, pdfLayout.BiayaBeli_B)), New PointF(koordX, koordY))

            koordY += fontSize + 6
            .DrawString(pdfLayout.BiayaJual, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.BiayaJual_R, pdfLayout.BiayaBeli_G, pdfLayout.BiayaJual_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(txtRedemptionFee.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.BiayaJual_R, pdfLayout.BiayaJual_G, pdfLayout.BiayaJual_B)), New PointF(koordX, koordY))

            koordY += fontSize + 6
            .DrawString(pdfLayout.SwitchingFee, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.SwitchingFee_R, pdfLayout.SwitchingFee_G, pdfLayout.SwitchingFee_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(txtSwitchingFee.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Kode_R, pdfLayout.Kode_G, pdfLayout.Kode_B)), New PointF(koordX, koordY))

            koordY += fontSize + 6
            .DrawString(pdfLayout.Kode, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.Kode_R, pdfLayout.Kode_G, pdfLayout.Kode_B)), New PointF(koordX, koordY))
            koordY += fontSize + 1
            .DrawString(txtISIN.Text & " / " & txtBloomberg.Text, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Kode_R, pdfLayout.Kode_G, pdfLayout.Kode_B)), New PointF(koordX, koordY))

            Dim factorList As String() = {txtRisk1.Text, txtRisk2.Text,
                                          txtRisk3.Text, txtRisk4.Text,
                                          txtRisk5.Text, txtRisk6.Text, txtRisk7.Text}
            koordY += fontSize + 6
            rc = New RectangleF(koordX, koordY, 140, 20)
            .DrawString(pdfLayout.FaktorRisiko, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.FaktorRisiko_R, pdfLayout.FaktorRisiko_G, pdfLayout.FaktorRisiko_B)), rc)
            koordY += fontSize
            For i As Integer = 0 To factorList.Length - 1
                rc = New RectangleF(koordX, koordY, 120, 50)
                If factorList(i) IsNot "" Then
                    .DrawString("- " & factorList(i), New Font(fontType, fontSize),
                            New SolidBrush(Color.FromArgb(pdfLayout.FaktorRisiko_R, pdfLayout.FaktorRisiko_G, pdfLayout.FaktorRisiko_B)), rc)
                    koordY = If(factorList(i).Length > 27, koordY + 2 * (fontSize + 1), koordY + fontSize)
                End If
            Next

            sf.LineAlignment = StringAlignment.Center
            sf.Alignment = StringAlignment.Center
            koordY += fontSize + 1
            .DrawString(pdfLayout.TotalReturn, New Font(fontType, fontSize, FontStyle.Bold),
                        New SolidBrush(Color.FromArgb(pdfLayout.TotalReturn_R, pdfLayout.TotalReturn_G, pdfLayout.TotalReturn_B)), New PointF(koordX, koordY))
            koordY += 12
            Dim fundBr1 = New SolidBrush(Color.FromArgb(60, 107, 178)),
                bestBr2 = New SolidBrush(Color.FromArgb(96, 126, 189)),
                worstBr3 = New SolidBrush(Color.FromArgb(130, 149, 203))
            .DrawString(pdfLayout.Fund, New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.Fund_R, pdfLayout.Fund_G, pdfLayout.Fund_B)), New Rectangle(koordX + 10, koordY, 35, fontSize + 2))
            .FillRectangle(fundBr1, New Rectangle(koordX + 85, koordY, 35, fontSize + 2))
            .DrawString(txtTotalReturn.Text, New Font(fontType, fontSize), Brushes.WhiteSmoke, New Rectangle(koordX + 85, koordY, 35, fontSize + 2), sf)
            .DrawString(pdfLayout.KinerjaTerbaik & "*", New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.KinerjaTerbaik_R, pdfLayout.KinerjaTerbaik_G, pdfLayout.KinerjaTerbaik_B)), New Rectangle(koordX + 10, koordY + 15, 60, fontSize + 2))
            .FillRectangle(bestBr2, New Rectangle(koordX + 85, koordY + 15, 35, fontSize + 2))
            .DrawString(txtBestMonth.Text, New Font(fontType, fontSize), Brushes.WhiteSmoke, New Rectangle(koordX + 85, koordY + 15, 35, fontSize + 2), sf)
            .DrawString(pdfLayout.KinerjaTerburuk & "**", New Font(fontType, fontSize),
                        New SolidBrush(Color.FromArgb(pdfLayout.KinerjaTerburuk_R, pdfLayout.KinerjaTerburuk_G, pdfLayout.KinerjaTerburuk_B)), New Rectangle(koordX + 10, koordY + 30, 60, fontSize + 2))
            .FillRectangle(worstBr3, New Rectangle(koordX + 85, koordY + 30, 35, fontSize + 2))
            .DrawString(txtWorstMonth.Text, New Font(fontType, fontSize), Brushes.WhiteSmoke, New Rectangle(koordX + 85, koordY + 30, 35, fontSize + 2), sf)
            .DrawString(" *" & txtBestMonthDate.Text, New Font(fontType, fontSize - 2),
                        New SolidBrush(Color.FromArgb(pdfLayout.KinerjaTerbaik_R, pdfLayout.KinerjaTerbaik_G, pdfLayout.KinerjaTerbaik_B)), New Rectangle(koordX + 10, koordY + 45, 35, fontSize + 2))
            .DrawString("**" & txtWorstMonthDate.Text, New Font(fontType, fontSize - 2),
                        New SolidBrush(Color.FromArgb(pdfLayout.KinerjaTerburuk_R, pdfLayout.KinerjaTerburuk_G, pdfLayout.KinerjaTerburuk_B)), New Rectangle(koordX + 10, koordY + 60, 35, fontSize + 2))

            koordY = 620
            .DrawLine(New Pen(Color.FromArgb(224, 225, 226), 2), New PointF(koordX, koordY), New PointF(koordX + 530, koordY))
            koordY += fontSize + 1
            .DrawString(txtTitleUlasan.Text, New Font(fontType, 13),
                        New SolidBrush(Color.FromArgb(pdfLayout.OutlookHolding_R, pdfLayout.OutlookHolding_G, pdfLayout.OutlookHolding_B)), New PointF(koordX, koordY))
            Dim c As Integer = 0
            koordY += 18

            rc = New RectangleF(koordX, koordY, 530, 100)
            .DrawString(txtUlasan.Text, New Font(fontType, 8),
                        New SolidBrush(Color.FromArgb(pdfLayout.OutlookHolding_R, pdfLayout.OutlookHolding_G, pdfLayout.OutlookHolding_B)), rc)

            koordY += 45
            .DrawString(txtMITitle.Text, New Font(fontType, 13),
                        New SolidBrush(Color.FromArgb(pdfLayout.TentangHolding_R, pdfLayout.TentangHolding_G, pdfLayout.TentangHolding_B)), New PointF(koordX, koordY))
            koordY += 18
            rc = New RectangleF(koordX, koordY, 530, 100)
            .DrawString(txtMIDesc.Text, New Font(fontType, 8),
                        New SolidBrush(Color.FromArgb(pdfLayout.TentangHolding_R, pdfLayout.TentangHolding_G, pdfLayout.TentangHolding_B)), rc)

            strFile = reportFileExists("ReportFS" & lblPortfolioCode.Text.Trim & dtAs.Value.ToString("yyyymmdd") & ".pdf")
            .Save(strFile)
            If Not isAttachment Then Process.Start(strFile)
        End With
        Return strFile
    End Function

    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click
        ReportSetting()
    End Sub
    Private Sub ReportSetting()
        Dim frm As New ReportFundSheetKPDSetting
        frm.frm = Me
        frm.Show()
        frm.FormLoad()
        frm.MdiParent = MDIMENU
    End Sub

#End Region
End Class