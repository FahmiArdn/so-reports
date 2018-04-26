Imports simpi.SimpiAccess
Imports simpi.ParameterSimpi
Imports simpi.GlobalUtilities
Imports System.Net.NetworkInformation
Imports simpi.Codeset

Module ModMain
    Friend objAccess As New UserAccess
    Friend objError As New ExceptionError
    Friend AppsCode As String = "so.simpi-pro"
    Friend AppsVersion As String = "1.0.0"
    Friend AppsTerminal As String = ""
    Friend LicenseKey As String = ""

    Friend objSimpi As New simpi.SimpiMaster.MasterSimpi
    Friend objMasterSimpi As New simpi.MasterSimpi.MasterSimpi
    Friend objUser As New simpi.SimpiMaster.SimpiUser
    Friend objMasterUser As New simpi.MasterSimpi.SimpiUser
    Friend objLayout As New simpi.SimpiMaster.SimpiCodeset

    Friend dtUserStatus As New DataTable
    Friend dtUserPortfolioCcy As New DataTable
    Friend dtMasterSales As New DataTable
    Friend dtSimpiCodeset As New DataTable
    Friend dtSimpiKYC As New DataTable
    Friend dtSimpiRisk As New DataTable

    Friend Function getString(ByVal stringData As String, ByVal ParamArray location() As Integer) As String
        Dim s As String = ""
        If location.Length <= 0 Then
            Return ""
            Exit Function
        End If
        For i As Integer = 0 To UBound(location, 1)
            s += stringData.Substring(location(i) - 1, 1)
        Next i
        Return s
    End Function

    Friend Function _key(ByVal TerminalID As String, ByVal LicenseKey As String) As String
        Dim stringKey As String = ""
        stringKey &= getString(LicenseKey, 12)
        stringKey &= getString(TerminalID, 1)
        stringKey &= getString(LicenseKey, 1)
        stringKey &= getString(LicenseKey, 8)
        stringKey &= getString(LicenseKey, 56)
        stringKey &= getString(LicenseKey, 101)
        stringKey &= getString(TerminalID, 2)
        stringKey &= getString(TerminalID, 4)
        stringKey &= getString(LicenseKey, 10)
        stringKey &= getString(TerminalID, 5)
        stringKey &= getString(TerminalID, 7)
        stringKey &= getString(LicenseKey, 94)
        stringKey &= getString(TerminalID, 8)
        stringKey &= getString(LicenseKey, 22)
        stringKey &= getString(LicenseKey, 125)
        stringKey &= getString(TerminalID, 10)
        stringKey &= getString(LicenseKey, 11)
        stringKey &= getString(TerminalID, 11)
        stringKey &= getString(LicenseKey, 33)
        stringKey &= getString(TerminalID, 13)
        stringKey &= getString(LicenseKey, 31)
        stringKey &= getString(TerminalID, 14)
        stringKey &= getString(TerminalID, 16)
        stringKey &= getString(TerminalID, 17)
        Return stringKey
    End Function

    Friend Function _iv(ByVal TerminalID As String, ByVal LicenseKey As String) As String
        Dim stringIV As String = ""
        stringIV &= getString(TerminalID, 4)
        stringIV &= getString(LicenseKey, 12)
        stringIV &= getString(TerminalID, 8)
        stringIV &= getString(LicenseKey, 94)
        stringIV &= getString(LicenseKey, 33)
        stringIV &= getString(LicenseKey, 10)
        stringIV &= getString(TerminalID, 13)
        stringIV &= getString(LicenseKey, 56)
        Return stringIV
    End Function

    Friend Function GetMacAddress() As String
        Dim nics() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces
        Dim s As String = ""
        For Each nc As NetworkInterface In nics
            If nc.OperationalStatus = OperationalStatus.Up Then
                s = nc.GetPhysicalAddress.ToString
                Exit For
            End If
        Next
        Dim t As String = ""
        If s.Length > 0 Then
            t = s.Substring(0, 2)
            For i As Integer = 2 To s.Length - 1
                t &= "-" & s.Substring(i, 2)
                i += 1
            Next
        End If
        Return t
    End Function

    Friend Function GetMacAddressList() As String
        Dim nics() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces
        Dim s As String = ""
        Dim t As String = ""
        Dim list As String = "List of TerminalID:" & Environment.NewLine
        Dim j As Integer = 1
        For Each nc As NetworkInterface In nics
            s = nc.GetPhysicalAddress.ToString
            If s.Length > 0 Then
                t = s.Substring(0, 2)
                For i As Integer = 2 To s.Length - 1
                    t &= "-" & s.Substring(i, 2)
                    i += 1
                Next
                list &= j & ". " & t & Environment.NewLine
                j += 1
            End If
        Next
        Return list
    End Function

    Friend Sub GetComboInit(ByVal obj As Object, ByVal cmb As C1.Win.C1InputPanel.InputComboBox, ByVal Value As String, ByVal Display As String)
        Dim tbl As New DataTable
        obj.Clear()
        obj.UserAccess = objAccess
        tbl = obj.Search("")
        cmb.Items.Clear()
        If obj.ErrID = 0 Then
            cmb.DataSource = Nothing
            cmb.DisplayMember = Display
            cmb.ValueMember = Value
            cmb.DataSource = tbl.Select(Display & " <> ''", Display).CopyToDataTable
        End If
    End Sub

    Friend Sub GetComboInitAll(ByVal obj As Object, ByVal cmb As C1.Win.C1InputPanel.InputComboBox, ByVal Value As String, ByVal Display As String)
        Dim tbl As New DataTable
        obj.Clear()
        obj.UserAccess = objAccess
        tbl = obj.Search("")
        cmb.Items.Clear()
        If obj.ErrID = 0 Then
            cmb.DataSource = Nothing
            cmb.DisplayMember = Display
            cmb.ValueMember = Value
            cmb.DataSource = tbl
            Dim row As DataRow = tbl.NewRow
            row(Display) = "ALL"
            row(Value) = 0
            tbl.Rows.Add(row)
            cmb.Text = tbl.Rows(tbl.Rows.Count - 1).Item(Display)
        End If
    End Sub

    Friend Sub GetComboDatatable(ByVal tbl As DataTable, ByVal cmb As C1.Win.C1InputPanel.InputComboBox,
                            ByVal Value As String, ByVal Display As String, ByVal AddBlank As Boolean)
        cmb.Items.Clear()
        If tbl.Rows.Count > 0 Then
            cmb.DataSource = Nothing
            cmb.DisplayMember = Display
            cmb.ValueMember = Value
            cmb.DataSource = tbl
            If AddBlank Then
                Dim row As DataRow = tbl.NewRow
                row(Display) = "All"
                row(Value) = 0
                tbl.Rows.Add(row)
                cmb.Text = tbl.Rows(tbl.Rows.Count - 1).Item(Display)
            End If
        End If
    End Sub

    Friend Sub CheckEmptyID(ByVal objError As ExceptionError, ByVal data As String, ByVal id As Double)
        If id <= 0 Then Throw New ExceptionData(1, objError.Message(1) & " " & data)
    End Sub

    Friend Sub CheckEmptyString(ByVal objError As ExceptionError, ByVal data As String, ByVal str As String)
        If Trim(str) = "" Then Throw New ExceptionData(1, objError.Message(1) & " " & data)
    End Sub

    Public Sub GetMasterSimpi()
        If objSimpi.GetSimpiID = 0 Then
            objSimpi.UserAccess = objAccess
            objMasterSimpi.UserAccess = objAccess
            objSimpi.Load()
            objMasterSimpi.Load(objAccess.GetSimpiEmail)

            objUser.UserAccess = objAccess
            objMasterUser.UserAccess = objAccess
            objUser.Load(objAccess.GetUserLogin)
            objMasterUser.Load(objAccess.GetUserLogin)

            Dim objSales As New simpi.MasterSales.MasterSales
            objSales.UserAccess = objAccess
            dtMasterSales = objSales.SearchAll(objMasterSimpi, "")

            objLayout.UserAccess = objAccess
        End If
    End Sub

    Friend Function GetSimpiUser(ByVal strKeyword As String) As DataTable
        Dim tbl As New DataTable
        Dim objUser As New simpi.SimpiMaster.SimpiUser
        objUser.UserAccess = objAccess
        tbl = objUser.Search(strKeyword)
        Return tbl
    End Function

    Friend Function GetParameterUserStatus(ByVal strKeyword As String) As DataTable
        If dtUserStatus IsNot Nothing AndAlso dtUserStatus.Rows.Count > 0 Then
            Return dtUserStatus
        Else
            Dim objStatus As New ParameterUserStatus
            objStatus.UserAccess = objAccess
            dtUserStatus = objStatus.Search(strKeyword)
            dtUserStatus.CaseSensitive = False
            Return dtUserStatus
        End If
    End Function

    'Friend Function GetUserPortfolioCcy() As DataTable
    '    If dtUserPortfolioCcy IsNot Nothing AndAlso dtUserPortfolioCcy.Rows.Count > 0 Then
    '        Return dtUserPortfolioCcy
    '    Else
    '        Dim objNAV As New simpi.CoreData.PortfolioNAV
    '        objNAV.UserAccess = objAccess
    '        dtUserPortfolioCcy = objNAV.SearchCcy()
    '        Return dtUserStatus
    '    End If
    'End Function

    Friend Function GetSimpiCodesetField() As DataTable
        If dtSimpiCodeset IsNot Nothing AndAlso dtSimpiCodeset.Rows.Count > 0 Then
            Return dtSimpiCodeset
        Else
            Dim objField As New CodesetSimpiField
            objField.UserAccess = objAccess
            dtSimpiCodeset = objField.Search("")
            dtSimpiCodeset.CaseSensitive = False
            Return dtSimpiCodeset
        End If
    End Function

    Friend Sub GetSimpiKYCField(ByVal TypeID As Integer)
        Dim objKYC As New simpi.MasterSimpi.SimpiKYC
        objKYC.UserAccess = objAccess
        dtSimpiKYC = objKYC.Search(objMasterSimpi, TypeID, "")
        dtSimpiKYC.CaseSensitive = False
    End Sub

    Friend Function GetSimpiRisk(ByVal TypeID As Integer) As DataTable
        If dtSimpiRisk IsNot Nothing AndAlso dtSimpiRisk.Rows.Count > 0 Then
            Return dtSimpiRisk
        Else
            Dim objRisk As New simpi.MasterSimpi.SimpiRiskLevel
            objRisk.UserAccess = objAccess
            dtSimpiRisk = objRisk.Question_Search(objMasterSimpi, TypeID)
            dtSimpiRisk.CaseSensitive = False
            Return dtSimpiRisk
        End If
    End Function

End Module
