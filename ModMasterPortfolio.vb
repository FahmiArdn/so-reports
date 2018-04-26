Imports simpi.MasterPortfolio
Imports simpi.ParameterPortfolio
Imports simpi.ParameterFA
Imports simpi.Codeset

Module ModMasterPortfolio
    Friend dtPortfolioSimpi As New DataTable
    Friend dtPortfolioUser As New DataTable
    Friend dtBankAccountSimpi As New DataTable
    Friend dtBankAccountUser As New DataTable
    Friend dtPortfolioCodeset As New DataTable
    Friend dtPortfolioExternalID As New DataTable

    Friend dtParameterPortfolioCodeset As New DataTable
    Friend dtParameterPortfolioPricing As New DataTable
    Friend dtParameterPortfolioFee As New DataTable

    Friend dtParameterPortfolioStatus As New DataTable
    Friend dtParameterPortfolioAsset As New DataTable
    Friend dtParameterPortfolioType As New DataTable

    Friend Function GetPortfolioUser() As DataTable
        If dtPortfolioUser IsNot Nothing AndAlso dtPortfolioUser.Rows.Count > 0 Then Return dtPortfolioUser Else Return GetPortfolioUserUpdate()
    End Function

    Friend Function GetPortfolioUserUpdate() As DataTable
        Dim objPortfolio As New MasterPortfolio
        objPortfolio.UserAccess = objAccess
        dtPortfolioUser = objPortfolio.Search("")
        dtPortfolioUser.CaseSensitive = False
        Return dtPortfolioUser
    End Function

    Friend Function GetPortfolioSimpi() As DataTable
        If dtPortfolioSimpi IsNot Nothing AndAlso dtPortfolioSimpi.Rows.Count > 0 Then Return dtPortfolioSimpi Else Return GetPortfolioSimpiUpdate()
    End Function

    Friend Function GetPortfolioSimpiUpdate() As DataTable
        Dim objPortfolio As New MasterPortfolio
        objPortfolio.UserAccess = objAccess
        dtPortfolioSimpi = objPortfolio.Search(objMasterSimpi, "")
        dtPortfolioSimpi.CaseSensitive = False
        Return dtPortfolioSimpi
    End Function

    Friend Function GetBankAccountSimpi() As DataTable
        If dtBankAccountSimpi IsNot Nothing AndAlso dtBankAccountSimpi.Rows.Count > 0 Then Return dtBankAccountSimpi Else Return GetBankAccountSimpiUpdate()
    End Function

    Friend Function GetBankAccountSimpiUpdate() As DataTable
        Dim objAccount As New PortfolioBankAccount
        objAccount.UserAccess = objAccess
        dtBankAccountSimpi = objAccount.SearchAccount(objMasterSimpi, "")
        dtBankAccountSimpi.CaseSensitive = False
        Return dtBankAccountSimpi
    End Function

    Friend Function GetBankAccountUser() As DataTable
        If dtBankAccountUser IsNot Nothing AndAlso dtBankAccountUser.Rows.Count > 0 Then Return dtBankAccountUser Else Return GetBankAccountUserUpdate()
    End Function

    Friend Function GetBankAccountUserUpdate() As DataTable
        Dim objAccount As New PortfolioBankAccount
        objAccount.UserAccess = objAccess
        dtBankAccountUser = objAccount.SearchAccount()
        dtBankAccountUser.CaseSensitive = False
        Return dtBankAccountUser
    End Function

    Friend Function GetPortfolioCodeset() As DataTable
        If dtPortfolioCodeset IsNot Nothing AndAlso dtPortfolioCodeset.Rows.Count > 0 Then Return dtPortfolioCodeset Else Return GetPortfolioCodetsetUpdate()
    End Function

    Friend Function GetPortfolioCodetsetUpdate() As DataTable
        Dim objCodeset As New PortfolioCodeset
        objCodeset.UserAccess = objAccess
        dtPortfolioCodeset = objCodeset.Search(objMasterSimpi, "")
        dtPortfolioCodeset.CaseSensitive = False
        Return dtPortfolioCodeset
    End Function

    Friend Function GetPortfolioExternalID() As DataTable
        If dtPortfolioExternalID IsNot Nothing AndAlso dtPortfolioExternalID.Rows.Count > 0 Then Return dtPortfolioExternalID Else Return GetPortfolioExternalIDUpdate()
    End Function

    Friend Function GetPortfolioExternalIDUpdate() As DataTable
        Dim objPortfolio As New MasterPortfolio
        objPortfolio.UserAccess = objAccess
        dtPortfolioExternalID = objPortfolio.ExternalID_Search(objMasterSimpi)
        dtPortfolioExternalID.CaseSensitive = False
        Return dtPortfolioExternalID
    End Function

    Friend Function GetParameterPortfolioFee() As DataTable
        If dtParameterPortfolioFee IsNot Nothing AndAlso dtParameterPortfolioFee.Rows.Count > 0 Then Return dtParameterPortfolioFee Else Return GetParameterPortfolioFeeUpdate()
    End Function

    Friend Function GetParameterPortfolioFeeUpdate() As DataTable
        Dim objFee As New ParameterCharges
        objFee.UserAccess = objAccess
        dtParameterPortfolioFee = objFee.Search("")
        dtParameterPortfolioFee.CaseSensitive = False
        Return dtParameterPortfolioFee
    End Function

    Friend Function GetParameterPortfolioCodesetField() As DataTable
        If dtParameterPortfolioCodeset IsNot Nothing AndAlso dtParameterPortfolioCodeset.Rows.Count > 0 Then Return dtParameterPortfolioCodeset Else Return GetParameterPortfolioCodesetFieldUpdate()
    End Function

    Friend Function GetParameterPortfolioCodesetFieldUpdate() As DataTable
        Dim objCodeset As New CodesetPortfolioField
        objCodeset.UserAccess = objAccess
        dtParameterPortfolioCodeset = objCodeset.Search("")
        dtParameterPortfolioCodeset.CaseSensitive = False
        Return dtParameterPortfolioCodeset
    End Function

    Friend Function GetParameterPortfolioPricing() As DataTable
        If dtParameterPortfolioPricing IsNot Nothing AndAlso dtParameterPortfolioPricing.Rows.Count > 0 Then Return dtParameterPortfolioPricing Else Return GetParameterPortfolioPricingUpdate()
    End Function

    Friend Function GetParameterPortfolioPricingUpdate() As DataTable
        Dim objPricing As New ParameterPortfolioPricing
        objPricing.UserAccess = objAccess
        dtParameterPortfolioPricing = objPricing.Search("")
        dtParameterPortfolioPricing.CaseSensitive = False
        Return dtParameterPortfolioPricing
    End Function

    Friend Function GetParameterPortfolioStatus() As DataTable
        If dtParameterPortfolioStatus IsNot Nothing AndAlso dtParameterPortfolioStatus.Rows.Count > 0 Then Return dtParameterPortfolioStatus Else Return GetParameterPortfolioStatusUpdate()
    End Function

    Friend Function GetParameterPortfolioStatusUpdate() As DataTable
        Dim objStatus As New ParameterPortfolioStatus
        objStatus.UserAccess = objAccess
        dtParameterPortfolioStatus = objStatus.Search("")
        dtParameterPortfolioStatus.CaseSensitive = False
        Return dtParameterPortfolioStatus
    End Function

    Friend Function GetParameterPortfolioType() As DataTable
        If dtParameterPortfolioType IsNot Nothing AndAlso dtParameterPortfolioType.Rows.Count > 0 Then Return dtParameterPortfolioType Else Return GetParameterPortfolioTypeUpdate()
    End Function

    Friend Function GetParameterPortfolioTypeUpdate() As DataTable
        Dim objType As New ParameterPortfolioType
        objType.UserAccess = objAccess
        dtParameterPortfolioType = objType.Search("")
        dtParameterPortfolioType.CaseSensitive = False
        Return dtParameterPortfolioType
    End Function

    Friend Function GetParameterPortfolioAssetType() As DataTable
        If dtParameterPortfolioAsset IsNot Nothing AndAlso dtParameterPortfolioAsset.Rows.Count > 0 Then Return dtParameterPortfolioAsset Else Return GetParameterPortfolioAssetTypeUpdate()
    End Function

    Friend Function GetParameterPortfolioAssetTypeUpdate() As DataTable
        Dim objAsset As New ParameterPortfolioAssetType
        objAsset.UserAccess = objAccess
        dtParameterPortfolioAsset = objAsset.Search("")
        dtParameterPortfolioAsset.CaseSensitive = False
        Return dtParameterPortfolioAsset
    End Function

End Module
