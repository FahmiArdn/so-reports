Imports simpi.CoreCompliance

Module ModCompliance
    Friend dtCompliancePortfolioAssetAllocation As New DataTable
    Friend dtCompliancePortfolioDeposit As New DataTable
    Friend dtCompliancePortfolioCompany As New DataTable
    Friend dtCompliancePortfolioCurrency As New DataTable

    Friend dtComplianceSimpiAssetAllocation As New DataTable
    Friend dtComplianceSimpiDeposit As New DataTable
    Friend dtComplianceSimpiCompany As New DataTable
    Friend dtComplianceSimpiCurrency As New DataTable

    Friend Function GetCompliancePortfolioAssetAllocation() As DataTable
        If dtCompliancePortfolioAssetAllocation IsNot Nothing AndAlso dtCompliancePortfolioAssetAllocation.Rows.Count > 0 Then Return dtCompliancePortfolioAssetAllocation Else Return GetCompliancePortfolioAssetAllocationUpdate()
    End Function

    Friend Function GetCompliancePortfolioAssetAllocationUpdate() As DataTable
        Dim objCompliance As New PortfolioAssetAllocation
        objCompliance.UserAccess = objAccess
        dtCompliancePortfolioAssetAllocation = objCompliance.Search()
        dtCompliancePortfolioAssetAllocation.CaseSensitive = False
        Return dtCompliancePortfolioAssetAllocation
    End Function

    Friend Function GetCompliancePortfolioDeposit() As DataTable
        If dtCompliancePortfolioDeposit IsNot Nothing AndAlso dtCompliancePortfolioDeposit.Rows.Count > 0 Then Return dtCompliancePortfolioDeposit Else Return GetCompliancePortfolioDepositUpdate()
    End Function

    Friend Function GetCompliancePortfolioDepositUpdate() As DataTable
        Dim objCompliance As New PortfolioDeposit
        objCompliance.UserAccess = objAccess
        dtCompliancePortfolioDeposit = objCompliance.Search()
        dtCompliancePortfolioDeposit.CaseSensitive = False
        Return dtCompliancePortfolioDeposit
    End Function

    Friend Function GetCompliancePortfolioCompany() As DataTable
        If dtCompliancePortfolioCompany IsNot Nothing AndAlso dtCompliancePortfolioCompany.Rows.Count > 0 Then Return dtCompliancePortfolioCompany Else Return GetCompliancePortfolioCompanyUpdate()
    End Function

    Friend Function GetCompliancePortfolioCompanyUpdate() As DataTable
        Dim objCompliance As New PortfolioCompany
        objCompliance.UserAccess = objAccess
        dtCompliancePortfolioCompany = objCompliance.Search()
        dtCompliancePortfolioCompany.CaseSensitive = False
        Return dtCompliancePortfolioCompany
    End Function

    Friend Function GetCompliancePortfolioCurrency() As DataTable
        If dtCompliancePortfolioCurrency IsNot Nothing AndAlso dtCompliancePortfolioCurrency.Rows.Count > 0 Then Return dtCompliancePortfolioCurrency Else Return GetCompliancePortfolioCurrencyUpdate()
    End Function

    Friend Function GetCompliancePortfolioCurrencyUpdate() As DataTable
        Dim objCompliance As New PortfolioCcy
        objCompliance.UserAccess = objAccess
        dtCompliancePortfolioCurrency = objCompliance.Search()
        dtCompliancePortfolioCurrency.CaseSensitive = False
        Return dtCompliancePortfolioCurrency
    End Function

    Friend Function GetComplianceSimpiAssetAllocation() As DataTable
        If dtComplianceSimpiAssetAllocation IsNot Nothing AndAlso dtComplianceSimpiAssetAllocation.Rows.Count > 0 Then Return dtComplianceSimpiAssetAllocation Else Return GetComplianceSimpiAssetAllocationUpdate()
    End Function

    Friend Function GetComplianceSimpiAssetAllocationUpdate() As DataTable
        Dim objCompliance As New SimpiAssetAllocation
        objCompliance.UserAccess = objAccess
        dtComplianceSimpiAssetAllocation = objCompliance.Search(objMasterSimpi)
        dtComplianceSimpiAssetAllocation.CaseSensitive = False
        Return dtComplianceSimpiAssetAllocation
     End Function

    Friend Function GetComplianceSimpiDeposit()
        If dtComplianceSimpiDeposit IsNot Nothing AndAlso dtComplianceSimpiDeposit.Rows.Count > 0 Then Return dtComplianceSimpiDeposit Else Return GetComplianceSimpiDepositUpdate()
    End Function

    Friend Function GetComplianceSimpiDepositUpdate()
        Dim objCompliance As New SimpiDeposit
        objCompliance.UserAccess = objAccess
        dtComplianceSimpiDeposit = objCompliance.Search(objMasterSimpi)
        dtComplianceSimpiDeposit.CaseSensitive = False
        Return dtComplianceSimpiDeposit
    End Function

    Friend Function GetComplianceSimpiCompany()
        If dtComplianceSimpiCompany IsNot Nothing AndAlso dtComplianceSimpiCompany.Rows.Count > 0 Then Return dtComplianceSimpiCompany Else Return GetComplianceSimpiCompanyUpdate()
    End Function

    Friend Function GetComplianceSimpiCompanyUpdate()
        Dim objCompliance As New SimpiCompany
        objCompliance.UserAccess = objAccess
        dtComplianceSimpiCompany = objCompliance.Search(objMasterSimpi)
        dtComplianceSimpiCompany.CaseSensitive = False
        Return dtComplianceSimpiCompany
    End Function

    Friend Function GetComplianceSimpiCurrency()
        If dtComplianceSimpiCurrency IsNot Nothing AndAlso dtComplianceSimpiCurrency.Rows.Count > 0 Then Return dtComplianceSimpiCurrency Else Return GetComplianceSimpiCurrencyUpdate()
    End Function

    Friend Function GetComplianceSimpiCurrencyUpdate()
        Dim objCompliance As New SimpiCcy
        objCompliance.UserAccess = objAccess
        dtComplianceSimpiCurrency = objCompliance.Search(objMasterSimpi)
        dtComplianceSimpiCurrency.CaseSensitive = False
        Return dtComplianceSimpiCurrency
    End Function

End Module
