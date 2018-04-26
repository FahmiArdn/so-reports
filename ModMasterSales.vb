Imports simpi.MasterSales

Module ModMasterSales
    Friend dtSalesMaster As New DataTable
    Friend dtSalesCodeset As New DataTable

    Friend Function GetSalesMaster() As DataTable
        If dtSalesMaster IsNot Nothing AndAlso dtSalesMaster.Rows.Count > 0 Then Return dtSalesMaster Else Return GetSalesMasterUpdate()
    End Function

    Friend Function GetSalesMasterUpdate() As DataTable
        Dim objSales As New MasterSales
        objSales.UserAccess = objAccess
        dtSalesMaster = objSales.SearchAll(objMasterSimpi, "")
        dtSalesMaster.CaseSensitive = False
        Return dtSalesMaster
    End Function

    Friend Function GetSalesCodeset() As DataTable
        If dtSalesCodeset IsNot Nothing AndAlso dtSalesCodeset.Rows.Count > 0 Then Return dtSalesCodeset Else Return GetSalesCodesetUpdate()
    End Function

    Friend Function GetSalesCodesetUpdate() As DataTable
        Dim objCodeset As New SalesCodeset
        objCodeset.UserAccess = objAccess
        dtSalesCodeset = objCodeset.Search(objMasterSimpi, "")
        dtSalesCodeset.CaseSensitive = False
        Return dtSalesCodeset
    End Function

End Module
