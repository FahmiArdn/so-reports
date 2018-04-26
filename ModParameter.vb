Imports simpi.MasterSecurities
Imports simpi.ParameterFA
Imports simpi.MarketCompany
Imports simpi.ParameterSecurities

Module ModParameter
    Friend dtParameterCountry As New DataTable
    Friend dtParameterCity As New DataTable
    Friend dtParameterBenchmarkClass As New DataTable
    Friend dtParameterSectorClass As New DataTable
    Friend dtParameterInstrumentType As New DataTable
    Friend dtParameterInstrumentSubType As New DataTable
    Friend dtParameterTATrxType1 As New DataTable
    Friend dtParameterBroker As New DataTable
    Friend dtParameterFACharges As New DataTable
    Friend dtParameterExternalSystem As New DataTable

    Friend Function GetParameterCountry() As DataTable
        If dtParameterCountry IsNot Nothing AndAlso dtParameterCountry.Rows.Count > 0 Then Return dtParameterCountry Else Return GetParameterCountryUpdate()
    End Function

    Friend Function GetParameterCountryUpdate() As DataTable
        Dim objCcy As New ParameterCountry
        objCcy.UserAccess = objAccess
        dtParameterCountry = objCcy.Search("")
        dtParameterCountry.CaseSensitive = False
        Return dtParameterCountry
    End Function

    Friend Function GetParameterCity() As DataTable
        If dtParameterCity IsNot Nothing AndAlso dtParameterCity.Rows.Count > 0 Then Return dtParameterCity Else Return GetParameterCityUpdate()
    End Function

    Friend Function GetParameterCityUpdate() As DataTable
        Dim objCity As New ParameterCity
        objCity.UserAccess = objAccess
        dtParameterCity = objCity.Search(0, 0, "")
        dtParameterCity.CaseSensitive = False
        Return dtParameterCity
    End Function

    Friend Function GetParameterBenchmarkClass() As DataTable
        If dtParameterBenchmarkClass IsNot Nothing AndAlso dtParameterBenchmarkClass.Rows.Count > 0 Then Return dtParameterBenchmarkClass Else Return GetParameterBenchmarkClassUpdate()
    End Function

    Friend Function GetParameterBenchmarkClassUpdate() As DataTable
        Dim objClass As New ParameterBenchmarkClass
        objClass.UserAccess = objAccess
        dtParameterBenchmarkClass = objClass.Search("")
        dtParameterBenchmarkClass.CaseSensitive = False
        Return dtParameterBenchmarkClass
    End Function

    Friend Function GetParameterSectorClass() As DataTable
        If dtParameterSectorClass IsNot Nothing AndAlso dtParameterSectorClass.Rows.Count > 0 Then Return dtParameterSectorClass Else Return GetParameterSectorClassUpdate()
    End Function

    Friend Function GetParameterSectorClassUpdate() As DataTable
        Dim objClass As New ParameterSectorClass
        objClass.UserAccess = objAccess
        dtParameterSectorClass = objClass.Search("")
        dtParameterSectorClass.CaseSensitive = False
        Return dtParameterSectorClass
    End Function

    Friend Function GetParameterInstrumentType() As DataTable
        If dtParameterInstrumentType IsNot Nothing AndAlso dtParameterInstrumentType.Rows.Count > 0 Then Return dtParameterInstrumentType Else Return GetParameterInstrumentTypeUpdate()
    End Function

    Friend Function GetParameterInstrumentTypeUpdate() As DataTable
        Dim objType As New ParameterInstrumentType
        objType.UserAccess = objAccess
        dtParameterInstrumentType = objType.Search("")
        dtParameterInstrumentType.CaseSensitive = False
        Return dtParameterInstrumentType
    End Function

    Friend Function GetParameterInstrumentSubType() As DataTable
        If dtParameterInstrumentSubType IsNot Nothing AndAlso dtParameterInstrumentSubType.Rows.Count > 0 Then Return dtParameterInstrumentSubType Else Return GetParameterInstrumentSubTypeUpdate()
    End Function

    Friend Function GetParameterInstrumentSubTypeUpdate() As DataTable
        Dim objSub As New ParameterInstrumentSubType
        objSub.UserAccess = objAccess
        dtParameterInstrumentSubType = objSub.Search("")
        dtParameterInstrumentSubType.CaseSensitive = False
        Return dtParameterInstrumentSubType
    End Function

    Friend Function GetParameterExternalSystem() As DataTable
        If dtParameterExternalSystem IsNot Nothing AndAlso dtParameterExternalSystem.Rows.Count > 0 Then Return dtParameterExternalSystem Else Return GetParameterExternalSystemUpdate()
    End Function

    Friend Function GetParameterExternalSystemUpdate() As DataTable
        Dim objSystem As New ParameterExternalSystem
        objSystem.UserAccess = objAccess
        dtParameterExternalSystem = objSystem.Search("")
        dtParameterExternalSystem.CaseSensitive = False
        Return dtParameterExternalSystem
    End Function

    Friend Function GetParameterFACharges() As DataTable
        If dtParameterFACharges IsNot Nothing AndAlso dtParameterFACharges.Rows.Count > 0 Then Return dtParameterFACharges Else Return GetParameterFAChargesUpdate()
    End Function

    Friend Function GetParameterFAChargesUpdate() As DataTable
        Dim objFee As New ParameterCharges
        objFee.UserAccess = objAccess
        dtParameterFACharges = objFee.Search("")
        dtParameterFACharges.CaseSensitive = False
        Return dtParameterFACharges
    End Function

    Friend Function GetParameterBroker(ByVal CountryID As Integer, ByVal strKeyword As String) As DataTable
        If dtParameterBroker IsNot Nothing AndAlso dtParameterBroker.Rows.Count > 0 Then Return dtParameterBroker Else Return GetParameterBrokerUpdate()
    End Function

    Friend Function GetParameterBrokerUpdate() As DataTable
        Dim objBroker As New MasterCounterPart
        objBroker.UserAccess = objAccess
        dtParameterBroker = objBroker.Search(0, "")
        dtParameterBroker.CaseSensitive = False
        Return dtParameterBroker
    End Function

    Friend Function GetParameterTATrxType1() As DataTable
        If dtParameterTATrxType1 IsNot Nothing AndAlso dtParameterTATrxType1.Rows.Count > 0 Then Return dtParameterTATrxType1 Else Return GetParameterTATrxType1Update()
    End Function

    Friend Function GetParameterTATrxType1Update() As DataTable
        Dim objType As New simpi.ParameterTA.ParameterTransactionTrxType1
        objType.UserAccess = objAccess
        dtParameterTATrxType1 = objType.Search("")
        dtParameterTATrxType1.CaseSensitive = False
        Return dtParameterTATrxType1
    End Function

End Module
