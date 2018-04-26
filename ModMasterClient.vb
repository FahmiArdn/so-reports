Imports simpi.ParameterClient
Imports simpi.MasterClient.ParameterClient
Imports simpi.MasterClient

Module ModMasterClient
    Friend dtClientMaster As New DataTable
    Friend dtClientIndividu As New DataTable
    Friend dtClientInstitusi As New DataTable
    Friend dtClientDistributor As New DataTable
    Friend dtClientIndividuSub As New DataTable
    Friend dtClientInstitusiSub As New DataTable
    Friend dtClientKYC As New DataTable
    Friend dtClientKYCIndividu As New DataTable
    Friend dtClientKYCInstitusi As New DataTable

    Friend dtParameterClientStatus As New DataTable
    Friend dtParameterClientType As New DataTable
    Friend dtParameterClientRisk As New DataTable
    Friend dtParameterClientReligion As New DataTable
    Friend dtParameterClientEducationLevel As New DataTable
    Friend dtParameterClientOccupation As New DataTable
    Friend dtParameterClientMarital As New DataTable
    Friend dtParameterClientDocumentType As New DataTable
    Friend dtParameterClientBusinessType As New DataTable
    Friend dtParameterClientBusinessActivity As New DataTable

    Friend Function GetClientMaster() As DataTable
        If dtClientMaster IsNot Nothing AndAlso dtClientMaster.Rows.Count > 0 Then Return dtClientMaster Else Return GetClientMasterUpdate()
    End Function

    Friend Function GetClientMasterUpdate() As DataTable
        Dim objClient As New MasterClient
        objClient.UserAccess = objAccess
        dtClientMaster = objClient.Search(objMasterSimpi, "")
        dtClientMaster.CaseSensitive = False
        Return dtClientMaster
    End Function

    Friend Function GetClientIndividu() As DataTable
        If dtClientIndividu IsNot Nothing AndAlso dtClientIndividu.Rows.Count > 0 Then Return dtClientIndividu Else Return GetClientIndividuUpdate()
    End Function

    Friend Function GetClientIndividuUpdate() As DataTable
        Dim objClient As New MasterClientIndividu
        objClient.UserAccess = objAccess
        dtClientMaster = objClient.Search(objMasterSimpi, "")
        dtClientIndividu.CaseSensitive = False
        Return dtClientIndividu
    End Function

    Friend Function GetClientInstitusi() As DataTable
        If dtClientInstitusi IsNot Nothing AndAlso dtClientInstitusi.Rows.Count > 0 Then Return dtClientInstitusi Else Return GetClientInstitusiUpdate()
    End Function

    Friend Function GetClientInstitusiUpdate() As DataTable
        Dim objClient As New MasterClientInstitution
        objClient.UserAccess = objAccess
        dtClientMaster = objClient.Search(objMasterSimpi, "")
        dtClientInstitusi.CaseSensitive = False
        Return dtClientInstitusi
    End Function

    Friend Function GetClientIndividuSub() As DataTable
        If dtClientIndividuSub IsNot Nothing AndAlso dtClientIndividuSub.Rows.Count > 0 Then Return dtClientIndividuSub Else Return GetClientIndividuSubUpdate()
    End Function

    Friend Function GetClientIndividuSubUpdate() As DataTable
        Dim objClient As New MasterClientIndividuSub
        objClient.UserAccess = objAccess
        dtClientMaster = objClient.Search(objMasterSimpi, "")
        dtClientIndividu.CaseSensitive = False
        Return dtClientIndividu
    End Function

    Friend Function GetClientInstitusiSub() As DataTable
        If dtClientInstitusiSub IsNot Nothing AndAlso dtClientInstitusiSub.Rows.Count > 0 Then Return dtClientInstitusiSub Else Return GetClientInstitusiSubUpdate()
    End Function

    Friend Function GetClientInstitusiSubUpdate() As DataTable
        Dim objClient As New MasterClientInstitutionSub
        objClient.UserAccess = objAccess
        dtClientMaster = objClient.Search(objMasterSimpi, "")
        dtClientInstitusi.CaseSensitive = False
        Return dtClientInstitusi
    End Function

    Friend Function GetClientDistributor() As DataTable
        If dtClientDistributor IsNot Nothing AndAlso dtClientDistributor.Rows.Count > 0 Then Return dtClientDistributor Else Return GetClientDistributorUpdate()
    End Function

    Friend Function GetClientDistributorUpdate() As DataTable
        Dim objClient As New MasterClientDistributor
        objClient.UserAccess = objAccess
        dtClientMaster = objClient.Search(objMasterSimpi, "")
        dtClientDistributor.CaseSensitive = False
        Return dtClientDistributor
    End Function

    Friend Function GetClientKYC() As DataTable
        If dtClientKYC IsNot Nothing AndAlso dtClientKYC.Rows.Count > 0 Then Return dtClientKYC Else Return GetClientKYCUpdate()
    End Function

    Friend Function GetClientKYCUpdate() As DataTable
        Dim objClient As New ClientKYC
        objClient.UserAccess = objAccess
        dtClientKYC = objClient.Search(objMasterSimpi, "")
        dtClientKYC.CaseSensitive = False
        Return dtClientKYC
    End Function

    Friend Function GetClientKYCIndividu() As DataTable
        If dtClientKYCIndividu IsNot Nothing AndAlso dtClientKYCIndividu.Rows.Count > 0 Then Return dtClientKYCIndividu Else Return GetClientKYCUpdateIndividu()
    End Function

    Friend Function GetClientKYCUpdateIndividu() As DataTable
        Dim objClient As New ClientKYC
        objClient.UserAccess = objAccess
        dtClientKYCIndividu = objClient.Search(objMasterSimpi, "", SetIndividu)
        dtClientKYCIndividu.CaseSensitive = False
        Return dtClientKYCIndividu
    End Function

    Friend Function GetClientKYCInstitusi() As DataTable
        If dtClientKYCInstitusi IsNot Nothing AndAlso dtClientKYCInstitusi.Rows.Count > 0 Then Return dtClientKYCInstitusi Else Return GetClientKYCUpdateInstitusi()
    End Function

    Friend Function GetClientKYCUpdateInstitusi() As DataTable
        Dim objClient As New ClientKYC
        objClient.UserAccess = objAccess
        dtClientKYCInstitusi = objClient.Search(objMasterSimpi, "", SetInstitution)
        dtClientKYCInstitusi.CaseSensitive = False
        Return dtClientKYCInstitusi
    End Function

    Friend Sub GetComboGender(ByVal cmb As C1.Win.C1InputPanel.InputComboBox)
        cmb.Items.Clear()
        cmb.Items.AddText("M")
        cmb.Items.AddText("F")
    End Sub

    Friend Function GetParameterClientStatus() As DataTable
        If dtParameterClientStatus IsNot Nothing AndAlso dtParameterClientStatus.Rows.Count > 0 Then Return dtParameterClientStatus Else Return GetParameterClientStatusUpdate()
    End Function

    Friend Function GetParameterClientStatusUpdate() As DataTable
        Dim objStatus As New ParameterClientStatus
        objStatus.UserAccess = objAccess
        dtParameterClientStatus = objStatus.Search("")
        dtParameterClientStatus.CaseSensitive = False
        Return dtParameterClientStatus
    End Function

    Friend Function GetParameterClientType() As DataTable
        If dtParameterClientType IsNot Nothing AndAlso dtParameterClientType.Rows.Count > 0 Then Return dtParameterClientType Else Return GetParameterClientTypeUpdate()
    End Function

    Friend Function GetParameterClientTypeUpdate() As DataTable
        Dim objAsset As New ParameterClientType
        objAsset.UserAccess = objAccess
        dtParameterClientType = objAsset.Search("")
        dtParameterClientType.CaseSensitive = False
        Return dtParameterClientType
    End Function

    Friend Function GetParameterClientRisk() As DataTable
        If dtParameterClientRisk IsNot Nothing AndAlso dtParameterClientRisk.Rows.Count > 0 Then Return dtParameterClientRisk Else Return GetParameterClientRiskUpdate()
    End Function

    Friend Function GetParameterClientRiskUpdate() As DataTable
        Dim objRisk As New ParameterClientRisk
        objRisk.UserAccess = objAccess
        dtParameterClientRisk = objRisk.Search("")
        dtParameterClientRisk.CaseSensitive = False
        Return dtParameterClientRisk
    End Function

    Friend Function GetParameterClientReligion() As DataTable
        If dtParameterClientReligion IsNot Nothing AndAlso dtParameterClientReligion.Rows.Count > 0 Then Return dtParameterClientReligion Else Return GetParameterClientReligionUpdate()
    End Function

    Friend Function GetParameterClientReligionUpdate() As DataTable
        Dim objReligion As New ParameterReligion
        objReligion.UserAccess = objAccess
        dtParameterClientReligion = objReligion.Search("")
        dtParameterClientReligion.CaseSensitive = False
        Return dtParameterClientReligion
    End Function

    Friend Function GetParameterClientEducationLevel() As DataTable
        If dtParameterClientEducationLevel IsNot Nothing AndAlso dtParameterClientEducationLevel.Rows.Count > 0 Then Return dtParameterClientEducationLevel Else Return GetParameterClientEducationLevelUpdate()
    End Function

    Friend Function GetParameterClientEducationLevelUpdate() As DataTable
        Dim objType As New ParameterEducationLevel
        objType.UserAccess = objAccess
        dtParameterClientEducationLevel = objType.Search("")
        dtParameterClientEducationLevel.CaseSensitive = False
        Return dtParameterClientEducationLevel
    End Function

    Friend Function GetParameterClientOccupation() As DataTable
        If dtParameterClientOccupation IsNot Nothing AndAlso dtParameterClientOccupation.Rows.Count > 0 Then Return dtParameterClientOccupation Else Return GetParameterClientOccupationUpdate()
    End Function

    Friend Function GetParameterClientOccupationUpdate() As DataTable
        Dim objType As New ParameterOccupation
        objType.UserAccess = objAccess
        dtParameterClientOccupation = objType.Search("")
        dtParameterClientOccupation.CaseSensitive = False
        Return dtParameterClientOccupation
    End Function

    Friend Function GetParameterClientMarital() As DataTable
        If dtParameterClientMarital IsNot Nothing AndAlso dtParameterClientMarital.Rows.Count > 0 Then Return dtParameterClientMarital Else Return GetParameterClientMaritalUpdate()
    End Function

    Friend Function GetParameterClientMaritalUpdate() As DataTable
        Dim objType As New ParameterMaritalStatus
        objType.UserAccess = objAccess
        dtParameterClientMarital = objType.Search("")
        dtParameterClientMarital.CaseSensitive = False
        Return dtParameterClientMarital
    End Function

    Friend Function GetParameterClientDocumentType() As DataTable
        If dtParameterClientDocumentType IsNot Nothing AndAlso dtParameterClientDocumentType.Rows.Count > 0 Then Return dtParameterClientDocumentType Else Return GetParameterClientDocumentTypeUpdate()
    End Function

    Friend Function GetParameterClientDocumentTypeUpdate() As DataTable
        Dim objType As New ParameterDocumentType
        objType.UserAccess = objAccess
        dtParameterClientDocumentType = objType.Search("")
        dtParameterClientDocumentType.CaseSensitive = False
        Return dtParameterClientDocumentType
    End Function

    Friend Function GetParameterClientBusinessType() As DataTable
        If dtParameterClientBusinessType IsNot Nothing AndAlso dtParameterClientBusinessType.Rows.Count > 0 Then Return dtParameterClientBusinessType Else Return GetParameterClientBusinessTypeUpdate()
    End Function

    Friend Function GetParameterClientBusinessTypeUpdate() As DataTable
        Dim objType As New ParameterBusinessType
        objType.UserAccess = objAccess
        dtParameterClientBusinessType = objType.Search("")
        dtParameterClientBusinessType.CaseSensitive = False
        Return dtParameterClientBusinessType
    End Function

    Friend Function GetParameterClientBusinessActivity() As DataTable
        If dtParameterClientBusinessActivity IsNot Nothing AndAlso dtParameterClientBusinessActivity.Rows.Count > 0 Then Return dtParameterClientBusinessActivity Else Return GetParameterClientBusinessActivityUpdate()
    End Function

    Friend Function GetParameterClientBusinessActivityUpdate() As DataTable
        Dim objType As New ParameterBusinessActivity
        objType.UserAccess = objAccess
        dtParameterClientBusinessActivity = objType.Search("")
        dtParameterClientBusinessActivity.CaseSensitive = False
        Return dtParameterClientBusinessActivity
    End Function

End Module
