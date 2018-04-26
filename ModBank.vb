Imports simpi.ParameterBank

Module ModBank
    Friend dtParameterBankTDTerm As New DataTable

    Friend Function GetParameterBankTDTerm() As DataTable
        If dtParameterBankTDTerm IsNot Nothing AndAlso dtParameterBankTDTerm.Rows.Count > 0 Then Return dtParameterBankTDTerm Else Return GetParameterBankTDTermUpdate()
    End Function

    Friend Function GetParameterBankTDTermUpdate() As DataTable
        Dim objTerm As New ParameterBankTDTerm
        objTerm.UserAccess = objAccess
        dtParameterBankTDTerm = objTerm.Search("")
        dtParameterBankTDTerm.CaseSensitive = False
        Return dtParameterBankTDTerm
    End Function

End Module
