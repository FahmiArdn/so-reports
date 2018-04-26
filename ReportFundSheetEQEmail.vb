Imports simpi.globalconnection
Imports simpi.globalutilities

Public Class ReportFundSheetEmail
    Public frm As ReportFundSheetEQ
    Dim emailGateway As New GlobalEmailSMTP
    Dim IsHolding As Boolean

    Private Sub ReportAccountStatementEmail_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblFromName.Text = objAccess.GetUserName
        lblFromEmail.Text = objAccess.GetUserLogin
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Close()
    End Sub

    Private Sub btnAttach_Click(sender As Object, e As EventArgs) Handles btnAttach.Click
        GenerateAttachment()
    End Sub

    Private Sub GenerateAttachment()
        Cursor = Cursors.WaitCursor
        lblAttachment.Text = frm.ExportPDF(True)
        Cursor = Cursors.Default
    End Sub

    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        EmailSend()
    End Sub

    Private Sub EmailSend()
        Try
            If txtTo.Text.Trim = "" Then
                ExceptionMessage.Show(objError.Message(81) & " To email address", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If lblAttachment.Text.Trim = "" Then
                ExceptionMessage.Show(objError.Message(81) & " report attachment", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            Dim cc As String = ""
            If txtCC.Text.Trim = "" Then cc = lblFromEmail.Text Else cc = txtCC.Text & "," & lblFromEmail.Text
            Dim msg As String
            msg = FromRtf(txtMessage)
            Cursor = Cursors.WaitCursor
            emailGateway.Send2(lblFromEmail.Text, lblFromName.Text, txtTo.Text, txtSubject.Text,
                               msg, cc, txtBCC.Text, lblAttachment.Text)
            Cursor = Cursors.Default
            Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

End Class