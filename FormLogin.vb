Imports simpi.GlobalUtilities
Imports System.IO
Imports simpi.GlobalConnection

Public Class FormLogin
    Private Sub FormLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtTerminalID.Text = GetMacAddress()
        iniLoad()
    End Sub

    Private Sub cmdSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubmit.Click
        LogOn()
    End Sub

    Private Sub txtPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Return Then LogOn()
    End Sub

    Private Sub LogOn()
        Try
            'ReadToken()
            Submit()
        Catch ex As Exception
            ExceptionMessage.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
        End Try
    End Sub

    'Private Sub ReadToken()
    '    Dim strFile As String = simpiFile("simpi.token")
    '    AppsTerminal = GetMacAddress()
    '    If AppsTerminal.Trim.Length < 1 Then Throw New ExceptionDataNotFound("internet connection/terminal MAC Address")
    '    strFile &= "." & AppsTerminal
    '    If Not GlobalFileWindows.FileExists(strFile) Then Throw New ExceptionDataNotFound("simpi token file")

    '    Dim objReader As New StreamReader(strFile)
    '    Dim aes As New GlobalEncryption
    '    LicenseKey = objReader.ReadLine
    '    Dim key, iv As String
    '    key = _key(AppsTerminal, LicenseKey)
    '    iv = _iv(AppsTerminal, LicenseKey)
    '    Dim TermimalMacAddress As String = aes.TDES_Decrypt(key, iv, objReader.ReadLine)
    '    If AppsTerminal <> TermimalMacAddress Then
    '        ExceptionMessage.Show("system detect simpi token is not valid for this terminal. Please contact your simpi administrator to resolve it",
    '                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Application.Exit()
    '    End If
    '    DatabaseServer = getString(aes.TDES_Decrypt(key, iv, objReader.ReadLine), 1, 99, 86, 65, 82, 66, 91, 67, 81)
    '    DatabasePort = getString(aes.TDES_Decrypt(key, iv, objReader.ReadLine), 88, 54, 55, 51)
    '    DatabaseName = getString(aes.TDES_Decrypt(key, iv, objReader.ReadLine), 84, 29, 13, 88, 38)
    '    DatabaseUser = getString(aes.TDES_Decrypt(key, iv, objReader.ReadLine), 21, 69, 28, 36)
    '    DatabasePassword = getString(aes.TDES_Decrypt(key, iv, objReader.ReadLine), 71, 85)
    '    DatabaseType = getString(aes.TDES_Decrypt(key, iv, objReader.ReadLine), 96, 4, 50, 23, 130)
    '    objReader.Close()
    '    objReader.Dispose()

    '    If DatabaseServer.Trim = "" Or DatabasePort = 0 Or DatabaseName.Trim = "" Or DatabaseUser.Trim = "" Then
    '        ExceptionMessage.Show("system detect invalid simpi token, can not read any database connection." &
    '                              " Please contact your simpi administrator to resolve it",
    '                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Application.Exit()
    '    End If
    'End Sub

    Private Sub Submit()
        With MDIMENU
            .Show()
            buttonSet(False)
            If rbMySQL.Checked Then .DatabaseType = "MySQL" Else .DatabaseType = "MSSQL"
            .UserConnection = "Server=" & txtDatabaseServer.Text.Trim & ";Port=" & txtDatabasePort.Text.Trim &
                              ";Database=simpi;uid=" & txtDatabaseUser.Text.Trim
            If txtDatabasePassword.Text.Trim = "" Then
                .UserConnection = .UserConnection & ";Pwd=''"
            Else
                .UserConnection = .UserConnection & ";Pwd=" & txtDatabasePassword.Text.Trim
            End If
            .UserLogin = txtUserLogin.Text
            .UserPassword = txtPassword.Text
            .server_connect()
            Close()
        End With
    End Sub

    Public Sub buttonSet(state As Boolean)
        cmdSubmit.Enabled = state
        cmdClose.Enabled = state
        btnFolder.Enabled = state
        btnNetwork.Enabled = state
        If Not state Then ighDatabase.Collapsed = True
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Application.Exit()
    End Sub

    Private Sub cmdClose_KeyDown(sender As Object, e As KeyEventArgs) Handles cmdClose.KeyDown
        Application.Exit()
    End Sub

    Private Sub btnFolder_Click(sender As Object, e As EventArgs) Handles btnFolder.Click
        OpenFolder()
    End Sub

    Private Sub OpenFolder()
        Process.Start("explorer.exe", String.Format("/n, /e, {0}", My.Application.Info.DirectoryPath))
    End Sub

    Private Sub btnNetwork_Click(sender As Object, e As EventArgs) Handles btnNetwork.Click
        ExceptionMessage.Show(GetMacAddressList, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ighProxy_CollapsedChanged(sender As Object, e As EventArgs) Handles ighDatabase.CollapsedChanged
        If ighDatabase.Collapsed Then Height = 275 Else Height = 375
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        iniSave()
    End Sub

    Private Sub iniSave()
        Try
            If txtDatabaseServer.Text.Trim = "" Then
                ExceptionMessage.Show("System can not detect database server name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If Not IsNumeric(txtDatabasePort.Text) Then
                ExceptionMessage.Show("System can not detect server port", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If txtDatabaseUser.Text.Trim = "" Then
                ExceptionMessage.Show("System can not detect user", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim strFile As String = simpiFile("simpi.ini")
            Dim file As New GlobalINI(strFile)
            Dim section As String = "Database Server"
            file.WriteString(section, "Server Name", txtDatabaseServer.Text.Trim)
            file.WriteInteger(section, "Server Port", txtDatabasePort.Text.Trim)
            file.WriteString(section, "User", txtDatabaseUser.Text.Trim)
            If txtDatabasePassword.Text.Trim <> "" Then
                Dim aes As New GlobalEncryption
                file.WriteString(section, "Password", aes.TDES_Encrypt(txtDatabasePassword.Text.Trim))
            End If
            If rbMySQL.Checked Then
                file.WriteString(section, "Database Type", "MySQL")
            ElseIf rbSQLServer.Checked Then
                file.WriteString(section, "Database Type", "MSSQL")
            End If
            ExceptionMessage.Show("Database connection setting had been saved")
        Catch ex As Exception
            ExceptionMessage.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub iniLoad()
        Try
            Dim strFile As String = simpiFile("simpi.ini")
            If GlobalFileWindows.FileExists(strFile) Then
                Dim file As New GlobalINI(strFile)
                Dim section As String = "Database Server"
                txtDatabaseServer.Text = file.GetString(section, "Server Name", "")
                txtDatabasePort.Text = file.GetInteger(section, "Server Port", 0)
                txtDatabaseUser.Text = file.GetString(section, "User", "")
                Dim str As String = file.GetString(section, "Password", "")
                If str.Trim <> "" Then
                    Dim aes As New GlobalEncryption
                    txtDatabasePassword.Text = aes.TDES_Decrypt(str)
                Else
                    txtDatabasePassword.Text = ""
                End If
                If simpi.GlobalUtilities.GlobalString.GetDatabaseType(file.GetString(section, "Database Type", "")) = 0 Then
                    rbSQLServer.Checked = True
                Else
                    rbMySQL.Checked = True
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

End Class