Imports simpi.GlobalUtilities
Imports simpi.SimpiLogin

Public Class MDIMENU
    Public UserPassword As String = ""
    Public UserLogin As String = ""
    Public UserConnection As String = ""
    Public DatabaseType As String = ""
    Dim objLogin As New UserLogin


#Region "windows"
    Private Sub fileExit_Click(ByVal sender As Object, ByVal e As EventArgs)
        ApplicationExit()
    End Sub

    Private Sub fileWindowsCascade_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub fileWindowsTileVertical_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub fileWindowsTileHorizontal_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub fileWindowsArrangeAll_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub fileWindowsCloseAll_Click(ByVal sender As Object, ByVal e As EventArgs)
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private Sub tsbExit_Click(sender As Object, e As EventArgs) Handles tsbExit.Click
        ApplicationExit()
    End Sub

#End Region
    Public Sub server_connect()
        lblStatus.Text = "LOCAL CONNECTION"
        lblStatus.ForeColor = Color.Green
        logon()
    End Sub

#Region "application access"
    Private Sub logon()
        If objLogin.Logon(UserConnection, DatabaseType, UserLogin, UserPassword, AppsCode, AppsVersion) Then
            If objLogin.ErrID = 0 Then
                objAccess.DatabaseConnection = UserConnection
                objAccess.DatabaseType = DatabaseType
                objAccess.GetSimpiLogin = objLogin
                lblUser.Text = "Master Simpi: " & objAccess.GetSimpiName.ToString() + ", UserName: " + objAccess.GetUserName.ToString() +
                               " | " + objAccess.GetUserTitle.ToString()
                lblAsOf.Text = "As of: " + objAccess.GetServerDate().ToString("dd-MMM-yyyy")
                lblVersion.Text = "version: " & AppsVersion
                lblTerminal.Text = "terminal: " & AppsTerminal
                objError.Clear()
                objError.UserAccess = objAccess
                GetMasterSimpi()
                cipLogin.Visible = False
                msMenu.Enabled = True
                tsbUserProfile.Enabled = True
            Else
                ExceptionMessage.Show(objLogin.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cipLogin.Visible = True
                txtUserLogin.Focus()
            End If
        Else
            ExceptionMessage.Show(objLogin.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cipLogin.Visible = True
            txtUserLogin.Focus()
        End If
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Close()
    End Sub

    Private Sub cmdSubmit_Click(sender As Object, e As EventArgs) Handles cmdSubmit.Click
        logon2()
    End Sub

    Private Sub txtPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Return Then
            logon2()
        End If
    End Sub

    Private Sub logon2()
        UserLogin = txtUserLogin.Text
        UserPassword = txtPassword.Text
        If objLogin.Logon(UserConnection, DatabaseType, UserLogin, UserPassword, AppsCode, AppsVersion) Then
            If objLogin.ErrID = 0 Then
                objAccess.DatabaseConnection = UserConnection
                objAccess.DatabaseType = DatabaseType
                objAccess.GetSimpiLogin = objLogin
                lblUser.Text = "Master Simpi: " & objAccess.GetSimpiName.ToString() + ", UserName: " + objAccess.GetUserName.ToString() +
                               " | " + objAccess.GetUserTitle.ToString()
                lblAsOf.Text = "As of: " + objAccess.GetServerDate().ToString("dd-MMM-yyyy")
                lblVersion.Text = "version: " & AppsVersion
                lblTerminal.Text = "terminal: " & AppsTerminal
                objError.Clear()
                objError.UserAccess = objAccess
                GetMasterSimpi()
                cipLogin.Visible = False
                msMenu.Enabled = True
                tsbUserProfile.Enabled = True
            Else
                ExceptionMessage.Show(objLogin.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtUserLogin.Focus()
            End If
        Else
            ExceptionMessage.Show(objLogin.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtUserLogin.Focus()
        End If
    End Sub

    Private Sub ApplicationExit()
        objAccess.Logout()
        Application.Exit()
    End Sub

#End Region

    Private Sub MenuCheck(ByVal strMenu As String)
        'If Not objAccess.AccessMenu(strMenu) Then
        '    Throw New Exception("System detect user: " & objAccess.GetUserName & " did not had access to menu: " & strMenu)
        'End If
    End Sub

    Private Sub ReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportToolStripMenuItem.Click
        Try
            'MenuCheck("Account Maintenance")
            Dim frm As New ReportFundSheetEQ
            frm.Left = 0
            frm.Top = 30
            frm.MdiParent = Me
            frm.Show()
        Catch ex As Exception
            ExceptionMessage.Show(ex.Message)
        End Try
    End Sub
End Class
