<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormLogin
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormLogin))
        Me.cipLogin = New C1.Win.C1InputPanel.C1InputPanel()
        Me.InputImage1 = New C1.Win.C1InputPanel.InputImage()
        Me.InputLabel1 = New C1.Win.C1InputPanel.InputLabel()
        Me.txtUserLogin = New C1.Win.C1InputPanel.InputTextBox()
        Me.InputLabel2 = New C1.Win.C1InputPanel.InputLabel()
        Me.txtPassword = New C1.Win.C1InputPanel.InputTextBox()
        Me.cmdSubmit = New C1.Win.C1InputPanel.InputButton()
        Me.cmdClose = New C1.Win.C1InputPanel.InputButton()
        Me.InputLabel3 = New C1.Win.C1InputPanel.InputLabel()
        Me.InputSeparator2 = New C1.Win.C1InputPanel.InputSeparator()
        Me.btnNetwork = New C1.Win.C1InputPanel.InputButton()
        Me.btnFolder = New C1.Win.C1InputPanel.InputButton()
        Me.InputSeparator4 = New C1.Win.C1InputPanel.InputSeparator()
        Me.InputLabel4 = New C1.Win.C1InputPanel.InputLabel()
        Me.txtTerminalID = New C1.Win.C1InputPanel.InputTextBox()
        Me.ighDatabase = New C1.Win.C1InputPanel.InputGroupHeader()
        Me.InputLabel5 = New C1.Win.C1InputPanel.InputLabel()
        Me.txtDatabaseServer = New C1.Win.C1InputPanel.InputTextBox()
        Me.InputLabel7 = New C1.Win.C1InputPanel.InputLabel()
        Me.txtDatabaseUser = New C1.Win.C1InputPanel.InputTextBox()
        Me.InputLabel6 = New C1.Win.C1InputPanel.InputLabel()
        Me.txtDatabasePort = New C1.Win.C1InputPanel.InputTextBox()
        Me.InputLabel8 = New C1.Win.C1InputPanel.InputLabel()
        Me.txtDatabasePassword = New C1.Win.C1InputPanel.InputTextBox()
        Me.InputSeparator1 = New C1.Win.C1InputPanel.InputSeparator()
        Me.InputLabel9 = New C1.Win.C1InputPanel.InputLabel()
        Me.rbMySQL = New C1.Win.C1InputPanel.InputRadioButton()
        Me.rbSQLServer = New C1.Win.C1InputPanel.InputRadioButton()
        Me.btnSave = New C1.Win.C1InputPanel.InputButton()
        CType(Me.cipLogin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cipLogin
        '
        Me.cipLogin.BackColor = System.Drawing.Color.White
        Me.cipLogin.BorderThickness = 1
        Me.cipLogin.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cipLogin.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.cipLogin.Items.Add(Me.InputImage1)
        Me.cipLogin.Items.Add(Me.InputLabel1)
        Me.cipLogin.Items.Add(Me.txtUserLogin)
        Me.cipLogin.Items.Add(Me.InputLabel2)
        Me.cipLogin.Items.Add(Me.txtPassword)
        Me.cipLogin.Items.Add(Me.cmdSubmit)
        Me.cipLogin.Items.Add(Me.cmdClose)
        Me.cipLogin.Items.Add(Me.InputLabel3)
        Me.cipLogin.Items.Add(Me.InputSeparator2)
        Me.cipLogin.Items.Add(Me.btnNetwork)
        Me.cipLogin.Items.Add(Me.btnFolder)
        Me.cipLogin.Items.Add(Me.InputSeparator4)
        Me.cipLogin.Items.Add(Me.InputLabel4)
        Me.cipLogin.Items.Add(Me.txtTerminalID)
        Me.cipLogin.Items.Add(Me.ighDatabase)
        Me.cipLogin.Items.Add(Me.InputLabel5)
        Me.cipLogin.Items.Add(Me.txtDatabaseServer)
        Me.cipLogin.Items.Add(Me.InputLabel7)
        Me.cipLogin.Items.Add(Me.txtDatabaseUser)
        Me.cipLogin.Items.Add(Me.InputLabel6)
        Me.cipLogin.Items.Add(Me.txtDatabasePort)
        Me.cipLogin.Items.Add(Me.InputLabel8)
        Me.cipLogin.Items.Add(Me.txtDatabasePassword)
        Me.cipLogin.Items.Add(Me.InputSeparator1)
        Me.cipLogin.Items.Add(Me.InputLabel9)
        Me.cipLogin.Items.Add(Me.rbMySQL)
        Me.cipLogin.Items.Add(Me.rbSQLServer)
        Me.cipLogin.Items.Add(Me.btnSave)
        Me.cipLogin.Location = New System.Drawing.Point(0, 0)
        Me.cipLogin.Name = "cipLogin"
        Me.cipLogin.Size = New System.Drawing.Size(483, 240)
        Me.cipLogin.TabIndex = 0
        Me.cipLogin.VisualStyle = C1.Win.C1InputPanel.VisualStyle.Office2010Blue
        '
        'InputImage1
        '
        Me.InputImage1.Height = 80
        Me.InputImage1.Image = CType(resources.GetObject("InputImage1.Image"), System.Drawing.Image)
        Me.InputImage1.Name = "InputImage1"
        Me.InputImage1.Width = 202
        '
        'InputLabel1
        '
        Me.InputLabel1.Break = C1.Win.C1InputPanel.BreakType.Row
        Me.InputLabel1.Name = "InputLabel1"
        Me.InputLabel1.Text = "User: "
        Me.InputLabel1.Width = 55
        '
        'txtUserLogin
        '
        Me.txtUserLogin.Name = "txtUserLogin"
        Me.txtUserLogin.ToolTipText = "<i>Username</i>"
        Me.txtUserLogin.Width = 202
        '
        'InputLabel2
        '
        Me.InputLabel2.Break = C1.Win.C1InputPanel.BreakType.Row
        Me.InputLabel2.Name = "InputLabel2"
        Me.InputLabel2.Text = "Password: "
        Me.InputLabel2.Width = 55
        '
        'txtPassword
        '
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Width = 202
        '
        'cmdSubmit
        '
        Me.cmdSubmit.Break = C1.Win.C1InputPanel.BreakType.None
        Me.cmdSubmit.Height = 35
        Me.cmdSubmit.Name = "cmdSubmit"
        Me.cmdSubmit.Text = "SIGN-IN"
        Me.cmdSubmit.Width = 100
        '
        'cmdClose
        '
        Me.cmdClose.Break = C1.Win.C1InputPanel.BreakType.Column
        Me.cmdClose.Height = 35
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Text = "CLOSE"
        Me.cmdClose.Width = 100
        '
        'InputLabel3
        '
        Me.InputLabel3.Break = C1.Win.C1InputPanel.BreakType.Row
        Me.InputLabel3.Font = New System.Drawing.Font("Comic Sans MS", 17.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InputLabel3.Height = 121
        Me.InputLabel3.HorizontalAlign = C1.Win.C1InputPanel.InputContentAlignment.Center
        Me.InputLabel3.Name = "InputLabel3"
        Me.InputLabel3.Text = "so.simpi-pro"
        Me.InputLabel3.VerticalAlign = C1.Win.C1InputPanel.InputContentAlignment.Center
        Me.InputLabel3.Width = 250
        '
        'InputSeparator2
        '
        Me.InputSeparator2.Name = "InputSeparator2"
        Me.InputSeparator2.Width = 250
        '
        'btnNetwork
        '
        Me.btnNetwork.Break = C1.Win.C1InputPanel.BreakType.None
        Me.btnNetwork.Height = 35
        Me.btnNetwork.Name = "btnNetwork"
        Me.btnNetwork.Text = "TERMINAL LIST"
        Me.btnNetwork.Width = 122
        '
        'btnFolder
        '
        Me.btnFolder.Height = 35
        Me.btnFolder.Name = "btnFolder"
        Me.btnFolder.Text = "APPS. FOLDER"
        Me.btnFolder.Width = 122
        '
        'InputSeparator4
        '
        Me.InputSeparator4.Name = "InputSeparator4"
        Me.InputSeparator4.Width = 246
        '
        'InputLabel4
        '
        Me.InputLabel4.Name = "InputLabel4"
        Me.InputLabel4.Text = "TerminalID: "
        '
        'txtTerminalID
        '
        Me.txtTerminalID.Break = C1.Win.C1InputPanel.BreakType.Column
        Me.txtTerminalID.Name = "txtTerminalID"
        Me.txtTerminalID.ReadOnly = True
        Me.txtTerminalID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtTerminalID.Width = 178
        '
        'ighDatabase
        '
        Me.ighDatabase.Collapsed = True
        Me.ighDatabase.Collapsible = True
        Me.ighDatabase.Name = "ighDatabase"
        Me.ighDatabase.Text = "LOCAL CONNECTION SETTING"
        '
        'InputLabel5
        '
        Me.InputLabel5.HorizontalAlign = C1.Win.C1InputPanel.InputContentAlignment.Far
        Me.InputLabel5.Name = "InputLabel5"
        Me.InputLabel5.Text = "Server: "
        Me.InputLabel5.Width = 65
        '
        'txtDatabaseServer
        '
        Me.txtDatabaseServer.Break = C1.Win.C1InputPanel.BreakType.None
        Me.txtDatabaseServer.Name = "txtDatabaseServer"
        Me.txtDatabaseServer.Width = 150
        '
        'InputLabel7
        '
        Me.InputLabel7.HorizontalAlign = C1.Win.C1InputPanel.InputContentAlignment.Far
        Me.InputLabel7.Name = "InputLabel7"
        Me.InputLabel7.Text = "Username:"
        Me.InputLabel7.Width = 65
        '
        'txtDatabaseUser
        '
        Me.txtDatabaseUser.Name = "txtDatabaseUser"
        Me.txtDatabaseUser.Width = 174
        '
        'InputLabel6
        '
        Me.InputLabel6.HorizontalAlign = C1.Win.C1InputPanel.InputContentAlignment.Far
        Me.InputLabel6.Name = "InputLabel6"
        Me.InputLabel6.Text = "Port: "
        Me.InputLabel6.Width = 65
        '
        'txtDatabasePort
        '
        Me.txtDatabasePort.Break = C1.Win.C1InputPanel.BreakType.None
        Me.txtDatabasePort.Name = "txtDatabasePort"
        Me.txtDatabasePort.Padding = New System.Windows.Forms.Padding(0, 0, 100, 0)
        Me.txtDatabasePort.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtDatabasePort.Width = 150
        '
        'InputLabel8
        '
        Me.InputLabel8.HorizontalAlign = C1.Win.C1InputPanel.InputContentAlignment.Far
        Me.InputLabel8.Name = "InputLabel8"
        Me.InputLabel8.Text = "Password: "
        Me.InputLabel8.Width = 65
        '
        'txtDatabasePassword
        '
        Me.txtDatabasePassword.Name = "txtDatabasePassword"
        Me.txtDatabasePassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtDatabasePassword.Width = 174
        '
        'InputSeparator1
        '
        Me.InputSeparator1.Name = "InputSeparator1"
        Me.InputSeparator1.Width = 464
        '
        'InputLabel9
        '
        Me.InputLabel9.Height = 35
        Me.InputLabel9.HorizontalAlign = C1.Win.C1InputPanel.InputContentAlignment.Far
        Me.InputLabel9.Name = "InputLabel9"
        Me.InputLabel9.Text = "Database: "
        Me.InputLabel9.Width = 65
        '
        'rbMySQL
        '
        Me.rbMySQL.Break = C1.Win.C1InputPanel.BreakType.None
        Me.rbMySQL.Checked = True
        Me.rbMySQL.Height = 35
        Me.rbMySQL.Name = "rbMySQL"
        Me.rbMySQL.Text = "MySQL/MariaDB"
        Me.rbMySQL.VerticalAlign = C1.Win.C1InputPanel.InputContentAlignment.Near
        '
        'rbSQLServer
        '
        Me.rbSQLServer.Break = C1.Win.C1InputPanel.BreakType.None
        Me.rbSQLServer.Height = 35
        Me.rbSQLServer.Name = "rbSQLServer"
        Me.rbSQLServer.Text = "Microsoft SQL Server"
        Me.rbSQLServer.VerticalAlign = C1.Win.C1InputPanel.InputContentAlignment.Near
        Me.rbSQLServer.Width = 177
        '
        'btnSave
        '
        Me.btnSave.Height = 35
        Me.btnSave.Image = CType(resources.GetObject("btnSave.Image"), System.Drawing.Image)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Text = "SETTING"
        Me.btnSave.Width = 100
        '
        'FormLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(483, 240)
        Me.ControlBox = False
        Me.Controls.Add(Me.cipLogin)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "LOG-ON"
        Me.TopMost = True
        CType(Me.cipLogin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cipLogin As C1.Win.C1InputPanel.C1InputPanel
    Friend WithEvents InputLabel1 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents txtUserLogin As C1.Win.C1InputPanel.InputTextBox
    Friend WithEvents InputLabel2 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents txtPassword As C1.Win.C1InputPanel.InputTextBox
    Friend WithEvents cmdSubmit As C1.Win.C1InputPanel.InputButton
    Friend WithEvents cmdClose As C1.Win.C1InputPanel.InputButton
    Friend WithEvents InputImage1 As C1.Win.C1InputPanel.InputImage
    Friend WithEvents InputLabel3 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents InputSeparator2 As C1.Win.C1InputPanel.InputSeparator
    Friend WithEvents InputLabel4 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents txtTerminalID As C1.Win.C1InputPanel.InputTextBox
    Friend WithEvents btnFolder As C1.Win.C1InputPanel.InputButton
    Friend WithEvents InputSeparator4 As C1.Win.C1InputPanel.InputSeparator
    Friend WithEvents InputLabel5 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents txtDatabaseServer As C1.Win.C1InputPanel.InputTextBox
    Friend WithEvents InputLabel6 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents txtDatabasePort As C1.Win.C1InputPanel.InputTextBox
    Friend WithEvents InputLabel7 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents txtDatabaseUser As C1.Win.C1InputPanel.InputTextBox
    Friend WithEvents InputLabel8 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents txtDatabasePassword As C1.Win.C1InputPanel.InputTextBox
    Friend WithEvents ighDatabase As C1.Win.C1InputPanel.InputGroupHeader
    Friend WithEvents btnSave As C1.Win.C1InputPanel.InputButton
    Friend WithEvents btnNetwork As C1.Win.C1InputPanel.InputButton
    Friend WithEvents InputSeparator1 As C1.Win.C1InputPanel.InputSeparator
    Friend WithEvents InputLabel9 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents rbMySQL As C1.Win.C1InputPanel.InputRadioButton
    Friend WithEvents rbSQLServer As C1.Win.C1InputPanel.InputRadioButton
End Class
