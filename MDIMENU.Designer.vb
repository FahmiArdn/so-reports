<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MDIMENU
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MDIMENU))
        Me.msMenu = New System.Windows.Forms.MenuStrip()
        Me.tsUser = New System.Windows.Forms.ToolStrip()
        Me.tsbExit = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator10 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnStatus = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbUserProfile = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.lblAsOf = New System.Windows.Forms.ToolStripLabel()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.lblStatus = New System.Windows.Forms.ToolStripLabel()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.lblSpeed = New System.Windows.Forms.ToolStripLabel()
        Me.tmr = New System.Windows.Forms.Timer(Me.components)
        Me.bw = New System.ComponentModel.BackgroundWorker()
        Me.C1StatusBar1 = New C1.Win.C1Ribbon.C1StatusBar()
        Me.lblVersion = New C1.Win.C1Ribbon.RibbonLabel()
        Me.RibbonSeparator1 = New C1.Win.C1Ribbon.RibbonSeparator()
        Me.lblUser = New C1.Win.C1Ribbon.RibbonLabel()
        Me.RibbonSeparator2 = New C1.Win.C1Ribbon.RibbonSeparator()
        Me.lblTerminal = New C1.Win.C1Ribbon.RibbonLabel()
        Me.cipLogin = New C1.Win.C1InputPanel.C1InputPanel()
        Me.InputImage1 = New C1.Win.C1InputPanel.InputImage()
        Me.InputLabel1 = New C1.Win.C1InputPanel.InputLabel()
        Me.txtUserLogin = New C1.Win.C1InputPanel.InputTextBox()
        Me.InputLabel2 = New C1.Win.C1InputPanel.InputLabel()
        Me.txtPassword = New C1.Win.C1InputPanel.InputTextBox()
        Me.cmdSubmit = New C1.Win.C1InputPanel.InputButton()
        Me.cmdClose = New C1.Win.C1InputPanel.InputButton()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.msMenu.SuspendLayout()
        Me.tsUser.SuspendLayout()
        CType(Me.C1StatusBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cipLogin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'msMenu
        '
        Me.msMenu.Enabled = False
        Me.msMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportToolStripMenuItem})
        Me.msMenu.Location = New System.Drawing.Point(0, 0)
        Me.msMenu.Name = "msMenu"
        Me.msMenu.Size = New System.Drawing.Size(1064, 24)
        Me.msMenu.TabIndex = 5
        Me.msMenu.Text = "MenuStrip"
        '
        'tsUser
        '
        Me.tsUser.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbExit, Me.ToolStripSeparator10, Me.btnStatus, Me.ToolStripSeparator9, Me.tsbUserProfile, Me.ToolStripSeparator1, Me.lblAsOf, Me.ToolStripSeparator3, Me.lblStatus, Me.ToolStripSeparator4, Me.lblSpeed})
        Me.tsUser.Location = New System.Drawing.Point(0, 24)
        Me.tsUser.Name = "tsUser"
        Me.tsUser.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.tsUser.Size = New System.Drawing.Size(1064, 25)
        Me.tsUser.TabIndex = 7
        Me.tsUser.Text = "ToolStrip1"
        '
        'tsbExit
        '
        Me.tsbExit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbExit.Image = CType(resources.GetObject("tsbExit.Image"), System.Drawing.Image)
        Me.tsbExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbExit.Name = "tsbExit"
        Me.tsbExit.Size = New System.Drawing.Size(23, 22)
        Me.tsbExit.Text = "Log-Out"
        '
        'ToolStripSeparator10
        '
        Me.ToolStripSeparator10.Name = "ToolStripSeparator10"
        Me.ToolStripSeparator10.Size = New System.Drawing.Size(6, 25)
        '
        'btnStatus
        '
        Me.btnStatus.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnStatus.Image = CType(resources.GetObject("btnStatus.Image"), System.Drawing.Image)
        Me.btnStatus.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnStatus.Name = "btnStatus"
        Me.btnStatus.Size = New System.Drawing.Size(23, 22)
        Me.btnStatus.Text = "Re-Connect"
        '
        'ToolStripSeparator9
        '
        Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
        Me.ToolStripSeparator9.Size = New System.Drawing.Size(6, 25)
        '
        'tsbUserProfile
        '
        Me.tsbUserProfile.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbUserProfile.Enabled = False
        Me.tsbUserProfile.Image = CType(resources.GetObject("tsbUserProfile.Image"), System.Drawing.Image)
        Me.tsbUserProfile.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbUserProfile.Name = "tsbUserProfile"
        Me.tsbUserProfile.Size = New System.Drawing.Size(23, 22)
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'lblAsOf
        '
        Me.lblAsOf.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Italic)
        Me.lblAsOf.Name = "lblAsOf"
        Me.lblAsOf.Size = New System.Drawing.Size(85, 22)
        Me.lblAsOf.Text = "dd-MMM-yyyy"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = False
        Me.lblStatus.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.lblStatus.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(125, 22)
        Me.lblStatus.Text = "LOCAL CONNECTION"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 25)
        '
        'lblSpeed
        '
        Me.lblSpeed.Name = "lblSpeed"
        Me.lblSpeed.Size = New System.Drawing.Size(141, 22)
        Me.lblSpeed.Text = "connection speed: xxx ms"
        '
        'tmr
        '
        Me.tmr.Interval = 1000
        '
        'C1StatusBar1
        '
        Me.C1StatusBar1.LeftPaneItems.Add(Me.lblVersion)
        Me.C1StatusBar1.LeftPaneItems.Add(Me.RibbonSeparator1)
        Me.C1StatusBar1.LeftPaneItems.Add(Me.lblUser)
        Me.C1StatusBar1.LeftPaneItems.Add(Me.RibbonSeparator2)
        Me.C1StatusBar1.LeftPaneItems.Add(Me.lblTerminal)
        Me.C1StatusBar1.Location = New System.Drawing.Point(0, 431)
        Me.C1StatusBar1.Name = "C1StatusBar1"
        Me.C1StatusBar1.Size = New System.Drawing.Size(1064, 22)
        '
        'lblVersion
        '
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Text = "Label"
        '
        'RibbonSeparator1
        '
        Me.RibbonSeparator1.Name = "RibbonSeparator1"
        '
        'lblUser
        '
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Text = "Label"
        '
        'RibbonSeparator2
        '
        Me.RibbonSeparator2.Name = "RibbonSeparator2"
        '
        'lblTerminal
        '
        Me.lblTerminal.Name = "lblTerminal"
        Me.lblTerminal.Text = "Label"
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
        Me.cipLogin.Location = New System.Drawing.Point(0, 49)
        Me.cipLogin.Name = "cipLogin"
        Me.cipLogin.Size = New System.Drawing.Size(1064, 382)
        Me.cipLogin.TabIndex = 9
        Me.cipLogin.Visible = False
        Me.cipLogin.VisualStyle = C1.Win.C1InputPanel.VisualStyle.Office2010Blue
        '
        'InputImage1
        '
        Me.InputImage1.Break = C1.Win.C1InputPanel.BreakType.Column
        Me.InputImage1.Height = 120
        Me.InputImage1.Image = CType(resources.GetObject("InputImage1.Image"), System.Drawing.Image)
        Me.InputImage1.Name = "InputImage1"
        Me.InputImage1.Width = 280
        '
        'InputLabel1
        '
        Me.InputLabel1.Break = C1.Win.C1InputPanel.BreakType.Row
        Me.InputLabel1.Name = "InputLabel1"
        Me.InputLabel1.Text = "User: "
        '
        'txtUserLogin
        '
        Me.txtUserLogin.Name = "txtUserLogin"
        Me.txtUserLogin.Width = 205
        '
        'InputLabel2
        '
        Me.InputLabel2.Break = C1.Win.C1InputPanel.BreakType.Row
        Me.InputLabel2.Name = "InputLabel2"
        Me.InputLabel2.Text = "Password: "
        '
        'txtPassword
        '
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Width = 205
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
        Me.cmdClose.Height = 35
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Text = "CLOSE"
        Me.cmdClose.Width = 100
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.ReportToolStripMenuItem.Text = "Report"
        '
        'MDIMENU
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1064, 453)
        Me.Controls.Add(Me.cipLogin)
        Me.Controls.Add(Me.C1StatusBar1)
        Me.Controls.Add(Me.tsUser)
        Me.Controls.Add(Me.msMenu)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.msMenu
        Me.Name = "MDIMENU"
        Me.Text = "APPLIKASI TEST"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.msMenu.ResumeLayout(False)
        Me.msMenu.PerformLayout()
        Me.tsUser.ResumeLayout(False)
        Me.tsUser.PerformLayout()
        CType(Me.C1StatusBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cipLogin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents msMenu As System.Windows.Forms.MenuStrip
    Friend WithEvents tsUser As ToolStrip
    Friend WithEvents tmr As Timer
    Friend WithEvents bw As System.ComponentModel.BackgroundWorker
    Friend WithEvents tsbUserProfile As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents lblAsOf As ToolStripLabel
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents lblStatus As ToolStripLabel
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents lblSpeed As ToolStripLabel
    Friend WithEvents C1StatusBar1 As C1.Win.C1Ribbon.C1StatusBar
    Friend WithEvents cipLogin As C1.Win.C1InputPanel.C1InputPanel
    Friend WithEvents InputImage1 As C1.Win.C1InputPanel.InputImage
    Friend WithEvents InputLabel1 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents txtUserLogin As C1.Win.C1InputPanel.InputTextBox
    Friend WithEvents InputLabel2 As C1.Win.C1InputPanel.InputLabel
    Friend WithEvents txtPassword As C1.Win.C1InputPanel.InputTextBox
    Friend WithEvents cmdSubmit As C1.Win.C1InputPanel.InputButton
    Friend WithEvents cmdClose As C1.Win.C1InputPanel.InputButton
    Friend WithEvents lblVersion As C1.Win.C1Ribbon.RibbonLabel
    Friend WithEvents RibbonSeparator1 As C1.Win.C1Ribbon.RibbonSeparator
    Friend WithEvents lblUser As C1.Win.C1Ribbon.RibbonLabel
    Friend WithEvents RibbonSeparator2 As C1.Win.C1Ribbon.RibbonSeparator
    Friend WithEvents lblTerminal As C1.Win.C1Ribbon.RibbonLabel
    Friend WithEvents tsbExit As ToolStripButton
    Friend WithEvents ToolStripSeparator10 As ToolStripSeparator
    Friend WithEvents btnStatus As ToolStripButton
    Friend WithEvents ToolStripSeparator9 As ToolStripSeparator
    Friend WithEvents ReportToolStripMenuItem As ToolStripMenuItem
End Class
