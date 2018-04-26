Imports C1.Win.C1TrueDBGrid
Imports simpi.MasterPortfolio.ParameterPortfolio
Imports simpi.GlobalUtilities
Imports simpi.GlobalUtilities.GlobalString

Public Class SelectMasterPortfolio
    Public lblCode As C1.Win.C1InputPanel.InputLabel
    Public lblName As C1.Win.C1InputPanel.InputLabel
    Public lblSimpiName As C1.Win.C1InputPanel.InputLabel
    Public lblSimpiEmail As C1.Win.C1InputPanel.InputLabel

    Private Sub FormSelectSimpiPortfolio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetPortfolioUser()
        GetParameterCountry()
        GetParameterPortfolioStatus()
        GetParameterPortfolioType()
        btnSelect.Enabled = False
        DBGPortfolio.FetchRowStyles = True
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        PortfolioSearch()
    End Sub

    Private Sub txtKeyword_KeyDown(sender As Object, e As KeyEventArgs) Handles txtKeyword.KeyDown
        If e.KeyCode = Keys.Enter Then
            PortfolioSearch()
        End If
    End Sub

    Private Sub PortfolioSearch()
        Dim strFilter As String = PortfolioKeyword(SQLKeyword(txtKeyword.Text, IIf(rbNone.Checked, SetNone, IIf(rbAnd.Checked, SetAnd, IIf(rbOr.Checked, SetOr, SetCombine)))))
        Dim rData As DataRow() = dtPortfolioUser.Select(strFilter)
        If rData IsNot Nothing AndAlso rData.Length > 0 Then
            Dim query = From p In rData.AsEnumerable
                        Join c In dtParameterCountry.AsEnumerable On p.Field(Of Integer)("CcyID") Equals c.Field(Of Integer)("CountryID")
                        Join s In dtParameterPortfolioStatus.AsEnumerable On p.Field(Of Integer)("StatusID") Equals s.Field(Of Integer)("StatusID")
                        Join l In dtParameterPortfolioType.AsEnumerable On p.Field(Of Integer)("TypeID") Equals l.Field(Of Integer)("TypeID")
                        Select
                             ID = p.Field(Of Integer)("PortfolioID"),
                             simpi = p.Field(Of String)("simpiName"),
                             simpiEmail = p.Field(Of String)("simpiEmail"),
                             Code = p.Field(Of String)("PortfolioCode"),
                             Shortname = p.Field(Of String)("PortfolioNameshort"),
                             Ccy = c.Field(Of String)("Ccy"),
                             Status = s.Field(Of String)("StatusCode"),
                             Level = l.Field(Of String)("TypeCode")

            With DBGPortfolio
                .AllowAddNew = False
                .AllowDelete = False
                .AllowUpdate = False
                .Style.WrapText = False
                .Columns.Clear()
                .DataSource = query.ToList

                For Each column As C1DisplayColumn In DBGPortfolio.Splits(0).DisplayColumns
                    If column.Name.Trim = "simpiEmail" Then column.Visible = False Else column.AutoSize()
                Next

            End With
        Else
            DBGPortfolio.Columns.Clear()
            ExceptionMessage.Show(objError.Message(2) & " master portfolio", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub DBGPortfolio_FetchRowStyle(sender As Object, e As FetchRowStyleEventArgs) Handles DBGPortfolio.FetchRowStyle
        If e.Row Mod 2 = 0 Then e.CellStyle.BackColor = Color.LemonChiffon
    End Sub

    Private Sub DBGPortfolio_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DBGPortfolio.Click
        If DBGPortfolio.RowCount > 0 Then
            DBGPortfolio.MarqueeStyle = MarqueeEnum.HighlightRow
            btnSelect.Enabled = True
        End If
    End Sub

    Private Sub DBGSimpi_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DBGPortfolio.DoubleClick
        PortfolioSelect()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Close()
    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        PortfolioSelect()
    End Sub

    Private Sub PortfolioSelect()
        With DBGPortfolio
            If .RowCount > 0 Then
                lblCode.Text = .Columns("Code").Text
                lblName.Text = .Columns("Shortname").Text
                lblSimpiName.Text = .Columns("simpi").Text
                lblSimpiEmail.Text = .Columns("simpiEmail").Text
                Close()
            End If
        End With
    End Sub

End Class