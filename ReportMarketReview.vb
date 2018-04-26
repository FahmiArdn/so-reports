Imports simpi.MasterPortfolio
Imports simpi.GlobalUtilities.GlobalString
Imports C1.Win.C1FlexGrid
Imports simpi.GlobalUtilities

Public Class ReportMarketReview
    Public lblID As C1.Win.C1InputPanel.InputLabel
    Public lblDate As C1.Win.C1InputPanel.InputLabel
    Public txtReview As C1.Win.C1InputPanel.InputTextBox
    Public objPortfolio As MasterPortfolio

    Dim objReview As New PortfolioMarketReview
    Private Sub ReportMarketReview_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtAs.Value = Now
        objReview.UserAccess = objAccess
        fgReview.DrawMode = DrawModeEnum.OwnerDraw
    End Sub

    Private Sub reviewInit()
        With fgReview
            .Rows.Count = 1
            .Cols.Count = 4
            .ExtendLastCol = False
            fgReview(0, 0) = "ID"
            fgReview(0, 1) = "Date"
            fgReview(0, 2) = "Portfolio"
            fgReview(0, 3) = "Review"
            .AllowResizing = AllowResizingEnum.Columns
            .SelectionMode = SelectionModeEnum.Row
            .AutoSizeCols()
        End With
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        reviewSearch
    End Sub

    Private Sub reviewSearch()
        Dim tbl As New DataTable
        objReview.Clear()
        tbl = objReview.Search(objPortfolio, dtAs.Value)
        reviewInit()
        If tbl IsNot Nothing And tbl.Rows.Count > 0 Then
            With fgReview
                For i = 0 To tbl.Rows.Count - 1
                    .Rows.Add()
                    fgReview(.Rows.Count - 1, 0) = GetNullData(tbl.Rows(i)("ReviewID"), 1)
                    fgReview(.Rows.Count - 1, 1) = CDate(GetNullData(tbl.Rows(i)("ReviewDate"), 2)).ToString("dd-MMM-yyyy")
                    fgReview(.Rows.Count - 1, 2) = IIf(GetNullData(tbl.Rows(i)("PortfolioID"), 1) = 0, "", objPortfolio.GetPortfolioCode)
                    fgReview(.Rows.Count - 1, 3) = GetNullData(tbl.Rows(i)("ReviewText"), 0)
                    .Rows(.Rows.Count - 1).Height=50
                Next
            End With
        End If
    End Sub

    Private Sub fgReview_OwnerDrawCell(sender As Object, e As OwnerDrawCellEventArgs) Handles fgReview.OwnerDrawCell
        Dim s As CellStyle = fgReview.Styles.Add("RowStyle")
        s.BackColor = Color.LemonChiffon
        If e.Row > 0 And e.Row Mod 2 = 0 Then fgReview.Rows(e.Row).Style = fgReview.Styles("RowStyle")
    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        reviewSelect()
    End Sub

    Private Sub reviewSelect()
        With fgReview
            If .Rows.Count > 1 AndAlso .Row > 0 Then
                If fgReview(.Row, 2).ToString.Trim <> "" Then
                    lblID.Text = fgReview(.Row, 0)
                    lblDate.Text = fgReview(.Row, 1)
                    txtReview.Text = fgReview(.Row, 3)
                    Close()
                Else
                    objReview.Clear()
                    objReview.Add(objPortfolio, CInt(fgReview(.Row, 0)))
                    If objReview.ErrID = 0 Then
                        lblID.Text = fgReview(.Row, 0)
                        lblDate.Text = fgReview(.Row, 1)
                        txtReview.Text = fgReview(.Row, 3)
                        Close()
                    Else
                        ExceptionMessage.Show(objReview.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        reviewRemove
    End Sub

    Private Sub reviewRemove()
        With fgReview
            If .Rows.Count > 1 AndAlso .Row > 0 Then
                objReview.Clear()
                objReview.Remove(objMasterSimpi, CInt(fgReview(.Row, 0)))
                If objReview.ErrID = 0 Then
                    reviewSearch()
                Else
                    ExceptionMessage.Show(objReview.ErrMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End With
    End Sub

End Class