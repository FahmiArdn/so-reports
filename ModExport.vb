Imports Microsoft.Office.Interop
Imports simpi.GlobalConnection

Module ModExport
    Friend Function simpiFile(strFile As String) As String
        If My.Application.Info.DirectoryPath.Length = 3 Then
            Return My.Application.Info.DirectoryPath.Trim & strFile.Trim
        Else
            Return My.Application.Info.DirectoryPath.Trim & "\" & strFile.Trim
        End If
    End Function

    Friend Function RGBWrite(ByVal type As String, ByVal color As Integer)
        Return type.Trim & ": " & color
    End Function

    Friend Function RGBRead(ByVal str As String) As Integer
        If str.Length > 3 Then Return CInt(Mid(str.Trim, 3)) Else Return 0
    End Function

    Friend Function RGBCheck(ByVal v As Integer) As Integer
        If v < 0 Or v > 255 Then Return 0 Else Return v
    End Function

    Friend Function CMYKCheck(ByVal i As Double) As Double
        If i < 0 Or i > 100 Then Return 0 Else Return i
    End Function

    Friend Sub RGBConvertCMYK(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer,
        ByRef C As Double, ByRef M As Double, ByRef Y As Double, ByRef K As Double)
        Dim r1, g1, b1, max As Double
        r1 = r / 255
        g1 = g / 255
        b1 = b / 255
        If r1 > g1 Then max = r1 Else max = g1
        If max < b1 Then max = b1
        K = 1 - max
        C = (1 - r1 - K) / (1 - K) * 100
        M = (1 - g1 - K) / (1 - K) * 100
        Y = (1 - b1 - K) / (1 - K) * 100
    End Sub

    Friend Sub CMYKConvertRGB(ByVal C As Double, ByVal M As Double, ByVal Y As Double, ByVal K As Double,
                              ByRef r As Integer, ByRef g As Integer, ByRef b As Integer)
        K = K / 100
        C = C / 100
        Y = Y / 100
        M = M / 100
        r = 255 * (1 - C) * (1 - K)
        g = 255 * (1 - M) * (1 - K)
        b = 255 * (1 - Y) * (1 - K)
    End Sub

    Friend Sub simpiLogo(ByVal pdfReport As C1.C1Pdf.C1PdfDocument, ByVal rc As RectangleF)
        Dim simpi As Image = My.Resources.simpi3
        pdfReport.DrawString("powered by", New Font("Calibri", 6), New SolidBrush(Color.Black),
                             New PointF(rc.Width - (0.3 * simpi.Size.Width) - 40, rc.Height - (0.3 * simpi.Size.Height) - 10))
        pdfReport.DrawImage(simpi,
                  New RectangleF(New PointF(rc.Width - (0.3 * simpi.Size.Width) - 10, rc.Height - (0.3 * simpi.Size.Height) - 10),
                  New SizeF(0.3 * simpi.Size.Width, 0.3 * simpi.Size.Height)))
    End Sub

    Friend Function FromRtf(ByVal rtf As RichTextBox) As String
        Dim fontfamily As String = "Arial"
        Dim fontsize As Integer = 12
        Dim htmlstr As String = String.Format("<html>{0}<body>{0}<div style=""text-align: left;""><span style=""font-family: Arial; font-size: 12pt;"">", vbCrLf)
        Dim x As Integer = 0
        Dim curchar As String
        While x < rtf.Text.Length
            rtf.Select(x, 1)
            curchar = rtf.SelectedText
            Select Case curchar
                Case vbCr, vbLf : htmlstr &= "<br/>"
                Case Else : htmlstr &= curchar
            End Select
            x += 1
        End While
        Return htmlstr & String.Format("</span>{0}</body>{0}</html>", vbCrLf)
    End Function

    Friend Function reportFileExists(ByVal fileName As String) As String
        Dim strPath As String = My.Application.Info.DirectoryPath() & "\Temp"
        If Not GlobalFileWindows.FolderExist(strPath) Then GlobalFileWindows.FolderCreate(strPath)
        Dim strFile As String = strPath.Trim & "\" & fileName.Trim
        If GlobalFileWindows.FileExists(strFile) Then GlobalFileWindows.FileDelete(strFile)
        Return strFile
    End Function

    Friend Function reportPDFPortrait(ByVal Codeset As String) As String
        Dim strPath As String = My.Application.Info.DirectoryPath() & "\Template"
        If Not GlobalFileWindows.FolderExist(strPath) Then GlobalFileWindows.FolderCreate(strPath)
        Dim strFile As String = strPath.Trim & "\" & objLayout.GetFieldData(Codeset)
        Return strFile
    End Function

    Friend Sub PrintExcel(ByVal DBG As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
        If DBG.RowCount > 0 Then
            Dim xls As Excel.Application
            Dim xlsWorkBook As Excel.Workbook
            Dim xlsWorkSheet As Excel.Worksheet
            xls = New Excel.Application
            xls.Visible = False
            xlsWorkBook = xls.Workbooks.Add
            xlsWorkSheet = xlsWorkBook.Sheets("sheet1")
            xls.ActiveWindow.DisplayGridlines = False
            For j = 0 To DBG.Columns.Count - 1
                xlsWorkSheet.Cells(1, j + 1) = DBG.Columns(j).Caption
            Next
            For i = 0 To DBG.RowCount - 1
                DBG.Row = i
                For j = 0 To DBG.Columns.Count - 1
                    xlsWorkSheet.Cells(i + 2, j + 1) = DBG.Columns(j).Text
                Next
            Next
            xlsWorkSheet.Columns("A:Z").EntireColumn.AutoFit
            xls.Visible = True
        End If
    End Sub

    Friend Sub PrintExcel(ByVal fg As C1.Win.C1FlexGrid.C1FlexGrid)
        If fg.Rows.Count > 1 Then
            Dim xls As Excel.Application
            Dim xlsWorkBook As Excel.Workbook
            Dim xlsWorkSheet As Excel.Worksheet
            xls = New Excel.Application
            xls.Visible = False
            xlsWorkBook = xls.Workbooks.Add
            xlsWorkSheet = xlsWorkBook.Sheets("sheet1")
            xls.ActiveWindow.DisplayGridlines = False
            For i = 0 To fg.Rows.Count - 1
                For j = 0 To fg.Cols.Count - 1
                    xlsWorkSheet.Cells(i + 1, j + 1) = fg(i, j).ToString
                Next
            Next
            xlsWorkSheet.Columns("A:Z").EntireColumn.AutoFit
            xls.Visible = True
        End If
    End Sub

    Friend Function PrintExcel(ByVal isAttachment As Boolean, ByVal DBG As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal reportFile As String) As String
        Dim strFile As String = ""
        If DBG.RowCount > 0 Then
            Dim xls As Excel.Application
            Dim xlsWorkBook As Excel.Workbook
            Dim xlsWorkSheet As Excel.Worksheet
            xls = New Excel.Application
            xls.Visible = False
            xlsWorkBook = xls.Workbooks.Add
            xlsWorkSheet = xlsWorkBook.Sheets("sheet1")
            xls.ActiveWindow.DisplayGridlines = False

            For j = 0 To DBG.Columns.Count - 1
                xlsWorkSheet.Cells(1, j + 1) = DBG.Columns(j).Caption
            Next
            For i = 0 To DBG.RowCount - 1
                DBG.Row = i
                For j = 0 To DBG.Columns.Count - 1
                    xlsWorkSheet.Cells(i + 2, j + 1) = DBG.Columns(j).Text
                Next
            Next
            xlsWorkSheet.Columns("A:Z").EntireColumn.AutoFit

            If isAttachment Then
                strFile = reportFileExists(reportFile)
                xlsWorkSheet.SaveAs(strFile)
                xlsWorkBook.Close()
                xls.Quit()
                releaseObject(xls)
                releaseObject(xlsWorkBook)
                releaseObject(xlsWorkSheet)
            Else
                xls.Visible = True
            End If

        End If
        Return strFile
    End Function

    Friend Function PrintExcel(ByVal isAttachment As Boolean, ByVal fg As C1.Win.C1FlexGrid.C1FlexGrid, ByVal reportFile As String) As String
        Dim strFile As String = ""
        If fg.Rows.Count > 1 Then
            Dim xls As Excel.Application
            Dim xlsWorkBook As Excel.Workbook
            Dim xlsWorkSheet As Excel.Worksheet
            xls = New Excel.Application
            xls.Visible = False
            xlsWorkBook = xls.Workbooks.Add
            xlsWorkSheet = xlsWorkBook.Sheets("sheet1")
            xls.ActiveWindow.DisplayGridlines = False

            For i = 0 To fg.Rows.Count - 1
                For j = 0 To fg.Cols.Count - 1
                    xlsWorkSheet.Cells(i + 1, j + 1) = fg(i, j).ToString
                Next
            Next
            xlsWorkSheet.Columns("A:Z").EntireColumn.AutoFit

            If isAttachment Then
                strFile = reportFileExists(reportFile)
                xlsWorkSheet.SaveAs(strFile)
                xlsWorkBook.Close()
                xls.Quit()
                releaseObject(xls)
                releaseObject(xlsWorkBook)
                releaseObject(xlsWorkSheet)
            Else
                xls.Visible = True
            End If

        End If
        Return strFile
    End Function

    Private Sub releaseObject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Module
