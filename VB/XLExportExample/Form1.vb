Imports DevExpress.Export.Xl
Imports DevExpress.XtraExport.Csv
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

Namespace XLExportExample
    Partial Public Class Form1
        Inherits Form

        Private sales As List(Of SalesData) = SalesRepository.GetSalesData()
        Private headerRowFormatting As XlCellFormatting
        Private evenRowFormatting As XlCellFormatting
        Private oddRowFormatting As XlCellFormatting
        Private totalRowFormatting As XlCellFormatting

        Public Sub New()
            InitializeComponent()
            InitializeFormatting()
        End Sub

        Private Sub InitializeFormatting()
            ' Specify formatting settings for the even rows.
            evenRowFormatting = New XlCellFormatting()
            evenRowFormatting.Font = New XlFont()
            evenRowFormatting.Font.Name = "Century Gothic"
            evenRowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
            evenRowFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.General, XlVerticalAlignment.Center)

            ' Specify formatting settings for the odd rows.
            oddRowFormatting = New XlCellFormatting()
            oddRowFormatting.CopyFrom(evenRowFormatting)
            oddRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light1, -0.15))

            ' Specify formatting settings for the header row.
            headerRowFormatting = New XlCellFormatting()
            headerRowFormatting.CopyFrom(evenRowFormatting)
            headerRowFormatting.Font.Bold = True
            headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
            headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0))
            headerRowFormatting.Border = New XlBorder()
            headerRowFormatting.Border.TopColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0)
            headerRowFormatting.Border.TopLineStyle = XlBorderLineStyle.Medium
            headerRowFormatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0)
            headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Medium

            ' Specify formatting settings for the total row.
            totalRowFormatting = New XlCellFormatting()
            totalRowFormatting.CopyFrom(evenRowFormatting)
            totalRowFormatting.Font.Bold = True
        End Sub

        ' Export the document to XLSX format.
        Private Sub btnExportToXLSX_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToXLSX.Click
            Dim fileName As String = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            If ExportToFile(fileName, XlDocumentFormat.Xlsx) Then
                ShowFile(fileName)
            End If
        End Sub

        ' Export the document to XLS format.
        Private Sub btnExportToXLS_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToXLS.Click
            Dim fileName As String = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            If ExportToFile(fileName, XlDocumentFormat.Xls) Then
                ShowFile(fileName)
            End If
        End Sub

        ' Export the document to CSV format.
        Private Sub btnExportToCSV_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToCSV.Click
            Dim fileName As String = GetSaveFileName("CSV (Comma delimited files)(*.csv)|*.csv", "Document.csv")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            If ExportToFile(fileName, XlDocumentFormat.Csv) Then
                ShowFile(fileName)
            End If
        End Sub

        Private Function GetSaveFileName(ByVal filter As String, ByVal defaulName As String) As String
            saveFileDialog1.Filter = filter
            saveFileDialog1.FileName = defaulName
            If saveFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
                Return Nothing
            End If
            Return saveFileDialog1.FileName
        End Function

        Private Sub ShowFile(ByVal fileName As String)
            If Not File.Exists(fileName) Then
                Return
            End If
            Dim dResult As DialogResult = MessageBox.Show(String.Format("Do you want to open the resulting file?", fileName), Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If dResult = System.Windows.Forms.DialogResult.Yes Then
                Process.Start(fileName)
            End If
        End Sub

        Private Function ExportToFile(ByVal fileName As String, ByVal documentFormat As XlDocumentFormat) As Boolean
            Try
                Using stream As New FileStream(fileName, FileMode.Create)
                    ' Create an exporter instance. 
                    Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
                    ' Create a new document and begin to write it to the specified stream.
                    Using document As IXlDocument = exporter.CreateDocument(stream)
                        ' Generate the document content.
                        GenerateDocument(document)
                    End Using
                End Using
                Return True
            Catch ex As Exception
                MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

        Private Sub GenerateDocument(ByVal document As IXlDocument)
            ' Specify the document culture.
            document.Options.Culture = CultureInfo.CurrentCulture

            ' Specify options for exporting the document in CSV format.
            Dim csvOptions As CsvDataAwareExporterOptions = TryCast(document.Options, CsvDataAwareExporterOptions)
            If csvOptions IsNot Nothing Then
                csvOptions.WritePreamble = True
                csvOptions.UseCellNumberFormat = False
                csvOptions.NewlineAfterLastRow = True
            End If

            ' Add a new worksheet to the document.
            Using sheet As IXlSheet = document.CreateSheet()
                ' Specify the worksheet name.
                sheet.Name = "Annual Sales"

                ' Specify page settings.
                sheet.PageSetup = New XlPageSetup()
                ' Scale the print area to fit to one page wide.
                sheet.PageSetup.FitToPage = True
                sheet.PageSetup.FitToWidth = 1
                sheet.PageSetup.FitToHeight = 0

                ' Generate worksheet columns.
                GenerateColumns(sheet)

                ' Add the title to the documents exported to the XLSX and XLS formats.
                If document.Options.DocumentFormat <> XlDocumentFormat.Csv Then
                    GenerateTitle(sheet)
                End If

                ' Create the header row.
                GenerateHeaderRow(sheet)

                Dim firstDataRowIndex As Integer = sheet.CurrentRowIndex

                ' Create the data rows.
                For i As Integer = 0 To sales.Count - 1
                    GenerateDataRow(sheet, sales(i), (i + 1) = sales.Count)
                Next i

                ' Create the total row.
                GenerateTotalRow(sheet, firstDataRowIndex)

                ' Specify the data range to be printed.
                sheet.PrintArea = sheet.DataRange

                ' Create conditional formatting rules to be applied to worksheet data.  
                GenerateConditionalFormatting(sheet, firstDataRowIndex)
            End Using
        End Sub

        Private Sub GenerateColumns(ByVal sheet As IXlSheet)
            Dim numberFormat As XlNumberFormat ="#,##0,,""M"""

            ' Create the "State" column and set its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 140
            End Using

            ' Create the "Sales" column, adjust its width and set the specific number format for its cells.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 140
                column.ApplyFormatting(numberFormat)
            End Using

            ' Create the "Sales vs Target" column, adjust its width and format its cells as percentage values.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 120
                column.ApplyFormatting(XlNumberFormat.Percentage2)
            End Using

            ' Create the "Profit" column, adjust its width and set the specific number format for its cells.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 140
                column.ApplyFormatting(numberFormat)
            End Using

            ' Create the "Market Share" column, adjust its width and format its cells as percentage values.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 120
                column.ApplyFormatting(XlNumberFormat.Percentage)
            End Using
        End Sub

        Private Sub GenerateTitle(ByVal sheet As IXlSheet)
            ' Specify formatting settings for the document title.
            Dim formatting As New XlCellFormatting()
            formatting.Font = New XlFont()
            formatting.Font.Name = "Calibri Light"
            formatting.Font.SchemeStyle = XlFontSchemeStyles.None
            formatting.Font.Size = 24
            formatting.Font.Color = XlColor.FromTheme(XlThemeColor.Dark1, 0.35)
            formatting.Border = New XlBorder()
            formatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.35)
            formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium

            ' Add the document title.
            Using row As IXlRow = sheet.CreateRow()
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = "SALES 2014"
                    cell.Formatting = formatting
                End Using
                ' Create four empty cells with the title formatting applied.
                For i As Integer = 0 To 3
                    Using cell As IXlCell = row.CreateCell()
                        cell.Formatting = formatting
                    End Using
                Next i
            End Using

            ' Skip one row before starting to generate data rows.
            sheet.SkipRows(1)

        End Sub

        Private Sub GenerateHeaderRow(ByVal sheet As IXlSheet)
            Dim columnNames() As String = { "State", "Sales", "Sales vs Target", "Profit", "Market Share" }
            ' Create the header row and set its height.
            Using row As IXlRow = sheet.CreateRow()
                row.HeightInPixels = 25
                ' Create required cells in the header row, assign values from the columnNames array to them and apply specific formatting settings. 
                row.BulkCells(columnNames, headerRowFormatting)
            End Using
        End Sub

        Private Sub GenerateDataRow(ByVal sheet As IXlSheet, ByVal data As SalesData, ByVal isLastRow As Boolean)
            ' Create the data row to display sales information for the specific state.
            Using row As IXlRow = sheet.CreateRow()
                row.HeightInPixels = 25

                ' Specify formatting settings to be applied to the data rows to shade alternate rows. 
                Dim formatting As New XlCellFormatting()
                formatting.CopyFrom(If(row.RowIndex Mod 2 = 0, evenRowFormatting, oddRowFormatting))
                ' Set the bottom border for the last data row.
                If isLastRow Then
                    formatting.Border = New XlBorder()
                    formatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0)
                    formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium
                End If

                ' Create the cell containing the state name. 
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = data.State
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell containing sales data.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = data.ActualSales
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell that displays the difference between the actual and target sales.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = data.ActualSales / data.TargetSales - 1
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell containing the state profit. 
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = data.Profit
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell containing the percentage of a total market.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = data.MarketShare
                    cell.ApplyFormatting(formatting)
                End Using
            End Using
        End Sub

        Private Sub GenerateTotalRow(ByVal sheet As IXlSheet, ByVal firstDataRowIndex As Integer)
            ' Create the total row and set its height.
            Using row As IXlRow = sheet.CreateRow()
                row.HeightInPixels = 25

                ' Create the first cell in the row and apply specific formatting settings to this cell.
                Using cell As IXlCell = row.CreateCell()
                    cell.ApplyFormatting(totalRowFormatting)
                End Using

                ' Create the second cell in the total row and assign the SUBTOTAL function to it to calculate the average of the subtotal of the cells located in the "Sales" column.
                Using cell As IXlCell = row.CreateCell()
                    cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(1, firstDataRowIndex, 1, row.RowIndex - 1), XlSummary.Average, False))
                    cell.ApplyFormatting(totalRowFormatting)
                    cell.ApplyFormatting(CType("""Avg=""#,##0,,""M""", XlNumberFormat))
                End Using

                ' Create the third cell in the row and apply specific formatting settings to this cell.
                Using cell As IXlCell = row.CreateCell()
                    cell.ApplyFormatting(totalRowFormatting)
                End Using

                ' Create the fourth cell in the total row and assign the SUBTOTAL function to it to calculate the sum of the subtotal of the cells located in the "Profit" column.
                Using cell As IXlCell = row.CreateCell()
                    cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(3, firstDataRowIndex, 3, row.RowIndex - 1), XlSummary.Sum, False))
                    cell.ApplyFormatting(totalRowFormatting)
                    cell.ApplyFormatting(CType("""Sum=""#,##0,,""M""", XlNumberFormat))
                End Using
            End Using
        End Sub

        Private Sub GenerateConditionalFormatting(ByVal sheet As IXlSheet, ByVal firstDataRowIndex As Integer)
            ' Create an instance of the XlConditionalFormatting class to define a new rule.
            Dim formatting As New XlConditionalFormatting()
            ' Specify the cell range to which the conditional formatting rule should be applied (B4:B38).
            formatting.Ranges.Add(XlCellRange.FromLTRB(1, firstDataRowIndex, 1, sheet.CurrentRowIndex - 2))
            ' Create the rule to compare values in the "Sales" column using data bars. 
            Dim rule1 As New XlCondFmtRuleDataBar()
            ' Specify the color of data bars. 
            rule1.FillColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.4)
            ' Set the solid fill type.
            rule1.GradientFill = False
            formatting.Rules.Add(rule1)
            ' Add the specified rule to the worksheet collection of conditional formatting rules.
            sheet.ConditionalFormattings.Add(formatting)

            ' Create an instance of the XlConditionalFormatting class to define new rules.
            formatting = New XlConditionalFormatting()
            ' Specify the cell range to which the conditional formatting rules should be applied (C4:C38).
            formatting.Ranges.Add(XlCellRange.FromLTRB(2, firstDataRowIndex, 2, sheet.CurrentRowIndex - 2))
            ' Create the rule to identify negative values in the "Sales vs Target" column.
            Dim rule2 As New XlCondFmtRuleCellIs()
            ' Specify the relational operator to be used in the conditional formatting rule.
            rule2.Operator = XlCondFmtOperator.LessThan
            ' Set the threshold value.
            rule2.Value = 0
            ' Specify formatting options to be applied to cells if the condition is true.
            ' Set the font color to dark red.
            rule2.Formatting = New XlFont() With {.Color = Color.DarkRed}
            formatting.Rules.Add(rule2)
            ' Create the rule to identify top five values in the "Sales vs Target" column.
            Dim rule3 As New XlCondFmtRuleTop10()
            rule3.Rank = 5
            ' Specify formatting options to be applied to cells if the condition is true.
            ' Set the font color to dark green.
            rule3.Formatting = New XlFont() With {.Color = Color.DarkGreen}
            formatting.Rules.Add(rule3)
            ' Add the specified rules to the worksheet collection of conditional formatting rules.
            sheet.ConditionalFormattings.Add(formatting)

            ' Create an instance of the XlConditionalFormatting class to define a new rule.
            formatting = New XlConditionalFormatting()
            ' Specify the cell range to which the conditional formatting rules should be applied (D4:D38).
            formatting.Ranges.Add(XlCellRange.FromLTRB(3, firstDataRowIndex, 3, sheet.CurrentRowIndex - 2))
            ' Create the rule to compare values in the "Profit" column using data bars. 
            Dim rule4 As New XlCondFmtRuleDataBar()
            ' Specify the color of data bars. 
            rule4.FillColor = Color.FromArgb(99, 195, 132)
            ' Specify the positive bar border color.
            rule4.BorderColor = Color.FromArgb(99, 195, 132)
            ' Specify the negative bar fill color.
            rule4.NegativeFillColor = Color.FromArgb(255, 85, 90)
            ' Specify the negative bar border color.
            rule4.NegativeBorderColor = Color.FromArgb(255, 85, 90)
            ' Specify the solid fill type.
            rule4.GradientFill = False
            formatting.Rules.Add(rule4)
            ' Add the specified rule to the worksheet collection of conditional formatting rules.
            sheet.ConditionalFormattings.Add(formatting)

            ' Create an instance of the XlConditionalFormatting class to define a new rule.
            formatting = New XlConditionalFormatting()
            ' Specify the cell range to which the conditional formatting rules should be applied (E4:E38).
            formatting.Ranges.Add(XlCellRange.FromLTRB(4, firstDataRowIndex, 4, sheet.CurrentRowIndex - 2))
            ' Create the rule to apply a specific icon from the three traffic lights icon set to each cell in the "Market Share" column based on its value. 
            Dim rule5 As New XlCondFmtRuleIconSet()
            rule5.IconSetType = XlCondFmtIconSetType.TrafficLights3
            formatting.Rules.Add(rule5)
            ' Add the specified rule to the worksheet collection of conditional formatting rules.
            sheet.ConditionalFormattings.Add(formatting)
        End Sub
    End Class
End Namespace
