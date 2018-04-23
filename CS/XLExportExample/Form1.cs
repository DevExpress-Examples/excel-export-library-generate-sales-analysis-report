using DevExpress.Export.Xl;
using DevExpress.XtraExport.Csv;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace XLExportExample {
    public partial class Form1 : Form {
        List<SalesData> sales = SalesRepository.GetSalesData();
        XlCellFormatting headerRowFormatting;
        XlCellFormatting evenRowFormatting;
        XlCellFormatting oddRowFormatting;
        XlCellFormatting totalRowFormatting;

        public Form1() {
            InitializeComponent();
            InitializeFormatting();
        }

        void InitializeFormatting() {
            // Specify formatting settings for the even rows.
            evenRowFormatting = new XlCellFormatting();
            evenRowFormatting.Font = new XlFont();
            evenRowFormatting.Font.Name = "Century Gothic";
            evenRowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;
            evenRowFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.General, XlVerticalAlignment.Center);

            // Specify formatting settings for the odd rows.
            oddRowFormatting = new XlCellFormatting();
            oddRowFormatting.CopyFrom(evenRowFormatting);
            oddRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light1, -0.15));

            // Specify formatting settings for the header row.
            headerRowFormatting = new XlCellFormatting();
            headerRowFormatting.CopyFrom(evenRowFormatting);
            headerRowFormatting.Font.Bold = true;
            headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
            headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));
            headerRowFormatting.Border = new XlBorder();
            headerRowFormatting.Border.TopColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0);
            headerRowFormatting.Border.TopLineStyle = XlBorderLineStyle.Medium;
            headerRowFormatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0);
            headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Medium;

            // Specify formatting settings for the total row.
            totalRowFormatting = new XlCellFormatting();
            totalRowFormatting.CopyFrom(evenRowFormatting);
            totalRowFormatting.Font.Bold = true;
        }

        // Export the document to XLSX format.
        void btnExportToXLSX_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Xlsx))
                ShowFile(fileName);
        }

        // Export the document to XLS format.
        void btnExportToXLS_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Xls))
                ShowFile(fileName);
        }

        // Export the document to CSV format.
        void btnExportToCSV_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("CSV (Comma delimited files)(*.csv)|*.csv", "Document.csv");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Csv))
                ShowFile(fileName);
        }

        string GetSaveFileName(string filter, string defaulName) {
            saveFileDialog1.Filter = filter;
            saveFileDialog1.FileName = defaulName;
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                return null;
            return saveFileDialog1.FileName;
        }

        void ShowFile(string fileName) {
            if (!File.Exists(fileName))
                return;
            DialogResult dResult = MessageBox.Show(String.Format("Do you want to open the resulting file?", fileName),
                this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dResult == DialogResult.Yes)
                Process.Start(fileName);
        }

        bool ExportToFile(string fileName, XlDocumentFormat documentFormat) {
            try {
                using (FileStream stream = new FileStream(fileName, FileMode.Create)) {
                    // Create an exporter instance. 
                    IXlExporter exporter = XlExport.CreateExporter(documentFormat);
                    // Create a new document and begin to write it to the specified stream.
                    using (IXlDocument document = exporter.CreateDocument(stream)) {
                        // Generate the document content.
                        GenerateDocument(document);
                    }
                }
                return true;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        void GenerateDocument(IXlDocument document) {
            // Specify the document culture.
            document.Options.Culture = CultureInfo.CurrentCulture;

            // Specify options for exporting the document in CSV format.
            CsvDataAwareExporterOptions csvOptions = document.Options as CsvDataAwareExporterOptions;
            if (csvOptions != null) {
                csvOptions.WritePreamble = true;
                csvOptions.UseCellNumberFormat = false;
                csvOptions.NewlineAfterLastRow = true;
            }

            // Add a new worksheet to the document.
            using (IXlSheet sheet = document.CreateSheet()) {
                // Specify the worksheet name.
                sheet.Name = "Annual Sales";

                // Specify page settings.
                sheet.PageSetup = new XlPageSetup();
                // Scale the print area to fit to one page wide.
                sheet.PageSetup.FitToPage = true;
                sheet.PageSetup.FitToWidth = 1;
                sheet.PageSetup.FitToHeight = 0;

                // Generate worksheet columns.
                GenerateColumns(sheet);

                // Add the title to the documents exported to the XLSX and XLS formats.
                if (document.Options.DocumentFormat != XlDocumentFormat.Csv)
                    GenerateTitle(sheet);

                // Create the header row.
                GenerateHeaderRow(sheet);

                int firstDataRowIndex = sheet.CurrentRowIndex;

                // Create the data rows.
                for (int i = 0; i < sales.Count; i++)
                    GenerateDataRow(sheet, sales[i], (i + 1) == sales.Count);

                // Create the total row.
                GenerateTotalRow(sheet, firstDataRowIndex);

                // Specify the data range to be printed.
                sheet.PrintArea = sheet.DataRange;

                // Create conditional formatting rules to be applied to worksheet data.  
                GenerateConditionalFormatting(sheet, firstDataRowIndex);
            }
        }

        void GenerateColumns(IXlSheet sheet) {
            XlNumberFormat numberFormat =@"#,##0,,""M""";

            // Create the "State" column and set its width.
            using (IXlColumn column = sheet.CreateColumn())
                column.WidthInPixels = 140;

            // Create the "Sales" column, adjust its width and set the specific number format for its cells.
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 140;
                column.ApplyFormatting(numberFormat);
            }

            // Create the "Sales vs Target" column, adjust its width and format its cells as percentage values.
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 120;
                column.ApplyFormatting(XlNumberFormat.Percentage2);
            }

            // Create the "Profit" column, adjust its width and set the specific number format for its cells.
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 140;
                column.ApplyFormatting(numberFormat);
            }

            // Create the "Market Share" column, adjust its width and format its cells as percentage values.
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 120;
                column.ApplyFormatting(XlNumberFormat.Percentage);
            }
        }

        void GenerateTitle(IXlSheet sheet) {
            // Specify formatting settings for the document title.
            XlCellFormatting formatting = new XlCellFormatting();
            formatting.Font = new XlFont();
            formatting.Font.Name = "Calibri Light";
            formatting.Font.SchemeStyle = XlFontSchemeStyles.None;
            formatting.Font.Size = 24;
            formatting.Font.Color = XlColor.FromTheme(XlThemeColor.Dark1, 0.35);
            formatting.Border = new XlBorder();
            formatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.35);
            formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium;

            // Add the document title.
            using (IXlRow row = sheet.CreateRow()) {
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = "SALES 2014";
                    cell.Formatting = formatting;
                }
                // Create four empty cells with the title formatting applied.
                for (int i = 0; i < 4; i++) {
                    using (IXlCell cell = row.CreateCell())
                        cell.Formatting = formatting;
                }
            }

            // Skip one row before starting to generate data rows.
            sheet.SkipRows(1);

        }

        void GenerateHeaderRow(IXlSheet sheet) {
            string[] columnNames = new string[] { "State", "Sales", "Sales vs Target", "Profit", "Market Share" };
            // Create the header row and set its height.
            using (IXlRow row = sheet.CreateRow()) {
                row.HeightInPixels = 25;
                // Create required cells in the header row, assign values from the columnNames array to them and apply specific formatting settings. 
                row.BulkCells(columnNames, headerRowFormatting);
            }
        }

        void GenerateDataRow(IXlSheet sheet, SalesData data, bool isLastRow) {
            // Create the data row to display sales information for the specific state.
            using (IXlRow row = sheet.CreateRow()) {
                row.HeightInPixels = 25;

                // Specify formatting settings to be applied to the data rows to shade alternate rows. 
                XlCellFormatting formatting = new XlCellFormatting();
                formatting.CopyFrom((row.RowIndex % 2 == 0) ? evenRowFormatting : oddRowFormatting);
                // Set the bottom border for the last data row.
                if (isLastRow) {
                    formatting.Border = new XlBorder();
                    formatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0);
                    formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium;
                }

                // Create the cell containing the state name. 
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = data.State;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell containing sales data.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = data.ActualSales;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell that displays the difference between the actual and target sales.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = data.ActualSales / data.TargetSales - 1;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell containing the state profit. 
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = data.Profit;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell containing the percentage of a total market.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = data.MarketShare;
                    cell.ApplyFormatting(formatting);
                }
            }
        }

        void GenerateTotalRow(IXlSheet sheet, int firstDataRowIndex) {
            // Create the total row and set its height.
            using (IXlRow row = sheet.CreateRow()) {
                row.HeightInPixels = 25;

                // Create the first cell in the row and apply specific formatting settings to this cell.
                using (IXlCell cell = row.CreateCell())
                    cell.ApplyFormatting(totalRowFormatting);

                // Create the second cell in the total row and assign the SUBTOTAL function to it to calculate the average of the subtotal of the cells located in the "Sales" column.
                using (IXlCell cell = row.CreateCell()) {
                    cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(1, firstDataRowIndex, 1, row.RowIndex - 1), XlSummary.Average, false));
                    cell.ApplyFormatting(totalRowFormatting);
                    cell.ApplyFormatting((XlNumberFormat)@"""Avg=""#,##0,,""M""");
                }

                // Create the third cell in the row and apply specific formatting settings to this cell.
                using (IXlCell cell = row.CreateCell())
                    cell.ApplyFormatting(totalRowFormatting);

                // Create the fourth cell in the total row and assign the SUBTOTAL function to it to calculate the sum of the subtotal of the cells located in the "Profit" column.
                using (IXlCell cell = row.CreateCell()) {
                    cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(3, firstDataRowIndex, 3, row.RowIndex - 1), XlSummary.Sum, false));
                    cell.ApplyFormatting(totalRowFormatting);
                    cell.ApplyFormatting((XlNumberFormat)@"""Sum=""#,##0,,""M""");
                }
            }
        }

        void GenerateConditionalFormatting(IXlSheet sheet, int firstDataRowIndex) {
            // Create an instance of the XlConditionalFormatting class to define a new rule.
            XlConditionalFormatting formatting = new XlConditionalFormatting();
            // Specify the cell range to which the conditional formatting rule should be applied (B4:B38).
            formatting.Ranges.Add(XlCellRange.FromLTRB(1, firstDataRowIndex, 1, sheet.CurrentRowIndex - 2));
            // Create the rule to compare values in the "Sales" column using data bars. 
            XlCondFmtRuleDataBar rule1 = new XlCondFmtRuleDataBar();
            // Specify the color of data bars. 
            rule1.FillColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.4);
            // Set the solid fill type.
            rule1.GradientFill = false;
            formatting.Rules.Add(rule1);
            // Add the specified rule to the worksheet collection of conditional formatting rules.
            sheet.ConditionalFormattings.Add(formatting);

            // Create an instance of the XlConditionalFormatting class to define new rules.
            formatting = new XlConditionalFormatting();
            // Specify the cell range to which the conditional formatting rules should be applied (C4:C38).
            formatting.Ranges.Add(XlCellRange.FromLTRB(2, firstDataRowIndex, 2, sheet.CurrentRowIndex - 2));
            // Create the rule to identify negative values in the "Sales vs Target" column.
            XlCondFmtRuleCellIs rule2 = new XlCondFmtRuleCellIs();
            // Specify the relational operator to be used in the conditional formatting rule.
            rule2.Operator = XlCondFmtOperator.LessThan;
            // Set the threshold value.
            rule2.Value = 0;
            // Specify formatting options to be applied to cells if the condition is true.
            // Set the font color to dark red.
            rule2.Formatting = new XlFont() { Color = Color.DarkRed };
            formatting.Rules.Add(rule2);
            // Create the rule to identify top five values in the "Sales vs Target" column.
            XlCondFmtRuleTop10 rule3 = new XlCondFmtRuleTop10();
            rule3.Rank = 5;
            // Specify formatting options to be applied to cells if the condition is true.
            // Set the font color to dark green.
            rule3.Formatting = new XlFont() { Color = Color.DarkGreen };
            formatting.Rules.Add(rule3);
            // Add the specified rules to the worksheet collection of conditional formatting rules.
            sheet.ConditionalFormattings.Add(formatting);

            // Create an instance of the XlConditionalFormatting class to define a new rule.
            formatting = new XlConditionalFormatting();
            // Specify the cell range to which the conditional formatting rules should be applied (D4:D38).
            formatting.Ranges.Add(XlCellRange.FromLTRB(3, firstDataRowIndex, 3, sheet.CurrentRowIndex - 2));
            // Create the rule to compare values in the "Profit" column using data bars. 
            XlCondFmtRuleDataBar rule4 = new XlCondFmtRuleDataBar();
            // Specify the color of data bars. 
            rule4.FillColor = Color.FromArgb(99, 195, 132);
            // Specify the positive bar border color.
            rule4.BorderColor = Color.FromArgb(99, 195, 132);
            // Specify the negative bar fill color.
            rule4.NegativeFillColor = Color.FromArgb(255, 85, 90);
            // Specify the negative bar border color.
            rule4.NegativeBorderColor = Color.FromArgb(255, 85, 90);
            // Specify the solid fill type.
            rule4.GradientFill = false;
            formatting.Rules.Add(rule4);
            // Add the specified rule to the worksheet collection of conditional formatting rules.
            sheet.ConditionalFormattings.Add(formatting);

            // Create an instance of the XlConditionalFormatting class to define a new rule.
            formatting = new XlConditionalFormatting();
            // Specify the cell range to which the conditional formatting rules should be applied (E4:E38).
            formatting.Ranges.Add(XlCellRange.FromLTRB(4, firstDataRowIndex, 4, sheet.CurrentRowIndex - 2));
            // Create the rule to apply a specific icon from the three traffic lights icon set to each cell in the "Market Share" column based on its value. 
            XlCondFmtRuleIconSet rule5 = new XlCondFmtRuleIconSet();
            rule5.IconSetType = XlCondFmtIconSetType.TrafficLights3;
            formatting.Rules.Add(rule5);
            // Add the specified rule to the worksheet collection of conditional formatting rules.
            sheet.ConditionalFormattings.Add(formatting);
        }
    }
}
