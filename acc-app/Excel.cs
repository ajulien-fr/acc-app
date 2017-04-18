using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace acc_app
{
    class Excel
    {
        private static class Constants
        {
            public static string Password { get { return "acc"; } }
            public static string Currency { get { return "# ##0,00 €"; } }
            public static string Date { get { return "jj/mm/aaaa"; } }
            public static string Text { get { return "@"; } }
        }

        private String fileName;
        private InteropExcel.Application excelApp;
        private InteropExcel.Workbook excelWb;

        private bool isOpen;
        public bool IsOpen { get => isOpen; set => isOpen = value; }

        public Excel(String fileName)
        {
            isOpen = false;
            this.fileName = fileName;
            this.excelApp = new InteropExcel.Application();
        }

        ~Excel()
        {
            if (this.isOpen == true)
            {
                this.excelWb.Close(true); // true dans Close() pour sauvegarder...
                isOpen = false;
            }
        }

        public void Open()
        {
            try
            {
                this.excelWb = this.excelApp.Workbooks.Open(
                    this.fileName,
                    Missing.Value,
                    false,
                    Missing.Value,
                    Missing.Value,
                    Constants.Password,
                    true,
                    Missing.Value,
                    Missing.Value,
                    true,
                    false,
                    Missing.Value,
                    false,
                    true,
                    Missing.Value);
            }
            catch
            {
                throw;
            }
            finally
            {
                isOpen = true;
            }
        }

        public void Create()
        {
            try
            {
                this.excelWb = this.excelApp.Workbooks.Add();

                this.excelWb.SaveAs(
                    this.fileName,
                    InteropExcel.XlFileFormat.xlOpenXMLWorkbook,
                    Missing.Value,
                    Constants.Password, // on protège en écriture avec un password...
                    false,
                    false,
                    InteropExcel.XlSaveAsAccessMode.xlNoChange,
                    InteropExcel.XlSaveConflictResolution.xlUserResolution,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value);

                CreateSheets();
            }
            catch
            {
                throw; // on fait suivre à l'appelant...
            }
            finally
            {
                isOpen = true;
            }
        }

        private void CreateSheets()
        {
            String[] names = { "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre" };

            try
            {
                InteropExcel.Worksheet sheet = this.excelWb.ActiveSheet;
                sheet.Name = names[0];
                CreateTables(sheet);

                for (int i = 1; i < names.Length; i++)
                {
                    sheet = this.excelWb.Worksheets.Add(Missing.Value, this.excelWb.Worksheets[this.excelWb.Worksheets.Count]);
                    sheet.Name = names[i];
                    CreateTables(sheet);
                }
            }
            catch
            {
                throw; // on fait suivre à l'appelant...
            }
        }

        private void CreateTables(InteropExcel.Worksheet sheet)
        {
            try
            {
                CreateTableRecettes(sheet);
                CreateTableDepenses(sheet);
                sheet.Columns.AutoFit();
                sheet.Rows.AutoFit();
            }
            catch
            {
                throw; // on fait suivre à l'appelant...
            }
        }

        private void CreateTableRecettes(InteropExcel.Worksheet sheet)
        {
            InteropExcel.Range range;
            String[] names;
            InteropExcel.ListObject table;
            List<string> list;

            try
            {
                // RECETTES TABLE TITLE
                range = sheet.Range["A1", "E2"];
                range.Cells.Merge();
                range.Value = "RECETTES";
                range.Style = "Accent6";
                range.Cells.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;
                range.Cells.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;

                // TABLE RECETTES HEADER
                names = new String[] { "DATES", "MODES", "LIBELLES", "PROVENANCES", "MONTANTS" };
                range = sheet.Range["A3", "E3"];
                range.Value = names;
                table = range.Worksheet.ListObjects.Add(InteropExcel.XlListObjectSourceType.xlSrcRange, range, Missing.Value, InteropExcel.XlYesNoGuess.xlYes, Missing.Value);
                table.Name = String.Format("TableRecettes{0}", sheet.Name);
                table.HeaderRowRange.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;
                table.HeaderRowRange.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;
                table.ShowTotals = true;
                table.ShowHeaders = true;
                table.TableStyle = "TableStyleLight14";
                InteropExcel.ListRow row = table.ListRows.Add();
                row.Range.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;
                row.Range.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;

                // PROPERTIES OF TABLE COLUMN
                table.ListColumns["DATES"].DataBodyRange.NumberFormatLocal = Constants.Date;

                list = new List<string>
                {
                    "ADHESION",
                    "DON",
                    "ADOPTION",
                    "VENTE",
                    "MANIFESTATION",
                    "MAIRIE",
                    "FONDATION",
                    "NOURRITURE",
                };
                table.ListColumns["PROVENANCES"].DataBodyRange.Validation.Add(InteropExcel.XlDVType.xlValidateList, InteropExcel.XlDVAlertStyle.xlValidAlertInformation, InteropExcel.XlFormatConditionOperator.xlBetween, string.Join(";", list.ToArray()), Missing.Value);
                table.ListColumns["PROVENANCES"].DataBodyRange.Validation.IgnoreBlank = true;
                table.ListColumns["PROVENANCES"].DataBodyRange.Validation.InCellDropdown = true;

                list = new List<string>
                {
                    "CHEQUE",
                    "ESPECE",
                    "VIREMENT",
                };
                table.ListColumns["MODES"].DataBodyRange.Validation.Add(InteropExcel.XlDVType.xlValidateList, InteropExcel.XlDVAlertStyle.xlValidAlertInformation, InteropExcel.XlFormatConditionOperator.xlBetween, string.Join(";", list.ToArray()), Missing.Value);
                table.ListColumns["MODES"].DataBodyRange.Validation.IgnoreBlank = true;
                table.ListColumns["MODES"].DataBodyRange.Validation.InCellDropdown = true;

                table.ListColumns["LIBELLES"].DataBodyRange.NumberFormatLocal = Constants.Text;

                table.ListColumns["MONTANTS"].DataBodyRange.NumberFormatLocal = Constants.Currency;
            }
            catch
            {
                throw; // on fait suivre à l'appelant...
            }
        }

        private void CreateTableDepenses(InteropExcel.Worksheet sheet)
        {
            InteropExcel.Range range;
            String[] names;
            InteropExcel.ListObject table;
            List<string> list;

            try
            {
                // DEPENSES TABLE TITLE
                range = sheet.Range["G1", "K2"];
                range.Cells.Merge();
                range.Value = "DEPENSES";
                range.Style = "Accent5";
                range.Cells.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;
                range.Cells.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;

                // TABLE RECETTES HEADER
                names = new String[] { "DATES", "MODES", "LIBELLES", "DESTINATIONS", "MONTANTS" };
                range = sheet.Range["G3", "K3"];
                range.Value = names;
                table = range.Worksheet.ListObjects.Add(InteropExcel.XlListObjectSourceType.xlSrcRange, range, Missing.Value, InteropExcel.XlYesNoGuess.xlYes, Missing.Value);
                table.Name = String.Format("TableDepenses{0}", sheet.Name);
                table.HeaderRowRange.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;
                table.HeaderRowRange.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;
                table.ShowTotals = true;
                table.ShowHeaders = true;
                table.TableStyle = "TableStyleLight13";
                InteropExcel.ListRow row = table.ListRows.Add();
                row.Range.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;
                row.Range.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;

                // PROPERTIES OF TABLE COLUMN
                table.ListColumns["DATES"].DataBodyRange.NumberFormatLocal = Constants.Date;

                list = new List<string>
                {
                    "VETO",
                    "FOURNITURE",
                    "ACHAT",
                    "MANIFESTATION",
                    "ASSURANCE",
                    "NOURRITURE",
                };
                table.ListColumns["DESTINATIONS"].DataBodyRange.Validation.Add(InteropExcel.XlDVType.xlValidateList, InteropExcel.XlDVAlertStyle.xlValidAlertInformation, InteropExcel.XlFormatConditionOperator.xlBetween, string.Join(";", list.ToArray()), Missing.Value);
                table.ListColumns["DESTINATIONS"].DataBodyRange.Validation.IgnoreBlank = true;
                table.ListColumns["DESTINATIONS"].DataBodyRange.Validation.InCellDropdown = true;

                list = new List<string>
                {
                    "CHEQUE",
                    "ESPECE",
                    "VIREMENT",
                    "PRELEVEMENT",
                };
                table.ListColumns["MODES"].DataBodyRange.Validation.Add(InteropExcel.XlDVType.xlValidateList, InteropExcel.XlDVAlertStyle.xlValidAlertInformation, InteropExcel.XlFormatConditionOperator.xlBetween, string.Join(";", list.ToArray()), Missing.Value);
                table.ListColumns["MODES"].DataBodyRange.Validation.IgnoreBlank = true;
                table.ListColumns["MODES"].DataBodyRange.Validation.InCellDropdown = true;

                table.ListColumns["LIBELLES"].DataBodyRange.NumberFormatLocal = Constants.Text;

                table.ListColumns["MONTANTS"].DataBodyRange.NumberFormatLocal = Constants.Currency;
            }
            catch
            {
                throw; // on fait suivre à l'appelant...
            }
        }

        public void AddRecette(List<Object> values)
        {
            try
            {
                // on get le tableau des recettes
                InteropExcel.Worksheet sheet = this.excelWb.Worksheets[DateTime.Now.Month];
                InteropExcel.ListObject table = sheet.ListObjects[String.Format("TableRecettes{0}", sheet.Name)];
                InteropExcel.ListRow row = table.ListRows.Add();
                InteropExcel.Range cells = row.Range.Cells;

                cells.Value2 = values.ToArray(); // Value2 for set value to correct format
            }
            catch
            {
                throw;
            }
        }

        public void AddDepense(List<Object> values)
        {
            try
            {
                // on get le tableau des recettes
                InteropExcel.Worksheet sheet = this.excelWb.Worksheets[DateTime.Now.Month];
                InteropExcel.ListObject table = sheet.ListObjects[String.Format("TableDepenses{0}", sheet.Name)];
                InteropExcel.ListRow row = table.ListRows.Add();
                InteropExcel.Range cells = row.Range.Cells;

                cells.Value2 = values.ToArray(); // Value2 for set value to correct format
            }
            catch
            {
                throw;
            }
        }
    }
}
