using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn
{
    public static class ExcelServices
    {
        #region [ Methods ]
        public static Workbook GetCurrentExcel()
        {
            Application currentApp = Globals.ThisAddIn.Application;
            return currentApp.ActiveWorkbook;
        }

        public static List<String> GetAllWorksheetsName(Workbook document)
        {
            List<String> listWorksheetsName = null;

            if (document?.Worksheets?.Count > 0)
            {
                listWorksheetsName = new List<string>();

                foreach (Worksheet worksheet in document.Worksheets)
                {
                    if (!String.IsNullOrWhiteSpace(worksheet.Name))
                        listWorksheetsName.Add(worksheet.Name);
                }
            }

            return listWorksheetsName;
        }

        /// <summary>
        /// Get the selected worksheet
        /// </summary>
        /// <param name="document">Excel document</param>
        /// <param name="worksheetName">name of worksheet to find</param>
        /// <returns>The selected worksheet or null</returns>
        public static Worksheet GetSelectedWorksheet(Workbook document, String worksheetName)
        {
            Worksheet selectedWorksheet = null;

            if (document?.Worksheets?.Count > 0 && !String.IsNullOrWhiteSpace(worksheetName))
                selectedWorksheet = document.Worksheets.Cast<Worksheet>().FirstOrDefault(w => w != null && w.Name == worksheetName);

            return selectedWorksheet;
        }

        public static void DeleteWorksheetsWithName(Workbook document, String nameToDelete)
        {
            if (document?.Worksheets?.Count > 0 && !String.IsNullOrWhiteSpace(nameToDelete))
            {
                //The index of Worksheets starts to 1 (and not 0)
                for (int i = document.Worksheets.Count; i > 0; i--)
                {
                    //Compare the names and ignore the caps
                    if (document.Worksheets[i] is Worksheet worksheet && worksheet.Name?.ToLowerInvariant() == nameToDelete.ToLowerInvariant())
                        worksheet.Delete();
                }
            }
        }

        /// <summary>
        /// Search the column where is the text
        /// </summary>
        /// <param name="worksheet">Worksheet where realize the searching</param>
        /// <param name="rowIndex">Index of the row where search the text</param>
        /// <param name="textToSearch">Text to search</param>
        /// <returns></returns>
        public static int GetColumnIndex(Worksheet worksheet, int rowIndex, String textToSearch)
        {
            int columnFound = -1;

            if (worksheet != null && rowIndex > 0 && !String.IsNullOrWhiteSpace(textToSearch))
            {
                int currentColumnIndex = 1;
                while (columnFound == -1 && currentColumnIndex < 100)
                {
                    Range currentCell = worksheet.Cells[rowIndex, currentColumnIndex];

                    if (currentCell.Value2?.ToString() == textToSearch)
                        columnFound = currentColumnIndex;
                    else
                        currentColumnIndex++;
                }
            }

            return columnFound;
        }

        /// <summary>
        /// Search the row where is the text
        /// </summary>
        /// <param name="worksheet">Worksheet where realize the searching</param>
        /// <param name="columnIndex">Index of the column where search the text</param>
        /// <param name="textToSearch">Text to search</param>
        /// <returns></returns>
        public static int GetRowIndex(Worksheet worksheet, int columnIndex, String textToSearch)
        {
            int rowFound = -1;

            if (worksheet != null && columnIndex > 0 && !String.IsNullOrWhiteSpace(textToSearch))
            {
                int currentRowIndex = 1;
                while (rowFound == -1 && currentRowIndex < 100)
                {
                    Range currentCell = worksheet.Cells[currentRowIndex, columnIndex];

                    if (currentCell.Value2?.ToString() == textToSearch)
                        rowFound = currentRowIndex;
                    else
                        currentRowIndex++;
                }
            }

            return rowFound;
        }

        public static string GetExcelColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// Count the number of rows with the selected color
        /// </summary>
        /// <param name="worksheet">Worksheet to analyse</param>
        /// <param name="columnIndex">Index of the column to analyse</param>
        /// <param name="firstRowIndex">Index of the first row to count</param>
        /// <param name="lastRowIndex">Index of the last row to count</param>
        /// <param name="selectedColor">Selected color to search</param>
        /// <returns>The number of columns with the selected color</returns>
        public static int CountRowsWithColor(Worksheet worksheet, int columnIndex, int firstRowIndex, int lastRowIndex, Color selectedColor)
        {
            int count = 0;

            if (worksheet != null && columnIndex > 0 && firstRowIndex > 0 && lastRowIndex > 0 && selectedColor != null)
            {
                int oleColor = ColorTranslator.ToOle(selectedColor);

                for (int i = firstRowIndex; i <= lastRowIndex; i++)
                {
                    Range currentCell = worksheet.Cells[i, columnIndex];

                    if ((int)currentCell.Interior.Color == oleColor)
                        count++;
                }
            }

            return count;
        }
        #endregion [ Methods ]
    }
}
