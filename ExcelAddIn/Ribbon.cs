using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn
{
    public partial class Ribbon
    {
        #region [ Consts ]
        private const String TAB_NAME = "NEW TAB";
        #endregion [ Consts ]


        #region [ Events ]
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonLaunchTest_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook currentDocument = ExcelServices.GetCurrentExcel();

            //Create a new sheet
            Worksheet newTab = InitializeNewTab(currentDocument);

            if (newTab != null)
            {
                List<string> listSheetsName = ExcelServices.GetAllWorksheetsName(currentDocument);

                if (listSheetsName?.Count > 0)
                {
                    //Display the name of sheets
                    for (int i = 0; i < listSheetsName.Count; i++)
                        newTab.Cells[1, i + 1].Value = listSheetsName[i];

                    //Search a text
                    String textToSearch = TAB_NAME;
                    int indexColumn = ExcelServices.GetColumnIndex(newTab, 1, textToSearch);
                    String letterColumn = ExcelServices.GetExcelColumnName(indexColumn);

                    //Add numbers
                    for(int i = 1; i<=10; i++)
                        newTab.Cells[i+1, indexColumn].Value = i;

                    //Add a row Total
                    int indexRowTotal = 12;
                    newTab.Cells[indexRowTotal, 1].Value = "TOTAL";
                    newTab.Cells[indexRowTotal, indexColumn].Value = $"=SUM({letterColumn}2: {letterColumn}11)";

                    //Add colors
                    StylePage(newTab, indexRowTotal, indexColumn);
                }
            }
        }
        #endregion [ Events ]


        #region [ Functions ]
        /// <summary>
        /// Initialize the creation of reviewTab inside the current Excel document
        /// </summary>
        /// <param name="excelDoc">Excel document</param>
        /// <returns>Returns the Worksheet reviewTab initialized</returns>
        /// <returns></returns>
        private static Worksheet InitializeNewTab(Workbook excelDoc)
        {
            Worksheet newTab = null;

            if (excelDoc != null)
            {
                //Check if there is already a sheet review
                ExcelServices.DeleteWorksheetsWithName(excelDoc, TAB_NAME);

                if (excelDoc?.Worksheets?.Count > 0)
                {
                    Worksheet currentLastWorksheet = excelDoc?.Worksheets.Cast<Worksheet>().LastOrDefault();

                    //Add the new sheet at the end
                    newTab = excelDoc.Worksheets.Add(After: currentLastWorksheet);
                    newTab.Name = TAB_NAME;
                }
            }

            return newTab;
        }

        private static void StylePage(Worksheet reviewTab, int rowTotalIndex, int columnNumbersIndex)
        {
            //Style the page
            Color COLOR_HEADER_BACKGROUND = Color.FromArgb(75, 172, 198);
            Color COLOR_HEADER_FOREGROUND = Color.White;
            Color COLOR_NUMBERS_COLUMN_BACKGROUND = Color.FromArgb(49, 134, 155);
            Color COLOR_TOTAL_BLOC_BACKGROUND = Color.FromArgb(75, 172, 198);

            String columnNumbersLetter = ExcelServices.GetExcelColumnName(columnNumbersIndex);

            //Column A
            Range columnA = reviewTab.Range["A:A"];
            columnA.EntireColumn.ColumnWidth = 20;
            columnA.Font.Color = Color.Black;
            columnA.Font.Bold = true;
            columnA.HorizontalAlignment = XlHAlign.xlHAlignLeft;

            //Column Numbers Header
            Range rangeColumnNumbersHeader = reviewTab.Range[$"{columnNumbersLetter}1:{columnNumbersLetter}1"];
            rangeColumnNumbersHeader.Interior.Color = COLOR_HEADER_BACKGROUND;
            rangeColumnNumbersHeader.Font.Color = COLOR_HEADER_FOREGROUND;
            rangeColumnNumbersHeader.Font.Bold = true;
            rangeColumnNumbersHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeColumnNumbersHeader.Borders.LineStyle = XlLineStyle.xlContinuous;

            //Column Numbers
            Range rangeColumnNumbers = reviewTab.Range[$"{columnNumbersLetter}2:{columnNumbersLetter}{rowTotalIndex - 1}"];
            rangeColumnNumbers.Interior.Color = COLOR_NUMBERS_COLUMN_BACKGROUND;
            rangeColumnNumbers.Font.Color = COLOR_HEADER_FOREGROUND;
            rangeColumnNumbers.Font.Bold = true;
            rangeColumnNumbers.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeColumnNumbers.Borders.LineStyle = XlLineStyle.xlContinuous;
            rangeColumnNumbers.Borders.Weight = XlBorderWeight.xlMedium;
            rangeColumnNumbers.Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            //Cell Total
            Range rangeTotal = reviewTab.Range[$"{columnNumbersLetter}{rowTotalIndex}:{columnNumbersLetter}{rowTotalIndex}"];
            rangeTotal.Interior.Color = COLOR_HEADER_BACKGROUND;
            rangeTotal.Font.Color = COLOR_HEADER_FOREGROUND;
            rangeTotal.Font.Bold = true;
            rangeTotal.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeTotal.Borders.LineStyle = XlLineStyle.xlContinuous;
        }
        #endregion [ Functions ]
    }
}
