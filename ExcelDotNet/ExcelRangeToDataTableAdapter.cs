using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Drawing;

using Excel = Microsoft.Office.Interop.Excel;

namespace EasyExcel
{
    /// <summary>
    /// Facade for simplifying reading a Excel range into a DataTable.
    /// </summary>
    public class ExcelRangeToDataTableAdapter
    {
        protected Excel.Worksheet inputSheet;

        public ExcelRangeToDataTableAdapter(Excel.Worksheet inputSheet) 
        {
            this.inputSheet = inputSheet;
        }


        /// <summary>
        /// Reads in an Excel Range and Converts to a DataTabe
        /// </summary>
        /// <param name="rangeName">e.g "A1:A20" or "ResultsRange"</param>
        /// <returns>DataTable containing the data from the Excel Range</returns>
        public DataTable ReadTable(string rangeName, bool hasHeaders = true)
        {
            
            var tableToOutput = new DataTable();
            int currentCol = 0;
            
            Excel.Range tableRange = inputSheet.get_Range(rangeName, Type.Missing);

            
            
            var currentRow = tableToOutput.NewRow();


            //loops by row and column.
            foreach (Excel.Range cell in tableRange)
            {

                if (hasHeaders && 1 == cell.Row)
                {
                    ReadHeader(tableToOutput, cell);
                }
                else
                {
                    currentCol = ReadCell(tableToOutput, currentCol, currentRow, cell);

                    AddRowToTableWhenFinished(tableToOutput, ref currentCol, tableRange, ref currentRow);
                }
            }
            
            return tableToOutput;
        }

        private static int ReadCell(DataTable tableToOutput, int currentCol, DataRow currentRow, Excel.Range cell)
        {
            try
            {
                currentRow[currentCol++] = cell.Value2;
            }
            catch (IndexOutOfRangeException)
            {
                tableToOutput.Columns.Add();
                currentCol--;
                currentRow[currentCol++] = cell.Value2;
            }
            return currentCol;
        }

        private void ReadHeader(DataTable tableToOutput, Excel.Range cell)
        {
            try
            {
                tableToOutput.Columns[cell.Column - 1].ColumnName = cell.Value2.ToString();
            }
            catch (IndexOutOfRangeException)
            {
                tableToOutput.Columns.Add();
                tableToOutput.Columns[cell.Column - 1].ColumnName = cell.Value2.ToString();
            }
        }

        private void AddRowToTableWhenFinished(DataTable tableToOutput, ref int currentCol, Excel.Range tableRange, ref DataRow currentRow)
        {
            if (currentCol > tableRange.Columns.Count - 1)
            {
                currentCol = 0;
                tableToOutput.Rows.Add(currentRow);
                currentRow = tableToOutput.NewRow();
            }
        }

    

    }
}
