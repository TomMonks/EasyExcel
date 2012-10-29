using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace EasyExcel
{
    public class DataTableToExcelAdapter
    {
        protected Excel.Worksheet wkSheet;
        protected DataTable tableToOutput;

        public DataTableToExcelAdapter(Excel.Worksheet wkSheet, DataTable tableToOutput) 
        {
            this.wkSheet = wkSheet;
            this.tableToOutput = tableToOutput;
        }

        

        #region Output

        public void Write(Point topLeft)
        {
            var currentLocation = topLeft;

            currentLocation = WriteHeaders(currentLocation);

            currentLocation.Y++;
            currentLocation.X = topLeft.X;

            WriteData(ref topLeft, ref currentLocation);
        }

        private Point WriteHeaders(Point currentLocation)
        {
            foreach (DataColumn column in this.tableToOutput.Columns)
            {
                this.wkSheet.Cells[currentLocation.Y, currentLocation.X++] = column.ColumnName;
            }
            return currentLocation;
        }

        private void WriteData(ref Point topLeft, ref Point currentLocation)
        {
            foreach (DataRow row in this.tableToOutput.Rows)
            {
                for (int col = 0; col < this.tableToOutput.Columns.Count; col++)
                {
                    this.wkSheet.Cells[currentLocation.Y, currentLocation.X++] = row[col];

                }

                currentLocation.Y++;
                currentLocation.X = topLeft.X;
            }
        }

        #endregion


    }
}
