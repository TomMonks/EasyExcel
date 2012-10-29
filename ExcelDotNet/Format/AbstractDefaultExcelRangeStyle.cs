using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace EasyExcel.Format
{
    public abstract class AbstractDefaultExcelRangeStyle : IExcelRangeStyle 
    {
        protected Excel.Worksheet sheetToFormat;


        protected ExcelCellCoordinate topLeft;
        protected ExcelCellCoordinate bottomRight;

        public AbstractDefaultExcelRangeStyle(Excel.Worksheet sheetToFormat, ExcelCellCoordinate topLeft, ExcelCellCoordinate bottomRight)
        {
            this.sheetToFormat = sheetToFormat;
            this.topLeft = topLeft;
            this.bottomRight = bottomRight;
        }

        public void Execute()
        {
            var selectedRange = SelectRange();
            ApplyStyle(selectedRange);
        }

        protected virtual Excel.Range SelectRange()
        {
            return this.sheetToFormat.Range[this.sheetToFormat.Cells[this.topLeft.Row, this.topLeft.Col], 
                this.sheetToFormat.Cells[this.bottomRight.Row, this.bottomRight.Col]];
        }

        protected abstract void ApplyStyle(Excel.Range rng);
        
    }
}
