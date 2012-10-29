using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace EasyExcel.Format
{
    public class ExcelRangeTableHeaderStyle : AbstractDefaultExcelRangeStyle 
    {
        private ExcelRangeBackColourStyle backStyle;

        public ExcelRangeTableHeaderStyle(Excel.Worksheet sheetToFormat, ExcelCellCoordinate topLeft, ExcelCellCoordinate bottomRight) 
            : base(sheetToFormat, topLeft, bottomRight)
		{
            this.backStyle = new ExcelRangeBackColourStyle(this.sheetToFormat, this.topLeft, this.bottomRight) 
                { BackColour = Color.LightBlue };
		}

		protected override void ApplyStyle(Excel.Range rng)
		{
            rng.Font.Bold = true;
            backStyle.Execute();
		}
    }

}
