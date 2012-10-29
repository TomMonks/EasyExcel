using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace EasyExcel.Format
{
	public class ExcelRangeSolidBorderStyle : AbstractDefaultExcelRangeStyle 
	{

		public ExcelRangeSolidBorderStyle(Excel.Worksheet sheetToFormat, ExcelCellCoordinate topLeft, 
            ExcelCellCoordinate bottomRight) : base(sheetToFormat, topLeft, bottomRight)
		{
			
		}

		protected override void ApplyStyle(Excel.Range rng)
		{
			rng.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
		}
	}
}
