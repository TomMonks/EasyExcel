using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace EasyExcel.Format
{
    public class ExcelRangeBackColourStyle : AbstractDefaultExcelRangeStyle 
    {
        public Color BackColour { get; set; }

        public ExcelRangeBackColourStyle(Excel.Worksheet sheetToFormat, ExcelCellCoordinate topLeft, ExcelCellCoordinate bottomRight)
            : base(sheetToFormat, topLeft, bottomRight)
		{
            this.BackColour = Color.PeachPuff;
		}

		protected override void ApplyStyle(Excel.Range rng)
		{
            rng.Interior.Color = ColorTranslator.ToOle(this.BackColour);
		}

    }

}
