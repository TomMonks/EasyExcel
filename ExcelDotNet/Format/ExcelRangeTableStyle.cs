using System.Drawing;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace EasyExcel.Format
{
    public class ExcelRangeTableStyle: AbstractDefaultExcelRangeStyle
    {
        protected IList<IExcelRangeStyle> styles;
        public bool FirstRowContainHeaders{get;set;}

        public ExcelRangeTableStyle(Excel.Worksheet sheetToFormat, ExcelCellCoordinate topLeft, ExcelCellCoordinate bottomRight)
            : base(sheetToFormat, topLeft, bottomRight)
		{
            
		}

		protected override void ApplyStyle(Excel.Range rng)
		{
            LoadStyles();

            foreach (var xlStyle in styles)
            {
                xlStyle.Execute();
            }
		}

        #region Style Creation

        private void LoadStyles()
        {
            styles = new List<IExcelRangeStyle>();
            styles.Add(new ExcelRangeSolidBorderStyle(this.sheetToFormat, this.topLeft, this.bottomRight));
            styles.Add(new ExcelRangeBackColourStyle(this.sheetToFormat, this.topLeft, this.bottomRight));

            if (this.FirstRowContainHeaders)
            {
                styles.Add(new ExcelRangeTableHeaderStyle(this.sheetToFormat, this.topLeft,
                        new ExcelCellCoordinate(this.topLeft.Row, this.bottomRight.Col)));
            }

        }

        #endregion

    }


}
