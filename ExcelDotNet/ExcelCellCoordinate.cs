using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyExcel
{
    public struct ExcelCellCoordinate
    {
        private readonly int row;
        private readonly int col;

        public int Row { get { return this.row; } }
        public int Col { get { return this.col; } }

        public ExcelCellCoordinate(int row, int col)
        {
            this.row = row;
            this.col = col;
        }


        /// <summary>
        /// Converts the excel coordinate into a System.Drawing.Point object
        /// </summary>
        /// <returns>Point object representing coordinate</returns>
        public System.Drawing.Point AsPoint()
        {
            return new System.Drawing.Point(this.Col, this.Row);
        }

        /// <summary>
        /// Converts column to letter format.  Currently works up to ZZ.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

            int multiple = this.col / 26;
            int remainder = this.col % 26;

            string index = "";

            index += Convert.ToChar('A' + ((this.col - 1) % 26));

            if (multiple > 0 && remainder == 0)
            {
                try
                {
                    index = alpha[multiple - 2] + index;
                }
                catch (IndexOutOfRangeException)
                {
                    //not right.
                    index = index;
                }
            }
            else if (multiple > 0)
            {
                index = alpha[multiple - 1] + index;
            }
            

            //return String.Format("Multiple: {0}; Remainder: {1}, {2}", multiple, (this.col % 26), index);
            return string.Format("{0}{1}", index, this.row);
        }
    }
}
