/*
 * Created by SharpDevelop.
 * User: tm300
 * Date: 17/07/2012
 * Time: 14:22
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Drawing;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
	
namespace EasyExcel
{
	/// <summary>
	/// An easy to use interface for converting lists of values into tables in Excel.
	/// </summary>
	public class ListToExcelTableAdaptor<T>
	{
		protected Worksheet wkSheet;
		protected List<T> listToOutput;
		
		public ListToExcelTableAdaptor(Worksheet wkSheet, List<T> listToOutput)
		{
			this.wkSheet = wkSheet;
			this.listToOutput = listToOutput;
		}
		
		public void Write(ExcelCellCoordinate loc, int columns)
		{

            var currentRow = loc.Row;
            var currentCol = loc.Col;

			try {

				foreach (var val in this.listToOutput) {
					wkSheet.Cells[currentRow, currentCol++] = val;

					if (currentCol > (loc.Col + columns - 1)) {
						currentCol = loc.Col;
						currentRow++;
					}
				}

			} catch (System.NullReferenceException e) {
				Console.WriteLine(e.ToString());
				throw e;
			}


		}
	}
}
