/*
 * Created by SharpDevelop.
 * User: tm300
 * Date: 17/07/2012
 * Time: 10:52
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;

using System.Collections.Generic;
using EasyExcel;
using EasyExcel.Format;
using Microsoft.Office.Interop.Excel;
using System.Drawing;


namespace TestEasyExcel
{
	class Program
	{
		public static void Main(string[] args)
		{
			Console.WriteLine("Hello World!");
			
			// TODO: Implement Functionality Here

            
			//adaptor.OpenExisting(@"C:\Test Excel\Book1.xlsx");
			
			var data = new List<double>();
			data.Add(10);
			data.Add(10);
			data.Add(10);
			data.Add(10);
			data.Add(10);
			
			var data2 = new List<string>();
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");
			data2.Add("hello");


            var adaptor = new ExcelWorkBookAdaptor();
     
            adaptor.NewBook();

			var listAdapter = new ListToExcelTableAdaptor<string>(adaptor[0], data2);
			listAdapter.Write(new ExcelCellCoordinate(10, 10), 2);
			
            var format = new ExcelRangeTableStyle(adaptor[0], new ExcelCellCoordinate(10, 10), new ExcelCellCoordinate(15, 15)) { FirstRowContainHeaders = true };
			format.Execute();

            adaptor.Show();

			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
			
			adaptor.SaveAndClose(@"C:\Test Excel\Book2.xlsx");
		}
	}
}