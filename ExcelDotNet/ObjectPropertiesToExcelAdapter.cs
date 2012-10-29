using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace EasyExcel
{

    /// <summary>
    /// An easy to use interface for converting properties of objects in lists into tables in Excel.
    /// </summary>
    /// <typeparam name="T">The object type</typeparam>
    public class ObjectPropertiesToExcelAdapter<T>
    {

        protected Excel.Worksheet wkSheet;
        protected List<T> listToOutput;

        public ObjectPropertiesToExcelAdapter(Excel.Worksheet wkSheet, List<T> listToOutput)
        {
            this.wkSheet = wkSheet;
            this.listToOutput = listToOutput;
           
        }


        public void Write(ExcelCellCoordinate loc)
        {
            var adapter = new DataTableToExcelAdapter(wkSheet, ConvertListOfObjectsToDataTable(listToOutput));
            adapter.Write(loc);
        }


        private DataTable ConvertListOfObjectsToDataTable(List<T> results)
        {

            Type type = results[0].GetType();
           
            IList<PropertyInfo> props = new List<PropertyInfo>(type.GetProperties());

            var table = new DataTable();

            foreach (PropertyInfo prop in props)
            {

                table.Columns.Add(prop.Name);
            }


            foreach (var result in results)
            {

                var newRow = table.NewRow();

                foreach (PropertyInfo prop in props)
                {
                    object propValue = prop.GetValue(result, null);

                    newRow[prop.Name] = propValue;

                }

                table.Rows.Add(newRow);
            }

            return table;
        }

    }
}
