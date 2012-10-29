using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AnalysisIO;
using EasyExcel;
using EasyExcel.Format;

namespace TestEasyExcelForms
{
    public partial class Form1 : Form
    {
        AnalysisIO.EnhancedDataGrid grid = new EnhancedDataGrid();

        public Form1()
        {
            InitializeComponent();
        }

        public void ShowGrid()
        {
            grid.Location = new System.Drawing.Point(12, 12);
            grid.Name = "dataGridView1";
            grid.Size = new System.Drawing.Size(527, 346);
            grid.TabIndex = 0;

         
            this.Controls.Add(grid);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var adaptor = new ExcelWorkBookAdaptor();

            adaptor.NewBook();

            Point topLeft = new Point(1, 1);
            Point bottomRight = new Point(this.grid.ScenarioData.Columns.Count, this.grid.ScenarioData.Rows.Count + 1);

            var tableAdapter = new DataTableToExcelAdapter(adaptor[0], this.grid.ScenarioData);
            tableAdapter.Write(topLeft);

            var format = new ExcelRangeTableStyle(adaptor[0], topLeft, bottomRight) { FirstRowContainHeaders = true };
            format.Execute();
           
            adaptor.Show();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable tableOutput = null;
            var adapter = new ExcelWorkBookAdaptor();
            adapter.Open("C:/temp/Book1.xlsx");
            adapter.Show();
            var xlRangeAdapter = new ExcelRangeToDataTableAdapter(adapter[0]);
            Point topLeft = new Point(1, 1);
            Point bottomRight = new Point(2, 20);

            try
            {
                tableOutput = xlRangeAdapter.ReadTable("A1:B21");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                adapter.CloseNoSave();
            }


            if (null != tableOutput)
            {
                adapter = new ExcelWorkBookAdaptor();
                adapter.NewBook();
                adapter.Show();
                var tableAdapter = new DataTableToExcelAdapter(adapter[0], tableOutput);
                tableAdapter.Write(topLeft);

                //adapter.SaveAndClose("C:/temp/Book2.xlsx");
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            var coord = new ExcelCellCoordinate(5, Convert.ToInt32(this.txtCol.Text));
            Console.WriteLine(coord.ToString());

            
            
        }
    }
}
