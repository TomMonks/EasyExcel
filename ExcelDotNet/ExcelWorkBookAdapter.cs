using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace EasyExcel
{
    public class ExcelWorkBookAdaptor
    {
        protected Application xl;
        protected Workbook wbk;
        protected const string XL_NULL_ERROR = "Please open or create a new workbook before attempting this operation";

        public ExcelWorkBookAdaptor()
        {
           
        }

        #region Properties

        public Worksheet this[int index]
        {
            get
            {
                return (Worksheet)wbk.Worksheets[index + 1];
            }
        }

        #endregion

        #region Opening Files

        public void NewBook()
        {
            xl = new Application();
            wbk = xl.Workbooks.Add();
        }

        public void Open(string fileName)
        {
            xl = new Application();
            wbk = xl.Workbooks.Open(fileName, Type.Missing, false);
        }

        public void Show()
        {
            if (null != xl)
            {
                xl.Visible = true;
            }
            else
            {
                throw new NullReferenceException(XL_NULL_ERROR);
            }
        }

        #endregion

        #region Closing Files

        public void CloseNoSave()
        {
            if (null != xl)
            {
                wbk.Close(false);
                xl.Quit();
                xl = null;
            }else
            {
                throw new NullReferenceException(XL_NULL_ERROR);
            }
        }

        public void SaveAndClose(string fileName)
        {
            if (null != xl)
            {
                xl.DisplayAlerts = false;
                wbk.SaveAs(fileName);
                wbk.Close(false);
                xl.DisplayAlerts = true;
                xl.Quit();
                xl = null;

            }
            else
            {
                throw new NullReferenceException(XL_NULL_ERROR);
            }
        }

        #endregion

    }
}
