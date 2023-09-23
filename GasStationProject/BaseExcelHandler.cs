using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace GasStationProject
{
    public class BaseExcelHandler
    {
        protected string path;
        protected int sheetNumber;

        protected Excel.Application application;
        protected Excel.Workbook workbook;
        protected Excel.Worksheet worksheet;

        public BaseExcelHandler(List<string> firstRowNames, string path, int sheetNumber, bool isVisible = false)
        {
            this.path = path;
            this.sheetNumber = sheetNumber;

            application = new Excel.Application();
            application.Visible = isVisible;

            if (!File.Exists(path))
            {
                workbook = application.Workbooks.Add();
                worksheet = workbook.Worksheets[sheetNumber];

                for (int index = 0; index < firstRowNames.Count; index++)
                {
                    worksheet.Cells[1, index + 1].Value = firstRowNames[index];
                }

                workbook.SaveAs(path);

                CloseAndQuit();
            }
        }

        protected void CloseAndQuit()
        {
            workbook.Close();
            application.Quit();
        }

        public void Clear()
        {
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(application);
        }
    }
}
