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
    public class ProductExcelHandler : BaseExcelHandler
    {
        // Id
        // Name
        // Price
        public ProductExcelHandler(List<string> firstRowNames, string path, int sheetNumber, bool isVisible = false) : base(firstRowNames, path, sheetNumber, isVisible)
        {
            
        }

        public void Create(Product product)
        {
            workbook = application.Workbooks.Open(path);
            worksheet = workbook.Worksheets[sheetNumber];

            Excel.Range range = worksheet.UsedRange;
            int lastRow = range.Rows.Count + 1;

            worksheet.Cells[lastRow, 1] = lastRow;
            worksheet.Cells[lastRow, 2] = product.Name;
            worksheet.Cells[lastRow, 3] = product.Price;

            workbook.Save();

            CloseAndQuit();
        }

        public Product Get(int id)
        {
            Product result = null;
            foreach(Product product in GetList())
            {
                if (product.Id == id)
                {
                    result = product;
                    return result;
                }
            }
            return result;
        }
        
        public Product Get(string name)
        {
            Product result = null;
            foreach(Product product in GetList())
            {
                if (product.Name == name)
                {
                    result = product;
                    return result;
                }
            }
            return result;
        }

        public List<Product> GetList()
        {
            List<Product> products = new List<Product>();

            workbook = application.Workbooks.Open(path);
            worksheet = workbook.Worksheets[sheetNumber];

            Excel.Range range = worksheet.UsedRange;
            int lastRow = range.Rows.Count + 1;

            for (int i = 2; i < lastRow; i++)
            {
                products.Add(new Product()
                {
                    Id = int.Parse((range.Cells[i, 1] as Excel.Range).Text),
                    Name = (range.Cells[i, 2] as Excel.Range).Text,
                    Price = decimal.Parse((range.Cells[i, 3] as Excel.Range).Text),
                });
            }

            CloseAndQuit();

            return products;
        }
    }
}
