using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace GasStationProject
{
    public class SalesManagementHandler : BaseExcelHandler
    {
        // ProductName
        // ProductPrice
        // SalesQuantity
        // TotalPrice

        private ProductExcelHandler _productExcelHandler;

        public SalesManagementHandler(ProductExcelHandler productExcelHandler, List<string> firstRowNames, string path, int sheetNumber, bool isVisible = false) : base(firstRowNames, path, sheetNumber, isVisible)
        {
            _productExcelHandler = productExcelHandler;
            if (GetList().Count == 0)
            {
                CreateSalesRow();
            }
        }

        public bool SellProduct(string productName, int quantity)
        {
            List<ProductSales> productSales = GetList();

            foreach (ProductSales productSalesItem in productSales) { 
                if(productSalesItem.ProductName.ToLower() == productName.ToLower())
                {
                    productSalesItem.Quantity += quantity;
                    productSalesItem.TotalPrice = productSalesItem.Quantity * productSalesItem.Price;
                    UpdateSalesRow(productSales);
                    return true;
                }
            }
            return false;
        }

        public void UpdateSalesRow(List<ProductSales> productSales)
        {
            workbook = application.Workbooks.Open(path);
            worksheet = workbook.Worksheets[sheetNumber];

            int startRow = 2;

            foreach (ProductSales item in productSales)
            {
                worksheet.Cells[startRow, 1] = item.ProductName;
                worksheet.Cells[startRow, 2] = item.Price;
                worksheet.Cells[startRow, 3] = item.Quantity;
                worksheet.Cells[startRow, 4] = item.TotalPrice;
                startRow++;
            }

            workbook.Save();

            CloseAndQuit();
        }

        public void CreateSalesRow()
        {
            workbook = application.Workbooks.Open(path);
            worksheet = workbook.Worksheets[sheetNumber];

            int startRow = 2;

            foreach (Product product in _productExcelHandler.GetList())
            {
                worksheet.Cells[startRow, 1] = product.Name;
                worksheet.Cells[startRow, 2] = product.Price;
                worksheet.Cells[startRow, 3] = 0;
                worksheet.Cells[startRow, 4] = 0;
                startRow++;
            }

            workbook.Save();

            CloseAndQuit();
        }

        public List<ProductSales> GetList()
        {
            List<ProductSales> productSales = new List<ProductSales>();

            workbook = application.Workbooks.Open(path);
            worksheet = workbook.Worksheets[sheetNumber];

            Excel.Range range = worksheet.UsedRange;
            int lastRow = range.Rows.Count + 1;

            for (int i = 2; i < lastRow; i++)
            {
                productSales.Add(new ProductSales()
                {
                    ProductName = (range.Cells[i, 1] as Excel.Range).Text,
                    Price = decimal.Parse((range.Cells[i, 2] as Excel.Range).Text),
                    Quantity = int.Parse((range.Cells[i, 3] as Excel.Range).Text),
                    TotalPrice = decimal.Parse((range.Cells[i, 4] as Excel.Range).Text),
                });
            }

            CloseAndQuit();

            return productSales;
        }
    }
}
