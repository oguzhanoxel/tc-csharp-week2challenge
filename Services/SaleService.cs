using FileManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Services
{
    public class SaleService
    {
        private List<Sale> _saleList;
        private ProductService _productService;
        private ExcelFileManager<Sale> _excelFileManager;

        public SaleService(ProductService productService, string filePath, string sheetName)
        {
            _saleList = new List<Sale>();
            _productService = productService;
            _excelFileManager = new ExcelFileManager<Sale>(filePath, sheetName);
        }

        public bool SellProduct(string productName, int quantity)
        {
            Product product = _productService.GetProductByName(productName);

            if (product == null)
            {
                return false;
            }

            _saleList = _excelFileManager.ReadDataFromExcel();
            _saleList.Add(new Sale() { ProductName = product.Name, Quantity = quantity, TotalPrice = product.Price * quantity, SaleDate = DateTime.Now });
            _excelFileManager.WriteDataToExcel(_saleList);
            return true;
        }

        public List<Sale> GetSales()
        {
            return _excelFileManager.ReadDataFromExcel();
        }
    }
}
