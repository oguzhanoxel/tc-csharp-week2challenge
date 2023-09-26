using FileManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Services
{
    public class ProductService
    {
        private List<Product> _products;
        private ExcelFileManager<Product> _productExcelFileManager;

        public ProductService(string filePath, string sheetName)
        {
            _products = new List<Product>();
            _productExcelFileManager = new ExcelFileManager<Product>(filePath, sheetName);
        }

        public List<Product> GetProducts()
        {
            return _productExcelFileManager.ReadDataFromExcel();
        }

        public Product GetProductById(int id)
        {
            List<Product> products = _productExcelFileManager.ReadDataFromExcel();
            return products.FirstOrDefault(product =>  product.Id == id);
        }

        public Product GetProductByName(string name)
        {
            List<Product> products = _productExcelFileManager.ReadDataFromExcel();
            return products.FirstOrDefault(product => product.Name.ToLower() == name);
        }

        public void CreateProduct(Product product)
        {
            _products.Add(product);
            _productExcelFileManager.WriteDataToExcel(_products);
        }

        public void DeleteProduct(int id)
        {
            List<Product> products = _productExcelFileManager.ReadDataFromExcel();
            products.Remove(GetProductById(id));
            _productExcelFileManager.WriteDataToExcel(products);
        }
    }
}
