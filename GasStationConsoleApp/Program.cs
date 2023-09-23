using GasStationProject;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GasStationConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            int chose;
            ProductExcelHandler productExcelHandler = InitProductHandler();
            SalesManagementHandler salesManagementHandler = InitSalesManagementHandler(productExcelHandler);

            while (true)
            {
                Console.Clear();
                Console.WriteLine("1. Start sale");
                Console.WriteLine("2. Market");
                Console.WriteLine("0. Quit");
                Console.Write("Enter a number: ");

                GetUserChoose(out chose);

                switch (chose)
                {
                    case 1:
                        do
                        {
                            Console.Clear();
                            foreach (Product product in productExcelHandler.GetList())
                            {
                                Console.WriteLine($"Name: {product.Name}, Price: {product.Price}");
                            }
                            Console.Write("Enter a Product Name: ");
                            string productName = Console.ReadLine();
                            Console.Write("Enter a Quantity: ");
                            GetQuantity(out int quantity);
                            
                            salesManagementHandler.SellProduct(productName, quantity);

                            Console.WriteLine("press enter for another sale or enter 'q' for quit.");
                        } while ((Console.ReadLine()=="q")?false:true);
                        break;
                    case 2:
                        Console.Clear();
                        bool isWorking = true;
                        while (isWorking)
                        {
                            Console.Clear();
                            Console.WriteLine("1. Show Products");
                            Console.WriteLine("2. Show Sales");
                            Console.WriteLine("0. Quit");
                            Console.Write("Enter a number: ");

                            GetUserChoose(out chose);

                            switch (chose)
                            {
                                case 1:
                                    Console.Clear();
                                    Console.WriteLine();
                                    foreach (Product product in productExcelHandler.GetList())
                                    {
                                        Console.WriteLine($"Id: {product.Id}, Name: {product.Name}, Price: {product.Price}");
                                    }

                                    Console.WriteLine("\npress a key to return");
                                    Console.ReadKey();
                                    break;
                                case 2:
                                    Console.Clear();
                                    foreach (ProductSales productSales in salesManagementHandler.GetList())
                                    {
                                        Console.WriteLine($"Name: {productSales.ProductName}, Price: {productSales.Price}, Quantity: {productSales.Quantity}, TotalPrice: {productSales.TotalPrice}");
                                    }
                                    Console.WriteLine("\npress a key to return");
                                    Console.ReadKey();
                                    break;
                                case 0:
                                    isWorking = false;
                                    break;
                            }
                        }
                        break;
                    case 0:
                        salesManagementHandler.Clear();
                        Environment.Exit(0);
                        break;
                }
            }
        }

        private static void GetUserChoose(out int chose)
        {
            if (!int.TryParse(Console.ReadLine(), out chose))
            {
                Console.WriteLine("Invalid value.");
            }
        }

        private static void GetQuantity(out int quantity)
        {
            if (!int.TryParse(Console.ReadLine(), out quantity))
            {
                Console.WriteLine("Invalid value.");
            }
        }

        private static SalesManagementHandler InitSalesManagementHandler(ProductExcelHandler productExcelHandler)
        {
            List<string> salesManagementFirstRows = new List<string>() { "ProductName", "ProductPrice", "SalesQuantity", "TotalPrice" };

            string path = "C:\\Users\\oguzh\\Desktop\\test\\salesManagement.xlsx";

            SalesManagementHandler salesManagementHandler = new SalesManagementHandler(productExcelHandler, salesManagementFirstRows, path, 1);

            return salesManagementHandler;
        }

        private static ProductExcelHandler InitProductHandler()
        {
            List<string> productFirstRows = new List<string>() { "Id", "Name", "Price" };

            Product product0 = new Product() { Name = "Su", Price = 10 };
            Product product1 = new Product() { Name = "Kahve", Price = 15 };
            Product product2 = new Product() { Name = "Çikolata", Price = 18 };
            Product product3 = new Product() { Name = "Sandviç", Price = 30 };

            string path = "C:\\Users\\oguzh\\Desktop\\test\\products.xlsx";

            ProductExcelHandler productExcelHandler = new ProductExcelHandler(productFirstRows, path, 1);

            if (productExcelHandler.GetList().Count == 0)
            {
                productExcelHandler.Create(product0);
                productExcelHandler.Create(product1);
                productExcelHandler.Create(product2);
                productExcelHandler.Create(product3);
            }

            return productExcelHandler;
        }
    }
}
