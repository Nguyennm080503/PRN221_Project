using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Formats.Asn1;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography.X509Certificates;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Shapes;
using System.Xml;
using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;

namespace WPF_Project
{
    public partial class ControlVoice : Window
    {
        private List<Product> mProduct = null;
        public List<Product> Products
        {
            get { return mProduct; }
            set { mProduct = value; OnPropertyChanged(); }
        }
        public ControlVoice()
        {
            Products = new List<Product>();
            InitializeComponent();
            this.DataContext = this;
            
            //Products.Add(new Product() { ProductId = 1, ProductName = "Apple", Price = 10.5m, Quantity = 100, Status = "Available" });
            //Products.Add(new Product() { ProductId = 2, ProductName = "Banana", Price = 20.75m, Quantity = 50, Status = "Available" });
            //Products.Add(new Product() { ProductId = 3, ProductName = "Coconut", Price = 15.0m, Quantity = 75, Status = "Available" });
            //Products.Add(new Product() { ProductId = 4, ProductName = "Orange", Price = 12.0m, Quantity = 60, Status = "Out of Stock" });

        }
        #region OnpropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        #endregion

        public class Product
        {
            public int ProductId { get; set; }
            public string ProductName { get; set; }
            public decimal Price { get; set; }
            public int Quantity { get; set; }
            public string Status { get; set; }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            string filePath = @"D:\PRN_ASS\PRN211_ASS\ProductList.xlsx";
            ExportExcelSheet export = new ExportExcelSheet(ProductDataGrid, DateTime.Now, filePath);
        }

        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            Load_File();
        }

        private void Load_File()
        {
            string path = @"D:\PRN_ASS\PRN211_ASS\ProductList.xlsx";

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 
                    int rowCount = worksheet.Dimension.Rows; 
                    int colCount = worksheet.Dimension.Columns; 

                    
                    for (int row = 2; row <= rowCount; row++)
                    {
                        
                        Products.Add(new Product
                        {
                            ProductId = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                            ProductName = worksheet.Cells[row, 2].Value.ToString(),
                            Price = decimal.Parse(worksheet.Cells[row, 3].Value.ToString()),
                            Quantity = int.Parse(worksheet.Cells[row, 4].Value.ToString()),
                            Status = worksheet.Cells[row, 5].Value.ToString()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Error");
            }
        }
       
        
    }
}
