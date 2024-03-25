using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography.X509Certificates;
using System.Windows;
using System.Windows.Controls;
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

            Products.Add(new Product() { ProductName = "Apple", Price = 10000, Quantity = 100, Status = "Available" });
            Products.Add(new Product() { ProductName = "Banana", Price = 12000, Quantity = 50, Status = "Available" });
            Products.Add(new Product() { ProductName = "Coconut", Price = 20000, Quantity = 75, Status = "Available" });
            Products.Add(new Product() { ProductName = "Orange", Price = 30000, Quantity = 60, Status = "Out of Stock" });

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
            private static int nextProductId = 1; // biến static để theo dõi ProductId kế tiếp

            public int ProductId { get; private set; }
            public string ProductName { get; set; }
            public decimal Price { get; set; }
            public int Quantity { get; set; }
            public string Status { get; set; }

            public Product()
            {
                ProductId = nextProductId++; // Tự động tăng và gán ProductId cho mỗi sản phẩm
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            string filePath = @"D:\PRN221_Slide\ProductList.xlsx";
            ExportExcelSheet export = new ExportExcelSheet(ProductDataGrid, DateTime.Now, filePath);
        }

    }
}