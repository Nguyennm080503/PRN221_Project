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

            Products.Add(new Product() { ProductId = 1, ProductName = "Apple", Price = 10.5m, Quantity = 100 });
            Products.Add(new Product() { ProductId = 2, ProductName = "Banana", Price = 20.75m, Quantity = 50 });
            Products.Add(new Product() { ProductId = 3, ProductName = "Coconut", Price = 15.0m, Quantity = 75 });
            Products.Add(new Product() { ProductId = 4, ProductName = "Coconut", Price = 15.0m, Quantity = 75 });
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

        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            ExportExcelSheet export = new ExportExcelSheet(ProductDataGrid, DateTime.Now);
        }

    }
}
