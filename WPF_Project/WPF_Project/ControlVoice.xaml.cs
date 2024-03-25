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
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Lấy trang tính đầu tiên
                    int rowCount = worksheet.Dimension.Rows; // Số hàng trong trang tính
                    int colCount = worksheet.Dimension.Columns; // Số cột trong trang tính

                    // Bắt đầu từ hàng thứ hai (hàng đầu tiên thường là tiêu đề)
                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Đọc dữ liệu từ các ô trong mỗi hàng và thêm vào danh sách Products
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
        private async void Load_FileExcel()
        {
            string path = @"D:\PRN_ASS\PRN211_ASS\ProductList.xlsx";
            Stopwatch stopwatch = new Stopwatch();
            using (var reader = new StreamReader(path))
            {
                using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.CurrentCulture)))
                {
                    stopwatch.Start();
                    // Read and skip the header
                    await csv.ReadAsync();
                    csv.ReadHeader();

                    // Set up DataTable for batch processing
                    var dataTable = new DataTable();
                    dataTable.Columns.Add("ProductID", typeof(int));
                    dataTable.Columns.Add("ProductName", typeof(string));
                    dataTable.Columns.Add("Price", typeof(decimal));
                    dataTable.Columns.Add("Quantity", typeof(decimal));
                    dataTable.Columns.Add("Status", typeof(string));
                    int hasData = 0;
                    while (await csv.ReadAsync())
                    {
                        var fieldData = csv.Context.Record;
                        
                            hasData = 1;
                            var dataRow = dataTable.NewRow();
                            dataRow["ProductID"] = csv.Context.Row - 2;
                            dataRow["ProductName"] = string.IsNullOrEmpty(fieldData[0]);
                            dataRow["Price"] = string.IsNullOrEmpty(fieldData[1]) ? DBNull.Value : ConvertStringIntoDecimal(fieldData[2]);
                        dataRow["Quantity"] = string.IsNullOrEmpty(fieldData[2]) ? DBNull.Value : ConvertStringIntoDecimal(fieldData[2]);
                        dataRow["Status"] = string.IsNullOrEmpty(fieldData[3]);
                           
                            //dataRow["MaTinh"] = int.Parse(fieldData[11]);
                            dataTable.Rows.Add(dataRow);
                        ProductDataGrid.ItemsSource = (System.Collections.IEnumerable)dataTable;
                            if (dataTable.Rows.Count % 500 == 0)
                            {
                                dataTable.Rows.Clear();
                            }
                        
                        if (hasData == 1)
                        {
                            System.Windows.MessageBox.Show($"Da cuoi dong roi");
                            break;
                        }
                    }
                    
                    stopwatch.Stop();
                }
            }
            TimeSpan elapsedTime = stopwatch.Elapsed;
            System.Windows.MessageBox.Show($"Thời gian thực hiện: {elapsedTime.TotalSeconds} s");
        }
        private decimal? ConvertStringIntoDecimal(string input)
        {
            if (string.IsNullOrEmpty(input))
                return null;
            return decimal.Parse(input);
        }
    }
}
