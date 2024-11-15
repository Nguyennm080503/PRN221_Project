using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using static WPF_Project.ControlVoice;

namespace WPF_Project
{
    /// <summary>
    /// Interaction logic for AddProduct.xaml
    /// </summary>
    public partial class AddProduct : Window
    {
        public AddProduct()
        {
            InitializeComponent();
        }

        private void txt_Name_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void txt_Price_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void txt_Quantity_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void btn_Add_Click(object sender, RoutedEventArgs e)
        {
            CreatProduct();
        }

        private void CreatProduct()
        {
            string filePath = @"D:\PRN_ASS\PRN211_ASS\ProductList.xlsx";

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Lấy số hàng hiện có trong sheet
                    int rowCount = worksheet.Dimension.Rows;

                    // Lấy productId cuối cùng từ sheet
                    int lastProductId = rowCount > 1 ? int.Parse(worksheet.Cells[rowCount, 1].Value.ToString()) : 0;

                    // Tạo sản phẩm mới với ProductId được tăng dần từ productId cuối cùng
                    Product newProduct = new Product
                    {
                        ProductId = lastProductId + 1, // Tăng productId cuối cùng lên 1
                        ProductName = txt_Name.Text,
                        Price = Convert.ToDecimal(txt_Price.Text),
                        Quantity = Convert.ToInt32(txt_Quantity.Text),
                        Status = 1
                    };

                    worksheet.Cells[rowCount + 1, 1].Value = newProduct.ProductId;
                    worksheet.Cells[rowCount + 1, 2].Value = newProduct.ProductName;
                    worksheet.Cells[rowCount + 1, 3].Value = newProduct.Price;
                    worksheet.Cells[rowCount + 1, 4].Value = newProduct.Quantity;
                    worksheet.Cells[rowCount + 1, 5].Value = newProduct.Status;

                    // Lưu lại file Excel
                    package.Save();
                    System.Windows.MessageBox.Show("Da add thanh cong!");

                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Error");
            }
        }
        private void btn_cancel_Click(object sender, RoutedEventArgs e)
        {
            txt_Name.Text = "";
            txt_Price.Text = "";
            txt_Quantity.Text = "";

        }
    }
}
