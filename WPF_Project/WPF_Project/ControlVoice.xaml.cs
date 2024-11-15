using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Formats.Asn1;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography.X509Certificates;
using System.Speech.Recognition;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Shapes;
using System.Xml;
using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;
using static WPF_Project.ControlVoice;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace WPF_Project
{
    public partial class ControlVoice : Window
    {
        private SpeechRecognitionEngine recognizer;
        private string recognizedText;
        private ExcelWorksheet worksheet;
        private ExcelPackage package;
        private static int currentProductId = 1;
        private List<Product> mProduct = null;
        private string filePath = @"D:\PRN_ASS\PRN221_Ass\PRN211_ASS\ProductList.xlsx";
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
            public int Status { get; set; }

        }

        public class ProductWithRowIndex
        {
            public Product Product { get; set; }
            public int RowIndex { get; set; }
        }

        public class Order
        {
            public int OrderID { get; set; }
      
            public decimal Total { get; set; }

            public int Quantity { get; set; }

            public string? CreareDate { get; set; }

            public int Status { get; set; }

        }

        public class OrderDetail
        {
            public int ID { get; set; }
            public int OrderID { get; set; }

            public int ProductId { get; set; }

            public decimal Price { get; set; }
            public decimal SubPrice { get; set; }


            public string? CreareDate { get; set; }

            public int Status { get; set; }

            public int Quantity { get; set; }

        }


        public class Payment
        {
            public int ID { get; set; }

            public string Method { get; set; }

            public decimal Total { get; set; }

            public int OrderID { get; set; }

            public int Status { get; set; }
        }

            private async Task CreateOrderAndODetail(string text)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    // Đọc dữ liệu từ Sheet Order
                    ExcelWorksheet orderWorksheet = package.Workbook.Worksheets[1];
                    ExcelWorksheet productWorksheet = package.Workbook.Worksheets[0];
                    
                    int orderRowCount = orderWorksheet.Dimension.Rows;
                    
                    int lastOrderId = orderRowCount > 1 ? int.Parse(orderWorksheet.Cells[orderRowCount, 1].Value.ToString()) : 0;

                    //tao order
                    DateTime date = DateTime.Now;
                    Order order = new Order()
                    {
                        OrderID = lastOrderId + 1,
                        CreareDate = date.ToString(),
                        Status = 1
                    };

                    // Ghi thông tin của đơn hàng mới vào Sheet Order
                    orderWorksheet.Cells[orderRowCount + 1, 1].Value = order.OrderID;
                    orderWorksheet.Cells[orderRowCount + 1, 4].Value = order.Status;
                    orderWorksheet.Cells[orderRowCount + 1, 5].Value = order.CreareDate;
                    package.Save();
                    

                    //nhom theo cap quality va name
                    List<string> groupQualityAndNames = GroupQualityAndName(text);
                    int totalQuantity = 0;
                    decimal totalPrice = 0;
                    foreach (string groupQualityAndName in groupQualityAndNames)
                    {
                        ExcelWorksheet orderDetailWorksheet = package.Workbook.Worksheets[2];
                        int orderDetailRowCount = orderDetailWorksheet.Dimension.Rows;
                        int lastOrderDetailId = orderDetailRowCount > 1 ? int.Parse(orderDetailWorksheet.Cells[orderDetailRowCount, 1].Value.ToString()) : 0;
                        List<string> words = Tokenize(groupQualityAndName);
                        ProductWithRowIndex product = null;
                        int quantity = DetectQuantity(words);
                        if (quantity == 1)
                        {
                            product = DetectProductByOneProduct(words);
                        }
                        else
                        {
                            product = DetectProductByManyProduct(words);
                        }

                        int quantityOfProduct = product.Product.Quantity;
                        int status = 1;

                        if(product != null)
                        {
                            OrderDetail orderDetail = new OrderDetail()
                            {
                                ID = lastOrderDetailId + 1,
                                OrderID = lastOrderId + 1,
                                ProductId = product.Product.ProductId,
                                Price = product.Product.Price,
                                SubPrice = product.Product.Price * quantity,
                                Status = 1,
                                CreareDate = date.ToString(),
                                Quantity = quantity,

                            };
                            quantityOfProduct -= quantity;
                            
                            if (quantityOfProduct == 0)
                            {
                                status = 0;
                                totalQuantity += quantity;
                                totalPrice += orderDetail.SubPrice;

                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 1].Value = orderDetail.ID;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 2].Value = orderDetail.OrderID;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 3].Value = orderDetail.ProductId;

                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 4].Value = orderDetail.Price;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 5].Value = orderDetail.SubPrice;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 6].Value = orderDetail.Status;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 7].Value = orderDetail.CreareDate;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 8].Value = orderDetail.Quantity;

                                orderWorksheet.Cells[orderRowCount + 1, 3].Value = totalPrice;
                                orderWorksheet.Cells[orderRowCount + 1, 2].Value = totalQuantity;
                                productWorksheet.Cells[product.RowIndex, 4].Value = quantityOfProduct;
                                productWorksheet.Cells[product.RowIndex, 5].Value = status;

                                package.Save();
                                
                            }
                            else
                            {
                                totalQuantity += quantity;
                                totalPrice += orderDetail.SubPrice;

                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 1].Value = orderDetail.ID;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 2].Value = orderDetail.OrderID;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 3].Value = orderDetail.ProductId;

                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 4].Value = orderDetail.Price;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 5].Value = orderDetail.SubPrice;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 6].Value = orderDetail.Status;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 7].Value = orderDetail.CreareDate;
                                orderDetailWorksheet.Cells[orderDetailRowCount + 1, 8].Value = orderDetail.Quantity;

                                orderWorksheet.Cells[orderRowCount + 1, 3].Value = totalPrice;
                                orderWorksheet.Cells[orderRowCount + 1, 2].Value = totalQuantity;
                                productWorksheet.Cells[product.RowIndex, 4].Value = quantityOfProduct;
                                productWorksheet.Cells[product.RowIndex, 5].Value = status;

                                package.Save();
                                
                            }
                        }
                        else
                        {
                            System.Windows.MessageBox.Show("Không tìm thấy sản phẩm phù hợp.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    System.Windows.MessageBox.Show("Da add order thanh cong!");

                    
                    ExcelWorksheet paymentWorksheet = package.Workbook.Worksheets[3];

                    int paymentRowCount = paymentWorksheet.Dimension.Rows;

                    int lastPaymentId = paymentRowCount > 1 ? int.Parse(paymentWorksheet.Cells[paymentRowCount, 1].Value.ToString()) : 0;
                    Payment  payment = new Payment()
                    {
                        ID = lastPaymentId + 1,
                        OrderID = lastOrderId + 1,
                        Method = "QR Code",
                        Total = totalPrice,
                        Status = 0,
                    };

                    // Ghi thông tin của đơn hàng mới vào Sheet Payment
                    
                     btPayment(lastOrderId + 1);
                    
                        paymentWorksheet.Cells[paymentRowCount + 1, 1].Value = payment.ID;
                        paymentWorksheet.Cells[paymentRowCount + 1, 2].Value = payment.Method;
                        paymentWorksheet.Cells[paymentRowCount + 1, 3].Value = payment.Total;
                        paymentWorksheet.Cells[paymentRowCount + 1, 4].Value = payment.OrderID;
                        paymentWorksheet.Cells[paymentRowCount + 1, 5].Value = 1;
                        package.Save();

                    System.Windows.MessageBox.Show("Thanh toan thanh cong ", "Success", MessageBoxButton.OK);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private Bitmap ResizeImage(Bitmap image, int maxWidth, int maxHeight)
        {
            double ratioX = (double)maxWidth / image.Width;
            double ratioY = (double)maxHeight / image.Height;
            double ratio = Math.Min(ratioX, ratioY);

            int newWidth = (int)(image.Width * ratio);
            int newHeight = (int)(image.Height * ratio);

            Bitmap newImage = new Bitmap(newWidth, newHeight);
            Graphics g = Graphics.FromImage(newImage);
            g.DrawImage(image, 0, 0, newWidth, newHeight);
            return newImage;
        }

        private Bitmap AddLogoToQRCode(Bitmap qrCode, Bitmap logo)
        {
            int xPos = (qrCode.Width - logo.Width) / 2;
            int yPos = (qrCode.Height - logo.Height) / 2;
            using (Graphics g = Graphics.FromImage(qrCode))
            {
                g.DrawImage(logo, new System.Drawing.Point(xPos, yPos));
            }
            return qrCode;
        }

        private bool ProcessPayment()
        {
            bool paymentSuccess = false;


            return paymentSuccess;
        }
        private async void btPayment(int orderId)
        {
            var order = FindOrderById(orderId);
            string Phone = "0961545926";
            string Name = "Dương Thị Hiền";
            string Email = "thaohien1372002@gmail.com";
            string PayNumber = order.Total.ToString();
            string Datetimes = DateTime.Now.ToString("dd/MM/yyyy");

            
                    string orderInfo = orderId.ToString();
                    string Description = $"orderId {orderInfo} + Giá tiền {PayNumber} + {Datetimes}";
                    MomoQRCodeGenerator momoGenerator = new MomoQRCodeGenerator();
                    string merchantCode = $"2|99|{Phone}|{Name}|{Email}|0|0|{PayNumber}|{Description}";
                    Bitmap momoQRCode = momoGenerator.GenerateMomoQRCode(merchantCode);
                    Bitmap resizedLogo = ResizeImage(Properties.Resources.logo, 50, 50);
                    momoQRCode = AddLogoToQRCode(momoQRCode, resizedLogo);


            if (!string.IsNullOrWhiteSpace(PayNumber))
            {
                MomoQRScan scanQR = new MomoQRScan();
                scanQR.UpdateQRCode(momoQRCode);
                double windowWidth = 500;
                double windowHeight = 500;
                Window qrCodeWindow = new Window
                {
                    Content = scanQR,
                    Width = windowWidth,
                    Height = windowHeight,
                    WindowStyle = WindowStyle.None,
                    ResizeMode = ResizeMode.NoResize,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    Title = "ScanQR"
                };
                qrCodeWindow.Content = scanQR;

                qrCodeWindow.Show();
                await Task.Delay(60000);        // Check time 
                if (!ProcessPayment())
                {
                    qrCodeWindow.Close();
                    MessageBox.Show("QR code quá thời gian. Vui lòng thanh toán lại !!!", "Error", MessageBoxButton.OK);

                }



            }
            else
            {
                MessageBox.Show("Please enter a valid total price.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
                }
            


    

    private List<string> GroupQualityAndName(string textinput)
        {
            List<string> ListResult = new List<string>();
            MatchCollection matches = Regex.Matches(textinput, 
                @"\b(?:one|two|three|four|five|six|seven|eight|nine|ten)\s(?:apples|apple|oranges|orange|coconut|coconuts|banana|bananas)\b");

            foreach(Match match in matches)
            {
                ListResult.Add(match.Value);
            }
            return ListResult;
        }
        //private void CreateOrder(string text)
        //{
        //    try
        //    {
        //        using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
        //        {
        //            // Đọc dữ liệu từ Sheet Order
        //            ExcelWorksheet orderWorksheet = package.Workbook.Worksheets[1];
        //            ExcelWorksheet productWorksheet = package.Workbook.Worksheets[0];
        //            int orderRowCount = orderWorksheet.Dimension.Rows;
        //            int lastOrderDetailId = orderRowCount > 1 ? int.Parse(orderWorksheet.Cells[orderRowCount, 1].Value.ToString()) : 0;


        //            List<string> words = Tokenize(text);

        //            ProductWithRowIndex product = null;

        //            int quantity = DetectQuantity(words);
        //            if (quantity == 1)
        //            {
        //                product = DetectProductByOneProduct(words);
        //            }
        //            else
        //            {
        //                product = DetectProductByManyProduct(words);
        //            }
        //            DateTime date = DateTime.Now;
        //            int quanityOfProduct = product.Product.Quantity;
        //            int status = 1;
        //            if (product != null)
        //            {

        //                Order order = new Order()
        //                {
        //                    OrderID = orderRowCount + 1,
        //                    ProductID = product.Product.ProductId,
        //                    Price = product.Product.Price,
        //                    Total = product.Product.Price * quantity,
        //                    Quantity = quantity,
        //                    CreareDate = date.ToString(),
        //                    Status = 1
        //                };

        //                quanityOfProduct -= quantity;

        //                if (quanityOfProduct < 0)
        //                {
        //                    System.Windows.MessageBox.Show("Khong du san pham trong gio hang!");
        //                }
        //                else if (quanityOfProduct == 0)
        //                {
        //                    status = 0;
        //                    // Ghi thông tin của đơn hàng mới vào Sheet Order
        //                    orderWorksheet.Cells[orderRowCount + 1, 1].Value = order.OrderID;
        //                    orderWorksheet.Cells[orderRowCount + 1, 2].Value = order.ProductID;
        //                    orderWorksheet.Cells[orderRowCount + 1, 3].Value = order.Quantity;
        //                    orderWorksheet.Cells[orderRowCount + 1, 4].Value = order.Price;
        //                    orderWorksheet.Cells[orderRowCount + 1, 5].Value = order.Total;
        //                    orderWorksheet.Cells[orderRowCount + 1, 6].Value = order.Status;
        //                    orderWorksheet.Cells[orderRowCount + 1, 7].Value = order.CreareDate;
        //                    productWorksheet.Cells[product.RowIndex, 4].Value = quanityOfProduct;
        //                    productWorksheet.Cells[product.RowIndex, 5].Value = status;
        //                    package.Save();
        //                    System.Windows.MessageBox.Show("Da add order thanh cong!");
        //                }
        //                else
        //                {
        //                    // Ghi thông tin của đơn hàng mới vào Sheet Order
        //                    orderWorksheet.Cells[orderRowCount + 1, 1].Value = order.OrderID;
        //                    orderWorksheet.Cells[orderRowCount + 1, 2].Value = order.ProductID;
        //                    orderWorksheet.Cells[orderRowCount + 1, 3].Value = order.Quantity;
        //                    orderWorksheet.Cells[orderRowCount + 1, 4].Value = order.Price;
        //                    orderWorksheet.Cells[orderRowCount + 1, 5].Value = order.Total;
        //                    orderWorksheet.Cells[orderRowCount + 1, 6].Value = order.Status;
        //                    orderWorksheet.Cells[orderRowCount + 1, 7].Value = order.CreareDate;
        //                    productWorksheet.Cells[product.RowIndex, 4].Value = quanityOfProduct;
        //                    productWorksheet.Cells[product.RowIndex, 5].Value = status;
        //                    package.Save();
        //                    System.Windows.MessageBox.Show("Da add order thanh cong!");
        //                }

        //            }

        //            else
        //            {
        //                System.Windows.MessageBox.Show("Không tìm thấy sản phẩm phù hợp.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Windows.MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        //    }
        //}


        private ProductWithRowIndex FindProductByName(string productName)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string? currentProductName = worksheet.Cells[row, 2].Value.ToString();

                        if (currentProductName.Equals(productName, StringComparison.OrdinalIgnoreCase))
                        {
                            Product product = new Product
                            {
                                ProductId = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                                ProductName = currentProductName,
                                Price = decimal.Parse(worksheet.Cells[row, 3].Value.ToString()),
                                Quantity = int.Parse(worksheet.Cells[row, 4].Value.ToString()),
                                Status = int.Parse(worksheet.Cells[row, 5].Value.ToString())
                            };

                            return new ProductWithRowIndex { Product = product, RowIndex = row };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Error");
            }

            return null; 
        }

        private Order FindOrderById(int OrderId)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        int? currentOrderId = null;
                        var orderIdCell = worksheet.Cells[row, 1].Value;
                        if (orderIdCell != null && int.TryParse(orderIdCell.ToString(), out int parsedOrderId))
                        {
                            currentOrderId = parsedOrderId;
                        }

                        if (currentOrderId == OrderId)
                        {
                            int? total = null;
                            var totalCell = worksheet.Cells[row, 3].Value;
                            if (totalCell != null && int.TryParse(totalCell.ToString(), out int parsedTotal))
                            {
                                total = parsedTotal;
                            }

                            int? status = null;
                            var statusCell = worksheet.Cells[row, 4].Value;
                            if (statusCell != null && int.TryParse(statusCell.ToString(), out int parsedStatus))
                            {
                                status = parsedStatus;
                            }

                            Order order = new Order
                            {
                                OrderID = OrderId,
                                Total = total ?? 0,
                                Status = status ?? 0
                            };

                            return order;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Error");
            }

            return null;
        }



        private void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            txt_TextAdd.Text = txt_TextAdd.Text + e.Result.Text.ToString() + Environment.NewLine;
            recognizedText = e.Result.Text;

        }


        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            Load_File();
        }

        private void Load_File()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 
                    int rowCount = worksheet.Dimension.Rows; 
                    int colCount = worksheet.Dimension.Columns; 

                    
                    for (int row = 2; row <= rowCount; row++)
                    {
                        int status = int.Parse(worksheet.Cells[row, 5].Value.ToString());
                        if(status == 1)
                        {
                            Products.Add(new Product
                            {
                                ProductId = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                                ProductName = worksheet.Cells[row, 2].Value.ToString(),
                                Price = decimal.Parse(worksheet.Cells[row, 3].Value.ToString()),
                                Quantity = int.Parse(worksheet.Cells[row, 4].Value.ToString()),
                                Status = int.Parse(worksheet.Cells[row, 5].Value.ToString())
                            });
                        }
                        
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Error");
            }
        }

        private void btn_spAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                recognizer = new SpeechRecognitionEngine();
                recognizer.SetInputToDefaultAudioDevice();
                Grammar grammar = new DictationGrammar();
                recognizer.LoadGrammar(grammar);
                recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);
                recognizer.RecognizeAsync(RecognizeMode.Multiple);

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Error");
            }

        }

        private void btn_addSh_Click(object sender, RoutedEventArgs e)
        {
            AddProduct addProduct = new AddProduct();
            addProduct.ShowDialog();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
                //recognizer.RecognizeAsyncCancel();
                recognizer.Dispose();
            
            txt_TextAdd.Clear();
        }

        private void btn_Stop_Click(object sender, RoutedEventArgs e)
        {
            bool containsOrderWord = false;
            if (recognizer != null)
            {
                recognizer.RecognizeAsyncCancel();
                recognizer.Dispose();
            }

            List<string> words = Tokenize(recognizedText);
            string OrderWord = "order";
            foreach(string word in words)
            {
                if (OrderWord.Contains(word.ToLower()))
                {

                    CreateOrderAndODetail(recognizedText);
                    containsOrderWord = true;
                    break;
                }
            }

            if (!containsOrderWord)
            {
                System.Windows.MessageBox.Show("Không tìm thấy từ add", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        

        private List<string> Tokenize(string text)
        {
            return text.Split(new char[] { ' ', '.', ',', '!', '?' }, StringSplitOptions.RemoveEmptyEntries).ToList();
        }

        private ProductWithRowIndex DetectProductByManyProduct(List<string> words)
        {
            string[] products = { "apple", "banana", "orange", "coconut" };

            foreach (string word in words)
            {
                string trimmedWord = word.ToLower().TrimEnd('s');

                
                if (products.Contains(trimmedWord.ToLower()))
                {
                    var product = FindProductByName(trimmedWord);

                    return product;
                }
            }
            System.Windows.MessageBox.Show("Không tìm thấy sản phẩm với tên đã nhận dạng.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            return null;
        }

        private ProductWithRowIndex DetectProductByOneProduct(List<string> words)
        {
            string[] products = { "apple", "banana", "orange", "coconut" };

            foreach (string word in words)
            {
                if (products.Contains(word.ToLower()))
                {
                    var product = FindProductByName(word);

                    return product;
                }
            }
            System.Windows.MessageBox.Show("Không tìm thấy sản phẩm với tên đã nhận dạng.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            return null;
        }


        private int DetectQuantity(List<string> words)
        {
            Dictionary<string, int> numberWords = new Dictionary<string, int>()
              {
        {"one", 1},
        {"two", 2},
        {"three", 3},
        {"four", 4},
        {"five", 5},
                {"six", 6 },
                {"seven", 7 },
                {"eight", 8 },
                {"nine", 9 },
                {"ten", 10 }
                };

            int quantity = 0;
            
            foreach (string word in words)
            {
                if (numberWords.ContainsKey(word.ToLower()))
                {
                    quantity = numberWords[word.ToLower()];
                   
                }
            }

            return quantity;
        }

    }
}