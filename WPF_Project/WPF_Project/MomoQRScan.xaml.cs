using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Drawing;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WPF_Project
{
    /// <summary>
    /// Interaction logic for MomoQRScan.xaml
    /// </summary>
    public partial class MomoQRScan : UserControl
    {
        public MomoQRScan()
        {
            InitializeComponent();
        }

        public void UpdateQRCode(Bitmap qrCode)
        {
            if (qrCode != null)
            {
                using (MemoryStream memory = new MemoryStream())
                {
                    qrCode.Save(memory, ImageFormat.Png);
                    memory.Position = 0;

                    BitmapImage bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = memory;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();

                    qrCodeImageView.Source = bitmapImage;
                }
            }
        }
    }
}
