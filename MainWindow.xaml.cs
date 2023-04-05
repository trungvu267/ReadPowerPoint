using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Syncfusion.Drawing;
using Syncfusion.Presentation;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Drawing;
namespace TestReadPowerpoint
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<BitmapImage> Images { get; set; }
        public ObservableCollection<CustomShape> CustomShapes { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }
        public class Position
        {
            public double Top { get; set; }
            public double Left { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
        }
        public class CustomShape
        {
            public string TextShape { get; set; }
            public BitmapImage ImageShape { get; set; }
        }
        public async Task ReadPPTX(string path, ObservableCollection<CustomShape> customShapes)
        {
            using (IPresentation presentation = Presentation.Open(path))
            {
                CustomShapes = new ObservableCollection<CustomShape>();
                //Main Logic
                // Lặp qua các slide 
                foreach (ISlide slide in presentation.Slides)
                {
                    // Lặp qua các shape trong slided
                    // Loop through all the shapes in the slide
                    foreach (IShape shape in slide.Shapes)
                    {
                        CustomShape customShape = new CustomShape();
                        // Kiểm tra xem shape có phải là hình ảnh ko
                        if (shape is IPicture)
                        {
                            //Images.Add(GetSlideImage(shape));
                            customShape.ImageShape = (GetSlideImage(shape));
                        }
                        // Nếu ko phải hình ảnh sẽ chạy vào condition này
                        else if (shape is IShape)
                        {
                            IShape textBox = (IShape)shape;
                            string text = textBox.TextBody.Text;
                            customShape.TextShape = text;
                        }
                        CustomShapes.Add(customShape);
                    }
                }
            }
        }
        // Method mở ra dialog chọn file
        private async void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".pptx";
            dlg.Filter = "PowerPoint documents (.pptx)|*.pptx";

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                FilePathTextBox.Text = dlg.FileName;
                string filename = dlg.FileName;

                await ReadPPTX(filename, CustomShapes);
                MyListBox.ItemsSource = CustomShapes;
            }
        }
        // Lấy kiểu dữ liệu bitmap cho hình ảnh
        private BitmapImage GetSlideImage(IShape shape)
        {
            IPicture picture = (IPicture)shape;
            byte[] imageData = picture.ImageData;

            // Create a BitmapImage from the image data
            BitmapImage bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.StreamSource = new MemoryStream(imageData);
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.EndInit();

            return bitmap;

        }
        // Method này chưa dùng đến
        //GetShapePositions
        public Position GetShapePosition(IShape shape)
        {
            Position position = new Position();
            position.Left = shape.Left;
            position.Top = shape.Top;
            position.Width = shape.Width;
            position.Height = shape.Height;
            return position;
        }


    }
}