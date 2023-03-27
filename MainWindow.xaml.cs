using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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
using Microsoft.Win32;
using Syncfusion.Drawing;
using Syncfusion.Presentation;

namespace TestReadPowerpoint
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            ReadPPTX();
        }
        public ObservableCollection<BitmapImage> Images { get; set; }

        private void ReadPPTX()
        {

            // Load the PowerPoint file
            using (IPresentation presentation = Presentation.Open("C:\\Users\\vutru\\OneDrive\\Desktop\\nhom1.pptx"))
            {
                //Loop through each slide in the presentation
                //foreach (ISlide slide in presentation.Slides)
                //{
                //    // Loop through all the shapes in the slide
                //    foreach (IShape shape in slide.Shapes)
                //    {
                //        // Check if the shape is a text box
                //        if (shape is IShape)
                //        {
                //            // Get the text from the text box
                //            IShape textBox = (IShape)shape;
                //            string text = textBox.TextBody.Text;

                //            // Do something with the text
                //            Console.WriteLine(text);
                //            //MessageBox.Show($"{text}");

                //        }
                //        else if (shape is IPicture)
                //        {
                //        }
                //        else
                //        {
                //            MessageBox.Show($"dont get data {slide.SlideNumber}");

                //        }
                //    }
                //    //MessageBox.Show($"Author -  {slide.SlideNumber}");
                //}
                //ISlide slide = presentation.Slides[0];


                // Do something with the list of slide texts (e.g., display it in a message box)
                //MessageBox.Show("Author - {0}", presentation.BuiltInDocumentProperties.Title);
                // Get the first slide
                //ISlide slide = presentation.Slides[2];

                // Get the first slide
                ISlide slide3 = presentation.Slides[3];
                Images = new ObservableCollection<BitmapImage>();
           

                //Loop through all the shapes in the slide
                //foreach (IShape shape in slide3.Shapes)
                //{
                //    // Check if the shape is an image
                //    if (shape is IPicture)
                //    {

                //        // Set the source of the Image element in your WPF UI
                //        PowerPointImage.Source = GetSlideImage(shape);
                //        Images.Add(GetSlideImage(shape));
                //        MessageBox.Show($"{Images[0]}");


                //    }
                //}

                // lấy ảnh trong nhiều slides
                foreach (ISlide slide in presentation.Slides)
                {
                    foreach (IShape shape in slide.Shapes)
                    {
                        if (shape is IPicture)
                        {

                            Images.Add(GetSlideImage(shape));
                        }
                    }
                }

            }
        }
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            // Create an open file dialog
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PowerPoint Presentation (*.pptx)|*.pptx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Show the dialog and get the result
            bool? result = openFileDialog.ShowDialog();

            // If the user clicked "OK", update the file path text box
            if (result == true)
            {
                FilePathTextBox.Text = openFileDialog.FileName;
                ReadPPTX();
            }
        }



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

            // Set the source of the Image element in your WPF UI
            return bitmap;

        }
        //GetShapePositions
        public void GetShapePositions(ISlide slide)
        {
            foreach (IShape shape in slide.Shapes)
            {
                double left = shape.Left;
                double top = shape.Top;
                double width = shape.Width;
                double height = shape.Height;
            }
        }


    }
}
