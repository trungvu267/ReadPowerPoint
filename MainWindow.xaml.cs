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

        public MainWindow()
        {
            InitializeComponent();

            ReadPPTX();
        }
        public ObservableCollection<BitmapImage> Images { get; set; }
        public ObservableCollection<string> texts { get; set; }

        public ObservableCollection<CustomShape> CustomShapes { get; set; }
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
        private void ReadPPTX()
        {

            // Load the PowerPoint file
            using (IPresentation presentation = Presentation.Open("C:\\Users\\vutru\\OneDrive\\Desktop\\nhom1.pptx"))
            {
                DataContext = this;

                Images = new ObservableCollection<BitmapImage>();
                texts = new ObservableCollection<string>();
                CustomShapes = new ObservableCollection<CustomShape>();

                //Loop through each slide in the presentation

                //Main Logic
                foreach (ISlide slide in presentation.Slides)
                {
                    // Loop through all the shapes in the slide
                    foreach (IShape shape in slide.Shapes)
                    {
                        CustomShape customShape = new CustomShape();
                        // Check if the shape is a text box
                        if (shape is IPicture)
                        {
                            //Images.Add(GetSlideImage(shape));
                            customShape.ImageShape = (GetSlideImage(shape));


                        }
                        else if (shape is IShape )
                        {
                            // Get the text from the text box
                            IShape textBox = (IShape)shape;
                            string text = textBox.TextBody.Text;
                            customShape.TextShape = text;
                            // Do something with the text
                            //Console.WriteLine(text);
                            //MessageBox.Show($"{text}");

                        }
                        else
                        {
                            //MessageBox.Show($"dont get data {slide.SlideNumber}");
                            MessageBox.Show($"hi");


                        }
                        CustomShapes.Add(customShape);
                    }
                    //slide page
                    //MessageBox.Show($"Author -  {slide.SlideNumber}");
                    //MessageBox.Show($"{customShapes[2].ImageShape}");
                }
                //ISlide slide = presentation.Slides[0];

                // 

                // Do something with the list of slide texts (e.g., display it in a message box)
                //MessageBox.Show("Author - {0}", presentation.BuiltInDocumentProperties.Title);
                // Get the first slide
                //ISlide slide = presentation.Slides[2];

                // Get the 3nd slide for test
                //ISlide slide3 = presentation.Slides[3];

                // lấy ảnh trong nhiều slides
                //foreach (ISlide slide in presentation.Slides)
                //{
                //    foreach (IShape shape in slide.Shapes)
                //    {
                //        if (shape is IPicture)
                //        {

                //            Images.Add(GetSlideImage(shape));
                //        }
                //    }
                //}
                //for test
                //ISlide slide = presentation.Slides[0];

                //// Iterate through the shapes in the slide and get their positions
                //foreach (IShape shape in slide.Shapes)
                //{
                   
                   

                //    if (shape is IShape)
                //    {

                //        IShape textBox = (IShape)shape;
                //        string text = textBox.TextBody.Text;
                //        texts.Add(text);
                //        // Do something with the text
                     
                //    }
                //}





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
        public Position GetShapePosition(IShape shape)
        {
            Position position = new Position();
            position.Left= shape.Left;
            position.Top=   shape.Top;
            position.Width =  shape.Width;
            position.Height = shape.Height;
            return position;
        }


    }
}
