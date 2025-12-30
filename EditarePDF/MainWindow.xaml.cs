using System.Text;
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
using System.Drawing;
using System.IO;
using PdfiumViewer;
using System.Windows.Ink;

namespace EditarePDF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenPdf_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PDF files (*.pdf)|*.pdf";

            if (openFileDialog.ShowDialog() == true)
            {
                // 1. Încărcăm documentul PDF offline
                using (var document = PdfDocument.Load(openFileDialog.FileName))
                {
                    // 2. Randăm prima pagină (indice 0) la o rezoluție bună pentru OCR (300 DPI)
                    // Parametrii: pagina, lățime, înălțime, DPI x, DPI y, rotire, opțiuni
                    var image = document.Render(0, 300, 300, true);
                    var bitmapSource = ImageToBitmapSource(image);

                    PdfDisplayImage.Source = bitmapSource;

                    System.Windows.Controls.Image wpfImage = new System.Windows.Controls.Image();
                    wpfImage.Source = ImageToBitmapSource(image);

                    // Sincronizăm dimensiunea stratului de desen cu imaginea
                    EraserCanvas.Width = bitmapSource.Width;
                    EraserCanvas.Height = bitmapSource.Height;
                    EraserCanvas.Strokes.Clear();
                }
            }
        }

        // Funcție utilitară pentru a face conversia de format necesară WPF-ului
        private BitmapSource ImageToBitmapSource(System.Drawing.Image bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }

        // 1. Configurăm aspectul radierei
        private void SetupEraser()
        {
            DrawingAttributes da = new DrawingAttributes();
            da.Color = System.Windows.Media.Colors.White; // Culoarea albă acoperă elementele
            da.Height = 20; // Dimensiunea radierei
            da.Width = 20;
            da.StylusTip = StylusTip.Rectangle; // Radieră pătrată sau elipsă

            EraserCanvas.DefaultDrawingAttributes = da;
        }

        // 2. Activăm modul de desenare când se apasă butonul "Radieră"
        private void Eraser_Click(object sender, RoutedEventArgs e)
        {
            // Când acest mod este activ, mouse-ul va desena cu alb
            EraserCanvas.EditingMode = InkCanvasEditingMode.Ink;
            SetupEraser();
        }

        private void ConvertToWord_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement PDF to Word conversion logic here
            MessageBox.Show("Convertire Word funcționalitate nu este încă implementată.");
        }

        private void EditorCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // TODO: Add your mouse down logic here
        }

        private void EditorCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            // TODO: Add your mouse move logic here
        }
    }
}