using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // Pentru formate .xlsx
using NPOI.XWPF.UserModel;
using PdfiumViewer;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Tesseract;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.ObjectModel;

namespace EditarePDF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Variabilă la nivel de clasă pentru a păstra textul extras
        private string ultimulTextExtras = "";

        // Add a simple view model for page thumbnails
        public class PdfPageItem
        {
            public ImageSource Thumbnail { get; set; } = default!;
            public int PageNumber { get; set; }
        }

        private readonly ObservableCollection<PdfPageItem> _pages = new ObservableCollection<PdfPageItem>();
        private PdfDocument? _loadedDocument;

        public MainWindow()
        {
            InitializeComponent();
            PagesList.ItemsSource = _pages;
        }

        private void OpenPdf_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PDF files (*.pdf)|*.pdf";

            if (openFileDialog.ShowDialog() == true)
            {
                // 1. Încărcăm documentul PDF offline
                _loadedDocument?.Dispose();
                _loadedDocument = PdfDocument.Load(openFileDialog.FileName);

                _pages.Clear();

                // Render all pages as thumbnails and add page numbers
                for (int i = 0; i < _loadedDocument.PageCount; i++)
                {
                    using (var img = _loadedDocument.Render(i, 150, 150, true))
                    {
                        var thumb = ImageToBitmapSource(img);
                        _pages.Add(new PdfPageItem
                        {
                            Thumbnail = thumb,
                            PageNumber = i + 1
                        });
                    }
                }

                PagesList.ItemsSource = _pages;

                // Select and display first page
                if (_loadedDocument.PageCount > 0)
                {
                    if (PagesList is ListBox listBox)
                    {
                        listBox.SelectedIndex = 0;
                    }
                    DisplayPage(0);
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

        private RenderTargetBitmap CreateCleanImage()
        {
            // 1. Calculăm dimensiunile zonei pe care vrem să o salvăm
            int width = (int)EraserCanvas.Width;
            int height = (int)EraserCanvas.Height;

            // 2. Creăm un obiect bitmap care va "fotografia" interfața
            // Folosim 96 DPI pentru ecran, dar Tesseract va procesa pixelii rezultați
            RenderTargetBitmap renderBitmap = new RenderTargetBitmap(
                width, height, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);

            // 3. Randăm (desenăm) Grid-ul care conține și imaginea și radiera în acest bitmap
            renderBitmap.Render(ContainerGrid);

            return renderBitmap;
        }

        private System.Drawing.Bitmap BitmapSourceToBitmap(BitmapSource srs)
        {
            int width = srs.PixelWidth;
            int height = srs.PixelHeight;
            int stride = width * ((srs.Format.BitsPerPixel + 7) / 8);
            IntPtr ptr = System.Runtime.InteropServices.Marshal.AllocHGlobal(height * stride);

            srs.CopyPixels(new System.Windows.Int32Rect(0, 0, width, height), ptr, height * stride, stride);

            using (var bmap = new System.Drawing.Bitmap(width, height, stride,
                   System.Drawing.Imaging.PixelFormat.Format32bppPArgb, ptr))
            {
                // Creăm o copie pentru a elibera memoria pointer-ului immediat
                return new System.Drawing.Bitmap(bmap);
            }
        }

        private void PrepareForOcr_Click(object sender, RoutedEventArgs e)
        {
            RenderTargetBitmap cleanBitmapSource = CreateCleanImage();

            using (System.Drawing.Bitmap finalImage = BitmapSourceToBitmap(cleanBitmapSource))
            {
                string rezultatulText = ExtractTextFromImage(finalImage);

                if (!string.IsNullOrEmpty(rezultatulText))
                {
                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.Filter = "Word Document (*.docx)|*.docx";
                    saveDialog.FileName = "Document_Convertit.docx";

                    if (saveDialog.ShowDialog() == true)
                    {
                        ExportToWord(rezultatulText, saveDialog.FileName);
                        MessageBox.Show("Exportul în Word a fost finalizat cu succes!");
                    }
                }

                ultimulTextExtras = ExtractTextFromImage(finalImage);
                if (!string.IsNullOrEmpty(ultimulTextExtras))
                {
                    MessageBox.Show("Textul a fost extras și este gata pentru export!");
                }
            }
        }

        public string ExtractTextFromImage(System.Drawing.Bitmap cleanImage)
        {
            try
            {
                // "tessdata" este folderul creat mai sus, "ron+eng" înseamnă că va căuta ambele limbi
                // EngineMode.Default este cel mai echilibrat pentru viteză/precizie
                using (var engine = new TesseractEngine(@"./tessdata", "ron+eng", EngineMode.Default))
                {
                    // Tesseract funcționează cel mai bine cu imagini Pix
                    // Convertim Bitmap-ul nostru curățat într-un format înțeles de Tesseract
                    using (var img = PixConverter.ToPix(cleanImage))
                    {
                        using (var page = engine.Process(img))
                        {
                            // Extragem textul
                            string text = page.GetText();

                            // Putem obține și gradul de încredere (0.0 - 1.0)
                            float confidence = page.GetMeanConfidence();
                            Console.WriteLine($"Precizie estimată: {confidence:P}");

                            return text;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return "Eroare la procesarea OCR: " + ex.Message;
            }
        }

        public void ExportToWord(string extractedText, string filePath)
        {
            // Creăm un document Word nou (.docx)
            XWPFDocument doc = new XWPFDocument();

            // Creăm un paragraf
            XWPFParagraph p1 = doc.CreateParagraph();
            XWPFRun run = p1.CreateRun();

            // Inserăm textul extras
            // Înlocuim caracterele de linie nouă pentru a fi recunoscute de Word
            if (extractedText != null)
            {
                string[] lines = extractedText.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.None);
                foreach (var line in lines)
                {
                    run.SetText(line);
                    run.AddCarriageReturn(); // Adaugă enter după fiecare linie
                }
            }

            // Salvăm fișierul pe disc
            using (FileStream sw = new FileStream(filePath, FileMode.Create))
            {
                doc.Write(sw);
            }
        }

        private void ExportToExcel(string extractedText, string filePath)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Text Extras");

            string[] lines = extractedText.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < lines.Length; i++)
            {
                IRow row = sheet.CreateRow(i);

                // Încercăm să separăm coloanele prin tab-uri sau grupuri de spații (minim 2 spații)
                // Tesseract pune adesea mai multe spații între coloanele unui tabel
                string[] columns = System.Text.RegularExpressions.Regex.Split(lines[i], @"\t|\s{2,}");

                for (int j = 0; j < columns.Length; j++)
                {
                    NPOI.SS.UserModel.ICell cell = row.CreateCell(j);
                    cell.SetCellValue(columns[j].Trim());
                }
            }

            // Auto-ajustare lățime coloane
            for (int i = 0; i < 10; i++) { try { sheet.AutoSizeColumn(i); } catch { } }

            using (FileStream sw = new FileStream(filePath, FileMode.Create))
            {
                workbook.Write(sw);
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(ultimulTextExtras))
            {
                MessageBox.Show("Vă rugăm să apăsați 'Procesează OCR' mai întâi.");
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
            saveDialog.FileName = "Tabel_Document.xlsx";

            if (saveDialog.ShowDialog() == true)
            {
                ExportToExcel(ultimulTextExtras, saveDialog.FileName);
                MessageBox.Show("Fișierul Excel a fost salvat!");
            }
        }

        private void ExportToWord_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(ultimulTextExtras))
            {
                MessageBox.Show("Vă rugăm să apăsați 'Procesează OCR' mai întâi.");
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Word Document (*.docx)|*.docx";
            saveDialog.FileName = "Document_Convertit.docx";

            if (saveDialog.ShowDialog() == true)
            {
                ExportToWord(ultimulTextExtras, saveDialog.FileName);
                MessageBox.Show("Fișierul Word a fost salvat!");
            }
        }

        private void ExportToImage_Click(object sender, RoutedEventArgs e)
        {
            RenderTargetBitmap cleanBitmapSource = CreateCleanImage();

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "PNG Image (*.png)|*.png|JPEG Image (*.jpg)|*.jpg";
            saveDialog.FileName = "Document_Imagine";

            if (saveDialog.ShowDialog() == true)
            {
                BitmapEncoder encoder;
                string ext = System.IO.Path.GetExtension(saveDialog.FileName).ToLowerInvariant();
                if (ext == ".jpg" || ext == ".jpeg")
                {
                    encoder = new JpegBitmapEncoder();
                }
                else
                {
                    encoder = new PngBitmapEncoder();
                }

                encoder.Frames.Add(BitmapFrame.Create(cleanBitmapSource));
                using (var fs = new FileStream(saveDialog.FileName, FileMode.Create))
                {
                    encoder.Save(fs);
                }
                MessageBox.Show("Imaginea a fost salvată!");
            }
        }

        private void ExportToPowerPoint(string extractedText, string filePath)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
            {
                presentationDocument.AddPresentationPart();
                PresentationPart presentationPart = presentationDocument.PresentationPart!;
                presentationPart.Presentation = new Presentation();

                SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
                slideMasterPart.SlideMaster = new SlideMaster(new CommonSlideData(new ShapeTree()));

                SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
                slideLayoutPart.SlideLayout = new SlideLayout(new CommonSlideData(new ShapeTree()));

                SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));

                // Link layout
                slidePart.AddPart(slideLayoutPart);

                // Add slide id list
                presentationPart.Presentation.SlideIdList = new SlideIdList(new SlideId() { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

                // Ensure CommonSlideData and ShapeTree are not null
                var commonSlideData = slidePart.Slide.CommonSlideData;
                if (commonSlideData == null)
                {
                    commonSlideData = new CommonSlideData(new ShapeTree());
                    slidePart.Slide.CommonSlideData = commonSlideData;
                }

                var shapeTree = commonSlideData.ShapeTree;
                if (shapeTree == null)
                {
                    shapeTree = new ShapeTree();
                    commonSlideData.ShapeTree = shapeTree;
                }

                // Add a textbox shape
                var shape = new Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties() { Id = 2U, Name = "TextBox 1" },
                        new NonVisualShapeDrawingProperties(new DocumentFormat.OpenXml.Drawing.ShapeLocks() { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties()),
                    new ShapeProperties(
                        new DocumentFormat.OpenXml.Drawing.Transform2D(
                            new DocumentFormat.OpenXml.Drawing.Offset() { X = 50 * 9525, Y = 50 * 9525 },
                            new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 600 * 9525, Cy = 400 * 9525 })),
                    new TextBody(
                        new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                        new DocumentFormat.OpenXml.Drawing.ListStyle(),
                        new DocumentFormat.OpenXml.Drawing.Paragraph(
                            new DocumentFormat.OpenXml.Drawing.Run(
                                new DocumentFormat.OpenXml.Drawing.RunProperties() { Language = "en-US", FontSize = 1800 },
                                new DocumentFormat.OpenXml.Drawing.Text(extractedText))
                        ))
                );

                shapeTree.AppendChild(shape);

                slidePart.Slide.Save();
                presentationPart.Presentation.Save();
            }
        }

        private void ExportToPowerPoint_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(ultimulTextExtras))
            {
                MessageBox.Show("Vă rugăm să apăsați 'Procesează OCR' mai întâi.");
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "PowerPoint Presentation (*.pptx)|*.pptx";
            saveDialog.FileName = "Prezentare_Document.pptx";

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    ExportToPowerPoint(ultimulTextExtras, saveDialog.FileName);
                    MessageBox.Show("Prezentarea PowerPoint a fost salvată cu succes!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Eroare la export: " + ex.Message);
                }
            }
        }

        // Centralized display logic to render selected page and sync InkCanvas
        private void DisplayPage(int pageIndex)
        {
            if (_loadedDocument == null || pageIndex < 0 || pageIndex >= _loadedDocument.PageCount)
                return;

            using (var img = _loadedDocument.Render(pageIndex, 200, 200, true))
            {
                var bmp = ImageToBitmapSource(img);
                PdfDisplayImage.Source = bmp;

                // Sync canvas size with displayed page
                EraserCanvas.Width = bmp.Width;
                EraserCanvas.Height = bmp.Height;
                EraserCanvas.Strokes.Clear();
            }
        }

        // Handle selection change in the ListBox
        private void PagesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // ItemsControl does not have SelectedItem, but SelectionChangedEventArgs provides the selection
            if (e.AddedItems.Count > 0 && e.AddedItems[0] is PdfPageItem item)
            {
                // PageNumber is 1-based; convert to 0-based index
                int pageIndex = item.PageNumber - 1;
                DisplayPage(pageIndex);
            }
        }
    }
}