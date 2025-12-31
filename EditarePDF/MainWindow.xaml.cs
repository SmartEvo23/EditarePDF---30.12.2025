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
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

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
        private PdfiumViewer.PdfDocument? _loadedDocument;

        // Add this field to your MainWindow class if it doesn't exist
        private TextBox ExtractedTextBox;

        // Keep track of 1:1 scale based on DPI
        private double _actualScale = 1.0;

        // Add field to track current page
        private int _currentPageIndex = -1;

        public MainWindow()
        {
            InitializeComponent();
            PagesList.ItemsSource = _pages;
            ExtractedTextBox = (TextBox)this.FindName("ExtractedTextBox");
            Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Default to 100%
            SetZoom(1.0);
            SelectZoomComboItemByScale(1.0);
            UpdatePageSizeStatus();
        }

        private void OpenPdf_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PDF files (*.pdf)|*.pdf";

            if (openFileDialog.ShowDialog() == true)
            {
                // 1. Încărcăm documentul PDF offline
                _loadedDocument?.Dispose();
                _loadedDocument = PdfiumViewer.PdfDocument.Load(openFileDialog.FileName);

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

        private System.Drawing.Bitmap BitmapSourceToBitmap(BitmapSource srs)
        {
            // Ensure a supported pixel format for GDI+ and Tesseract (32bpp ARGB, not PArgb)
            var converted = srs.Format == PixelFormats.Bgra32
                ? srs
                : new FormatConvertedBitmap(srs, PixelFormats.Bgra32, null, 0);

            int width = converted.PixelWidth;
            int height = converted.PixelHeight;
            int stride = width * 4; // BGRA32

            IntPtr ptr = System.Runtime.InteropServices.Marshal.AllocHGlobal(height * stride);
            try
            {
                converted.CopyPixels(new System.Windows.Int32Rect(0, 0, width, height), ptr, height * stride, stride);

                using (var bmap = new System.Drawing.Bitmap(
                           width, height, stride,
                           System.Drawing.Imaging.PixelFormat.Format32bppArgb, ptr)) // <- Use Argb, not PArgb
                {
                    // Copy to a managed bitmap that owns its memory
                    return new System.Drawing.Bitmap(bmap);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FreeHGlobal(ptr);
            }
        }
        private void ProcessOcr_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Creăm o imagine curată a zonei de randare
                RenderTargetBitmap cleanBitmapSource = CreateCleanImage();
                // 2. Convertim RenderTargetBitmap în System.Drawing.Bitmap
                using System.Drawing.Bitmap cleanBitmap = BitmapSourceToBitmap(cleanBitmapSource);
                // 3. Extragem textul folosind Tesseract OCR
                string extractedText = ExtractTextFromImage(cleanBitmap);
                // 4. Afișăm textul extras în TextBox-ul din UI
                    ExtractedTextBox.Text = extractedText;
                // 5. Salvăm textul extras pentru utilizare ulterioară
                ultimulTextExtras = extractedText;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la procesarea OCR: " + ex.Message);
            }
        }
        private RenderTargetBitmap CreateCleanImage()
        {
            // Prefer the container grid size (or PdfDisplayImage if that’s your target)
            int width = (int)Math.Round(ContainerGrid.ActualWidth);
            int height = (int)Math.Round(ContainerGrid.ActualHeight);

            if (width <= 0 || height <= 0)
                throw new InvalidOperationException("Zona de randare nu este pregătită. Asigurați-vă că o pagină PDF este selectată și afișată.");

            var renderBitmap = new RenderTargetBitmap(width, height, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
            renderBitmap.Render(ContainerGrid);
            return renderBitmap;
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
            _currentPageIndex = pageIndex;
            RenderAndDisplayPage(pageIndex);
        }

        private void ReRenderCurrentPage()
        {
            if (_loadedDocument == null || _currentPageIndex < 0) return;
            RenderAndDisplayPage(_currentPageIndex);
        }

        // Render at DPI that matches current zoom for crisp display
        private void RenderAndDisplayPage(int pageIndex)
        {
            if (_loadedDocument == null || pageIndex < 0 || pageIndex >= _loadedDocument.PageCount) return;

            // Base DPI for good quality; scale with current zoom
            double baseDpi = 150; // change to 300 for very high quality
            double scale = ZoomTransform.ScaleX <= 0 ? 1.0 : ZoomTransform.ScaleX;
            int dpiX = (int)Math.Round(baseDpi * scale);
            int dpiY = (int)Math.Round(baseDpi * scale);

            using (var img = _loadedDocument.Render(pageIndex, dpiX, dpiY, true))
            {
                var bmp = ImageToBitmapSource(img);

                // Ensure best scaling quality in WPF
                RenderOptions.SetBitmapScalingMode(PdfDisplayImage, BitmapScalingMode.HighQuality);
                PdfDisplayImage.Source = bmp;

                // Match InkCanvas to image pixel size
                EraserCanvas.Width = bmp.Width;
                EraserCanvas.Height = bmp.Height;
                // EraserCanvas.Strokes.Freeze(); // minor perf improvement for static strokes
            }

            UpdatePageSizeStatus();
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

        private RenderTargetBitmap CreatePageComposite()
        {
            if (PdfDisplayImage.Source is not BitmapSource pageBmp)
                throw new InvalidOperationException("No page loaded.");

            int width = pageBmp.PixelWidth;
            int height = pageBmp.PixelHeight;

            // Ensure both Image and InkCanvas are laid out to the same pixel size
            PdfDisplayImage.Width = width;
            PdfDisplayImage.Height = height;
            EraserCanvas.Width = width;
            EraserCanvas.Height = height;

            // Arrange a temporary grid to render exactly the page + strokes
            var grid = new Grid { Width = width, Height = height };
            var img = new Image { Source = pageBmp, Stretch = Stretch.None, Width = width, Height = height };
            var ink = new InkCanvas { Background = Brushes.Transparent, Width = width, Height = height };
            // Copy existing strokes
            foreach (var s in EraserCanvas.Strokes)
                ink.Strokes.Add(s.Clone());

            grid.Children.Add(img);
            grid.Children.Add(ink);

            // Measure/arrange before render
            grid.Measure(new Size(width, height));
            grid.Arrange(new System.Windows.Rect(0, 0, width, height));
            grid.UpdateLayout();

            var rtb = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(grid);
            return rtb;
        }

        private void ExportErasedPageToWord(string filePath)
        {
            var composite = CreatePageComposite();
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(composite));
            using var ms = new MemoryStream();
            encoder.Save(ms);

            var doc = new XWPFDocument();
            var p = doc.CreateParagraph();
            var run = p.CreateRun();
            ms.Position = 0;
            run.AddPicture(ms, (int)NPOI.XWPF.UserModel.PictureType.PNG, "page.png", 600 * 9525, 800 * 9525); // adjust size
            using var fs = new FileStream(filePath, FileMode.Create);
            doc.Write(fs);
        }

        private void ExportErasedPageToPdf(string filePath)
        {
            var composite = CreatePageComposite();
            // Convert WPF BitmapSource to a System.Drawing.Bitmap first
            using var bmp = BitmapSourceToBitmap(composite);

            using var doc = new PdfSharp.Pdf.PdfDocument();
            var page = doc.AddPage();
            page.Width = XUnit.FromPoint(bmp.Width);
            page.Height = XUnit.FromPoint(bmp.Height);

            using var gfx = XGraphics.FromPdfPage(page);
            using var stream = new MemoryStream();
            bmp.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
            stream.Position = 0;
            var img = XImage.FromStream(stream);
            gfx.DrawImage(img, 0, 0, page.Width.Point, page.Height.Point);

            doc.Save(filePath);
        }

        // Add a unified Save As handler that performs OCR (if needed) and exports accordingly.
        private async void SaveAs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Ensure a page is loaded
                if (PdfDisplayImage.Source is not BitmapSource)
                {
                    MessageBox.Show("Nici o pagină PDF nu este încărcată sau afișată.");
                    return;
                }

                // Render the composite (page + eraser strokes)
                var composite = CreatePageComposite();

                // Run OCR if we don't already have text (FineReader-style: process before save)
                if (string.IsNullOrWhiteSpace(ultimulTextExtras))
                {
                    await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Loaded);
                    using (var finalImage = BitmapSourceToBitmap(composite))
                    {
                        // NOTE: fix Tesseract setup per earlier guidance if this crashes.
                        var text = ExtractTextFromImage(finalImage);
                        ultimulTextExtras = text;
                    }
                }

                // Unified Save As dialog, user picks output format in filter
                var saveDialog = new SaveFileDialog
                {
                    Title = "Save As",
                    FileName = "Document",
                    Filter =
                        "Searchable PDF (*.pdf)|*.pdf|" +
                        "Word Document (*.docx)|*.docx|" +
                        "Excel Workbook (*.xlsx)|*.xlsx|" +
                        "PowerPoint Presentation (*.pptx)|*.pptx|" +
                        "PNG Image (*.png)|*.png|" +
                        "JPEG Image (*.jpg;*.jpeg)|*.jpg;*.jpeg"
                };

                if (saveDialog.ShowDialog() != true)
                    return;

                var ext = System.IO.Path.GetExtension(saveDialog.FileName).ToLowerInvariant();
                switch (ext)
                {
                    case ".pdf":
                        // Export the erased page image to PDF (baseline). Optionally add a text layer later.
                        ExportErasedPageToPdf(saveDialog.FileName);
                        MessageBox.Show("PDF salvat!");
                        break;

                    case ".docx":
                        ExportToWord(ultimulTextExtras, saveDialog.FileName);
                        MessageBox.Show("Fișierul Word a fost salvat!");
                        break;

                    case ".xlsx":
                        ExportToExcel(ultimulTextExtras, saveDialog.FileName);
                        MessageBox.Show("Fișierul Excel a fost salvat!");
                        break;

                    case ".pptx":
                        ExportToPowerPoint(ultimulTextExtras, saveDialog.FileName);
                        MessageBox.Show("Prezentarea PowerPoint a fost salvată!");
                        break;

                    case ".png":
                    case ".jpg":
                    case ".jpeg":
                    {
                        BitmapEncoder encoder = ext is ".jpg" or ".jpeg"
                            ? new JpegBitmapEncoder()
                            : new PngBitmapEncoder();

                        encoder.Frames.Add(BitmapFrame.Create(composite));
                        using var fs = new FileStream(saveDialog.FileName, FileMode.Create);
                        encoder.Save(fs);
                        MessageBox.Show("Imaginea a fost salvată!");
                        break;
                    }

                    default:
                        MessageBox.Show("Format de fișier neacceptat.");
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la 'Save As': " + ex.Message);
            }
        }

        // Show the context menu directly under the button when clicked
        private void SaveMenuButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.ContextMenu is ContextMenu cm)
            {
                cm.PlacementTarget = btn;
                cm.IsOpen = true;
            }
        }

        // Each format prompts for a location using SaveFileDialog
        private void SaveAsPdf_Click(object sender, RoutedEventArgs e)
        {
            var sfd = new SaveFileDialog
            {
                Title = "Save As PDF",
                Filter = "Searchable PDF (*.pdf)|*.pdf",
                FileName = "Document.pdf"
            };
            if (sfd.ShowDialog() == true)
            {
                ExportErasedPageToPdf(sfd.FileName);
                MessageBox.Show("PDF salvat!");
            }
        }

        private void SaveAsDocx_Click(object sender, RoutedEventArgs e)
        {
            EnsureOcrText();
            var sfd = new SaveFileDialog
            {
                Title = "Save As Word",
                Filter = "Word Document (*.docx)|*.docx",
                FileName = "Document_Convertit.docx"
            };
            if (sfd.ShowDialog() == true)
            {
                ExportToWord(ultimulTextExtras, sfd.FileName);
                MessageBox.Show("Fișierul Word a fost salvat!");
            }
        }

        private void SaveAsXlsx_Click(object sender, RoutedEventArgs e)
        {
            EnsureOcrText();
            var sfd = new SaveFileDialog
            {
                Title = "Save As Excel",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "Tabel_Document.xlsx"
            };
            if (sfd.ShowDialog() == true)
            {
                ExportToExcel(ultimulTextExtras, sfd.FileName);
                MessageBox.Show("Fișierul Excel a fost salvat!");
            }
        }

        private void SaveAsPptx_Click(object sender, RoutedEventArgs e)
        {
            EnsureOcrText();
            var sfd = new SaveFileDialog
            {
                Title = "Save As PowerPoint",
                Filter = "PowerPoint Presentation (*.pptx)|*.pptx",
                FileName = "Prezentare_Document.pptx"
            };
            if (sfd.ShowDialog() == true)
            {
                ExportToPowerPoint(ultimulTextExtras, sfd.FileName);
                MessageBox.Show("Prezentarea PowerPoint a fost salvată!");
            }
        }

        private void SaveAsPng_Click(object sender, RoutedEventArgs e)
        {
            var composite = CreatePageComposite();
            var sfd = new SaveFileDialog
            {
                Title = "Save As PNG",
                Filter = "PNG Image (*.png)|*.png",
                FileName = "Document_Imagine.png"
            };
            if (sfd.ShowDialog() == true)
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(composite));
                using var fs = new FileStream(sfd.FileName, FileMode.Create);
                encoder.Save(fs);
                MessageBox.Show("Imaginea a fost salvată!");
            }
        }

        private void SaveAsJpeg_Click(object sender, RoutedEventArgs e)
        {
            var composite = CreatePageComposite();
            var sfd = new SaveFileDialog
            {
                Title = "Save As JPEG",
                Filter = "JPEG Image (*.jpg;*.jpeg)|*.jpg;*.jpeg",
                FileName = "Document_Imagine.jpg"
            };
            if (sfd.ShowDialog() == true)
            {
                var encoder = new JpegBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(composite));
                using var fs = new FileStream(sfd.FileName, FileMode.Create);
                encoder.Save(fs);
                MessageBox.Show("Imaginea a fost salvată!");
            }
        }

        // Ensure OCR text exists before exporting text formats
        private void EnsureOcrText()
        {
            if (string.IsNullOrWhiteSpace(ultimulTextExtras))
            {
                var composite = CreatePageComposite();
                using var finalImage = BitmapSourceToBitmap(composite);
                ultimulTextExtras = ExtractTextFromImage(finalImage);
                if (ExtractedTextBox != null)
                    ExtractedTextBox.Text = ultimulTextExtras;
            }
        }

        // Call this after loading a PDF page image into PdfDisplayImage.Source
        private void OnPdfImageChanged()
        {
            // Ensure InkCanvas matches the image pixel size for 1:1 alignment
            if (PdfDisplayImage.Source is BitmapSource bmp)
            {
                // Physical pixel dimensions
                EraserCanvas.Width = bmp.PixelWidth;
                EraserCanvas.Height = bmp.PixelHeight;

                // Compute 1:1 scale based on image DPI vs 96 WPF units
                var dpiX = bmp.DpiX > 0 ? bmp.DpiX : 96.0;
                var dpiY = bmp.DpiY > 0 ? bmp.DpiY : 96.0;

                // WPF device-independent units: 1 unit = 1/96 inch.
                // To display physical pixels 1:1, scale DIU to pixel ratio.
                var scaleX = dpiX / 96.0;
                var scaleY = dpiY / 96.0;

                // Use uniform scale (usually equal)
                _actualScale = Math.Max(scaleX, scaleY);
                SetZoom(_actualScale);
                SelectZoomComboItemByScale(1.0); // Show 100% as logical zoom; actual is accounted in baseline
                UpdatePageSizeStatus();
            }
        }

        private void SetZoom(double scale)
        {
            ZoomTransform.ScaleX = scale;
            ZoomTransform.ScaleY = scale;
            ReRenderCurrentPage();
            UpdatePageSizeStatus();
        }

        private void SelectZoomComboItemByScale(double scale)
        {
            // Tries to select the closest preset
            double[] presets = { 0.5, 0.75, 1.0, 1.25, 1.5, 2.0 };
            double closest = 1.0;
            double minDiff = double.MaxValue;
            foreach (var p in presets)
            {
                var d = Math.Abs(p - scale);
                if (d < minDiff)
                {
                    minDiff = d;
                    closest = p;
                }
            }

            foreach (var item in ZoomCombo.Items)
            {
                if (item is ComboBoxItem cbi && double.TryParse(cbi.Tag?.ToString(), out var tagScale))
                {
                    if (Math.Abs(tagScale - closest) < 0.0001)
                    {
                        ZoomCombo.SelectedItem = cbi;
                        break;
                    }
                }
            }
        }

        private void ZoomCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ZoomCombo.SelectedItem is ComboBoxItem cbi &&
                double.TryParse(cbi.Tag?.ToString(), out var requestedScale))
            {
                SetZoom(requestedScale);
            }
        }

        private void ActualSize_Click(object sender, RoutedEventArgs e)
        {
            // Restore true 1:1 pixel display based on image DPI baseline
            SetZoom(_actualScale);
            UpdatePageSizeStatus();
        }

        private void UpdatePageSizeStatus()
        {
            // Show percentage and page size info
            double percent = ZoomTransform.ScaleX * 100.0;
            string sizeText = "";

            if (PdfDisplayImage.Source is BitmapSource bmp)
            {
                var pxW = bmp.PixelWidth;
                var pxH = bmp.PixelHeight;
                var dpiX = bmp.DpiX > 0 ? bmp.DpiX : 96.0;
                var dpiY = bmp.DpiY > 0 ? bmp.DpiY : 96.0;

                // Physical size in inches
                double inchesW = pxW / dpiX;
                double inchesH = pxH / dpiY;

                // Convert to millimeters for readability
                double mmW = inchesW * 25.4;
                double mmH = inchesH * 25.4;

                sizeText = $" | Page: {pxW}×{pxH}px ({mmW:F1}×{mmH:F1} mm @ {dpiX:F0}×{dpiY:F0} DPI)";
            }

            PageSizeText.Text = $"{percent:F0}%{sizeText}";
        }

        // Example: wherever you set the PDF page as an image source
        private void SetPdfPageImage(BitmapSource bitmap)
        {
            PdfDisplayImage.Source = bitmap;

            // Ensure the Image control uses pixel-aligned presentation
            RenderOptions.SetBitmapScalingMode(PdfDisplayImage, BitmapScalingMode.NearestNeighbor);
            OnPdfImageChanged();
        }

        private void FitToWidth_Click(object sender, RoutedEventArgs e)
        {
            if (PdfDisplayImage.Source is BitmapSource bmp && ContainerGrid.ActualWidth > 0)
            {
                // Compute scale so the page width fits the container width
                double containerWidth = ContainerGrid.ActualWidth;
                double scale = containerWidth / bmp.PixelWidth;
                SetZoom(scale);
                UpdatePageSizeStatus();
            }
        }

        private void BestFit_Click(object sender, RoutedEventArgs e)
        {
            if (PdfDisplayImage.Source is BitmapSource bmp && ContainerGrid.ActualWidth > 0 && ContainerGrid.ActualHeight > 0)
            {
                // Compute scale so the page fits both width and height (uniform)
                double scaleX = ContainerGrid.ActualWidth / bmp.PixelWidth;
                double scaleY = ContainerGrid.ActualHeight / bmp.PixelHeight;
                double scale = Math.Min(scaleX, scaleY);
                SetZoom(scale);
                UpdatePageSizeStatus();
            }
        }
    }
}