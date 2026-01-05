using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
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
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;

namespace EditarePDF
{
    public partial class MainWindow : Window
    {
        private string ultimulTextExtras = "";

        // Optional: UI element to show extracted text if present in XAML
        private TextBox? ExtractedTextBox;

        public MainWindow()
        {
            InitializeComponent();
            ExtractedTextBox = (TextBox)FindName("ExtractedTextBox");
        }

        private void OpenPdf_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                CheckFileExists = true,
                Multiselect = false
            };

            if (ofd.ShowDialog(this) == true)
            {
                // Load directly into Syncfusion viewer
                pdfViewer.Load(ofd.FileName);
                // Size overlay to viewer
                pdfViewer.SizeChanged += (_, __) =>
                {
                    EraserCanvas.Width = pdfViewer.ActualWidth;
                    EraserCanvas.Height = pdfViewer.ActualHeight;
                };
            }
        }

        private void SetupEraser()
        {
            DrawingAttributes da = new DrawingAttributes
            {
                Color = Colors.White,
                Height = 20,
                Width = 20,
                StylusTip = StylusTip.Rectangle
            };
            EraserCanvas.DefaultDrawingAttributes = da;
        }

        private void Eraser_Click(object sender, RoutedEventArgs e)
        {
            EraserCanvas.EditingMode = InkCanvasEditingMode.Ink;
            SetupEraser();
        }

        private void ConvertToWord_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Convertire Word funcționalitate nu este încă implementată.");
        }

        private void EditorCanvas_MouseDown(object sender, MouseButtonEventArgs e) { }
        private void EditorCanvas_MouseMove(object sender, MouseEventArgs e) { }

        private System.Drawing.Bitmap BitmapSourceToBitmap(BitmapSource srs)
        {
            var converted = srs.Format == PixelFormats.Bgra32
                ? srs
                : new FormatConvertedBitmap(srs, PixelFormats.Bgra32, null, 0);

            int width = converted.PixelWidth;
            int height = converted.PixelHeight;
            int stride = width * 4;

            IntPtr ptr = System.Runtime.InteropServices.Marshal.AllocHGlobal(height * stride);
            try
            {
                converted.CopyPixels(new Int32Rect(0, 0, width, height), ptr, height * stride, stride);
                using (var bmap = new System.Drawing.Bitmap(
                           width, height, stride,
                           System.Drawing.Imaging.PixelFormat.Format32bppArgb, ptr))
                {
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
                RenderTargetBitmap cleanBitmapSource = CreateCleanImage();
                using System.Drawing.Bitmap cleanBitmap = BitmapSourceToBitmap(cleanBitmapSource);
                string extractedText = ExtractTextFromImage(cleanBitmap);
                if (ExtractedTextBox != null)
                {
                    ExtractedTextBox.Text = extractedText;
                }
                ultimulTextExtras = extractedText;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la procesarea OCR: " + ex.Message);
            }
        }

        private RenderTargetBitmap CreateCleanImage()
        {
            // Snapshot the viewer + ink overlay
            int width = (int)Math.Max(1, Math.Round(ContainerGrid.ActualWidth));
            int height = (int)Math.Max(1, Math.Round(ContainerGrid.ActualHeight));

            var rtb = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(ContainerGrid);
            return rtb;
        }

        public string ExtractTextFromImage(System.Drawing.Bitmap cleanImage)
        {
            try
            {
                using (var engine = new TesseractEngine(@"./tessdata", "ron+eng", EngineMode.Default))
                using (var img = PixConverter.ToPix(cleanImage))
                using (var page = engine.Process(img))
                {
                    string text = page.GetText();
                    float confidence = page.GetMeanConfidence();
                    Console.WriteLine($"Precizie estimată: {confidence:P}");
                    return text;
                }
            }
            catch (Exception ex)
            {
                return "Eroare la procesarea OCR: " + ex.Message;
            }
        }

        public void ExportToWord(string extractedText, string filePath)
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p1 = doc.CreateParagraph();
            XWPFRun run = p1.CreateRun();

            if (extractedText != null)
            {
                string[] lines = extractedText.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.None);
                foreach (var line in lines)
                {
                    run.SetText(line);
                    run.AddCarriageReturn();
                }
            }

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
                string[] columns = System.Text.RegularExpressions.Regex.Split(lines[i], @"\t|\s{2,}");

                for (int j = 0; j < columns.Length; j++)
                {
                    NPOI.SS.UserModel.ICell cell = row.CreateCell(j);
                    cell.SetCellValue(columns[j].Trim());
                }
            }

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

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "Tabel_Document.xlsx"
            };

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

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Word Document (*.docx)|*.docx",
                FileName = "Document_Convertit.docx"
            };

            if (saveDialog.ShowDialog() == true)
            {
                ExportToWord(ultimulTextExtras, saveDialog.FileName);
                MessageBox.Show("Fișierul Word a fost salvat!");
            }
        }

        private void ExportToImage_Click(object sender, RoutedEventArgs e)
        {
            RenderTargetBitmap cleanBitmapSource = CreateCleanImage();

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "PNG Image (*.png)|*.png|JPEG Image (*.jpg)|*.jpg",
                FileName = "Document_Imagine"
            };

            if (saveDialog.ShowDialog() == true)
            {
                BitmapEncoder encoder;
                string ext = System.IO.Path.GetExtension(saveDialog.FileName).ToLowerInvariant();
                encoder = (ext == ".jpg" || ext == ".jpeg") ? new JpegBitmapEncoder() : new PngBitmapEncoder();

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

                slidePart.AddPart(slideLayoutPart);
                presentationPart.Presentation.SlideIdList = new SlideIdList(new SlideId() { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

                var commonSlideData = slidePart.Slide.CommonSlideData ?? new CommonSlideData(new ShapeTree());
                slidePart.Slide.CommonSlideData = commonSlideData;

                var shapeTree = commonSlideData.ShapeTree ?? new ShapeTree();
                commonSlideData.ShapeTree = shapeTree;

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

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "PowerPoint Presentation (*.pptx)|*.pptx",
                FileName = "Prezentare_Document.pptx"
            };

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

        private RenderTargetBitmap CreatePageComposite()
        {
            // Composite of viewer + ink overlay
            return CreateCleanImage();
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
            run.AddPicture(ms, (int)NPOI.XWPF.UserModel.PictureType.PNG, "page.png", 600 * 9525, 800 * 9525);
            using var fs = new FileStream(filePath, FileMode.Create);
            doc.Write(fs);
        }

        private void ExportErasedPageToPdf(string filePath)
        {
            var composite = CreatePageComposite();
            using var bmp = BitmapSourceToBitmap(composite);

            using var doc = new PdfDocument();
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

        private async void SaveAs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var composite = CreatePageComposite();

                if (string.IsNullOrWhiteSpace(ultimulTextExtras))
                {
                    await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Loaded);
                    using (var finalImage = BitmapSourceToBitmap(composite))
                    {
                        var text = ExtractTextFromImage(finalImage);
                        ultimulTextExtras = text;
                    }
                }

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

        private void SaveMenuButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.ContextMenu is ContextMenu cm)
            {
                cm.PlacementTarget = btn;
                cm.IsOpen = true;
            }
        }

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

        // Keep handler to satisfy XAML hook; navigating pages via Syncfusion can be added if needed.
        private void PagesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Intentionally left empty; integrate with pdfViewer page navigation if/when you populate thumbnails.
        }
    }
}