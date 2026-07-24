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
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EditarePDF
{
    public partial class MainWindow : Window
    {
        // Ink strokes captured per PDF page index (0-based), so the "eraser" survives page navigation
        // and can be re-applied when exporting each page.
        private readonly Dictionary<int, StrokeCollection> _pageStrokes = new();
        private int _currentPageIndex = -1;
        private bool _isNavigatingProgrammatically = false;

        public MainWindow()
        {
            InitializeComponent();

            // Size overlay to viewer (subscribed once, not per file opened)
            pdfViewer.SizeChanged += (_, __) =>
            {
                EraserCanvas.Width = pdfViewer.ActualWidth;
                EraserCanvas.Height = pdfViewer.ActualHeight;
            };

            pdfViewer.DocumentLoaded += PdfViewer_DocumentLoaded;
            pdfViewer.CurrentPageChanged += PdfViewer_CurrentPageChanged;
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
                try
                {
                    pdfViewer.Load(ofd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Nu s-a putut deschide fișierul PDF: " + ex.Message);
                }
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

        // --- Per-page navigation & ink tracking ---

        private void PdfViewer_DocumentLoaded(object? sender, EventArgs e)
        {
            _pageStrokes.Clear();
            EraserCanvas.Strokes.Clear();
            _currentPageIndex = pdfViewer.CurrentPageIndex;

            PagesList.Items.Clear();
            for (int i = 1; i <= pdfViewer.PageCount; i++)
            {
                PagesList.Items.Add($"Pagina {i}");
            }

            if (PagesList.Items.Count > 0)
            {
                _isNavigatingProgrammatically = true;
                PagesList.SelectedIndex = Math.Max(0, _currentPageIndex);
                _isNavigatingProgrammatically = false;
            }
        }

        private void PdfViewer_CurrentPageChanged(object? sender, EventArgs e)
        {
            SwitchToPage(pdfViewer.CurrentPageIndex);
        }

        private void SwitchToPage(int newIndex)
        {
            if (newIndex < 0 || newIndex == _currentPageIndex)
                return;

            if (_currentPageIndex >= 0)
            {
                _pageStrokes[_currentPageIndex] = new StrokeCollection(EraserCanvas.Strokes);
            }

            _currentPageIndex = newIndex;

            EraserCanvas.Strokes.Clear();
            if (_pageStrokes.TryGetValue(newIndex, out var strokes))
            {
                EraserCanvas.Strokes.Add(strokes);
            }

            if (newIndex < PagesList.Items.Count && PagesList.SelectedIndex != newIndex)
            {
                _isNavigatingProgrammatically = true;
                PagesList.SelectedIndex = newIndex;
                _isNavigatingProgrammatically = false;
            }
        }

        private void PagesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isNavigatingProgrammatically) return;

            int index = PagesList.SelectedIndex;
            if (index < 0) return;

            pdfViewer.GoToPageAtIndex(index);
            SwitchToPage(index);
        }

        // --- Rendering helpers ---

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

        private RenderTargetBitmap CreateCleanImage()
        {
            // Snapshot the viewer + ink overlay for the page currently shown
            int width = (int)Math.Max(1, Math.Round(ContainerGrid.ActualWidth));
            int height = (int)Math.Max(1, Math.Round(ContainerGrid.ActualHeight));

            var rtb = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(ContainerGrid);
            return rtb;
        }

        private RenderTargetBitmap CreatePageComposite() => CreateCleanImage();

        public string ExtractTextFromImage(System.Drawing.Bitmap cleanImage)
        {
            try
            {
                string tessDataPath = System.IO.Path.Combine(AppContext.BaseDirectory, "tessdata");
                using (var engine = new TesseractEngine(tessDataPath, "ron+eng", EngineMode.Default))
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

        // Walks every page of the loaded document, swapping in that page's stored ink strokes
        // before snapshotting it, so exports include the eraser marks for the right page.
        private async Task<List<(BitmapSource Image, string Text)>> CaptureAllPagesAsync()
        {
            IsEnabled = false;
            try
            {
                var results = new List<(BitmapSource, string)>();
                int pageCount = pdfViewer.PageCount;
                int originalIndex = _currentPageIndex;

                for (int i = 0; i < pageCount; i++)
                {
                    if (i != _currentPageIndex)
                    {
                        pdfViewer.GoToPageAtIndex(i);
                        SwitchToPage(i);
                        // Give the viewer a chance to finish rendering the newly navigated page
                        await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Background);
                        await Task.Delay(150);
                    }

                    var composite = CreateCleanImage();
                    composite.Freeze();

                    using var bmp = BitmapSourceToBitmap(composite);
                    string text = ExtractTextFromImage(bmp);

                    results.Add((composite, text));
                }

                if (originalIndex >= 0 && originalIndex != _currentPageIndex)
                {
                    pdfViewer.GoToPageAtIndex(originalIndex);
                    SwitchToPage(originalIndex);
                }

                return results;
            }
            finally
            {
                IsEnabled = true;
            }
        }

        // --- Multi-page exporters ---

        private void ExportAllPagesToPdf(IEnumerable<BitmapSource> pages, string filePath)
        {
            using var doc = new PdfDocument();
            foreach (var pageImg in pages)
            {
                using var bmp = BitmapSourceToBitmap(pageImg);
                var page = doc.AddPage();
                page.Width = XUnit.FromPoint(bmp.Width);
                page.Height = XUnit.FromPoint(bmp.Height);

                using var gfx = XGraphics.FromPdfPage(page);
                using var stream = new MemoryStream();
                bmp.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                stream.Position = 0;
                var img = XImage.FromStream(stream);
                gfx.DrawImage(img, 0, 0, page.Width.Point, page.Height.Point);
            }
            doc.Save(filePath);
        }

        private void ExportAllPagesToWord(List<(BitmapSource Image, string Text)> pages, string filePath)
        {
            var doc = new XWPFDocument();
            for (int i = 0; i < pages.Count; i++)
            {
                var p = doc.CreateParagraph();
                var run = p.CreateRun();
                run.SetText($"--- Pagina {i + 1} ---");
                run.AddCarriageReturn();

                string[] lines = pages[i].Text.Split(new[] { "\n", "\r\n" }, StringSplitOptions.None);
                foreach (var line in lines)
                {
                    run.SetText(line);
                    run.AddCarriageReturn();
                }

                if (i < pages.Count - 1)
                {
                    run.AddBreak(BreakType.PAGE);
                }
            }

            using var fs = new FileStream(filePath, FileMode.Create);
            doc.Write(fs);
        }

        private void ExportAllPagesToExcel(List<(BitmapSource Image, string Text)> pages, string filePath)
        {
            IWorkbook workbook = new XSSFWorkbook();
            for (int p = 0; p < pages.Count; p++)
            {
                ISheet sheet = workbook.CreateSheet($"Pagina {p + 1}");
                string[] lines = pages[p].Text.Split(new[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

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
            }

            using var fs = new FileStream(filePath, FileMode.Create);
            workbook.Write(fs);
        }

        private void ExportAllPagesToPowerPoint(List<(BitmapSource Image, string Text)> pages, string filePath)
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

                var slideIds = new SlideIdList();
                uint slideId = 256;

                foreach (var page in pages)
                {
                    SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                    slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));
                    slidePart.AddPart(slideLayoutPart);

                    slideIds.AppendChild(new SlideId() { Id = slideId++, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

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
                                new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 8500000, Cy = 6000000 })),
                        new TextBody(
                            new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                            new DocumentFormat.OpenXml.Drawing.ListStyle(),
                            new DocumentFormat.OpenXml.Drawing.Paragraph(
                                new DocumentFormat.OpenXml.Drawing.Run(
                                    new DocumentFormat.OpenXml.Drawing.RunProperties() { Language = "en-US", FontSize = 1800 },
                                    new DocumentFormat.OpenXml.Drawing.Text(page.Text))
                            ))
                    );

                    shapeTree.AppendChild(shape);
                    slidePart.Slide.Save();
                }

                presentationPart.Presentation.SlideIdList = slideIds;
                presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
                    new SlideMasterId() { Id = 2147483648, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) });
                presentationPart.Presentation.SlideSize = new SlideSize() { Cx = 9144000, Cy = 6858000 };
                presentationPart.Presentation.NotesSize = new NotesSize() { Cx = 6858000, Cy = 9144000 };
                presentationPart.Presentation.Save();
            }
        }

        // --- Save As handlers (wired from the toolbar's "Salvare ca..." menu) ---

        private void SaveMenuButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.ContextMenu is ContextMenu cm)
            {
                cm.PlacementTarget = btn;
                cm.IsOpen = true;
            }
        }

        private bool EnsureDocumentLoaded()
        {
            if (pdfViewer.PageCount <= 0)
            {
                MessageBox.Show("Deschideți mai întâi un PDF.");
                return false;
            }
            return true;
        }

        private async void SaveAsPdf_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDocumentLoaded()) return;

            var sfd = new SaveFileDialog
            {
                Title = "Save As PDF",
                Filter = "PDF (imagine, *.pdf)|*.pdf",
                FileName = "Document.pdf"
            };
            if (sfd.ShowDialog() != true) return;

            try
            {
                var pages = await CaptureAllPagesAsync();
                ExportAllPagesToPdf(pages.Select(p => p.Image), sfd.FileName);
                MessageBox.Show("PDF salvat!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la salvare: " + ex.Message);
            }
        }

        private async void SaveAsDocx_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDocumentLoaded()) return;

            var sfd = new SaveFileDialog
            {
                Title = "Save As Word",
                Filter = "Word Document (*.docx)|*.docx",
                FileName = "Document_Convertit.docx"
            };
            if (sfd.ShowDialog() != true) return;

            try
            {
                var pages = await CaptureAllPagesAsync();
                ExportAllPagesToWord(pages, sfd.FileName);
                MessageBox.Show("Fișierul Word a fost salvat!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la salvare: " + ex.Message);
            }
        }

        private async void SaveAsXlsx_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDocumentLoaded()) return;

            var sfd = new SaveFileDialog
            {
                Title = "Save As Excel",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "Tabel_Document.xlsx"
            };
            if (sfd.ShowDialog() != true) return;

            try
            {
                var pages = await CaptureAllPagesAsync();
                ExportAllPagesToExcel(pages, sfd.FileName);
                MessageBox.Show("Fișierul Excel a fost salvat!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la salvare: " + ex.Message);
            }
        }

        private async void SaveAsPptx_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDocumentLoaded()) return;

            var sfd = new SaveFileDialog
            {
                Title = "Save As PowerPoint",
                Filter = "PowerPoint Presentation (*.pptx)|*.pptx",
                FileName = "Prezentare_Document.pptx"
            };
            if (sfd.ShowDialog() != true) return;

            try
            {
                var pages = await CaptureAllPagesAsync();
                ExportAllPagesToPowerPoint(pages, sfd.FileName);
                MessageBox.Show("Prezentarea PowerPoint a fost salvată!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la export: " + ex.Message);
            }
        }

        private void SaveAsPng_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDocumentLoaded()) return;

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
            if (!EnsureDocumentLoaded()) return;

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
    }
}
