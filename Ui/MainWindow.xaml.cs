using AForge.Video;
using AForge.Video.DirectShow;
using Microsoft.Win32;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using Tesseract;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Controls.Primitives;
using PdfSharp.Drawing;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using PdfSharp.Pdf;
using System.Windows.Shapes;
using AForge.Imaging.Filters;

namespace Ui
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private FilterInfoCollection videoDevices;
        private VideoCaptureDevice videoSource;
        private Bitmap currentImage;

        public MainWindow()
        {
            InitializeComponent();
            InitializeWebcam();
        }

        private void InitializeWebcam()
        {
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            if (videoDevices.Count == 0)
            {
                BtnCaptureWebcam.IsEnabled = false;
                TxtStatus.Text = "No webcam detected.";
            }
        }

        private void BtnCaptureWebcam_Click(object sender, RoutedEventArgs e)
        {
            if (videoDevices.Count == 0) return;

            videoSource = new VideoCaptureDevice(videoDevices[0].MonikerString);
            videoSource.NewFrame += VideoSource_NewFrame;
            videoSource.Start();
            TxtStatus.Text = "Webcam started. Click again to capture.";
            BtnCaptureWebcam.Click -= BtnCaptureWebcam_Click;
            BtnCaptureWebcam.Click += BtnCaptureFrame_Click;
        }

        private void BtnCaptureFrame_Click(object sender, RoutedEventArgs e)
        {
            if (videoSource != null && videoSource.IsRunning)
            {
                videoSource.SignalToStop();
                videoSource.WaitForStop();
                BtnCaptureWebcam.Click -= BtnCaptureFrame_Click;
                BtnCaptureWebcam.Click += BtnCaptureWebcam_Click;
                TxtStatus.Text = "Image captured from webcam.";
            }
        }

        private void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Dispatcher.Invoke(() =>
            {
                currentImage = (Bitmap)eventArgs.Frame.Clone();
                ImgPreview.Source = BitmapToImageSource(currentImage);
            });
        }

        private void BtnSelectImage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                currentImage = new Bitmap(openFileDialog.FileName);
                ImgPreview.Source = BitmapToImageSource(currentImage);
                TxtStatus.Text = "Image selected.";
            }
        }

        private void BtnScan_Click(object sender, RoutedEventArgs e)
        {
            if (currentImage == null)
            {
                MessageBox.Show("Please select or capture an image first.");
                return;
            }

            try
            {
                // Lưu Bitmap thành tệp tạm thời
                string tempImagePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ocr_temp_image.png");
                currentImage.Save(tempImagePath, System.Drawing.Imaging.ImageFormat.Png);

                // Xác định đường dẫn tuyệt đối đến thư mục tessdata
                string tessdataPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");

                // Kiểm tra xem thư mục tessdata có tồn tại không
                if (!Directory.Exists(tessdataPath))
                {
                    MessageBox.Show($"Thư mục tessdata không tồn tại tại: {tessdataPath}");
                    TxtStatus.Text = "Scan failed.";
                    return;
                }

                // Sử dụng Tesseract để xử lý từ tệp
                using (var engine = new TesseractEngine(tessdataPath, "vie", EngineMode.Default))
                {
                    using (var img = Pix.LoadFromFile(tempImagePath))
                    {
                        using (var page = engine.Process(img))
                        {
                            // Lấy HOCR để phân tích cấu trúc
                            string hocrText = page.GetHOCRText(0);
                            ProcessHocr(hocrText);
                            TxtStatus.Text = "Scan completed.";
                        }
                    }
                }

                // Xóa tệp tạm thời
                File.Delete(tempImagePath);
            }
            catch (TesseractException tex)
            {
                MessageBox.Show($"Tesseract Error: {tex.Message}\nStack Trace: {tex.StackTrace}");
                TxtStatus.Text = "Scan failed.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"General Error: {ex.Message}\nStack Trace: {ex.StackTrace}");
                TxtStatus.Text = "Scan failed.";
            }
        }

        private void ProcessHocr(string hocrText)
        {
            RtbOcrResult.Document.Blocks.Clear();
            var hocrDoc = XDocument.Parse(hocrText);

            // Lấy tất cả các phần tử ocr_line (dòng văn bản)
            var lines = hocrDoc.Descendants("span")
                .Where(x => x.Attribute("class")?.Value == "ocr_line")
                .Select(line => new
                {
                    Text = line.Descendants("span")
                        .Where(w => w.Attribute("class")?.Value == "ocrx_word")
                        .Select(w => w.Value)
                        .Aggregate((a, b) => a + " " + b),
                    Title = line.Attribute("title")?.Value
                })
                .ToList();

            // Phân tích dòng để xác định tiêu đề, danh sách, đoạn văn
            Paragraph currentParagraph = null;
            bool isListItem = false;

            foreach (var line in lines)
            {
                string text = line.Text.Trim();
                if (string.IsNullOrWhiteSpace(text)) continue;

                // Lấy thông tin tọa độ và kích thước từ title
                var match = Regex.Match(line.Title, @"bbox (\d+) (\d+) (\d+) (\d+);.*?baseline.*?(\d+\.?\d*)?");
                if (!match.Success) continue;

                int x = int.Parse(match.Groups[1].Value);
                int y = int.Parse(match.Groups[2].Value);
                int width = int.Parse(match.Groups[3].Value) - x;
                int height = int.Parse(match.Groups[4].Value) - y;
                float fontSize = height > 0 ? height : 12;

                // Xác định loại dòng dựa trên nội dung và vị trí
                Run run = new Run(text);
                run.FontSize = fontSize;

                // Tiêu đề: Dòng đầu tiên, chữ lớn, canh giữa
                if (y < 100 && fontSize > 16)
                {
                    run.FontWeight = FontWeights.Bold;
                    currentParagraph = new Paragraph(run)
                    {
                        TextAlignment = TextAlignment.Center,
                        Margin = new Thickness(0, 10, 0, 10)
                    };
                }
                // Danh sách gạch đầu dòng hoặc số thứ tự
                else if (text.StartsWith("◊") || text.StartsWith("♦") || Regex.IsMatch(text, @"^(Bước \d+:|\d+\.)"))
                {
                    run.FontWeight = FontWeights.Bold;
                    currentParagraph = new Paragraph(run)
                    {
                        Margin = new Thickness(20, 5, 0, 5)
                    };
                    isListItem = true;
                }
                // Đoạn văn thường
                else
                {
                    if (isListItem && x > 50)
                    {
                        // Tiếp tục danh sách
                        run.FontWeight = FontWeights.Normal;
                        currentParagraph.Inlines.Add(new LineBreak());
                        currentParagraph.Inlines.Add(run);
                    }
                    else
                    {
                        isListItem = false;
                        currentParagraph = new Paragraph(run)
                        {
                            Margin = new Thickness(0, 5, 0, 5)
                        };
                    }
                }

                RtbOcrResult.Document.Blocks.Add(currentParagraph);
            }
        }

        private void BtnExportPdf_Click(object sender, RoutedEventArgs e)
        {
            TextRange textRange = new TextRange(RtbOcrResult.Document.ContentStart, RtbOcrResult.Document.ContentEnd);
            if (string.IsNullOrWhiteSpace(textRange.Text.Trim()))
            {
                MessageBox.Show("No text to export.");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                FileName = "ScannedDocument.pdf"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    // Đảm bảo thư mục đích có quyền ghi
                    string directory = System.IO.Path.GetDirectoryName(saveFileDialog.FileName);
                    if (!Directory.Exists(directory))
                    {
                        MessageBox.Show("Destination directory does not exist.");
                        TxtStatus.Text = "PDF export failed.";
                        return;
                    }

                    // Tạo PDF với PdfSharp
                    using (var document = new PdfDocument())
                    {
                        var page = document.AddPage();
                        page.Width = 595; // A4 width in points
                        page.Height = 842; // A4 height in points

                        using (var gfx = XGraphics.FromPdfPage(page))
                        {
                            double yPosition = 20; // Vị trí bắt đầu
                            double xMargin = 20;

                            foreach (Block block in RtbOcrResult.Document.Blocks)
                            {
                                if (block is Paragraph paragraph)
                                {
                                    // Lấy canh lề đoạn
                                    TextAlignment alignment = paragraph.TextAlignment;
                                    XStringFormat stringFormat = new XStringFormat();
                                    switch (alignment)
                                    {
                                        case TextAlignment.Left:
                                            stringFormat.Alignment = XStringAlignment.Near;
                                            break;
                                        case TextAlignment.Center:
                                            stringFormat.Alignment = XStringAlignment.Center;
                                            break;
                                        case TextAlignment.Right:
                                            stringFormat.Alignment = XStringAlignment.Far;
                                            break;
                                        case TextAlignment.Justify:
                                            stringFormat.Alignment = XStringAlignment.Near; // PdfSharp không hỗ trợ justify
                                            break;
                                    }

                                    foreach (Inline inline in paragraph.Inlines)
                                    {
                                        if (inline is Run run)
                                        {
                                            // Lấy định dạng từ Run
                                            string text = run.Text;
                                            string fontFamily = run.FontFamily.ToString();
                                            double fontSize = run.FontSize;
                                            bool isBold = run.FontWeight == FontWeights.Bold;
                                            bool isItalic = run.FontStyle == FontStyles.Italic;
                                            var color = run.Foreground as SolidColorBrush;

                                            // Tạo font PdfSharp
                                            var fontStyle = XFontStyleEx.Regular;
                                            if (isBold) fontStyle |= XFontStyleEx.Bold;
                                            if (isItalic) fontStyle |= XFontStyleEx.Italic;

                                            // Sử dụng Times New Roman để hỗ trợ tiếng Việt
                                            fontFamily = "Times New Roman";
                                            var font = new XFont(fontFamily, fontSize, fontStyle);

                                            // Chuyển đổi màu WPF sang PdfSharp
                                            var xBrush = XBrushes.Black;
                                            if (color != null)
                                            {
                                                xBrush = new XSolidBrush(XColor.FromArgb(
                                                    color.Color.A, color.Color.R, color.Color.G, color.Color.B));
                                            }

                                            // Vẽ văn bản với canh lề
                                            gfx.DrawString(text, font, xBrush, new XRect(xMargin, yPosition, page.Width - 2 * xMargin, fontSize * 1.2), stringFormat);

                                            // Cập nhật vị trí y
                                            yPosition += fontSize * 1.2;
                                        }
                                    }
                                    // Thêm khoảng cách giữa các đoạn
                                    yPosition += 10;
                                }
                            }
                        }

                        document.Save(saveFileDialog.FileName);
                    }

                    TxtStatus.Text = "PDF exported successfully.";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"PDF Error: {ex.Message}\nStack Trace: {ex.StackTrace}");
                    TxtStatus.Text = "PDF export failed.";
                }
            }
        }

        private void BtnAddPictureToPdf_Click(object sender, RoutedEventArgs e)
        {
            if (currentImage == null)
            {
                MessageBox.Show("No image to add. Please select or capture an image first.");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                FileName = "ScannedDocumentWithImage.pdf"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    // Đảm bảo thư mục đích có quyền ghi
                    string directory = System.IO.Path.GetDirectoryName(saveFileDialog.FileName);
                    if (!Directory.Exists(directory))
                    {
                        MessageBox.Show("Destination directory does not exist.");
                        TxtStatus.Text = "PDF export failed.";
                        return;
                    }

                    // Xử lý hình ảnh để làm rõ chữ
                    Bitmap processedImage = ProcessImage(currentImage);

                    // Tạo PDF với PdfSharp
                    using (var document = new PdfDocument())
                    {
                        var page = document.AddPage();
                        page.Width = 595; // A4 width in points
                        page.Height = 842; // A4 height in points

                        using (var gfx = XGraphics.FromPdfPage(page))
                        {
                            // Chèn hình ảnh vào PDF
                            using (MemoryStream imageStream = new MemoryStream())
                            {
                                processedImage.Save(imageStream, System.Drawing.Imaging.ImageFormat.Png);
                                imageStream.Position = 0;
                                XImage xImage = XImage.FromStream(imageStream);

                                // Tính toán kích thước hình ảnh để vừa trang
                                double imgWidth = xImage.PixelWidth;
                                double imgHeight = xImage.PixelHeight;
                                double scale = Math.Min((page.Width - 40) / imgWidth, 1000 / imgHeight);
                                imgWidth *= scale;
                                imgHeight *= scale;

                                gfx.DrawImage(xImage, 20, 20, imgWidth, imgHeight);
                                double yPosition = 20 + imgHeight + 20; // Vị trí bắt đầu văn bản

                                // Chèn văn bản từ RichTextBox
                                TextRange textRange = new TextRange(RtbOcrResult.Document.ContentStart, RtbOcrResult.Document.ContentEnd);
                                if (!string.IsNullOrWhiteSpace(textRange.Text.Trim()))
                                {
                                    foreach (Block block in RtbOcrResult.Document.Blocks)
                                    {
                                        if (block is Paragraph paragraph)
                                        {
                                            // Lấy canh lề đoạn
                                            TextAlignment alignment = paragraph.TextAlignment;
                                            XStringFormat stringFormat = new XStringFormat();
                                            switch (alignment)
                                            {
                                                case TextAlignment.Left:
                                                    stringFormat.Alignment = XStringAlignment.Near;
                                                    break;
                                                case TextAlignment.Center:
                                                    stringFormat.Alignment = XStringAlignment.Center;
                                                    break;
                                                case TextAlignment.Right:
                                                    stringFormat.Alignment = XStringAlignment.Far;
                                                    break;
                                                case TextAlignment.Justify:
                                                    stringFormat.Alignment = XStringAlignment.Near;
                                                    break;
                                            }

                                            foreach (Inline inline in paragraph.Inlines)
                                            {
                                                if (inline is Run run)
                                                {
                                                    // Lấy định dạng từ Run
                                                    string text = run.Text;
                                                    string fontFamily = "Times New Roman";
                                                    double fontSize = run.FontSize;
                                                    bool isBold = run.FontWeight == FontWeights.Bold;
                                                    bool isItalic = run.FontStyle == FontStyles.Italic;
                                                    var color = run.Foreground as SolidColorBrush;

                                                    // Tạo font PdfSharp
                                                    var fontStyle = XFontStyleEx.Regular;
                                                    if (isBold) fontStyle |= XFontStyleEx.Bold;
                                                    if (isItalic) fontStyle |= XFontStyleEx.Italic;
                                                    var font = new XFont(fontFamily, fontSize, fontStyle);

                                                    // Chuyển đổi màu
                                                    var xBrush = XBrushes.Black;
                                                    if (color != null)
                                                    {
                                                        xBrush = new XSolidBrush(XColor.FromArgb(
                                                            color.Color.A, color.Color.R, color.Color.G, color.Color.B));
                                                    }

                                                    // Vẽ văn bản
                                                    gfx.DrawString(text, font, xBrush, new XRect(20, yPosition, page.Width - 40, fontSize * 1.2), stringFormat);
                                                    yPosition += fontSize * 1.2;
                                                }
                                            }
                                            yPosition += 10;
                                        }
                                    }
                                }
                            }
                        }

                        document.Save(saveFileDialog.FileName);
                    }

                    TxtStatus.Text = "PDF with image exported successfully.";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"PDF Error: {ex.Message}\nStack Trace: {ex.StackTrace}");
                    TxtStatus.Text = "PDF export failed.";
                }
            }
        }

        private Bitmap ProcessImage(Bitmap inputImage)
        {
            // Tạo bản sao để xử lý
            Bitmap processed = (Bitmap)inputImage.Clone();

            // Áp dụng bộ lọc tăng độ sắc nét
            Sharpen sharpenFilter = new Sharpen();
            processed = sharpenFilter.Apply(processed);

            // Áp dụng bộ lọc tăng độ tương phản
            ContrastStretch contrastFilter = new ContrastStretch();
            processed = contrastFilter.Apply(processed);

            // Điều chỉnh độ sáng (tùy chọn)
            BrightnessCorrection brightnessFilter = new BrightnessCorrection(10); // Tăng nhẹ độ sáng
            processed = brightnessFilter.Apply(processed);

            return processed;
        }

        private BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                memory.Position = 0;
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }
    }
}