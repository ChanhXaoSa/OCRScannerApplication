using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using Tesseract;
using System.Windows.Documents;
using System.Windows.Media;
using PdfSharp.Drawing;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using PdfSharp.Pdf;
using OpenCvSharp;

namespace Ui
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private VideoCapture? videoCapture;
        private Mat? currentImage;

        public MainWindow()
        {
            InitializeComponent();
            InitializeWebcam();
        }

        private void InitializeWebcam()
        {
            try
            {
                videoCapture = new VideoCapture(0);
                if (!videoCapture.IsOpened())
                {
                    BtnCaptureWebcam.IsEnabled = false;
                    TxtStatus.Text = "No webcam detected.";
                }
            }
            catch
            {
                BtnCaptureWebcam.IsEnabled = false;
                TxtStatus.Text = "Error initializing webcam.";
            }
        }

        private void BtnCaptureWebcam_Click(object sender, RoutedEventArgs e)
        {
            if (videoCapture == null || !videoCapture.IsOpened()) return;

            BtnCaptureWebcam.Content = "Capture Frame";
            BtnCaptureWebcam.Click -= BtnCaptureWebcam_Click;
            BtnCaptureWebcam.Click += BtnCaptureFrame_Click;

            // Bắt đầu hiển thị video
            CompositionTarget.Rendering += RenderWebcamFrame;
            TxtStatus.Text = "Webcam started. Click 'Capture Frame' to capture.";
        }

        private void BtnCaptureFrame_Click(object sender, RoutedEventArgs e)
        {
            if (videoCapture != null && videoCapture.IsOpened())
            {
                CompositionTarget.Rendering -= RenderWebcamFrame;
                BtnCaptureWebcam.Content = "Capture from Webcam";
                BtnCaptureWebcam.Click -= BtnCaptureFrame_Click;
                BtnCaptureWebcam.Click += BtnCaptureWebcam_Click;
                TxtStatus.Text = "Image captured from webcam.";
            }
        }

        private void RenderWebcamFrame(object? sender, EventArgs e)
        {
            if (videoCapture == null || !videoCapture.IsOpened()) return;

            using Mat frame = new();
            videoCapture.Read(frame);
            if (!frame.Empty())
            {
                currentImage = frame.Clone();
                ImgPreview.Source = BitmapSourceFromMat(frame);
            }
        }

        private void BtnSelectImage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                currentImage = Cv2.ImRead(openFileDialog.FileName);
                ImgPreview.Source = BitmapSourceFromMat(currentImage);
                TxtStatus.Text = "Image selected.";
            }
        }

        private void BtnScan_Click(object sender, RoutedEventArgs e)
        {
            if (currentImage == null || currentImage.Empty())
            {
                MessageBox.Show("Please select or capture an image first.");
                return;
            }

            try
            {
                // Lưu Bitmap thành tệp tạm thời
                string tempImagePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ocr_temp_image.png");
                Cv2.ImWrite(tempImagePath, currentImage);

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
            if (currentImage == null || currentImage.Empty())
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
                    using Mat processedImage = ProcessImage(currentImage);

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
                                processedImage.WriteToStream(imageStream, ".png");
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

        private Mat ProcessImage(Mat inputImage)
        {
            // Tạo bản sao để xử lý
            Mat processed = inputImage.Clone();

            // Chuyển sang grayscale để tăng độ tương phản
            Mat gray = new();
            Cv2.CvtColor(processed, gray, ColorConversionCodes.BGR2GRAY);

            // Áp dụng bộ lọc tăng độ sắc nét (Laplacian)
            Mat sharpened = new();
            //Cv2.Laplacian(gray, sharpened, MatType.CV_16S);
            Mat sharpened8bit = new();
            //Cv2.ConvertScaleAbs(sharpened, sharpened8bit);
            //Cv2.AddWeighted(gray, 1.5, sharpened8bit, -0.5, 0, gray);

            // Tăng độ tương phản
            Cv2.Normalize(gray, gray, 0, 255, NormTypes.MinMax);

            // Điều chỉnh độ sáng
            Mat brightened = new();
            Cv2.ConvertScaleAbs(gray, brightened, 1.1, 10); // Tăng độ sáng nhẹ

            // Chuyển lại sang BGR để lưu
            Cv2.CvtColor(brightened, processed, ColorConversionCodes.GRAY2BGR);

            gray.Dispose();
            sharpened.Dispose();
            sharpened8bit.Dispose();
            brightened.Dispose();

            return processed;
        }

        private BitmapSource BitmapSourceFromMat(Mat mat)
        {
            using var stream = new MemoryStream();
            mat.WriteToStream(stream, ".png");
            stream.Position = 0;
            BitmapImage bitmapImage = new();
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = stream;
            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
            bitmapImage.EndInit();
            bitmapImage.Freeze();
            return bitmapImage;
        }
    }
}