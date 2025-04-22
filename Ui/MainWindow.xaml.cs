using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using OpenCvSharp;
using System.Diagnostics;
using System.Windows.Media;
using System.Windows.Controls;

namespace Ui
{
    public partial class MainWindow : System.Windows.Window
    {
        private VideoCapture? videoCapture;
        private Mat? currentImage;
        private List<int> availableWebcamIndices = [];

        public MainWindow()
        {
            InitializeComponent();
            InitializeWebcam();
        }

        private void InitializeWebcam()
        {
            availableWebcamIndices = [];
            CmbWebcamDevices.Items.Clear();

            // Kiểm tra các thiết bị webcam khả dụng (index từ 0 đến 9)
            for (int i = 0; i < 10; i++)
            {
                using var tempCapture = new VideoCapture(i);
                if (tempCapture.IsOpened())
                {
                    availableWebcamIndices.Add(i);
                    CmbWebcamDevices.Items.Add($"Webcam {i}");
                }
            }

            if (availableWebcamIndices.Count > 0)
            {
                CmbWebcamDevices.SelectedIndex = 0; // Chọn webcam đầu tiên mặc định
                BtnCaptureWebcam.IsEnabled = true;
                TxtStatus.Text = "Webcam devices detected.";
            }
            else
            {
                BtnCaptureWebcam.IsEnabled = false;
                TxtStatus.Text = "No webcam detected.";
            }
        }

        private void CmbWebcamDevices_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Đóng webcam hiện tại nếu đang mở
            if (videoCapture != null)
            {
                CompositionTarget.Rendering -= RenderWebcamFrame;
                videoCapture.Release();
                videoCapture.Dispose();
                videoCapture = null;
            }

            // Nếu đang ở trạng thái chụp, khởi động lại webcam mới
            if (BtnCaptureWebcam.Content.ToString() == "Capture Frame")
            {
                StartWebcam();
            }
        }

        private void BtnCaptureWebcam_Click(object sender, RoutedEventArgs e)
        {
            if (availableWebcamIndices.Count == 0) return;

            BtnCaptureWebcam.Content = "Capture Frame";
            BtnCaptureWebcam.Click -= BtnCaptureWebcam_Click;
            BtnCaptureWebcam.Click += BtnCaptureFrame_Click;

            StartWebcam();
            TxtStatus.Text = "Webcam started. Click 'Capture Frame' to capture.";
        }

        private void StartWebcam()
        {
            if (CmbWebcamDevices.SelectedIndex < 0) return;

            int selectedIndex = availableWebcamIndices[CmbWebcamDevices.SelectedIndex];
            videoCapture = new VideoCapture(selectedIndex);

            if (!videoCapture.IsOpened())
            {
                MessageBox.Show($"Failed to open webcam {selectedIndex}.");
                BtnCaptureWebcam.Content = "Capture from Webcam";
                BtnCaptureWebcam.Click -= BtnCaptureFrame_Click;
                BtnCaptureWebcam.Click += BtnCaptureWebcam_Click;
                TxtStatus.Text = "Failed to start webcam.";
                return;
            }

            CompositionTarget.Rendering += RenderWebcamFrame;
        }

        private void BtnCaptureFrame_Click(object sender, RoutedEventArgs e)
        {
            if (videoCapture != null && videoCapture.IsOpened())
            {
                CompositionTarget.Rendering -= RenderWebcamFrame;
                videoCapture.Release();
                videoCapture.Dispose();
                videoCapture = null;

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
            OpenFileDialog openFileDialog = new()
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
            MessageBox.Show("Scan functionality is disabled. Use 'Add Picture to PDF' to export the image directly.");
            TxtStatus.Text = "Scan disabled.";
        }

        private void BtnExportPdf_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("This function is disabled. Use 'Add Picture to PDF' to export the image directly.");
            TxtStatus.Text = "PDF export (text) disabled.";
        }

        private void BtnAddPictureToPdf_Click(object sender, RoutedEventArgs e)
        {
            if (currentImage == null || currentImage.Empty())
            {
                MessageBox.Show("No image to add. Please select or capture an image first.");
                return;
            }

            SaveFileDialog saveFileDialog = new()
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                FileName = "ScannedDocumentWithImage.pdf"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    string directory = Path.GetDirectoryName(saveFileDialog.FileName) ?? "";
                    if (!Directory.Exists(directory))
                    {
                        MessageBox.Show("Destination directory does not exist.");
                        TxtStatus.Text = "PDF export failed.";
                        return;
                    }

                    using Mat paperRegion = CropPaperRegion(currentImage);
                    if (paperRegion.Empty())
                    {
                        MessageBox.Show("Could not detect paper region in the image.");
                        TxtStatus.Text = "PDF export failed.";
                        return;
                    }

                    using Mat processedImage = ProcessImage(paperRegion);
                    using Mat imageWithWhiteBackground = PlaceOnWhiteBackground(processedImage);

                    using var document = new PdfDocument();
                    var page = document.AddPage();
                    page.Width = XUnit.FromPoint(595); // A4 width in points
                    page.Height = XUnit.FromPoint(842); // A4 height in points

                    using (var gfx = XGraphics.FromPdfPage(page))
                    {
                        using MemoryStream imageStream = new();
                        imageWithWhiteBackground.WriteToStream(imageStream, ".png");
                        imageStream.Position = 0;
                        XImage xImage = XImage.FromStream(imageStream);

                        double imgWidth = imageWithWhiteBackground.Width;
                        double imgHeight = imageWithWhiteBackground.Height;
                        double scale = Math.Min((page.Width.Point - 40) / imgWidth, (page.Height.Point - 40) / imgHeight);
                        imgWidth *= scale;
                        imgHeight *= scale;

                        double xPosition = (page.Width.Point - imgWidth) / 2; // Center horizontally
                        double yPosition = (page.Height.Point - imgHeight) / 2; // Center vertically

                        gfx.DrawImage(xImage, xPosition, yPosition, imgWidth, imgHeight);
                    }

                    document.Save(saveFileDialog.FileName);
                    TxtStatus.Text = "PDF with image on white background exported successfully.";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"PDF Error: {ex.Message}\nStack Trace: {ex.StackTrace}");
                    TxtStatus.Text = "PDF export failed.";
                }
            }
        }

        private void BtnExportWord_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This function is disabled.");
            TxtStatus.Text = "Word export disabled.";
        }

        private static Mat CropPaperRegion(Mat inputImage)
        {
            Mat processed = inputImage.Clone();

            // Convert to HSV color space for better paper detection
            using Mat hsv = new();
            Cv2.CvtColor(processed, hsv, ColorConversionCodes.BGR2HSV);

            // Create a mask to detect bright regions (typically white paper)
            using Mat mask = new();
            Cv2.InRange(hsv, new Scalar(0, 0, 150), new Scalar(180, 50, 255), mask);

            // Smooth the mask to reduce noise
            Cv2.GaussianBlur(mask, mask, new OpenCvSharp.Size(9, 9), 0);

            // Apply morphological operations to clean up the mask
            using Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(7, 7));
            Cv2.MorphologyEx(mask, mask, MorphTypes.Close, kernel, iterations: 2);

            // Find contours
            Cv2.FindContours(mask, out OpenCvSharp.Point[][] contours, out _, RetrievalModes.List, ContourApproximationModes.ApproxSimple);

            double maxArea = 0;
            OpenCvSharp.Rect? paperRect = null;
            foreach (var contour in contours)
            {
                double area = Cv2.ContourArea(contour);
                if (area > maxArea && area > (processed.Width * processed.Height * 0.2))
                {
                    maxArea = area;
                    paperRect = Cv2.BoundingRect(contour);
                }
            }

            if (paperRect == null)
            {
                processed.Dispose();
                return new Mat();
            }

            int padding = 5; // Add small padding to ensure no content is cropped
            int x = Math.Max(0, paperRect.Value.X - padding);
            int y = Math.Max(0, paperRect.Value.Y - padding);
            int width = Math.Min(processed.Width - x, paperRect.Value.Width + 2 * padding);
            int height = Math.Min(processed.Height - y, paperRect.Value.Height + 2 * padding);

            OpenCvSharp.Rect roi = new(x, y, width, height);
            Mat cropped = new(processed, roi);

            processed.Dispose();
            return cropped;
        }

        private static Mat ProcessImage(Mat inputImage)
        {
            Mat processed = inputImage.Clone();

            // Loại bỏ bóng bằng bộ lọc song phương
            Mat bilateral = new();
            Cv2.BilateralFilter(processed, bilateral, 11, 17, 17);

            // Tách thành các kênh màu BGR
            Mat[] channels = Cv2.Split(bilateral);

            // Tăng độ tương phản trên từng kênh
            for (int i = 0; i < channels.Length; i++)
            {
                Cv2.Normalize(channels[i], channels[i], 0, 255, NormTypes.MinMax);
            }

            // Tăng độ sắc nét trên từng kênh
            Mat[] sharpenedChannels = new Mat[channels.Length];
            for (int i = 0; i < channels.Length; i++)
            {
                Mat sharpened = new();
                Cv2.Laplacian(channels[i], sharpened, MatType.CV_16S);
                Mat sharpened8bit = new();
                Cv2.ConvertScaleAbs(sharpened, sharpened8bit);
                Cv2.AddWeighted(channels[i], 1.8, sharpened8bit, -0.6, 0, channels[i]);
                sharpenedChannels[i] = channels[i];
                sharpened.Dispose();
                sharpened8bit.Dispose();
            }

            // Điều chỉnh độ sáng để làm nổi bật chữ và con dấu
            for (int i = 0; i < channels.Length; i++)
            {
                Cv2.ConvertScaleAbs(channels[i], channels[i], 1.2, 15);
            }

            // Gộp các kênh lại thành ảnh màu
            Cv2.Merge(channels, processed);

            // Giải phóng tài nguyên
            foreach (var channel in channels)
            {
                channel.Dispose();
            }
            bilateral.Dispose();

            return processed;
        }

        private static Mat PlaceOnWhiteBackground(Mat inputImage)
        {
            // Kích thước nền trắng (A4 ở 300 DPI)
            int backgroundWidth = 2480; // 8.27 inch * 300 DPI
            int backgroundHeight = 3508; // 11.69 inch * 300 DPI

            // Tạo hình ảnh nền trắng
            Mat background = new(backgroundHeight, backgroundWidth, MatType.CV_8UC3, new Scalar(255, 255, 255));

            // Tính toán tỷ lệ để giữ nguyên tỷ lệ khung hình của hình ảnh gốc
            double scale = Math.Min((double)(backgroundWidth - 80) / inputImage.Width, (double)(backgroundHeight - 80) / inputImage.Height);
            int newWidth = (int)(inputImage.Width * scale);
            int newHeight = (int)(inputImage.Height * scale);

            // Thay đổi kích thước hình ảnh gốc
            Mat resizedImage = new();
            Cv2.Resize(inputImage, resizedImage, new OpenCvSharp.Size(newWidth, newHeight));

            // Tính vị trí để đặt hình ảnh vào giữa nền trắng
            int xOffset = (backgroundWidth - newWidth) / 2;
            int yOffset = (backgroundHeight - newHeight) / 2;

            // Đặt hình ảnh vào giữa nền trắng
            OpenCvSharp.Rect roi = new(xOffset, yOffset, newWidth, newHeight);
            resizedImage.CopyTo(new Mat(background, roi));

            resizedImage.Dispose();
            return background;
        }

        private static BitmapImage BitmapSourceFromMat(Mat mat)
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