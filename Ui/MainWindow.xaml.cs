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
using PdfSharp.Drawing.Layout;
using Tesseract;

namespace Ui
{
    public partial class MainWindow : System.Windows.Window
    {
        private VideoCapture? videoCapture;
        private Mat? currentImage;
        private List<int> availableWebcamIndices = [];
        private OpenCvSharp.Point[]? lastDetectedInnerRect;

        public MainWindow()
        {
            InitializeComponent();
            InitializeWebcam();
        }

        private void InitializeWebcam()
        {
            availableWebcamIndices = new List<int>();
            CmbWebcamDevices.Items.Clear();

            for (int i = 0; i < 10; i++)
            {
                using (var tempCapture = new VideoCapture(i))
                {
                    if (tempCapture.IsOpened())
                    {
                        availableWebcamIndices.Add(i);
                        CmbWebcamDevices.Items.Add($"Webcam {i}");
                    }
                }
            }

            if (availableWebcamIndices.Count > 0)
            {
                CmbWebcamDevices.SelectedIndex = 0;
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
            if (videoCapture != null)
            {
                CompositionTarget.Rendering -= RenderWebcamFrame;
                videoCapture.Release();
                videoCapture.Dispose();
                videoCapture = null;
            }

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

            videoCapture.Set(VideoCaptureProperties.FrameWidth, 1280);
            videoCapture.Set(VideoCaptureProperties.FrameHeight, 720);

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

                if (lastDetectedInnerRect != null && currentImage != null && !currentImage.Empty())
                {
                    // Tạo mặt nạ cho tứ giác
                    Mat mask = new Mat(currentImage.Size(), MatType.CV_8UC1, Scalar.Black);
                    Cv2.FillPoly(mask, new[] { lastDetectedInnerRect }, Scalar.White);
                    // Áp dụng mặt nạ để cắt vùng tứ giác
                    Mat cropped = new Mat(currentImage.Size(), currentImage.Type());
                    Cv2.BitwiseAnd(currentImage, currentImage, cropped, mask);
                    // Cắt hình chữ nhật bao quanh tứ giác để loại bỏ phần thừa
                    OpenCvSharp.Rect boundingRect = Cv2.BoundingRect(lastDetectedInnerRect);
                    currentImage = new Mat(cropped, boundingRect);
                    mask.Dispose();
                    cropped.Dispose();
                }

                BtnCaptureWebcam.Content = "Capture from Webcam";
                BtnCaptureWebcam.Click -= BtnCaptureFrame_Click;
                BtnCaptureWebcam.Click += BtnCaptureWebcam_Click;
                TxtStatus.Text = "Image captured and cropped from webcam.";
            }
        }

        private void RenderWebcamFrame(object? sender, EventArgs e)
        {
            if (videoCapture == null || !videoCapture.IsOpened()) return;

            using Mat frame = new();
            videoCapture.Read(frame);
            if (frame.Empty()) return;

            lastDetectedInnerRect = DetectInnerRectangle(frame);

            using Mat displayFrame = frame.Clone();

            if (lastDetectedInnerRect != null)
            {
                Cv2.FillPoly(displayFrame, new[] { lastDetectedInnerRect }, new Scalar(255, 255, 200, 128));
                Cv2.AddWeighted(displayFrame, 0.5, frame, 0.5, 0, displayFrame);
                Cv2.Polylines(displayFrame, new[] { lastDetectedInnerRect }, true, new Scalar(0, 255, 0), 2);
            }

            ImgPreview.Source = BitmapSourceFromMat(displayFrame);
            currentImage = frame.Clone();
        }

        private void BtnSelectImage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                Mat selectedImage = Cv2.ImRead(openFileDialog.FileName);
                if (selectedImage.Empty())
                {
                    MessageBox.Show("Failed to load the selected image.");
                    TxtStatus.Text = "Image loading failed.";
                    return;
                }

                AdjustImageWindow adjustWindow = new AdjustImageWindow(selectedImage);
                if (adjustWindow.ShowDialog() == true)
                {
                    currentImage = adjustWindow.CroppedImage;
                    if (currentImage != null && !currentImage.Empty())
                    {
                        ImgPreview.Source = BitmapSourceFromMat(currentImage);
                        TxtStatus.Text = "Image selected and cropped.";
                    }
                    else
                    {
                        TxtStatus.Text = "No region selected for cropping.";
                    }
                }
                else
                {
                    TxtStatus.Text = "Image selection canceled.";
                }
            }
        }

        private async void BtnScan_Click(object sender, RoutedEventArgs e)
        {
            if (currentImage == null || currentImage.Empty())
            {
                MessageBox.Show("No image to scan. Please select or capture an image first.");
                return;
            }

            ScanProgressBar.Visibility = Visibility.Visible;
            ScanStatusText.Visibility = Visibility.Visible;
            ScanProgressBar.Value = 0;
            ScanStatusText.Text = "Scanning started...";

            try
            {
                await Task.Run(() =>
                {
                    int totalSteps = 5;
                    for (int step = 1; step <= totalSteps; step++)
                    {
                        Thread.Sleep(500);
                        Dispatcher.Invoke(() =>
                        {
                            ScanProgressBar.Value = (step * 100) / totalSteps;
                            ScanStatusText.Text = $"Scanning... {ScanProgressBar.Value}% completed.";
                        });
                    }

                    using Mat scannedContent = ExtractContent(currentImage);
                    Dispatcher.Invoke(() =>
                    {
                        if (!scannedContent.Empty())
                        {
                            ImgPreview.Source = BitmapSourceFromMat(scannedContent);
                            ScanStatusText.Text = "Scanning completed successfully.";
                        }
                        else
                        {
                            ScanStatusText.Text = "No content detected during scanning.";
                        }
                    });
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during scanning: {ex.Message}");
                ScanStatusText.Text = "Scanning failed.";
            }
            finally
            {
                await Task.Delay(1000);
                ScanProgressBar.Visibility = Visibility.Collapsed;
                ScanStatusText.Visibility = Visibility.Collapsed;
            }
        }

        private async void BtnScanAndDisplayProgress_Click(object sender, RoutedEventArgs e)
        {
            if (currentImage == null || currentImage.Empty())
            {
                MessageBox.Show("No image to scan. Please select or capture an image first.");
                return;
            }

            ScanProgressBar.Visibility = Visibility.Visible;
            ScanStatusText.Visibility = Visibility.Visible;
            ScanProgressBar.Value = 0;
            ScanStatusText.Text = "Scanning started...";

            try
            {
                await Task.Run(() =>
                {
                    int totalSteps = 5;
                    for (int step = 1; step <= totalSteps; step++)
                    {
                        Thread.Sleep(500);
                        Dispatcher.Invoke(() =>
                        {
                            ScanProgressBar.Value = (step * 100) / totalSteps;
                            ScanStatusText.Text = $"Scanning... {ScanProgressBar.Value}% completed.";
                        });
                    }

                    using Mat scannedContent = ExtractContent(currentImage);
                    Dispatcher.Invoke(() =>
                    {
                        if (!scannedContent.Empty())
                        {
                            ImgPreview.Source = BitmapSourceFromMat(scannedContent);
                            ScanStatusText.Text = "Scanning completed successfully.";
                        }
                        else
                        {
                            ScanStatusText.Text = "No content detected during scanning.";
                        }
                    });
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during scanning: {ex.Message}");
                ScanStatusText.Text = "Scanning failed.";
            }
            finally
            {
                await Task.Delay(1000);
                ScanProgressBar.Visibility = Visibility.Collapsed;
                ScanStatusText.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnExportPdf_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("This function is disabled. Use 'Add Picture to PDF' to export the image directly.");
            TxtStatus.Text = "PDF export (text) disabled.";
        }

        private Mat ExtractContent(Mat inputImage)
        {
            Mat gray = new();
            Cv2.CvtColor(inputImage, gray, ColorConversionCodes.BGR2GRAY);

            Mat binary = new();
            Cv2.AdaptiveThreshold(gray, binary, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 15, 5);

            OpenCvSharp.Point[][] contours;
            HierarchyIndex[] hierarchy;
            Cv2.FindContours(binary, out contours, out hierarchy, RetrievalModes.External, ContourApproximationModes.ApproxSimple);

            OpenCvSharp.Rect? contentRect = null;
            foreach (var contour in contours)
            {
                OpenCvSharp.Rect rect = Cv2.BoundingRect(contour);
                contentRect = contentRect.HasValue ? contentRect.Value.Union(rect) : rect;
            }

            if (contentRect != null)
            {
                int padding = 10;
                contentRect = new OpenCvSharp.Rect(
                    Math.Max(0, contentRect.Value.X - padding),
                    Math.Max(0, contentRect.Value.Y - padding),
                    Math.Min(inputImage.Width - contentRect.Value.X, contentRect.Value.Width + 2 * padding),
                    Math.Min(inputImage.Height - contentRect.Value.Y, contentRect.Value.Height + 2 * padding)
                );

                Mat content = new(inputImage, contentRect.Value);
                gray.Dispose();
                binary.Dispose();
                return content;
            }

            gray.Dispose();
            binary.Dispose();
            return new Mat();
        }

        private void BtnAddPictureToPdf_Click(object sender, RoutedEventArgs e)
        {
            if (currentImage == null || currentImage.Empty())
            {
                MessageBox.Show("No image to process. Please select or capture an image first.");
                return;
            }

            SaveFileDialog saveFileDialog = new()
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                FileName = "ScannedDocument.pdf"
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

                    // Process the image to enhance quality
                    using Mat processedImage = ProcessImage(currentImage);
                    // Extract inner content, excluding paper borders
                    using Mat contentImage = ExtractInnerContent(processedImage);
                    if (contentImage.Empty())
                    {
                        MessageBox.Show("Could not detect content in the image.");
                        TxtStatus.Text = "PDF export failed.";
                        return;
                    }

                    // Place the content image on a white background
                    using Mat finalImage = PlaceOnWhiteBackground(contentImage);

                    // Create PDF with the content image
                    using var document = new PdfDocument();
                    var page = document.AddPage();
                    page.Width = 595;
                    page.Height = 842;

                    using (var gfx = XGraphics.FromPdfPage(page))
                    {
                        using MemoryStream imageStream = new();
                        finalImage.WriteToStream(imageStream, ".png");
                        imageStream.Position = 0;
                        XImage xImage = XImage.FromStream(imageStream);

                        double imgWidth = finalImage.Width;
                        double imgHeight = finalImage.Height;
                        double scale = Math.Min((page.Width - 40) / imgWidth, (page.Height - 40) / imgHeight);
                        imgWidth *= scale;
                        imgHeight *= scale;

                        double xPosition = (page.Width - imgWidth) / 2;
                        double yPosition = (page.Height - imgHeight) / 2;

                        gfx.DrawImage(xImage, xPosition, yPosition, imgWidth, imgHeight);
                    }

                    document.Save(saveFileDialog.FileName);
                    TxtStatus.Text = "PDF with scanned content image exported successfully.";
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

        private static Mat ProcessImage(Mat inputImage)
        {
            Mat processed = inputImage.Clone();

            // Apply bilateral filter to reduce noise while preserving edges
            Mat bilateral = new();
            Cv2.BilateralFilter(processed, bilateral, 11, 17, 17);

            // Convert to grayscale for thresholding
            Mat gray = new();
            Cv2.CvtColor(bilateral, gray, ColorConversionCodes.BGR2GRAY);

            // Apply adaptive thresholding to make text darker and more distinct
            Mat binary = new();
            Cv2.AdaptiveThreshold(gray, binary, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 15, 10);

            // Merge binary image back to color image to retain color for signatures/seals
            Mat[] channels = new Mat[3];
            for (int i = 0; i < 3; i++)
            {
                channels[i] = binary.Clone();
            }
            Cv2.Merge(channels, processed);

            // Enhance contrast and sharpness
            Mat[] splitChannels = Cv2.Split(processed);
            for (int i = 0; i < splitChannels.Length; i++)
            {
                // Normalize to increase contrast
                Cv2.Normalize(splitChannels[i], splitChannels[i], 0, 255, NormTypes.MinMax);

                // Sharpen using Laplacian
                Mat sharpened = new();
                Cv2.Laplacian(splitChannels[i], sharpened, MatType.CV_16S);
                Mat sharpened8bit = new();
                Cv2.ConvertScaleAbs(sharpened, sharpened8bit);
                Cv2.AddWeighted(splitChannels[i], 1.8, sharpened8bit, -0.6, 0, splitChannels[i]);

                // Increase brightness slightly
                Cv2.ConvertScaleAbs(splitChannels[i], splitChannels[i], 1.1, 10);

                sharpened.Dispose();
                sharpened8bit.Dispose();
            }

            Cv2.Merge(splitChannels, processed);

            // Clean up
            foreach (var channel in splitChannels)
            {
                channel.Dispose();
            }
            foreach (var channel in channels)
            {
                channel.Dispose();
            }
            bilateral.Dispose();
            gray.Dispose();
            binary.Dispose();

            return processed;
        }

        private static Mat PlaceOnWhiteBackground(Mat inputImage)
        {
            int backgroundWidth = 2480;
            int backgroundHeight = 3508;

            Mat background = new(backgroundHeight, backgroundWidth, MatType.CV_8UC3, new Scalar(255, 255, 255));

            double scale = Math.Min((double)(backgroundWidth - 80) / inputImage.Width, (double)(backgroundHeight - 80) / inputImage.Height);
            int newWidth = (int)(inputImage.Width * scale);
            int newHeight = (int)(inputImage.Height * scale);

            Mat resizedImage = new();
            Cv2.Resize(inputImage, resizedImage, new OpenCvSharp.Size(newWidth, newHeight));

            int xOffset = (backgroundWidth - newWidth) / 2;
            int yOffset = (backgroundHeight - newHeight) / 2;

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

        private OpenCvSharp.Point[]? DetectInnerRectangle(Mat inputImage)
        {
            Mat processed = inputImage.Clone();
            Mat hsv = new();
            Cv2.CvtColor(processed, hsv, ColorConversionCodes.BGR2HSV);

            // Adjust HSV range to better detect white paper
            Mat mask = new();
            Cv2.InRange(hsv, new Scalar(0, 0, 180), new Scalar(180, 40, 255), mask);

            // Apply stronger Gaussian blur to reduce noise
            Cv2.GaussianBlur(mask, mask, new OpenCvSharp.Size(9, 9), 0);

            // Apply morphological operations to close gaps and clean up
            Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(7, 7));
            Cv2.MorphologyEx(mask, mask, MorphTypes.Close, kernel, iterations: 3);
            Cv2.MorphologyEx(mask, mask, MorphTypes.Open, kernel, iterations: 2);

            // Find contours
            OpenCvSharp.Point[][] contours;
            HierarchyIndex[] hierarchy;
            Cv2.FindContours(mask, out contours, out hierarchy, RetrievalModes.External, ContourApproximationModes.ApproxSimple);

            //tim duong vien lon nhat
            double maxArea = 0;
            OpenCvSharp.Point[]? paperQuad = null;
            foreach (var contour in contours)
            {
                double area = Cv2.ContourArea(contour);
                if (area > maxArea && area > (inputImage.Width * inputImage.Height * 0.15))
                {
                    // Xấp xỉ đường viền thành đa giác
                    double peri = Cv2.ArcLength(contour, true);
                    OpenCvSharp.Point[] approx = Cv2.ApproxPolyDP(contour, 0.02 * peri, true);
                    // Kiểm tra nếu đa giác có 4 đỉnh (tứ giác)
                    if (approx.Length == 4)
                    {
                        maxArea = area;
                        paperQuad = approx;
                    }
                }
            }

            hsv.Dispose();
            mask.Dispose();
            kernel.Dispose();
            processed.Dispose();

            if (paperQuad == null) return null;
            // Tính tâm của tứ giác
            OpenCvSharp.Point center = new OpenCvSharp.Point(
                (paperQuad[0].X + paperQuad[1].X + paperQuad[2].X + paperQuad[3].X) / 4,
                (paperQuad[0].Y + paperQuad[1].Y + paperQuad[2].Y + paperQuad[3].Y) / 4
            );
            // Thu nhỏ tứ giác bằng cách di chuyển các đỉnh về phía tâm
            float shrinkFactor = 0.9f; // Thu nhỏ 10%
            OpenCvSharp.Point[] shrunkQuad = new OpenCvSharp.Point[4];
            for (int i = 0; i < 4; i++)
            {
                shrunkQuad[i] = new OpenCvSharp.Point(
                    (int)(center.X + shrinkFactor * (paperQuad[i].X - center.X)),
                    (int)(center.Y + shrinkFactor * (paperQuad[i].Y - center.Y))
                );
                // Đảm bảo đỉnh nằm trong giới hạn hình ảnh
                shrunkQuad[i].X = Math.Max(0, Math.Min(inputImage.Width - 1, shrunkQuad[i].X));
                shrunkQuad[i].Y = Math.Max(0, Math.Min(inputImage.Height - 1, shrunkQuad[i].Y));
            }
            return shrunkQuad;
        }

        private Mat ExtractInnerContent(Mat inputImage)
        {
            Mat gray = new();
            Cv2.CvtColor(inputImage, gray, ColorConversionCodes.BGR2GRAY);

            Mat binary = new();
            Cv2.AdaptiveThreshold(gray, binary, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 15, 5);

            OpenCvSharp.Point[][] contours;
            HierarchyIndex[] hierarchy;
            Cv2.FindContours(binary, out contours, out hierarchy, RetrievalModes.List, ContourApproximationModes.ApproxSimple);

            OpenCvSharp.Rect? contentRect = null;
            foreach (var contour in contours)
            {
                double area = Cv2.ContourArea(contour);
                if (area > 1000) // Filter small noise contours
                {
                    OpenCvSharp.Rect rect = Cv2.BoundingRect(contour);
                    contentRect = contentRect.HasValue ? contentRect.Value.Union(rect) : rect;
                }
            }

            if (contentRect != null)
            {
                // Increase border margin to exclude paper edges and lighting artifacts
                int borderMargin = 30;
                int newX = Math.Max(contentRect.Value.X + borderMargin, 0);
                int newY = Math.Max(contentRect.Value.Y + borderMargin, 0);
                int newWidth = Math.Min(contentRect.Value.Width - 2 * borderMargin, inputImage.Width - newX);
                int newHeight = Math.Min(contentRect.Value.Height - 2 * borderMargin, inputImage.Height - newY);

                if (newWidth > 0 && newHeight > 0)
                {
                    Mat content = new(inputImage, new OpenCvSharp.Rect(newX, newY, newWidth, newHeight));
                    gray.Dispose();
                    binary.Dispose();
                    return content;
                }
            }

            gray.Dispose();
            binary.Dispose();
            return new Mat();
        }
    }
}