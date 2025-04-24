using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using OpenCvSharp;

namespace Ui
{
    public partial class AdjustImageWindow : System.Windows.Window
    {
        private Mat originalImage;
        private Mat displayImage;
        private OpenCvSharp.Rect detectedRect;
        private OpenCvSharp.Rect adjustedRect;
        private bool isDragging = false;
        private int selectedCorner = -1; // 0: TopLeft, 1: TopRight, 2: BottomRight, 3: BottomLeft
        private const int cornerSize = 20; // Tăng kích thước góc để dễ nhìn hơn
        private double scaleX, scaleY;

        public Mat CroppedImage { get; private set; }

        public AdjustImageWindow(Mat image)
        {
            InitializeComponent();
            originalImage = image.Clone();
            DetectPaperRegion();
            UpdateDisplayImage();
        }

        private void DetectPaperRegion()
        {
            Mat processed = originalImage.Clone();
            Mat hsv = new();
            Cv2.CvtColor(processed, hsv, ColorConversionCodes.BGR2HSV);
            Mat mask = new();
            Cv2.InRange(hsv, new Scalar(0, 0, 150), new Scalar(180, 50, 255), mask);
            Cv2.GaussianBlur(mask, mask, new OpenCvSharp.Size(9, 9), 0);
            Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(7, 7));
            Cv2.MorphologyEx(mask, mask, MorphTypes.Close, kernel, iterations: 3);

            OpenCvSharp.Point[][] contours;
            HierarchyIndex[] hierarchy;
            Cv2.FindContours(mask, out contours, out hierarchy, RetrievalModes.List, ContourApproximationModes.ApproxSimple);

            double maxArea = 0;
            OpenCvSharp.Rect? paperRect = null;
            foreach (var contour in contours)
            {
                double area = Cv2.ContourArea(contour);
                if (area > maxArea && area > (originalImage.Width * originalImage.Height * 0.2))
                {
                    maxArea = area;
                    paperRect = Cv2.BoundingRect(contour);
                }
            }

            hsv.Dispose();
            mask.Dispose();
            kernel.Dispose();
            processed.Dispose();

            if (paperRect == null)
            {
                detectedRect = new OpenCvSharp.Rect(0, 0, originalImage.Width, originalImage.Height);
            }
            else
            {
                int shrinkFactor = 10;
                int newWidth = (int)(paperRect.Value.Width * (1 - shrinkFactor / 100.0));
                int newHeight = (int)(paperRect.Value.Height * (1 - shrinkFactor / 100.0));
                int newX = paperRect.Value.X + (paperRect.Value.Width - newWidth) / 2;
                int newY = paperRect.Value.Y + (paperRect.Value.Height - newHeight) / 2;

                newX = Math.Max(0, newX);
                newY = Math.Max(0, newY);
                newWidth = Math.Min(originalImage.Width - newX, newWidth);
                newHeight = Math.Min(originalImage.Height - newY, newHeight);

                detectedRect = new OpenCvSharp.Rect(newX, newY, newWidth, newHeight);
            }

            adjustedRect = detectedRect;
            Console.WriteLine($"Detected Rect: {detectedRect}");
        }

        private void CalculateScale()
        {
            double displayWidth = ImgAdjustPreview.ActualWidth;
            double displayHeight = ImgAdjustPreview.ActualHeight;
            double imageWidth = originalImage.Width;
            double imageHeight = originalImage.Height;

            if (displayWidth > 0 && displayHeight > 0)
            {
                scaleX = imageWidth / displayWidth;
                scaleY = imageHeight / displayHeight;
            }
            else
            {
                scaleX = 1;
                scaleY = 1;
            }
            Console.WriteLine($"ScaleX: {scaleX}, ScaleY: {scaleY}, Display: {displayWidth}x{displayHeight}, Image: {imageWidth}x{imageHeight}");
        }

        private void ImgAdjustPreview_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            CalculateScale();
            UpdateDisplayImage();
        }

        private void UpdateDisplayImage()
        {
            displayImage = originalImage.Clone();
            Mat overlay = displayImage.Clone();

            Cv2.Rectangle(overlay, adjustedRect, new Scalar(255, 255, 200, 128), -1);
            Cv2.AddWeighted(overlay, 0.5, displayImage, 0.5, 0, displayImage);

            Cv2.Rectangle(displayImage, adjustedRect, new Scalar(0, 255, 0), 2);

            OpenCvSharp.Point topLeft = new OpenCvSharp.Point(adjustedRect.X, adjustedRect.Y);
            OpenCvSharp.Point topRight = new OpenCvSharp.Point(adjustedRect.X + adjustedRect.Width, adjustedRect.Y);
            OpenCvSharp.Point bottomRight = new OpenCvSharp.Point(adjustedRect.X + adjustedRect.Width, adjustedRect.Y + adjustedRect.Height);
            OpenCvSharp.Point bottomLeft = new OpenCvSharp.Point(adjustedRect.X, adjustedRect.Y + adjustedRect.Height);

            Cv2.Circle(displayImage, topLeft, cornerSize, new Scalar(0, 0, 255), -1);
            Cv2.Circle(displayImage, topRight, cornerSize, new Scalar(0, 0, 255), -1);
            Cv2.Circle(displayImage, bottomRight, cornerSize, new Scalar(0, 0, 255), -1);
            Cv2.Circle(displayImage, bottomLeft, cornerSize, new Scalar(0, 0, 255), -1);

            Console.WriteLine($"Drawing corners - TopLeft: {topLeft}, TopRight: {topRight}, BottomRight: {bottomRight}, BottomLeft: {bottomLeft}");

            ImgAdjustPreview.Source = BitmapSourceFromMat(displayImage);
            overlay.Dispose();
        }

        private void ImgAdjustPreview_MouseDown(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Point mousePos = e.GetPosition(ImgAdjustPreview);
            OpenCvSharp.Point imagePos = new OpenCvSharp.Point((int)(mousePos.X * scaleX), (int)(mousePos.Y * scaleY));

            OpenCvSharp.Point topLeft = new OpenCvSharp.Point(adjustedRect.X, adjustedRect.Y);
            OpenCvSharp.Point topRight = new OpenCvSharp.Point(adjustedRect.X + adjustedRect.Width, adjustedRect.Y);
            OpenCvSharp.Point bottomRight = new OpenCvSharp.Point(adjustedRect.X + adjustedRect.Width, adjustedRect.Y + adjustedRect.Height);
            OpenCvSharp.Point bottomLeft = new OpenCvSharp.Point(adjustedRect.X, adjustedRect.Y + adjustedRect.Height);

            if (IsPointNearCorner(imagePos, topLeft)) selectedCorner = 0;
            else if (IsPointNearCorner(imagePos, topRight)) selectedCorner = 1;
            else if (IsPointNearCorner(imagePos, bottomRight)) selectedCorner = 2;
            else if (IsPointNearCorner(imagePos, bottomLeft)) selectedCorner = 3;
            else selectedCorner = -1;

            Console.WriteLine($"MouseDown at ImagePos: {imagePos}, SelectedCorner: {selectedCorner}");

            if (selectedCorner >= 0)
            {
                isDragging = true;
                e.Handled = true;
            }
        }

        private void ImgAdjustPreview_MouseMove(object sender, MouseEventArgs e)
        {
            if (!isDragging || selectedCorner < 0) return;

            System.Windows.Point mousePos = e.GetPosition(ImgAdjustPreview);
            OpenCvSharp.Point imagePos = new OpenCvSharp.Point((int)(mousePos.X * scaleX), (int)(mousePos.Y * scaleY));

            imagePos.X = Math.Max(0, Math.Min(originalImage.Width - 1, imagePos.X));
            imagePos.Y = Math.Max(0, Math.Min(originalImage.Height - 1, imagePos.Y));

            int x = adjustedRect.X;
            int y = adjustedRect.Y;
            int width = adjustedRect.Width;
            int height = adjustedRect.Height;

            switch (selectedCorner)
            {
                case 0: // TopLeft
                    width = (adjustedRect.X + adjustedRect.Width) - imagePos.X;
                    height = (adjustedRect.Y + adjustedRect.Height) - imagePos.Y;
                    x = imagePos.X;
                    y = imagePos.Y;
                    break;
                case 1: // TopRight
                    width = imagePos.X - adjustedRect.X;
                    height = (adjustedRect.Y + adjustedRect.Height) - imagePos.Y;
                    y = imagePos.Y;
                    break;
                case 2: // BottomRight
                    width = imagePos.X - adjustedRect.X;
                    height = imagePos.Y - adjustedRect.Y;
                    break;
                case 3: // BottomLeft
                    width = (adjustedRect.X + adjustedRect.Width) - imagePos.X;
                    height = imagePos.Y - adjustedRect.Y;
                    x = imagePos.X;
                    break;
            }

            if (width > 0 && height > 0)
            {
                x = Math.Max(0, x);
                y = Math.Max(0, y);
                width = Math.Min(originalImage.Width - x, width);
                height = Math.Min(originalImage.Height - y, height);
                adjustedRect = new OpenCvSharp.Rect(x, y, width, height);
                UpdateDisplayImage();
            }

            e.Handled = true;
        }

        private void ImgAdjustPreview_MouseUp(object sender, MouseButtonEventArgs e)
        {
            isDragging = false;
            selectedCorner = -1;
            e.Handled = true;
        }

        private bool IsPointNearCorner(OpenCvSharp.Point point, OpenCvSharp.Point corner)
        {
            return Math.Abs(point.X - corner.X) <= cornerSize && Math.Abs(point.Y - corner.Y) <= cornerSize;
        }

        private void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            if (adjustedRect.Width > 0 && adjustedRect.Height > 0)
            {
                CroppedImage = new Mat(originalImage, adjustedRect);
                DialogResult = true;
            }
            else
            {
                MessageBox.Show("Invalid region selected.");
                DialogResult = false;
            }
            Close();
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