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
        private System.Windows.Point topLeft, topRight, bottomLeft, bottomRight;
        private readonly Mat originalImage;
        public Mat CroppedImage { get; private set; }
        private double currentScale = 1.0; // Initial scale

        public AdjustImageWindow(Mat image)
        {
            InitializeComponent();
            originalImage = image.Clone();
            DisplayedImage.Source = BitmapSourceFromMat(originalImage);

            // Initialize corner positions
            topLeft = new System.Windows.Point(50, 50);
            topRight = new System.Windows.Point(200, 50);
            bottomLeft = new System.Windows.Point(50, 200);
            bottomRight = new System.Windows.Point(200, 200);

            UpdateCornerPositions();
            FitImageToWindow();
        }

        private void FitImageToWindow()
        {
            // Calculate the scale to fit the image within the window
            double scaleX = ImageCanvas.ActualWidth / originalImage.Width;
            double scaleY = ImageCanvas.ActualHeight / originalImage.Height;
            currentScale = Math.Min(scaleX, scaleY);

            // Apply the scale
            ImageScaleTransform.ScaleX = currentScale;
            ImageScaleTransform.ScaleY = currentScale;
        }

        private void UpdateCornerPositions()
        {
            Canvas.SetLeft(TopLeftCorner, topLeft.X);
            Canvas.SetTop(TopLeftCorner, topLeft.Y);

            Canvas.SetLeft(TopRightCorner, topRight.X);
            Canvas.SetTop(TopRightCorner, topRight.Y);

            Canvas.SetLeft(BottomLeftCorner, bottomLeft.X);
            Canvas.SetTop(BottomLeftCorner, bottomLeft.Y);

            Canvas.SetLeft(BottomRightCorner, bottomRight.X);
            Canvas.SetTop(BottomRightCorner, bottomRight.Y);
        }

        private void CornerButton_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                Button corner = sender as Button;
                System.Windows.Point mousePosition = e.GetPosition(ImageCanvas);

                if (corner == TopLeftCorner) topLeft = mousePosition;
                else if (corner == TopRightCorner) topRight = mousePosition;
                else if (corner == BottomLeftCorner) bottomLeft = mousePosition;
                else if (corner == BottomRightCorner) bottomRight = mousePosition;

                UpdateCornerPositions();
            }
        }

        private void ZoomInButton_Click(object sender, RoutedEventArgs e)
        {
            currentScale += 0.1; // Increase scale
            ImageScaleTransform.ScaleX = currentScale;
            ImageScaleTransform.ScaleY = currentScale;
        }

        private void ZoomOutButton_Click(object sender, RoutedEventArgs e)
        {
            currentScale = Math.Max(0.1, currentScale - 0.1); // Decrease scale, but not below 0.1
            ImageScaleTransform.ScaleX = currentScale;
            ImageScaleTransform.ScaleY = currentScale;
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            // Crop the image based on the adjusted corners
            var srcPoints = new[]
            {
                    new Point2f((float)topLeft.X, (float)topLeft.Y),
                    new Point2f((float)topRight.X, (float)topRight.Y),
                    new Point2f((float)bottomRight.X, (float)bottomRight.Y),
                    new Point2f((float)bottomLeft.X, (float)bottomLeft.Y)
                };

            var dstPoints = new[]
            {
                    new Point2f(0, 0),
                    new Point2f(originalImage.Width - 1, 0),
                    new Point2f(originalImage.Width - 1, originalImage.Height - 1),
                    new Point2f(0, originalImage.Height - 1)
                };

            Mat perspectiveTransform = Cv2.GetPerspectiveTransform(srcPoints, dstPoints);
            CroppedImage = new Mat();
            Cv2.WarpPerspective(originalImage, CroppedImage, perspectiveTransform, new OpenCvSharp.Size(originalImage.Width, originalImage.Height));

            DialogResult = true;
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
