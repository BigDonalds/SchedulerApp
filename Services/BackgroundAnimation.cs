using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace SchedulerApp.Services
{
    /// <summary>
    /// Manages the background particle animation on the main window.
    /// Supports multiple themes selectable by index:
    ///   0 = Snow
    ///   1 = Starry Night
    /// </summary>
    public class BackgroundAnimation
    {
        private readonly Canvas _canvas;
        private readonly Random _random = new Random();
        private DispatcherTimer _timer;
        private int _currentTheme = 0;
        private readonly List<UIElement> _activeParticles = new List<UIElement>();
        private readonly List<UIElement> _staticParticles = new List<UIElement>();
        private Brush _originalCanvasBackground;
        private DateTime _lastShootingStarTime = DateTime.MinValue;

        public BackgroundAnimation(Canvas canvas)
        {
            _canvas = canvas;
            _originalCanvasBackground = canvas.Background;
        }

        /// <summary>
        /// Sets the active animation theme by index (0-1).
        /// Stops the current theme, clears particles, and starts the new one.
        /// </summary>
        public void SetTheme(int themeNumber)
        {
            _currentTheme = themeNumber;
            Stop();
            ClearParticles();
            RestoreCanvasBackground();
            Start();
        }

        public int CurrentTheme => _currentTheme;

        public void Start()
        {
            if (_timer != null) return;

            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromMilliseconds(150);
            _timer.Tick += Timer_Tick;
            _timer.Start();

            // Some themes need static elements drawn once
            if (_currentTheme == 1) // Starry Night
                CreateStarryNight();
        }

        public void Stop()
        {
            if (_timer != null)
            {
                _timer.Stop();
                _timer.Tick -= Timer_Tick;
                _timer = null;
            }
        }

        private void ClearParticles()
        {
            foreach (var particle in _activeParticles)
            {
                _canvas.Children.Remove(particle);
            }
            _activeParticles.Clear();

            foreach (var particle in _staticParticles)
            {
                _canvas.Children.Remove(particle);
            }
            _staticParticles.Clear();
        }

        private void RestoreCanvasBackground()
        {
            _canvas.Background = _originalCanvasBackground;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (_canvas == null || !_canvas.IsVisible) return;

            double canvasWidth = _canvas.ActualWidth;
            double canvasHeight = _canvas.ActualHeight;
            if (canvasWidth < 10 || canvasHeight < 10) return;

            switch (_currentTheme)
            {
                case 0: SpawnSnowflake(canvasWidth, canvasHeight); break;
                case 1: SpawnShootingStar(canvasWidth, canvasHeight); break;
            }
        }

        // ==================== THEME 0: SNOW ====================
        private void SpawnSnowflake(double canvasWidth, double canvasHeight)
        {
            // Larger flakes get a detailed 6-pointed crystal shape,
            // smaller ones stay as soft round dots for depth.
            double size = _random.Next(3, 9);
            bool isCrystal = size >= 5;

            UIElement flake;
            if (isCrystal)
            {
                var crystal = CreateSnowflakeShape(size);
                crystal.Opacity = 0;
                flake = crystal;
            }
            else
            {
                var dot = new Ellipse
                {
                    Width = size,
                    Height = size,
                    Fill = new SolidColorBrush(Color.FromArgb(
                        (byte)_random.Next(120, 200), 255, 255, 255)),
                    Opacity = 0
                };
                flake = dot;
            }

            double startX = _random.Next(0, (int)canvasWidth);
            double startY = -10;
            double fallDuration = _random.Next(3, 8);
            double swayDistance = _random.Next(15, 40);

            Canvas.SetLeft(flake, startX);
            Canvas.SetTop(flake, startY);
            Canvas.SetZIndex(flake, 100);
            _canvas.Children.Add(flake);
            _activeParticles.Add(flake);

            var fallAnimation = new DoubleAnimation
            {
                From = startY,
                To = canvasHeight + 10,
                Duration = TimeSpan.FromSeconds(fallDuration)
            };
            fallAnimation.Completed += (s, args) => RemoveParticle(flake);
            flake.BeginAnimation(Canvas.TopProperty, fallAnimation);

            var swayAnimation = new DoubleAnimation
            {
                From = startX,
                To = startX + (_random.Next(0, 200) < 100 ? -swayDistance : swayDistance),
                Duration = TimeSpan.FromSeconds(fallDuration * 0.8),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever
            };
            flake.BeginAnimation(Canvas.LeftProperty, swayAnimation);

            // Crystals slowly rotate as they fall
            if (isCrystal)
            {
                var rotateTransform = new RotateTransform(0);
                flake.RenderTransform = rotateTransform;
                flake.RenderTransformOrigin = new Point(0.5, 0.5);
                var rotation = new DoubleAnimation
                {
                    From = 0,
                    To = _random.Next(0, 200) < 100 ? 360 : -360,
                    Duration = TimeSpan.FromSeconds(fallDuration),
                    RepeatBehavior = RepeatBehavior.Forever
                };
                rotateTransform.BeginAnimation(RotateTransform.AngleProperty, rotation);
            }

            var fadeIn = new DoubleAnimation
            {
                From = 0,
                To = 0.8,
                Duration = TimeSpan.FromMilliseconds(500)
            };
            flake.BeginAnimation(UIElement.OpacityProperty, fadeIn);
        }

        /// <summary>
        /// Creates a 6-pointed snowflake crystal shape using a polygon.
        /// </summary>
        private Polygon CreateSnowflakeShape(double size)
        {
            // 6-pointed star: outer points at 0/60/120/180/240/300 degrees,
            // inner concave points at 30/90/150/210/270/330 degrees.
            var points = new PointCollection
            {
                new Point(0.5, 0),        // outer top
                new Point(0.625, 0.2835), // inner
                new Point(0.933, 0.25),   // outer
                new Point(0.75, 0.5),     // inner
                new Point(0.933, 0.75),   // outer
                new Point(0.625, 0.7165), // inner
                new Point(0.5, 1),        // outer bottom
                new Point(0.375, 0.7165), // inner
                new Point(0.067, 0.75),   // outer
                new Point(0.25, 0.5),     // inner
                new Point(0.067, 0.25),   // outer
                new Point(0.375, 0.2835)  // inner
            };

            // Scale points to the flake size
            for (int i = 0; i < points.Count; i++)
            {
                points[i] = new Point(points[i].X * size, points[i].Y * size);
            }

            return new Polygon
            {
                Points = points,
                Fill = new SolidColorBrush(Color.FromArgb(
                    (byte)_random.Next(150, 220), 255, 255, 255)),
                Stroke = new SolidColorBrush(Color.FromArgb(
                    (byte)_random.Next(120, 180), 255, 255, 255)),
                StrokeThickness = 0.5
            };
        }
        // ==================== THEME 1: STARRY NIGHT ====================
        private void CreateStarryNight()
        {
            if (_canvas == null) return;
            double canvasWidth = _canvas.ActualWidth;
            double canvasHeight = _canvas.ActualHeight;
            if (canvasWidth < 10 || canvasHeight < 10) return;

            // Pitch black background
            _canvas.Background = new SolidColorBrush(Color.FromRgb(2, 2, 8));

            // Create a dense field of stars with varying sizes and brightness
            int starCount = Math.Max(60, (int)(canvasWidth * canvasHeight / 9000));
            for (int i = 0; i < starCount; i++)
            {
                double size = _random.Next(1, 5) * 0.8; // 0.8 to 4.0
                bool isLarge = _random.Next(0, 100) < 12; // ~12% large stars
                if (isLarge) size = _random.Next(5, 9);

                var star = CreateStarShape(size);

                double x = _random.Next(0, (int)canvasWidth);
                double y = _random.Next(0, (int)canvasHeight);

                Canvas.SetLeft(star, x);
                Canvas.SetTop(star, y);
                Canvas.SetZIndex(star, 50);
                _canvas.Children.Add(star);
                _staticParticles.Add(star);

                // Some stars shine brighter (twinkle), others stay steady
                if (_random.Next(0, 100) < 55)
                {
                    double baseOpacity = _random.Next(3, 8) / 10.0;
                    double peakOpacity = Math.Min(1.0, baseOpacity + _random.Next(2, 5) / 10.0);
                    star.Opacity = baseOpacity;

                    var twinkle = new DoubleAnimation
                    {
                        From = baseOpacity,
                        To = peakOpacity,
                        Duration = TimeSpan.FromMilliseconds(_random.Next(250, 800)),
                        AutoReverse = true,
                        RepeatBehavior = RepeatBehavior.Forever
                    };
                    star.BeginAnimation(UIElement.OpacityProperty, twinkle);
                }
                else
                {
                    star.Opacity = _random.Next(4, 9) / 10.0;
                }
            }

            // Add a subtle nebula glow in a few spots
            for (int i = 0; i < 3; i++)
            {
                double glowSize = _random.Next(150, 300);
                var glow = new Ellipse
                {
                    Width = glowSize,
                    Height = glowSize,
                    Fill = new RadialGradientBrush
                    {
                        GradientStops = new GradientStopCollection
                        {
                            new GradientStop(Color.FromArgb(25, 80, 100, 180), 0),
                            new GradientStop(Color.FromArgb(0, 80, 100, 180), 1)
                        }
                    },
                    Opacity = 0.6
                };

                Canvas.SetLeft(glow, _random.Next(-100, (int)canvasWidth));
                Canvas.SetTop(glow, _random.Next(-100, (int)canvasHeight));
                Canvas.SetZIndex(glow, 5);
                _canvas.Children.Add(glow);
                _staticParticles.Add(glow);
            }
        }

        /// <summary>
        /// Creates a 4-pointed sparkle star shape using a polygon.
        /// </summary>
        private Polygon CreateStarShape(double size)
        {
            // 4-pointed sparkle star: top, right, bottom, left with concave inner points
            var points = new PointCollection
            {
                new Point(0.5, 0),      // top
                new Point(0.62, 0.38),  // inner top-right
                new Point(1, 0.5),      // right
                new Point(0.62, 0.62),  // inner bottom-right
                new Point(0.5, 1),      // bottom
                new Point(0.38, 0.62),  // inner bottom-left
                new Point(0, 0.5),      // left
                new Point(0.38, 0.38)   // inner top-left
            };

            // Scale points to the star size
            for (int i = 0; i < points.Count; i++)
            {
                points[i] = new Point(points[i].X * size, points[i].Y * size);
            }

            // Warm white with slight blue tint for realism
            byte warmth = (byte)_random.Next(220, 256);
            var star = new Polygon
            {
                Points = points,
                Fill = new SolidColorBrush(Color.FromArgb(255, warmth, warmth, 255)),
                Stroke = new SolidColorBrush(Color.FromArgb(200, warmth, warmth, 255)),
                StrokeThickness = 0.5
            };

            return star;
        }

        private void SpawnShootingStar(double canvasWidth, double canvasHeight)
        {
            // Only spawn once every 60 seconds
            if ((DateTime.Now - _lastShootingStarTime).TotalSeconds < 60) return;
            _lastShootingStarTime = DateTime.Now;

            double length = 60;
            double angle = 3; // nearly horizontal

            // Start position in the upper portion of the canvas
            double startX = canvasWidth * 0.000005;
            double startY = canvasHeight * 0.5;
            double travelX = canvasWidth * 0.9;
            double travelY = travelX * Math.Tan(angle * Math.PI / 180) * 0.5;

            // the star stays within the canvas vertically
            if (startY + travelY > canvasHeight - 5)
                travelY = canvasHeight - 5 - startY;
            if (travelY < 0) travelY = 0;

            var trail = new Line
            {
                X1 = 0,
                Y1 = 0,
                X2 = -length,
                Y2 = length * Math.Tan(angle * Math.PI / 180) * 0.5,
                Stroke = new LinearGradientBrush
                {
                    StartPoint = new Point(0, 0),
                    EndPoint = new Point(1, 0),
                    GradientStops = new GradientStopCollection
                    {
                        new GradientStop(Color.FromArgb(0, 255, 255, 255), 0),
                        new GradientStop(Color.FromArgb(255, 255, 255, 255), 1)
                    }
                },
                StrokeThickness = 1,
                Opacity = 0
            };

            Canvas.SetLeft(trail, startX);
            Canvas.SetTop(trail, startY);
            Canvas.SetZIndex(trail, 70);
            _canvas.Children.Add(trail);
            _activeParticles.Add(trail);

            // Smooth fade in
            var fadeIn = new DoubleAnimation
            {
                From = 0,
                To = 1,
                Duration = TimeSpan.FromMilliseconds(1000)
            };
            trail.BeginAnimation(UIElement.OpacityProperty, fadeIn);

            // Move horizontally across the canvas
            var moveTrail = new DoubleAnimation
            {
                From = startX,
                To = startX + travelX,
                Duration = TimeSpan.FromSeconds(6.0)
            };
            moveTrail.Completed += (s, args) =>
            {
                RemoveParticle(trail);
            };
            trail.BeginAnimation(Canvas.LeftProperty, moveTrail);

            // Parabolic arc motion using a loop of discrete keyframes:
            // the star moves forward, rises gradually to a peak,
            // then descends gradually back down.
            var arcY = new DoubleAnimationUsingKeyFrames
            {
                Duration = TimeSpan.FromSeconds(6.0)
            };

            int steps = 30;
            for (int s = 0; s <= steps; s++)
            {
                double t = (double)s / steps; // 0.0 to 1.0
                // Parabola: peak rise of 2px at t=0.5, then settles down
                double yOffset = -2 * (1 - Math.Pow(2 * t - 1, 2)) + 1 * t;
                double y = startY + yOffset;
                arcY.KeyFrames.Add(new LinearDoubleKeyFrame(y, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(6.0 * t))));
            }
            trail.BeginAnimation(Canvas.TopProperty, arcY);

            // Smooth fade out near the end
            var fadeOut = new DoubleAnimation
            {
                From = 1,
                To = 0,
                BeginTime = TimeSpan.FromMilliseconds(4500),
                Duration = TimeSpan.FromMilliseconds(1500)
            };
            trail.BeginAnimation(UIElement.OpacityProperty, fadeOut);
        }

        private void RemoveParticle(UIElement particle)
        {
            _canvas.Children.Remove(particle);
            _activeParticles.Remove(particle);
        }
    }
}