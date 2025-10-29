using Size = System.Drawing.Size;
using Point = System.Drawing.Point;
using System.Collections.ObjectModel;
using ClickType = PointAC.Services.MouseService.ClickType;
using MouseButton = PointAC.Services.MouseService.MouseButton;
using PointAC.Miscellaneous;

namespace PointAC.Core
{
    public class PointManager
    {
        private readonly Size PointSize = new Size(32, 32);

        public bool IsRuntimeMode { get; private set; } = false;
        public ObservableCollection<PointEntry> Points { get; } = new();

        public PointManager()
        {
            PointsOverlay.Instance.RendererRecreated += OnRendererRecreated;
        }

        #region Events
        private void OnRendererRecreated()
        {
            foreach (var point in Points)
            {
                var centeredPoint = new System.Drawing.Point(
                    point.Position.X - PointSize.Width / 2,
                    point.Position.Y - PointSize.Height / 2
                );

                point.Handle = PointsOverlay.Instance.Add("Assets/Target.png", centeredPoint, PointSize);
            }
        }
        #endregion

        #region Exposed Methods
        public Guid AddPoint(string? customImage, Point position, MouseButton button, ClickType clickType, int duration)
        {
            var centeredPoint = new Point(
                position.X - PointSize.Width / 2,
                position.Y - PointSize.Height / 2
            );

            string imageSource = RandomUtilities.IsValidBitmap(customImage)
                ? customImage!
                : "pack://application:,,,/PointAC;component/Assets/Target.png";

            var handle = PointsOverlay.Instance.Add(imageSource, centeredPoint, PointSize);
            
            int nextOrder = Points.Any() ? Points.Max(p => p.Order) + 1 : 0;

            var entry = new PointEntry(handle, button, clickType, position, duration, nextOrder)
            {
                OnOrderChanged = HandleOrderChanged
            };

            Points.Add(entry);
            return handle;
        }

        public bool RemoveNearest(Point target, int radius = 15)
        {
            if (Points.Count == 0)
                return false;

            var nearest = Points.OrderBy(p => Distance(p.Position, target)).First();
            if (Distance(nearest.Position, target) <= radius)
            {
                PointsOverlay.Instance.Remove(nearest.Handle);
                Points.Remove(nearest);
                return true;
            }

            return false;
        }

        public void ClearAll()
        {
            foreach (var p in Points)
                PointsOverlay.Instance.Remove(p.Handle);

            Points.Clear();
        }

        public void SetRuntimeMode(bool enabled)
        {
            IsRuntimeMode = enabled;

            foreach (var p in Points)
            {
                p.IsRuntimeLocked = enabled;
                if (enabled)
                    p.IsHovered = false;
            }
        }
        #endregion

        #region Private Methods
        private void HandleOrderChanged(PointEntry entry, int newOrder)
        {
            if (IsOrderUnique(newOrder, entry))
            {
                entry.SetOrderInternal(newOrder);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"Duplicate order {newOrder} ignored for {entry.Handle}");
            }
        }

        public bool IsOrderUnique(int order, PointEntry? exclude = null)
        {
            return !Points.Any(p => p != exclude && p.Order == order);
        }

        private static double Distance(Point a, Point b)
        {
            int dx = a.X - b.X;
            int dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }
        #endregion

        #region Exposed Events
        public event Action<PointEntry>? PointAdded;
        public event Action<PointEntry>? PointRemoved;
        #endregion
    }
}