using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;
using ClickType = PointAC.MouseHandler.ClickType;
using MouseButton = PointAC.MouseHandler.MouseButton;

namespace PointAC
{
    public class PointEntry : INotifyPropertyChanged
    {
        private int order;
        private bool isHovered;
        private bool isRuntimeLocked;

        [JsonIgnore]
        public Guid Handle { get; internal set; }

        [JsonIgnore]
        public MouseButton Button { get; internal set; }

        [JsonIgnore]
        public ClickType ClickType { get; internal set; }

        [JsonIgnore]
        public Action<PointEntry, int>? OnOrderChanged { get; set; }

        // Serializable properties
        public Point Position { get; set; }

        private int _duration;
        public int Duration
        {
            get => _duration;
            set
            {
                if (_duration != value)
                {
                    _duration = value;
                    OnPropertyChanged();
                }
            }
        }

        public int Order
        {
            get => order;
            set
            {
                if (order == value)
                    return;

                if (OnOrderChanged != null)
                    OnOrderChanged.Invoke(this, value);
                else
                {
                    order = value;
                    OnPropertyChanged();
                }
            }
        }
        public string ButtonString
        {
            get => MouseHandler.GetStringFromMouseButton(Button);
            set
            {
                var parsed = MouseHandler.GetMouseButtonFromString(value);
                if (Button != parsed)
                {
                    Button = parsed;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(ButtonString));
                }
            }
        }

        public string ClickTypeString
        {
            get => MouseHandler.GetStringFromClickType(ClickType);
            set
            {
                var parsed = MouseHandler.GetClickTypeFromString(value);
                if (ClickType != parsed)
                {
                    ClickType = parsed;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(ClickTypeString));
                }
            }
        }

        internal void SetOrderInternal(int newOrder)
        {
            if (order != newOrder)
            {
                order = newOrder;
                OnPropertyChanged(nameof(Order));
            }
        }

        [JsonIgnore]
        public bool IsHovered
        {
            get => isHovered;
            set
            {
                if (isHovered != value)
                {
                    isHovered = value;
                    OnPropertyChanged();
                }
            }
        }


        [JsonIgnore]
        public bool IsRuntimeLocked
        {
            get => isRuntimeLocked;
            set
            {
                if (isRuntimeLocked != value)
                {
                    isRuntimeLocked = value;
                    OnPropertyChanged();
                }
            }
        }

        public PointEntry(Guid handle, MouseButton button, ClickType clickType, Point position, int duration, int order)
        {
            Handle = handle;
            Button = button;
            ClickType = clickType;
            Position = position;
            Duration = duration;
            Order = order;
        }

        public PointEntry() { }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
