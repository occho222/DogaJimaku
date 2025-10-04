using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace DogaJimaku.Models
{
    public enum EditType
    {
        Cut,        // カット（削除）
        Trim,       // トリミング
        Split,      // 分割
        SpeedChange // 速度変更
    }

    public class VideoEdit : INotifyPropertyChanged
    {
        private EditType _type;
        private double _startTime;
        private double _endTime;
        private double _speedRatio = 1.0;
        private string _label = string.Empty;

        public EditType Type
        {
            get => _type;
            set
            {
                if (_type != value)
                {
                    _type = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(TypeLabel));
                }
            }
        }

        public double StartTime
        {
            get => _startTime;
            set
            {
                if (_startTime != value)
                {
                    _startTime = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(StartTimeFormatted));
                    OnPropertyChanged(nameof(Duration));
                }
            }
        }

        public double EndTime
        {
            get => _endTime;
            set
            {
                if (_endTime != value)
                {
                    _endTime = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(EndTimeFormatted));
                    OnPropertyChanged(nameof(Duration));
                }
            }
        }

        public double SpeedRatio
        {
            get => _speedRatio;
            set
            {
                if (_speedRatio != value)
                {
                    _speedRatio = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(SpeedLabel));
                }
            }
        }

        public string Label
        {
            get => _label;
            set
            {
                if (_label != value)
                {
                    _label = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StartTimeFormatted => TimeSpan.FromSeconds(_startTime).ToString(@"hh\:mm\:ss\.ff");
        public string EndTimeFormatted => TimeSpan.FromSeconds(_endTime).ToString(@"hh\:mm\:ss\.ff");
        public double Duration => _endTime - _startTime;

        public string TypeLabel => Type switch
        {
            EditType.Cut => "✂️ カット",
            EditType.Trim => "✂️ トリミング",
            EditType.Split => "✂️ 分割",
            EditType.SpeedChange => "⚡ 速度変更",
            _ => "編集"
        };

        public string SpeedLabel => $"{SpeedRatio:F1}x";

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class VideoClip
    {
        public string FilePath { get; set; } = string.Empty;
        public double StartTime { get; set; }
        public double EndTime { get; set; }
        public double SpeedRatio { get; set; } = 1.0;
    }
}
