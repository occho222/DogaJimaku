using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace DogaJimaku.Models
{
    public class Subtitle : INotifyPropertyChanged
    {
        private string _text = string.Empty;
        private double _startTime;
        private double _endTime;
        private SubtitlePosition _position = SubtitlePosition.BottomCenter;
        private double _fontSize = 48;
        private Color _textColor = Colors.Red;

        public string Text
        {
            get => _text;
            set
            {
                if (_text != value)
                {
                    _text = value;
                    OnPropertyChanged();
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
                }
            }
        }

        public SubtitlePosition Position
        {
            get => _position;
            set
            {
                if (_position != value)
                {
                    _position = value;
                    OnPropertyChanged();
                }
            }
        }

        public double FontSize
        {
            get => _fontSize;
            set
            {
                if (_fontSize != value)
                {
                    _fontSize = value;
                    OnPropertyChanged();
                }
            }
        }

        public Color TextColor
        {
            get => _textColor;
            set
            {
                if (_textColor != value)
                {
                    _textColor = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StartTimeFormatted => TimeSpan.FromSeconds(_startTime).ToString(@"hh\:mm\:ss\.ff");
        public string EndTimeFormatted => TimeSpan.FromSeconds(_endTime).ToString(@"hh\:mm\:ss\.ff");

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public enum SubtitlePosition
    {
        BottomCenter,
        TopCenter,
        Center,
        BottomLeft,
        BottomRight
    }
}
