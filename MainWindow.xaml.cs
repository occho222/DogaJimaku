using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;
using DogaJimaku.Models;
using DogaJimaku.Services;
using Microsoft.Win32;

namespace DogaJimaku
{
    public partial class MainWindow : Window
    {
        private ObservableCollection<Subtitle> _subtitles = new();
        private DispatcherTimer? _timer;
        private bool _isPlaying = false;
        private bool _isDraggingSeekBar = false;
        private Subtitle? _currentEditingSubtitle = null;
        private string? _currentVideoPath = null;
        private CancellationTokenSource? _exportCancellationTokenSource;

        public MainWindow()
        {
            InitializeComponent();
            SubtitleListBox.ItemsSource = _subtitles;

            // タイマー設定（字幕の同期と時刻表示用）
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromMilliseconds(50);
            _timer.Tick += Timer_Tick;
        }

        private void Timer_Tick(object? sender, EventArgs e)
        {
            if (!_isDraggingSeekBar && VideoPlayer.Source != null)
            {
                // シークバー更新
                if (VideoPlayer.NaturalDuration.HasTimeSpan)
                {
                    var totalSeconds = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
                    var currentSeconds = VideoPlayer.Position.TotalSeconds;
                    if (totalSeconds > 0)
                    {
                        SeekBar.Value = (currentSeconds / totalSeconds) * 100;
                    }
                }

                // 時刻表示更新
                TxtCurrentTime.Text = VideoPlayer.Position.ToString(@"hh\:mm\:ss");

                // 字幕表示更新
                UpdateSubtitleDisplay(VideoPlayer.Position.TotalSeconds);
            }
        }

        private void UpdateSubtitleDisplay(double currentTime)
        {
            var activeSubtitle = _subtitles.FirstOrDefault(s =>
                s.StartTime <= currentTime && s.EndTime >= currentTime);

            if (activeSubtitle != null)
            {
                SubtitleText.Text = activeSubtitle.Text;
                SubtitleText.FontSize = activeSubtitle.FontSize;
                SubtitleText.Foreground = new SolidColorBrush(activeSubtitle.TextColor);
                UpdateSubtitlePosition(activeSubtitle.Position);
                SubtitleText.Visibility = Visibility.Visible;
            }
            else
            {
                SubtitleText.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdateSubtitlePosition(SubtitlePosition position)
        {
            switch (position)
            {
                case SubtitlePosition.BottomCenter:
                    SubtitleText.VerticalAlignment = VerticalAlignment.Bottom;
                    SubtitleText.HorizontalAlignment = HorizontalAlignment.Center;
                    SubtitleText.Margin = new Thickness(40, 0, 40, 40);
                    break;
                case SubtitlePosition.TopCenter:
                    SubtitleText.VerticalAlignment = VerticalAlignment.Top;
                    SubtitleText.HorizontalAlignment = HorizontalAlignment.Center;
                    SubtitleText.Margin = new Thickness(40, 40, 40, 0);
                    break;
                case SubtitlePosition.Center:
                    SubtitleText.VerticalAlignment = VerticalAlignment.Center;
                    SubtitleText.HorizontalAlignment = HorizontalAlignment.Center;
                    SubtitleText.Margin = new Thickness(40, 0, 40, 0);
                    break;
                case SubtitlePosition.BottomLeft:
                    SubtitleText.VerticalAlignment = VerticalAlignment.Bottom;
                    SubtitleText.HorizontalAlignment = HorizontalAlignment.Left;
                    SubtitleText.Margin = new Thickness(40, 0, 40, 40);
                    break;
                case SubtitlePosition.BottomRight:
                    SubtitleText.VerticalAlignment = VerticalAlignment.Bottom;
                    SubtitleText.HorizontalAlignment = HorizontalAlignment.Right;
                    SubtitleText.Margin = new Thickness(40, 0, 40, 40);
                    break;
            }
        }

        private void BtnOpenVideo_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "動画ファイル|*.mp4;*.avi;*.mkv;*.mov;*.wmv;*.flv;*.webm|すべてのファイル|*.*",
                Title = "動画ファイルを選択"
            };

            if (dialog.ShowDialog() == true)
            {
                _currentVideoPath = dialog.FileName;
                VideoPlayer.Source = new Uri(_currentVideoPath);
                NoVideoMessage.Visibility = Visibility.Collapsed;
                BtnPlayPause.IsEnabled = true;
                BtnStop.IsEnabled = true;
                BtnAddSubtitle.IsEnabled = true;
                BtnExportVideo.IsEnabled = true;
                BtnBackward5.IsEnabled = true;
                BtnBackward1.IsEnabled = true;
                BtnForward1.IsEnabled = true;
                BtnForward5.IsEnabled = true;
                BtnSpeedSlow.IsEnabled = true;
                BtnSpeedNormal.IsEnabled = true;
                BtnSpeedFast.IsEnabled = true;
            }
        }

        private void VideoPlayer_MediaOpened(object sender, RoutedEventArgs e)
        {
            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                TxtDuration.Text = VideoPlayer.NaturalDuration.TimeSpan.ToString(@"hh\:mm\:ss");
                SeekBar.Maximum = 100;
            }
        }

        private void VideoPlayer_MediaEnded(object sender, RoutedEventArgs e)
        {
            _isPlaying = false;
            _timer?.Stop();
            BtnPlayPause.Content = "▶️ 再生";
            VideoPlayer.Position = TimeSpan.Zero;
        }

        private void BtnPlayPause_Click(object sender, RoutedEventArgs e)
        {
            if (_isPlaying)
            {
                VideoPlayer.Pause();
                _timer?.Stop();
                BtnPlayPause.Content = "▶️ 再生";
            }
            else
            {
                VideoPlayer.Play();
                _timer?.Start();
                BtnPlayPause.Content = "⏸️ 一時停止";
            }
            _isPlaying = !_isPlaying;
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            VideoPlayer.Stop();
            _timer?.Stop();
            _isPlaying = false;
            BtnPlayPause.Content = "▶️ 再生";
            VideoPlayer.Position = TimeSpan.Zero;
            SeekBar.Value = 0;
        }

        private void BtnBackward5_Click(object sender, RoutedEventArgs e)
        {
            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                var newPosition = VideoPlayer.Position - TimeSpan.FromSeconds(5);
                VideoPlayer.Position = newPosition < TimeSpan.Zero ? TimeSpan.Zero : newPosition;
            }
        }

        private void BtnBackward1_Click(object sender, RoutedEventArgs e)
        {
            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                var newPosition = VideoPlayer.Position - TimeSpan.FromSeconds(1);
                VideoPlayer.Position = newPosition < TimeSpan.Zero ? TimeSpan.Zero : newPosition;
            }
        }

        private void BtnForward1_Click(object sender, RoutedEventArgs e)
        {
            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                var newPosition = VideoPlayer.Position + TimeSpan.FromSeconds(1);
                var duration = VideoPlayer.NaturalDuration.TimeSpan;
                VideoPlayer.Position = newPosition > duration ? duration : newPosition;
            }
        }

        private void BtnForward5_Click(object sender, RoutedEventArgs e)
        {
            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                var newPosition = VideoPlayer.Position + TimeSpan.FromSeconds(5);
                var duration = VideoPlayer.NaturalDuration.TimeSpan;
                VideoPlayer.Position = newPosition > duration ? duration : newPosition;
            }
        }

        private void BtnSpeedSlow_Click(object sender, RoutedEventArgs e)
        {
            VideoPlayer.SpeedRatio = 0.5;
        }

        private void BtnSpeedNormal_Click(object sender, RoutedEventArgs e)
        {
            VideoPlayer.SpeedRatio = 1.0;
        }

        private void BtnSpeedFast_Click(object sender, RoutedEventArgs e)
        {
            VideoPlayer.SpeedRatio = 2.0;
        }

        private void SeekBar_DragStarted(object sender, DragStartedEventArgs e)
        {
            _isDraggingSeekBar = true;
        }

        private void SeekBar_DragCompleted(object sender, DragCompletedEventArgs e)
        {
            _isDraggingSeekBar = false;
            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                var totalSeconds = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
                var newPosition = TimeSpan.FromSeconds((SeekBar.Value / 100) * totalSeconds);
                VideoPlayer.Position = newPosition;
            }
        }

        private void SeekBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isDraggingSeekBar && VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                var totalSeconds = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
                var newPosition = TimeSpan.FromSeconds((SeekBar.Value / 100) * totalSeconds);
                TxtCurrentTime.Text = newPosition.ToString(@"hh\:mm\:ss");
            }
        }

        private void BtnAddSubtitle_Click(object sender, RoutedEventArgs e)
        {
            var currentTime = VideoPlayer.Position.TotalSeconds;
            var newSubtitle = new Subtitle
            {
                Text = "",
                StartTime = currentTime,
                EndTime = currentTime + 3,
                Position = SubtitlePosition.BottomCenter,
                FontSize = 48,
                TextColor = Colors.Red
            };

            _subtitles.Add(newSubtitle);
            var sortedSubtitles = _subtitles.OrderBy(s => s.StartTime).ToList();
            _subtitles.Clear();
            foreach (var sub in sortedSubtitles)
            {
                _subtitles.Add(sub);
            }

            SubtitleListBox.SelectedItem = newSubtitle;
            TxtSubtitleText.Focus();

            BtnSave.IsEnabled = true;
            BtnExport.IsEnabled = true;
        }

        private void SubtitleListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SubtitleListBox.SelectedItem is Subtitle subtitle)
            {
                _currentEditingSubtitle = subtitle;
                TxtSubtitleText.Text = subtitle.Text;
                TxtStartTime.Text = subtitle.StartTime.ToString("F2");
                TxtEndTime.Text = subtitle.EndTime.ToString("F2");
                CmbPosition.SelectedIndex = (int)subtitle.Position;
                SliderFontSize.Value = subtitle.FontSize;

                // 文字色を設定
                var colorName = subtitle.TextColor.ToString();
                for (int i = 0; i < CmbTextColor.Items.Count; i++)
                {
                    if (CmbTextColor.Items[i] is ComboBoxItem item && item.Tag?.ToString() == colorName)
                    {
                        CmbTextColor.SelectedIndex = i;
                        break;
                    }
                }
            }
        }

        private void TxtSubtitleText_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_currentEditingSubtitle != null)
            {
                _currentEditingSubtitle.Text = TxtSubtitleText.Text;
            }
        }

        private void TxtStartTime_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_currentEditingSubtitle != null && double.TryParse(TxtStartTime.Text, out var startTime))
            {
                _currentEditingSubtitle.StartTime = startTime;
                var sortedSubtitles = _subtitles.OrderBy(s => s.StartTime).ToList();
                _subtitles.Clear();
                foreach (var sub in sortedSubtitles)
                {
                    _subtitles.Add(sub);
                }
            }
        }

        private void TxtEndTime_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_currentEditingSubtitle != null && double.TryParse(TxtEndTime.Text, out var endTime))
            {
                _currentEditingSubtitle.EndTime = endTime;
            }
        }

        private void BtnSetStartTime_Click(object sender, RoutedEventArgs e)
        {
            TxtStartTime.Text = VideoPlayer.Position.TotalSeconds.ToString("F2");
        }

        private void BtnSetEndTime_Click(object sender, RoutedEventArgs e)
        {
            TxtEndTime.Text = VideoPlayer.Position.TotalSeconds.ToString("F2");
        }

        private void CmbPosition_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_currentEditingSubtitle != null)
            {
                _currentEditingSubtitle.Position = (SubtitlePosition)CmbPosition.SelectedIndex;
            }
        }

        private void SliderFontSize_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_currentEditingSubtitle != null)
            {
                _currentEditingSubtitle.FontSize = SliderFontSize.Value;
            }
            if (TxtFontSizeValue != null)
            {
                TxtFontSizeValue.Text = $"{SliderFontSize.Value:F0}px";
            }
        }

        private void CmbTextColor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_currentEditingSubtitle != null && CmbTextColor.SelectedItem is ComboBoxItem item)
            {
                var colorName = item.Tag?.ToString();
                if (!string.IsNullOrEmpty(colorName))
                {
                    var color = (Color)ColorConverter.ConvertFromString(colorName);
                    _currentEditingSubtitle.TextColor = color;
                }
            }
        }

        private void BtnDeleteSubtitle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is Subtitle subtitle)
            {
                _subtitles.Remove(subtitle);
                if (_currentEditingSubtitle == subtitle)
                {
                    _currentEditingSubtitle = null;
                    TxtSubtitleText.Text = "";
                    TxtStartTime.Text = "";
                    TxtEndTime.Text = "";
                }
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "字幕プロジェクトファイル|*.djsub|すべてのファイル|*.*",
                Title = "字幕プロジェクトを保存",
                DefaultExt = ".djsub"
            };

            if (dialog.ShowDialog() == true)
            {
                var projectData = new
                {
                    VideoPath = _currentVideoPath,
                    Subtitles = _subtitles.Select(s => new
                    {
                        s.Text,
                        s.StartTime,
                        s.EndTime,
                        Position = s.Position.ToString(),
                        s.FontSize
                    })
                };

                var json = JsonSerializer.Serialize(projectData, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(dialog.FileName, json);
                MessageBox.Show("保存しました", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "SRT字幕ファイル|*.srt|すべてのファイル|*.*",
                Title = "字幕を出力",
                DefaultExt = ".srt"
            };

            if (dialog.ShowDialog() == true)
            {
                using var writer = new StreamWriter(dialog.FileName);
                int index = 1;
                foreach (var subtitle in _subtitles.OrderBy(s => s.StartTime))
                {
                    writer.WriteLine(index);
                    var startTime = TimeSpan.FromSeconds(subtitle.StartTime);
                    var endTime = TimeSpan.FromSeconds(subtitle.EndTime);
                    writer.WriteLine($"{startTime:hh\\:mm\\:ss\\,fff} --> {endTime:hh\\:mm\\:ss\\,fff}");
                    writer.WriteLine(subtitle.Text);
                    writer.WriteLine();
                    index++;
                }

                MessageBox.Show("字幕ファイルを出力しました", "出力完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private async void BtnExportVideo_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentVideoPath))
            {
                MessageBox.Show("動画が読み込まれていません", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_subtitles.Count == 0)
            {
                var result = MessageBox.Show("字幕が追加されていませんが、動画を出力しますか?", "確認",
                    MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result != MessageBoxResult.Yes)
                    return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "MP4動画ファイル|*.mp4|すべてのファイル|*.*",
                Title = "字幕入り動画を出力",
                DefaultExt = ".mp4",
                FileName = Path.GetFileNameWithoutExtension(_currentVideoPath) + "_subtitled.mp4"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    // UIの状態を変更
                    ProgressPanel.Visibility = Visibility.Visible;
                    ProgressBar.Value = 0;
                    TxtProgress.Text = "動画を処理中...";

                    // ボタンを無効化
                    BtnExportVideo.IsEnabled = false;
                    BtnOpenVideo.IsEnabled = false;
                    BtnAddSubtitle.IsEnabled = false;
                    BtnSave.IsEnabled = false;
                    BtnExport.IsEnabled = false;
                    BtnPlayPause.IsEnabled = false;
                    BtnStop.IsEnabled = false;

                    // 動画を停止
                    if (_isPlaying)
                    {
                        VideoPlayer.Pause();
                        _timer?.Stop();
                        _isPlaying = false;
                        BtnPlayPause.Content = "▶️ 再生";
                    }

                    _exportCancellationTokenSource = new CancellationTokenSource();
                    var exportService = new VideoExportService();

                    // 進捗イベント
                    exportService.ProgressChanged += (s, progress) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            ProgressBar.Value = Math.Min(progress, 100);
                            TxtProgress.Text = $"処理中... {progress:F0}%";
                        });
                    };

                    var success = await exportService.ExportVideoWithSubtitlesAsync(
                        _currentVideoPath,
                        dialog.FileName,
                        _subtitles,
                        _exportCancellationTokenSource.Token
                    );

                    if (success)
                    {
                        ProgressBar.Value = 100;
                        TxtProgress.Text = "完了!";
                        MessageBox.Show($"字幕入り動画を出力しました:\n{dialog.FileName}",
                            "出力完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("動画の出力に失敗しました。FFmpegが正しくインストールされているか確認してください。",
                            "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"エラーが発生しました: {ex.Message}",
                        "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    // UIを元に戻す
                    ProgressPanel.Visibility = Visibility.Collapsed;
                    ProgressBar.Value = 0;

                    BtnExportVideo.IsEnabled = true;
                    BtnOpenVideo.IsEnabled = true;
                    BtnAddSubtitle.IsEnabled = true;
                    BtnSave.IsEnabled = _subtitles.Count > 0;
                    BtnExport.IsEnabled = _subtitles.Count > 0;
                    BtnPlayPause.IsEnabled = true;
                    BtnStop.IsEnabled = true;

                    _exportCancellationTokenSource?.Dispose();
                    _exportCancellationTokenSource = null;
                }
            }
        }
    }
}