using System.IO;
using System.Text;
using DogaJimaku.Models;
using Xabe.FFmpeg;

namespace DogaJimaku.Services
{
    public class VideoExportService
    {
        public event EventHandler<double>? ProgressChanged;

        public async Task<bool> ExportVideoWithSubtitlesAsync(
            string inputVideoPath,
            string outputVideoPath,
            IEnumerable<Subtitle> subtitles,
            CancellationToken cancellationToken = default)
        {
            try
            {
                // FFmpegバイナリの初期化
                await EnsureFFmpegAsync();

                // 一時的なASSファイルを作成
                var tempAssPath = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.ass");
                CreateAssFile(tempAssPath, subtitles);

                // FFmpegで字幕を焼き込み
                var mediaInfo = await FFmpeg.GetMediaInfo(inputVideoPath);
                var videoStream = mediaInfo.VideoStreams.FirstOrDefault();
                var audioStream = mediaInfo.AudioStreams.FirstOrDefault();

                if (videoStream == null)
                {
                    throw new Exception("動画ストリームが見つかりません");
                }

                var conversion = FFmpeg.Conversions.New();

                // 入力ファイルを追加
                conversion.AddParameter($"-i \"{inputVideoPath}\"");

                // 字幕フィルターを追加（Windowsパス対応）
                var escapedAssPath = tempAssPath.Replace("\\", "\\\\").Replace(":", "\\:");
                conversion.AddParameter($"-vf \"ass='{escapedAssPath}'\"");

                // ビデオとオーディオのコーデック設定
                conversion.AddParameter("-c:v libx264");
                conversion.AddParameter("-preset medium");
                conversion.AddParameter("-crf 23");

                if (audioStream != null)
                {
                    conversion.AddParameter("-c:a aac");
                    conversion.AddParameter("-b:a 192k");
                }

                // 出力ファイル
                conversion.AddParameter($"\"{outputVideoPath}\"");

                // 進捗イベント
                conversion.OnProgress += (sender, args) =>
                {
                    var progress = args.Duration.TotalSeconds / args.TotalLength.TotalSeconds * 100;
                    ProgressChanged?.Invoke(this, progress);
                };

                // 変換実行
                await conversion.Start(cancellationToken);

                // 一時ファイル削除
                if (File.Exists(tempAssPath))
                {
                    File.Delete(tempAssPath);
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Video export error: {ex.Message}");
                return false;
            }
        }

        private void CreateSrtFile(string path, IEnumerable<Subtitle> subtitles)
        {
            using var writer = new StreamWriter(path, false, Encoding.UTF8);
            int index = 1;

            foreach (var subtitle in subtitles.OrderBy(s => s.StartTime))
            {
                writer.WriteLine(index);
                var startTime = TimeSpan.FromSeconds(subtitle.StartTime);
                var endTime = TimeSpan.FromSeconds(subtitle.EndTime);
                writer.WriteLine($"{startTime:hh\\:mm\\:ss\\,fff} --> {endTime:hh\\:mm\\:ss\\,fff}");
                writer.WriteLine(subtitle.Text);
                writer.WriteLine();
                index++;
            }
        }

        private void CreateAssFile(string path, IEnumerable<Subtitle> subtitles)
        {
            using var writer = new StreamWriter(path, false, Encoding.UTF8);

            // ASSヘッダー
            writer.WriteLine("[Script Info]");
            writer.WriteLine("Title: DogaJimaku Subtitles");
            writer.WriteLine("ScriptType: v4.00+");
            writer.WriteLine("WrapStyle: 0");
            writer.WriteLine("PlayResX: 1920");
            writer.WriteLine("PlayResY: 1080");
            writer.WriteLine("ScaledBorderAndShadow: yes");
            writer.WriteLine();

            // スタイル定義（フォントサイズを2倍にしてプレビューと統一）
            writer.WriteLine("[V4+ Styles]");
            writer.WriteLine("Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding");

            // 下部中央 - フォントサイズ56 (28*2)
            writer.WriteLine("Style: BottomCenter,Yu Gothic UI,56,&H00FFFFFF,&H000000FF,&H00000000,&H00000000,-1,0,0,0,100,100,0,0,1,3,3,2,40,40,40,1");
            // 上部中央
            writer.WriteLine("Style: TopCenter,Yu Gothic UI,56,&H00FFFFFF,&H000000FF,&H00000000,&H00000000,-1,0,0,0,100,100,0,0,1,3,3,8,40,40,40,1");
            // 中央
            writer.WriteLine("Style: Center,Yu Gothic UI,56,&H00FFFFFF,&H000000FF,&H00000000,&H00000000,-1,0,0,0,100,100,0,0,1,3,3,5,40,40,40,1");
            // 左下
            writer.WriteLine("Style: BottomLeft,Yu Gothic UI,56,&H00FFFFFF,&H000000FF,&H00000000,&H00000000,-1,0,0,0,100,100,0,0,1,3,3,1,40,40,40,1");
            // 右下
            writer.WriteLine("Style: BottomRight,Yu Gothic UI,56,&H00FFFFFF,&H000000FF,&H00000000,&H00000000,-1,0,0,0,100,100,0,0,1,3,3,3,40,40,40,1");
            writer.WriteLine();

            // イベント
            writer.WriteLine("[Events]");
            writer.WriteLine("Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text");

            foreach (var subtitle in subtitles.OrderBy(s => s.StartTime))
            {
                var startTime = TimeSpan.FromSeconds(subtitle.StartTime);
                var endTime = TimeSpan.FromSeconds(subtitle.EndTime);

                var styleName = subtitle.Position switch
                {
                    SubtitlePosition.BottomCenter => "BottomCenter",
                    SubtitlePosition.TopCenter => "TopCenter",
                    SubtitlePosition.Center => "Center",
                    SubtitlePosition.BottomLeft => "BottomLeft",
                    SubtitlePosition.BottomRight => "BottomRight",
                    _ => "BottomCenter"
                };

                // フォントサイズをスタイルに反映（2倍にしてプレビューと統一）
                var fontSize = (int)(subtitle.FontSize * 2);

                // 文字色をASSフォーマットに変換（BGR形式）
                var color = subtitle.TextColor;
                var assColor = $"&H{color.B:X2}{color.G:X2}{color.R:X2}";

                // 常にインラインスタイルでフォントサイズと色を指定
                var text = $"{{\\fs{fontSize}\\c{assColor}}}{subtitle.Text}";
                writer.WriteLine($"Dialogue: 0,{startTime:h\\:mm\\:ss\\.ff},{endTime:h\\:mm\\:ss\\.ff},{styleName},,0,0,0,,{text}");
            }
        }

        private async Task EnsureFFmpegAsync()
        {
            var ffmpegPath = Path.Combine(Path.GetTempPath(), "FFmpeg");

            if (!Directory.Exists(ffmpegPath) || !File.Exists(Path.Combine(ffmpegPath, "ffmpeg.exe")))
            {
                Directory.CreateDirectory(ffmpegPath);
                await Xabe.FFmpeg.Downloader.FFmpegDownloader.GetLatestVersion(Xabe.FFmpeg.Downloader.FFmpegVersion.Official, ffmpegPath);
            }

            FFmpeg.SetExecutablesPath(ffmpegPath);
        }
    }
}
