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

        public async Task<bool> ProcessVideoWithEditsAsync(
            string inputVideoPath,
            string outputVideoPath,
            IEnumerable<VideoEdit> edits,
            CancellationToken cancellationToken = default)
        {
            try
            {
                // FFmpegバイナリの初期化
                await EnsureFFmpegAsync();

                // 編集リストを時間順にソート
                var sortedEdits = edits.OrderBy(e => e.StartTime).ToList();

                // 編集タイプごとに処理を分岐
                var hasSplit = sortedEdits.Any(e => e.Type == EditType.Split);
                var hasTrim = sortedEdits.Any(e => e.Type == EditType.Trim);
                var hasCut = sortedEdits.Any(e => e.Type == EditType.Cut);
                var hasSpeedChange = sortedEdits.Any(e => e.Type == EditType.SpeedChange);

                // トリミングが指定されている場合
                if (hasTrim)
                {
                    var trimEdit = sortedEdits.First(e => e.Type == EditType.Trim);
                    return await ProcessTrimAsync(inputVideoPath, outputVideoPath, trimEdit, cancellationToken);
                }

                // 分割が指定されている場合
                if (hasSplit)
                {
                    return await ProcessSplitAsync(inputVideoPath, outputVideoPath, sortedEdits.Where(e => e.Type == EditType.Split), cancellationToken);
                }

                // カットまたは速度変更の場合
                if (hasCut || hasSpeedChange)
                {
                    return await ProcessCutAndSpeedAsync(inputVideoPath, outputVideoPath, sortedEdits, cancellationToken);
                }

                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Video processing error: {ex.Message}");
                return false;
            }
        }

        private async Task<bool> ProcessTrimAsync(
            string inputVideoPath,
            string outputVideoPath,
            VideoEdit trimEdit,
            CancellationToken cancellationToken)
        {
            var mediaInfo = await FFmpeg.GetMediaInfo(inputVideoPath);
            var conversion = FFmpeg.Conversions.New();

            var startTime = TimeSpan.FromSeconds(trimEdit.StartTime);
            var duration = TimeSpan.FromSeconds(trimEdit.EndTime - trimEdit.StartTime);

            conversion.AddParameter($"-i \"{inputVideoPath}\"");
            conversion.AddParameter($"-ss {startTime:hh\\:mm\\:ss\\.ff}");
            conversion.AddParameter($"-t {duration:hh\\:mm\\:ss\\.ff}");
            conversion.AddParameter("-c:v libx264");
            conversion.AddParameter("-preset medium");
            conversion.AddParameter("-crf 23");
            conversion.AddParameter("-c:a aac");
            conversion.AddParameter("-b:a 192k");
            conversion.AddParameter($"\"{outputVideoPath}\"");

            conversion.OnProgress += (sender, args) =>
            {
                var progress = args.Duration.TotalSeconds / args.TotalLength.TotalSeconds * 100;
                ProgressChanged?.Invoke(this, progress);
            };

            await conversion.Start(cancellationToken);
            return true;
        }

        private async Task<bool> ProcessSplitAsync(
            string inputVideoPath,
            string outputVideoPath,
            IEnumerable<VideoEdit> splitEdits,
            CancellationToken cancellationToken)
        {
            var mediaInfo = await FFmpeg.GetMediaInfo(inputVideoPath);
            var duration = mediaInfo.Duration;

            // 分割点を取得
            var splitPoints = splitEdits.Select(e => e.StartTime).OrderBy(t => t).ToList();
            splitPoints.Insert(0, 0); // 開始点を追加
            splitPoints.Add(duration.TotalSeconds); // 終了点を追加

            var tempFiles = new List<string>();
            var outputDir = Path.GetDirectoryName(outputVideoPath) ?? Path.GetTempPath();
            var outputFileName = Path.GetFileNameWithoutExtension(outputVideoPath);

            // 各区間を個別のファイルとして出力
            for (int i = 0; i < splitPoints.Count - 1; i++)
            {
                var startTime = TimeSpan.FromSeconds(splitPoints[i]);
                var endTime = TimeSpan.FromSeconds(splitPoints[i + 1]);
                var segmentDuration = endTime - startTime;

                var tempFile = Path.Combine(outputDir, $"{outputFileName}_part{i + 1}.mp4");
                tempFiles.Add(tempFile);

                var conversion = FFmpeg.Conversions.New();
                conversion.AddParameter($"-i \"{inputVideoPath}\"");
                conversion.AddParameter($"-ss {startTime:hh\\:mm\\:ss\\.ff}");
                conversion.AddParameter($"-t {segmentDuration:hh\\:mm\\:ss\\.ff}");
                conversion.AddParameter("-c:v libx264");
                conversion.AddParameter("-preset medium");
                conversion.AddParameter("-crf 23");
                conversion.AddParameter("-c:a aac");
                conversion.AddParameter("-b:a 192k");
                conversion.AddParameter($"\"{tempFile}\"");

                var progress = (i * 100.0) / (splitPoints.Count - 1);
                ProgressChanged?.Invoke(this, progress);

                await conversion.Start(cancellationToken);
            }

            // 最初のファイルを出力ファイルにリネーム
            if (tempFiles.Count > 0)
            {
                File.Move(tempFiles[0], outputVideoPath, true);
                tempFiles.RemoveAt(0);
            }

            ProgressChanged?.Invoke(this, 100);
            return true;
        }

        private async Task<bool> ProcessCutAndSpeedAsync(
            string inputVideoPath,
            string outputVideoPath,
            List<VideoEdit> edits,
            CancellationToken cancellationToken)
        {
            var mediaInfo = await FFmpeg.GetMediaInfo(inputVideoPath);
            var duration = mediaInfo.Duration.TotalSeconds;

            // カット範囲と速度変更範囲を統合して、保持する区間を生成
            var segments = GenerateSegments(duration, edits);

            if (segments.Count == 0)
            {
                return false;
            }

            // 各セグメントを処理
            var tempFiles = new List<string>();
            for (int i = 0; i < segments.Count; i++)
            {
                var segment = segments[i];
                var tempFile = Path.Combine(Path.GetTempPath(), $"segment_{Guid.NewGuid()}.mp4");
                tempFiles.Add(tempFile);

                var conversion = FFmpeg.Conversions.New();
                conversion.AddParameter($"-i \"{inputVideoPath}\"");
                conversion.AddParameter($"-ss {segment.StartTime:F3}");
                conversion.AddParameter($"-to {segment.EndTime:F3}");

                // 速度変更
                if (Math.Abs(segment.SpeedRatio - 1.0) > 0.01)
                {
                    conversion.AddParameter($"-filter:v \"setpts={1.0 / segment.SpeedRatio}*PTS\"");
                    conversion.AddParameter($"-filter:a \"atempo={segment.SpeedRatio}\"");
                }

                conversion.AddParameter("-c:v libx264");
                conversion.AddParameter("-preset medium");
                conversion.AddParameter("-crf 23");
                conversion.AddParameter("-c:a aac");
                conversion.AddParameter("-b:a 192k");
                conversion.AddParameter($"\"{tempFile}\"");

                var progress = (i * 80.0) / segments.Count;
                ProgressChanged?.Invoke(this, progress);

                await conversion.Start(cancellationToken);
            }

            // セグメントを結合
            if (tempFiles.Count == 1)
            {
                File.Move(tempFiles[0], outputVideoPath, true);
            }
            else
            {
                await ConcatenateVideosAsync(tempFiles, outputVideoPath, cancellationToken);
            }

            // 一時ファイルを削除
            foreach (var tempFile in tempFiles)
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }

            ProgressChanged?.Invoke(this, 100);
            return true;
        }

        private List<VideoClip> GenerateSegments(double duration, List<VideoEdit> edits)
        {
            var segments = new List<VideoClip>();
            var currentTime = 0.0;

            // カット範囲をソート
            var cutEdits = edits.Where(e => e.Type == EditType.Cut).OrderBy(e => e.StartTime).ToList();

            while (currentTime < duration)
            {
                // 現在位置から次のカット開始までの区間を追加
                var nextCut = cutEdits.FirstOrDefault(e => e.StartTime >= currentTime);

                if (nextCut != null)
                {
                    // カット開始前まで
                    if (nextCut.StartTime > currentTime)
                    {
                        var segment = new VideoClip
                        {
                            StartTime = currentTime,
                            EndTime = nextCut.StartTime,
                            SpeedRatio = GetSpeedRatioForTime(currentTime, edits)
                        };
                        segments.Add(segment);
                    }
                    currentTime = nextCut.EndTime;
                }
                else
                {
                    // 最後の区間
                    var segment = new VideoClip
                    {
                        StartTime = currentTime,
                        EndTime = duration,
                        SpeedRatio = GetSpeedRatioForTime(currentTime, edits)
                    };
                    segments.Add(segment);
                    break;
                }
            }

            return segments;
        }

        private double GetSpeedRatioForTime(double time, List<VideoEdit> edits)
        {
            var speedEdit = edits.FirstOrDefault(e =>
                e.Type == EditType.SpeedChange &&
                e.StartTime <= time &&
                e.EndTime >= time);

            return speedEdit?.SpeedRatio ?? 1.0;
        }

        private async Task ConcatenateVideosAsync(List<string> inputFiles, string outputFile, CancellationToken cancellationToken)
        {
            // FFmpegのconcat demuxerを使用
            var concatFile = Path.Combine(Path.GetTempPath(), $"concat_{Guid.NewGuid()}.txt");
            using (var writer = new StreamWriter(concatFile))
            {
                foreach (var file in inputFiles)
                {
                    writer.WriteLine($"file '{file}'");
                }
            }

            var conversion = FFmpeg.Conversions.New();
            conversion.AddParameter($"-f concat");
            conversion.AddParameter($"-safe 0");
            conversion.AddParameter($"-i \"{concatFile}\"");
            conversion.AddParameter("-c copy");
            conversion.AddParameter($"\"{outputFile}\"");

            ProgressChanged?.Invoke(this, 90);

            await conversion.Start(cancellationToken);

            if (File.Exists(concatFile))
            {
                File.Delete(concatFile);
            }
        }
    }
}
