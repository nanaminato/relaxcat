using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using RelaxCat.Core.Export;
using RelaxCat.Core.Imaging;

namespace RelaxCat.App;

public sealed class MainWindow : Window
{
    private readonly List<ImportedImage> images = new();
    private readonly ListBox queue = new();
    private readonly Image preview = new() { Stretch = Stretch.Uniform, Margin = new Thickness(20) };
    private readonly TextBlock status = new() { TextWrapping = TextWrapping.Wrap };
    private readonly ProgressBar progress = new() { Height = 5, Minimum = 0, Maximum = 1 };
    private readonly StackPanel settings = new() { Spacing = 8, Margin = new Thickness(16) };
    private readonly StackPanel toolbar = new() { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new Thickness(12) };
    private readonly TextBox width = Input("1920"), height = Input("1080"), scale = Input("0.5"), enlargement = Input("1.5");
    private readonly TextBox horizontal = Input("12"), vertical = Input("12"), rowHeight = Input("18"), columnWidth = Input("10");
    private readonly TextBox template = Input("");
    private readonly TextBox rules = new() { AcceptsReturn = true, Height = 140, TextWrapping = TextWrapping.Wrap };
    private readonly CheckBox compact = new();
    private readonly ComboBox language = new() { ItemsSource = new[] { "简体中文", "日本語" }, SelectedIndex = 0, Width = 125 };
    private readonly ExcelSession excel = new();
    private readonly ImageExporter exporter = new();
    private CancellationTokenSource? operation;
    private Bitmap? selectedBitmap;
    private bool Japanese => language.SelectedIndex == 1;
    private string T(string zh, string ja) => Japanese ? ja : zh;
    private static TextBox Input(string text) => new() { Text = text };

    public MainWindow()
    {
        Title = "RelaxCat"; Width = 1180; Height = 800; MinWidth = 900; MinHeight = 650;
        Background = new SolidColorBrush(Color.Parse("#1E1F22"));
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*,Auto") };
        root.Children.Add(toolbar);
        var body = new Grid { ColumnDefinitions = new ColumnDefinitions("250,*,310") };
        Grid.SetRow(body, 1); root.Children.Add(body);
        queue.Margin = new Thickness(8); queue.SelectionChanged += (_, _) => ShowPreview(); body.Children.Add(queue);
        Grid.SetColumn(preview, 1); body.Children.Add(preview);
        var scroll = new ScrollViewer { Content = settings }; Grid.SetColumn(scroll, 2); body.Children.Add(scroll);
        var footer = new StackPanel { Margin = new Thickness(12), Spacing = 8, Children = { status, progress } };
        Grid.SetRow(footer, 2); root.Children.Add(footer); Content = root;
        language.SelectionChanged += (_, _) => RenderLabels();
        Closing += (_, e) => { if (operation is not null) { e.Cancel = true; operation.Cancel(); } };
        Closed += (_, _) => { selectedBitmap?.Dispose(); excel.Dispose(); };
        RenderLabels();
    }

    private void RenderLabels()
    {
        toolbar.Children.Clear(); settings.Children.Clear();
        Button AddButton(string zh, string ja, Action action)
        {
            var button = new Button { Content = T(zh, ja) };
            button.Click += (_, _) => action(); toolbar.Children.Add(button); return button;
        }
        AddButton("导入图片", "画像を追加", async () => await Import());
        AddButton("移除选中", "選択を削除", () => { if (operation is null && queue.SelectedIndex >= 0) { images.RemoveAt(queue.SelectedIndex); Refresh(); } });
        AddButton("导出到 Excel", "Excel に出力", async () => await Export());
        AddButton("取消", "キャンセル", () => operation?.Cancel());
        toolbar.Children.Add(language);
        void Field(string zh, string ja, Control control)
        { settings.Children.Add(new TextBlock { Text = T(zh, ja) }); settings.Children.Add(control); }
        Field("设计宽度（96 DPI 逻辑像素）", "設計幅（96 DPI 論理ピクセル）", width);
        Field("设计高度", "設計高さ", height);
        Field("设计缩放", "設計倍率", scale);
        Field("回退最大放大倍数", "既定の最大拡大率", enlargement);
        Field("有序规则（条件行 + scale = 数值）", "順序付きルール（条件行 + scale = 数値）", rules);
        compact.Content = T("节省空间排列", "省スペース配置"); settings.Children.Add(compact);
        Field("水平间距（pt）", "横間隔（pt）", horizontal); Field("垂直间距（pt）", "縦間隔（pt）", vertical);
        Field("Excel 行高（pt）", "Excel 行の高さ（pt）", rowHeight); Field("Excel 列宽（字符）", "Excel 列幅（文字）", columnWidth);
        Field("模板路径（空白使用 Excel 默认模板）", "テンプレートパス（空欄は Excel 標準）", template);
        status.Text = T("就绪 · 手动导入 → 预览 → Excel", "準備完了 · 画像追加 → プレビュー → Excel");
    }

    private void Refresh() => queue.ItemsSource = images.Select(i => Path.GetFileName(i.Path)).ToArray();
    private void ShowPreview()
    {
        preview.Source = null; selectedBitmap?.Dispose(); selectedBitmap = null;
        if (queue.SelectedIndex < 0 || queue.SelectedIndex >= images.Count) return;
        try
        {
            using var stream = File.OpenRead(images[queue.SelectedIndex].Path);
            selectedBitmap = Bitmap.DecodeToWidth(stream, 1200); preview.Source = selectedBitmap;
        }
        catch (Exception e) { status.Text = T("预览失败：", "プレビュー失敗：") + e.Message; }
    }

    private async Task Import()
    {
        if (operation is not null) return;
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            AllowMultiple = true, Title = T("选择图片", "画像を選択"),
            FileTypeFilter = new[] { new FilePickerFileType(T("图片", "画像")) { Patterns = new[] { "*.png", "*.jpg", "*.jpeg", "*.bmp" } } }
        });
        if (files.Count == 0 || operation is not null) return;
        operation = new CancellationTokenSource(); var token = operation.Token;
        var errors = new List<string>(); var count = 0;
        try
        {
            foreach (var file in files)
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    var path = file.TryGetLocalPath() ?? throw new IOException(T("仅支持本地图片", "ローカル画像のみ対応"));
                    var metadata = await Task.Run(() =>
                    {
                        using var bitmap = new Bitmap(path);
                        return new ImageMetadata(bitmap.PixelSize.Width, bitmap.PixelSize.Height, bitmap.Dpi.X, bitmap.Dpi.Y);
                    }, token);
                    token.ThrowIfCancellationRequested(); images.Add(new ImportedImage(path, metadata));
                }
                catch (Exception e) when (e is not OperationCanceledException) { errors.Add(file.Name + ": " + e.Message); }
                progress.Value = (double)++count / files.Count;
                status.Text = T("导入", "読み込み") + $" {count}/{files.Count}";
            }
            status.Text = errors.Count == 0 ? T("导入完成", "読み込み完了") : T("部分导入失败：", "一部の読み込みに失敗：") + string.Join("; ", errors);
        }
        catch (OperationCanceledException) { status.Text = T("已取消；已导入图片保留", "キャンセル済み。追加済み画像は保持されます"); }
        finally { operation.Dispose(); operation = null; Refresh(); if (images.Count > 0) queue.SelectedIndex = 0; }
    }

    private double Number(TextBox box, bool zero = false)
    {
        if (!double.TryParse(box.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) || !double.IsFinite(n) || (zero ? n < 0 : n <= 0))
            throw new ArgumentException(T("数值必须为有效正数（间距可为零）", "有効な正数を入力してください（間隔はゼロ可）"));
        return n;
    }

    private async Task Export()
    {
        if (operation is not null || images.Count == 0) return;
        try
        {
            var options = new ExportOptions(new SizingOptions(Number(width), Number(height), Number(scale), Number(enlargement), rules.Text ?? ""),
                compact.IsChecked == true, Number(horizontal, true), Number(vertical, true), Number(rowHeight), Number(columnWidth), template.Text);
            var snapshot = images.ToArray();
            operation = new CancellationTokenSource(); var token = operation.Token;
            status.Text = T("正在准备 Excel…", "Excel を準備中…"); progress.Value = 0;
            var reporting = new Progress<int>(n => { progress.Value = (double)n / snapshot.Length; status.Text = T("导出", "出力") + $" {n}/{snapshot.Length}"; });
            await excel.Run(() => exporter.Export(snapshot, options, excel, reporting, token));
            status.Text = T("导出完成；请在 Excel 中保存工作簿", "出力完了。Excel でブックを保存してください");
        }
        catch (OperationCanceledException) { status.Text = T("导出已取消；已插入图片保留", "出力をキャンセルしました。挿入済み画像は保持されます"); }
        catch (Exception e) { status.Text = T("导出失败（可重试）：", "出力失敗（再試行可）：") + e.Message; }
        finally { operation?.Dispose(); operation = null; }
    }
}
