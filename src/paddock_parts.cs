// ==============================================================================
// paddock_parts.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 画面の部品のうち、特定のタブ・ビューに属さないもの（台帳ファイルの検出・読み込み中の
// 覆いと回転・台帳の読み取り・空の状態・チップの幅・一時的な案内・詳細検索の入力欄と
// 条件チップ・日付のカレンダー・パネルの開閉・縦の開閉・コピー・口座なしの行）。
// PaddockApp の一部（partial class）。ツールの説明とビルドの方法は
// paddock.cs の冒頭にある。
// ==============================================================================

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using System.Xml;

public partial class PaddockApp
{
    // 台帳フォルダから pattern に当たる台帳CSVを検出し、年度（yearRegex の1つ目のグループ）→パスの辞書を返す
    private static Dictionary<int, string> DetectLedgerFiles(string folder, string pattern, Regex yearRegex)
    {
        var files = new Dictionary<int, string>();
        if (!Directory.Exists(folder)) return files;

        foreach (var path in Directory.GetFiles(folder, pattern))
        {
            var match = yearRegex.Match(System.IO.Path.GetFileName(path));
            if (!match.Success) continue;
            int year;
            if (int.TryParse(match.Groups[1].Value, out year))
                files[year] = path;
        }
        return files;
    }

    // 小さなリング（14px・太さ 2・アクセント色）を回し始めて返す。明細の読み込み中の行と
    // 候補ダイアログの「探している間」で共用する（rotation は止めるときに StopSpinner へ渡す）
    private static System.Windows.Shapes.Path MakeSmallSpinner(out RotateTransform rotation, Thickness margin)
    {
        var spinner = new System.Windows.Shapes.Path
        {
            Data = Geometry.Parse("M 7,1 A 6,6 0 1 1 1,7"),
            Stroke = BrushAccent,
            StrokeThickness = 2,
            StrokeStartLineCap = PenLineCap.Round,
            StrokeEndLineCap = PenLineCap.Round,
            Width = 14, Height = 14, Stretch = Stretch.None,
            RenderTransformOrigin = new Point(0.5, 0.5),
            Margin = margin
        };
        rotation = new RotateTransform();
        spinner.RenderTransform = rotation;
        StartSpinner(rotation);
        return spinner;
    }

    // 読み込みオーバーレイの表示（即時・スピナー回転開始）
    private void ShowLoadingOverlay(Grid overlay, RotateTransform rotation)
    {
        overlay.BeginAnimation(UIElement.OpacityProperty, null);
        overlay.Opacity = 1;
        overlay.Visibility = Visibility.Visible;
        StartSpinner(rotation);
    }

    // 読み込みオーバーレイの消灯（フェードアウト100ms）。
    // 一覧への反映が終わってから呼ぶ
    private void HideLoadingOverlay(Grid overlay, RotateTransform rotation)
    {
        var fadeOut = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(100)));
        fadeOut.Completed += delegate
        {
            overlay.Visibility = Visibility.Collapsed;
            overlay.BeginAnimation(UIElement.OpacityProperty, null);
            overlay.Opacity = 1;
            StopSpinner(rotation);
        };
        overlay.BeginAnimation(UIElement.OpacityProperty, fadeOut);
    }

    // スピナー回転アニメーションの開始/停止（一覧のオーバーレイと詳細パネルのミニスピナーで
    // 共用するため回転対象を引数に取る）
    private static void StartSpinner(RotateTransform rotation)
    {
        if (rotation == null) return;
        var anim = new DoubleAnimation();
        anim.By = 360;
        anim.Duration = new Duration(TimeSpan.FromSeconds(1));
        anim.RepeatBehavior = RepeatBehavior.Forever;
        rotation.BeginAnimation(RotateTransform.AngleProperty, anim);
    }

    private static void StopSpinner(RotateTransform rotation)
    {
        if (rotation == null) return;
        rotation.BeginAnimation(RotateTransform.AngleProperty, null);
    }

    // 台帳CSVの全文を読み込む（一覧の読み込みと他年度走査で共用）
    // 書き手（judge.ps1）が排他ロック中は開けないため、
    // 共有違反時は少し待ってリトライする（アトミック追記は短時間で完了する）
    private static string ReadLedgerText(string path)
    {
        for (int attempt = 1; ; attempt++)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(fs, Encoding.UTF8))
                    return reader.ReadToEnd();
            }
            catch (IOException)
            {
                if (attempt >= 3) throw;
                System.Threading.Thread.Sleep(200);
            }
        }
    }

    // 一覧の空状態メッセージを表示する
    private static void ShowEmptyState(TextBlock target, string message)
    {
        target.Text = message;
        target.Visibility = Visibility.Visible;
    }

    // 空文字列を null に変換する（ToolTip バインディング用）
    private static string NullIfEmpty(string value)
    {
        return string.IsNullOrEmpty(value) ? null : value;
    }

    // 検出チップ文字列の実測幅キャッシュ（distinct 単位。行数が増えてもチップ
    // 文字列の種類は限られるため、年度をまたいで持ち続ける。DPI はセッション中
    // 不変の前提）
    private readonly Dictionary<string, double> chipTextWidthCache = new Dictionary<string, double>();

    // チップ1個分の文字列幅を実測する（FontSize 11・UI フォント・現在の DPI。
    // XAML のチップと同条件で、画面描画と同じ DPI 基準の幅を得る）
    private double MeasureChipText(string text)
    {
        double width;
        if (chipTextWidthCache.TryGetValue(text, out width)) return width;
        width = MeasureUiText(text, 11);
        chipTextWidthCache[text] = width;
        return width;
    }

    // UI フォント（Meiryo UI・Normal）での文字列の描画幅（候補ダイアログのボタンの幅など、キャッシュしない実測）
    private double MeasureUiText(string text, double fontSize)
    {
        var formatted = new FormattedText(text, CultureInfo.CurrentUICulture,
            FlowDirection.LeftToRight, ChipTypeface, fontSize, Brushes.Black,
            VisualTreeHelper.GetDpi(window).PixelsPerDip);
        return Math.Ceiling(formatted.Width);
    }

    // 表示中のビューのフッター左に案内を一時表示し、数秒後に表示する前の文言へ戻す
    //（どのビューから開いた詳細パネルで操作しても、案内がいま見ている画面に出るように）
    private void ShowTransientStatus(string message)
    {
        if (statusTimer == null)
        {
            statusTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
            statusTimer.Tick += delegate { EndTransientStatus(); };
        }
        var target = GetVisibleStatusText();
        // 別の欄に一時表示中なら先に戻す。同じ欄に続けて出すときは最初に控えた文言を保つ
        if (statusTimer.IsEnabled && statusTransientTarget != target) EndTransientStatus();
        if (!statusTimer.IsEnabled)
        {
            statusTransientTarget = target;
            statusRestoreText = target.Text;
        }
        target.Text = message;
        statusTimer.Stop();
        statusTimer.Start();
    }

    // 一時表示を終えて、表示する前の文言へ戻す
    private void EndTransientStatus()
    {
        statusTimer.Stop();
        statusTransientTarget.Text = statusRestoreText;
    }

    // フッター左に通常の文言（件数など）を書く。その欄が一時表示中なら一時表示を打ち切る
    //（書いた文言が、一時表示の終わりに古い文言で上書きされないように）
    private void SetStatusText(TextBlock target, string text)
    {
        if (statusTimer != null && statusTimer.IsEnabled && statusTransientTarget == target) statusTimer.Stop();
        target.Text = text;
    }

    // 表示中のビューのフッター左の欄（設定ビューは保存の結果を自前の一時表示で出すため、
    // ここでは預金の一覧の欄を返す）
    private TextBlock GetVisibleStatusText()
    {
        if (insuranceTab.insuranceLedgerView.Visibility == Visibility.Visible) return insuranceTab.insuranceStatusLeft;
        if (recentView.Visibility == Visibility.Visible) return recentFooterText;
        if (seizureView.Visibility == Visibility.Visible) return seizureFooterText;
        return depositStatusLeft;
    }

    // 金額欄の読み取り（カンマ許容）。空=指定なし・非数は指定なし扱い+コーラル枠で提示
    private static long? ReadAmountInput(TextBox box, Border frame)
    {
        var text = box.Text.Trim().Replace(",", "");
        if (text == "") { frame.BorderBrush = BrushLine; return null; }
        long value;
        if (long.TryParse(text, out value)) { frame.BorderBrush = BrushLine; return value; }
        frame.BorderBrush = BrushChipCoralFg;
        return null;
    }

    // 日付欄の読み取り（yyyy/mm/dd）。空=指定なし・パース不能は指定なし扱い+コーラル枠で提示
    private static DateTime? ReadDateInput(TextBox box, Border frame)
    {
        var text = box.Text.Trim();
        if (text == "") { frame.BorderBrush = BrushLine; return null; }
        var value = BusinessLogic.ParseDisplayDate(text);
        frame.BorderBrush = value == null ? BrushChipCoralFg : BrushLine;
        return value;
    }

    // 日付範囲チップの文言（片側のみの指定は「〜」の片側を空ける）
    private static string FormatDateRangeChip(string label, DateTime? from, DateTime? to)
    {
        return label + " " +
            (from != null ? from.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture) : "") + "〜" +
            (to != null ? to.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture) : "");
    }

    // 条件チップ1個をチップ行へ追加する（クリック=そのグループの条件を解除）
    private void AddConditionChip(WrapPanel chipsRow, TextBox searchBox, string text, Action clear)
    {
        var chip = new Border
        {
            Background = BrushSelectedCardBg,
            BorderBrush = BrushSelectedLine,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(13),
            Padding = new Thickness(11, 4, 9, 5),
            Margin = new Thickness(0, 0, 6, 0),
            Cursor = Cursors.Hand,
            ToolTip = "クリックでこの条件を解除"
        };
        var stack = new StackPanel { Orientation = Orientation.Horizontal };
        stack.Children.Add(new TextBlock { Text = text, FontSize = 11, Foreground = BrushAccent });
        stack.Children.Add(new TextBlock
        {
            Text = "✕",
            FontSize = 9,
            Foreground = BrushSelectedLine,
            Margin = new Thickness(7, 1, 0, 0),
            VerticalAlignment = VerticalAlignment.Center
        });
        chip.Child = stack;
        chip.MouseEnter += delegate { chip.Background = BrushChipVioletBg; };
        chip.MouseLeave += delegate { chip.Background = BrushSelectedCardBg; };
        // キーで外したときは、組み直したチップ列の同じ位置（なければ検索欄）へフォーカスを移す
        Action remove = delegate
        {
            bool refocus = chip.IsKeyboardFocused;
            int index = chipsRow.Children.IndexOf(chip);
            clear();
            if (!refocus) return;
            FocusLater(delegate
            {
                int count = chipsRow.Children.Count;
                if (count > 0) ((UIElement)chipsRow.Children[Math.Min(index, count - 1)]).Focus();
                else searchBox.Focus();
            });
        };
        chip.MouseLeftButtonUp += delegate { remove(); };
        MakeKeyActivatable(chip, remove, keyFocusPillStyle);
        chipsRow.Children.Add(chip);
    }

    // 統合トグルの暗色ツールチップの中身（「表記ゆれを統合」+吸収した生値の列挙）
    private UIElement BuildStaffMergeTip(StaffCluster cluster)
    {
        var stack = new StackPanel { MaxWidth = 460 };
        stack.Children.Add(new TextBlock
        {
            Text = "表記ゆれを統合:",
            FontSize = 10.5,
            Foreground = BrushTipNeutralFg,
            Margin = new Thickness(0, 0, 0, 3)
        });
        stack.Children.Add(new TextBlock
        {
            Text = string.Join(" ／ ", cluster.RawValues.ToArray()),
            FontSize = 11,
            Foreground = BrushTipText,
            TextWrapping = TextWrapping.Wrap
        });
        return stack;
    }

    // 入力欄のプレースホルダ表示を配線する（検索ボックスと同じ方式）
    private static void WirePlaceholder(TextBox box, TextBlock placeholder)
    {
        box.TextChanged += delegate
        {
            placeholder.Visibility = box.Text.Length > 0 ? Visibility.Collapsed : Visibility.Visible;
        };
    }

    // 残高欄の入力整形を配線する。
    // 入力中: 半角数字・カンマ・先頭マイナスだけで構成される間はその場で3桁区切りに
    // 整え直す。IME変換中の未確定文字列を書き換えないよう、それ以外の文字を含む間は
    // 触らない（全角数字はフォーカスが離れた時に片寄せする）。
    // フォーカス喪失時: 全角を片寄せして整形し、書き換えた場合は条件の有効/無効が
    // 変わり得るため apply で即時反映する（blur は明示操作としてトグル系と同じ扱い）
    private static void WireAmountFormatting(TextBox box, Action apply)
    {
        bool reformatting = false;   // 自身の書き換えによる TextChanged の再入を防ぐ
        box.TextChanged += delegate
        {
            if (reformatting) return;
            var text = box.Text;
            var formatted = BusinessLogic.FormatAmountText(text);
            if (formatted == null || formatted == text) return;
            int digitsBefore = CountDigitsBefore(text, box.SelectionStart);
            reformatting = true;
            box.Text = formatted;
            reformatting = false;
            box.SelectionStart = CaretAfterDigits(formatted, digitsBefore);
        };
        box.LostFocus += delegate
        {
            var normalized = BusinessLogic.NormalizeAmountText(box.Text);
            if (normalized == null || normalized == box.Text) return;
            box.Text = normalized;
            apply();
        };
    }

    // 日付欄のフォーカス喪失時の補正を配線する。
    // 解釈できた入力だけを正規形 yyyy/MM/dd に書き直し、apply で即時反映する。
    // 解釈できない入力は書き換えず、コーラル枠の提示は ReadDateInput が行う
    private static void WireDateNormalize(TextBox box, Action apply)
    {
        box.LostFocus += delegate
        {
            var normalized = BusinessLogic.NormalizeDateText(box.Text);
            if (normalized == null || normalized == box.Text) return;
            box.Text = normalized;
            apply();
        };
    }

    // 日付欄のカレンダーを開く。共有ポップアップを対象欄の直下に配置し、
    // 常に日選択ビューで開く（前回の月/年選択ビューが残るのを防ぐ）。欄に有効な日付が
    // あればその日を選択・表示の起点に、なければ今日の月を表示する。apply は日付を確定したときの反映
    private void OpenAdvDatePopup(TextBox box, Action apply)
    {
        advDateTarget = box;
        advDateApply = apply;
        // Popup は開く瞬間にしか位置を計算しないため、別の欄からの切替に備えて
        // 一度閉じてから開き直す
        advDatePopup.IsOpen = false;
        // 入力枠の Border（TextBox → Grid → Border）を基準に置き、欄の左端に揃える
        advDatePopup.PlacementTarget = (UIElement)VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(box));
        var current = BusinessLogic.ParseDisplayDate(box.Text);
        advDateSuppress = true;
        advDateCalendar.DisplayMode = CalendarMode.Month;
        advDateCalendar.SelectedDate = current;
        advDateCalendar.DisplayDate = current ?? DateTime.Today;
        advDateSuppress = false;
        advDateTodayValue.Text = DateTime.Today.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
        advDatePopup.IsOpen = true;
    }

    // 日付欄1つにカレンダーを配線する: 欄のクリックで開き、打鍵を始めたら閉じて手打ちに譲る
    //（Esc は window 側で処理）。apply は日付を確定したときの反映
    private void WireAdvDateBox(TextBox box, Action apply)
    {
        box.PreviewMouseLeftButtonUp += delegate { OpenAdvDatePopup(box, apply); };
        box.PreviewKeyDown += delegate
        {
            if (advDatePopup.IsOpen) advDatePopup.IsOpen = false;
        };
    }

    // カレンダーからの日付確定。正規形で欄に設定して閉じ、フォーカスを欄に戻して
    // 絞り込みを即時反映する（明示操作のためデバウンスを介さない）
    private void CommitAdvCalendarDate(DateTime date)
    {
        if (advDateTarget == null) return;
        advDatePopup.IsOpen = false;
        advDateTarget.Text = date.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
        advDateTarget.Focus();
        advDateTarget.CaretIndex = advDateTarget.Text.Length;
        advDateApply();
    }

    // 整形の前後でキャレット位置を対応付ける（「位置より左にある数字の個数」で合わせる）
    private static int CountDigitsBefore(string text, int index)
    {
        int count = 0;
        for (int i = 0; i < index && i < text.Length; i++)
            if (char.IsDigit(text[i])) count++;
        return count;
    }

    private static int CaretAfterDigits(string text, int digits)
    {
        if (digits <= 0) return text.StartsWith("-") ? 1 : 0;   // 数字より左は記号（マイナス）の直後へ
        int seen = 0;
        for (int i = 0; i < text.Length; i++)
        {
            if (!char.IsDigit(text[i])) continue;
            seen++;
            if (seen == digits) return i + 1;
        }
        return text.Length;
    }

    // パネル幅の開閉アニメーション（中断時は現在幅から目標幅へ滑らかに継続する）
    private static void AnimateDetailPanel(Border panel, double toWidth)
    {
        var animation = new DoubleAnimation
        {
            To = toWidth,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };
        panel.BeginAnimation(FrameworkElement.WidthProperty, animation);
    }

    // 開閉アニメーションの標準（150ms）: 高さ 0⇔自然高+フェードで縦に開閉する。
    // 詳細検索パネル・設定の保存バー・メモの入力エリアで共用。開閉が交錯した場合は
    // 現在の見かけの高さから継続する
    private static void AnimateExpand(FrameworkElement element)
    {
        double from = element.Visibility == Visibility.Visible ? element.ActualHeight : 0;
        if (element.Visibility != Visibility.Visible) element.Opacity = 0;

        // 高さの固定を解いて自然な高さを確定させ、アニメーションの終点にする
        element.BeginAnimation(FrameworkElement.HeightProperty, null);
        element.Height = double.NaN;
        element.Visibility = Visibility.Visible;
        element.UpdateLayout();

        var grow = new DoubleAnimation(from, element.ActualHeight,
            new Duration(TimeSpan.FromMilliseconds(150)));
        grow.Completed += delegate
        {
            // 高さを Auto に戻し、以後の内容量の変化に追従させる
            element.BeginAnimation(FrameworkElement.HeightProperty, null);
            element.Height = double.NaN;
        };
        element.BeginAnimation(FrameworkElement.HeightProperty, grow);
        element.BeginAnimation(UIElement.OpacityProperty,
            new DoubleAnimation(1, new Duration(TimeSpan.FromMilliseconds(150))));
    }

    private static void AnimateCollapse(FrameworkElement element)
    {
        var shrink = new DoubleAnimation(element.ActualHeight, 0,
            new Duration(TimeSpan.FromMilliseconds(150)));
        shrink.Completed += delegate
        {
            element.Visibility = Visibility.Collapsed;
            element.BeginAnimation(FrameworkElement.HeightProperty, null);
            element.Height = double.NaN;
            element.BeginAnimation(UIElement.OpacityProperty, null);
            element.Opacity = 1;
        };
        element.BeginAnimation(FrameworkElement.HeightProperty, shrink);
        element.BeginAnimation(UIElement.OpacityProperty,
            new DoubleAnimation(0, new Duration(TimeSpan.FromMilliseconds(150))));
    }

    // 年度の一覧を「2024年度、2025年度」の形にする（昇順・フッターと脚注で共用）
    private static string FormatYearLabels(List<int> years)
    {
        var sorted = new List<int>(years);
        sorted.Sort();
        var labels = new List<string>();
        foreach (var year in sorted) labels.Add(BusinessLogic.FormatYearKey(year));
        return string.Join("、", labels.ToArray());
    }

    // 今年度の鍵（履歴の見出しで今年度の区画を見分ける。年度の判定で時計を読むのはここだけ）
    private static string GetCurrentYearKey()
    {
        return BusinessLogic.FormatYearKey(BusinessLogic.GetFiscalYear(DateTime.Now));
    }

    // 要素のダブルクリック、またはフォーカスして Enter で、getText の文字列をクリップボードへ写す
    //（TextBlock・Run の両方に付けられる。copyChip は写したことを知らせるチップ）
    private void WireCopy(IInputElement element, Func<string> getText, Border copyChip)
    {
        element.MouseLeftButtonDown += delegate(object s, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2) CopyPanelText(getText(), copyChip);
        };
        element.Focusable = true;
        var frameworkElement = element as FrameworkElement;
        if (frameworkElement != null) frameworkElement.FocusVisualStyle = keyFocusStyle;
        var contentElement = element as FrameworkContentElement;
        if (contentElement != null) contentElement.FocusVisualStyle = keyFocusStyle;
        element.KeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.Key != Key.Enter || Keyboard.Modifiers != ModifierKeys.None || e.OriginalSource != element) return;
            CopyPanelText(getText(), copyChip);
            e.Handled = true;
        };
    }

    // 文字列をクリップボードへ写し、ヘッダー右上の chip に「コピーしました」を1.5秒表示する。
    // 空文字は何もしない。クリップボードは他プロセスとの競合で書き込みが確定しないことがあるため
    // WinForms の再試行版（10回×100ms）で書き、それでも失敗したときは黙って何もしない
    private static void CopyPanelText(string text, Border chip)
    {
        if (string.IsNullOrEmpty(text)) return;
        try { System.Windows.Forms.Clipboard.SetDataObject(text, true, 10, 100); }
        catch { return; }

        // 表示を消すタイマーはチップごとに1つ持つ（チップの Tag に置く）
        var timer = chip.Tag as DispatcherTimer;
        if (timer == null)
        {
            timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(1500) };
            timer.Tick += delegate
            {
                timer.Stop();
                SetChip(chip, null, null, null);
            };
            chip.Tag = timer;
        }
        SetChip(chip, "コピーしました", BrushSelectedCardBg, BrushAccent);
        timer.Stop();
        timer.Start();
    }

    // 回次内の口座なし・契約なしの行（label のピル+金融機関名・保険会社名の列挙）を組み立てる
    private static UIElement BuildNoAccountLine(string label, List<string> names)
    {
        var grid = new Grid { Margin = new Thickness(1, 2, 0, 0) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var pill = new Border
        {
            Background = BrushChipNeutralBg,
            CornerRadius = new CornerRadius(9),
            Padding = new Thickness(8, 1, 8, 2),
            VerticalAlignment = VerticalAlignment.Top
        };
        pill.Child = new TextBlock { Text = label, FontSize = 10.5, Foreground = BrushTextMuted };
        grid.Children.Add(pill);

        var text = new TextBlock
        {
            Text = string.Join("、", names),
            FontSize = 11,
            Foreground = BrushTextMuted,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(8, 1, 0, 0)
        };
        Grid.SetColumn(text, 1);
        grid.Children.Add(text);
        return grid;
    }
}