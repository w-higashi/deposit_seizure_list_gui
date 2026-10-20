// ==============================================================================
// paddock_detail.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 明細を見る（口座カードの明細の読み取りと表示・グラフ・明細ウィンドウ）。
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
    // ==============================================================
    // 明細を見る
    // ==============================================================
    //
    // 口座カードの展開部に、照会ファイルから読んだ取引明細の概観（残高推移のスパークライン+
    // 検出ごとの集約）を出す。読み取りはカードの展開時に1本のワーカーで行い、人物ファイルの
    // 骨格・シートのセル・口座の結果をセッション内にキャッシュする。ファイルの解決は
    // 「照会ファイルを開く」と同じ規則だが対話は挟まず、特定できないときだけ退行表示から
    // 候補ダイアログへ渡す。判定・書き込みは一切しない

    // --- セクションの枠 ---

    // 展開部の末尾に置く明細セクションの枠: 破線（展開部の行があるとき）・⧉ の行・中身（Section）を縦に並べる。
    // 余白は UpdateDetailFrameSpacing が破線と ⧉ の行の表示に合わせて決める
    private static StackPanel BuildDetailSectionFrame(DepositCardExpansion expansion)
    {
        var frame = new StackPanel();
        if (expansion.Rule != null) frame.Children.Add(expansion.Rule);
        frame.Children.Add(expansion.Toolbar);
        frame.Children.Add(expansion.Section);
        return frame;
    }

    // 明細セクション先頭の ⧉ の行: 左に最終取引日、右に ⧉ の受け皿。高さは ⧉ に合わせて固定する
    //（読み終えて ⧉ が現れても最終取引日の位置が動かないように）
    private static Grid BuildDetailToolbar(string lastTx, out Decorator openSlot)
    {
        var toolbar = new Grid { Height = 18 };
        toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        if (lastTx != "")
        {
            var meta = MakeDepositLastTxText(lastTx);
            meta.VerticalAlignment = VerticalAlignment.Center;
            toolbar.Children.Add(meta);
        }
        openSlot = new Decorator { VerticalAlignment = VerticalAlignment.Center };
        Grid.SetColumn(openSlot, 1);
        toolbar.Children.Add(openSlot);
        return toolbar;
    }

    // 枠の余白を破線と ⧉ の行の表示に合わせて決める（余白はここだけで決める）。
    // 破線あり: 枠の上 8・破線の下 7／破線なし: 枠の上 2。⧉ の行は最終取引日か ⧉ があるときだけ出し、中身との間は 2
    private static void UpdateDetailFrameSpacing(DepositCardExpansion expansion)
    {
        bool ruleShown = expansion.Rule != null && expansion.Rule.Visibility == Visibility.Visible;
        bool toolbarShown = expansion.HasLastTx || expansion.OpenSlot.Child != null;
        double afterRule = ruleShown ? 7 : 0;
        expansion.Toolbar.Visibility = toolbarShown ? Visibility.Visible : Visibility.Collapsed;
        expansion.Frame.Margin = new Thickness(0, ruleShown ? 8 : 2, 0, 0);
        expansion.Toolbar.Margin = new Thickness(0, afterRule, 0, 0);
        expansion.Section.Margin = new Thickness(0, toolbarShown ? 2 : afterRule, 0, 0);
    }

    // ⧉ を受け皿に入れる（openWindow＝⧉ の処理）／外す（null）。⧉ の行の表示と余白も決め直す
    private void SetDetailOpenButton(DepositCardExpansion expansion, Action openWindow)
    {
        expansion.OpenWindow = openWindow;
        expansion.OpenSlot.Child = openWindow != null ? MakeDetailOpenButton(openWindow) : null;
        UpdateDetailFrameSpacing(expansion);
    }

    // 幅いっぱいの破線（1px・選択面の枠線色）。幅は描画時に確定するため SizeChanged で線を引き直す
    private static UIElement MakeDashedRule()
    {
        var host = new Canvas { Height = 1, ClipToBounds = true };
        var line = new System.Windows.Shapes.Line
        {
            X1 = 0, Y1 = 0.5, Y2 = 0.5,
            Stroke = BrushSelectedLine,
            StrokeThickness = 1,
            StrokeDashArray = new DoubleCollection { 3, 2 }
        };
        host.Children.Add(line);
        host.SizeChanged += delegate(object s, SizeChangedEventArgs e) { line.X2 = e.NewSize.Width; };
        return host;
    }

    // ⧉ の行の右端に置く ⧉（明細ウィンドウを開く）。見出し文字は置かない（グラフを見れば役割は分かる）
    private UIElement MakeDetailOpenButton(Action openWindow)
    {
        var button = new Border
        {
            Width = 22, Height = 18,
            CornerRadius = new CornerRadius(6),
            Background = Brushes.White,
            BorderBrush = BrushSelectedLine,
            BorderThickness = new Thickness(1),
            Cursor = Cursors.Hand,
            ToolTip = "明細ウィンドウで開く (Ctrl+M)"
        };
        button.Child = new TextBlock
        {
            Text = "⧉", FontSize = 12, Foreground = BrushAccent,
            HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, -1, 0, 0)
        };
        button.MouseEnter += delegate { button.Background = BrushSelectedCardBg; };
        button.MouseLeave += delegate { button.Background = Brushes.White; };
        button.MouseLeftButtonUp += delegate(object s, MouseButtonEventArgs e)
        {
            e.Handled = true;   // カードの開閉に伝えない
            openWindow();
        };
        MakeKeyActivatable(button, openWindow);
        return button;
    }

    // 退行表示の1行（読み込み中・特定不能・取引なし など）
    private static TextBlock MakeDetailNote(string text)
    {
        return new TextBlock
        {
            Text = text,
            FontSize = 11,
            Foreground = BrushTextMuted,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 4, 0, 0)
        };
    }

    // --- 読み取りの要求とワーカー ---

    // 展開されたカードの明細を要求する。結果がキャッシュにあれば即描き、なければ読み込み中を
    // 出してワーカーに積む（同じ行の要求が待っていれば1つにまとめる）
    private void RequestDetail(DepositCardExpansion expansion)
    {
        var rowKey = BusinessLogic.GetDepositLedgerRowKey(expansion.Fields, eraMapping);
        DetailResult cached;
        if (detailResultCache.TryGetValue(rowKey, out cached))
        {
            if (cached.Data != null) TouchDetailCache(cached.Data.SourcePath);
            RenderDetailSection(expansion, cached);
            return;
        }
        ShowDetailLoading(expansion);

        if (rowKey == detailRunningKey) return;
        foreach (var queued in detailQueue)
            if (queued.RowKey == rowKey) return;

        var fields = expansion.Fields;
        var request = new DetailRequest
        {
            Fields = fields,
            RowKey = rowKey,
            ExpectedSheet = GetExpectedSheetName(fields),
            ResolvedPaths = new Dictionary<string, string>(detailResolvedPaths),
            Workbooks = new Dictionary<string, InquiryWorkbook>(detailWorkbookCache, StringComparer.OrdinalIgnoreCase),
            Stamps = new Dictionary<string, DateTime>(detailWorkbookStamps, StringComparer.OrdinalIgnoreCase),
            Sheets = new Dictionary<string, SheetCells>(detailSheetCache, StringComparer.OrdinalIgnoreCase)
        };
        request.ResolveKey = GetDetailResolveKey(fields, request.ExpectedSheet);
        detailChosenPaths.TryGetValue(rowKey, out request.ChosenPath);
        detailQueue.Add(request);
        StartNextDetail();
    }

    // 期待シート名 {照会結果ファイル名(拡張子なし)}_{シート名} の安全化（ドリルダウンと同じ規則。作れなければ null）
    private static string GetExpectedSheetName(string[] fields)
    {
        return BusinessLogic.GetExpectedSheetName(fields[BusinessLogic.DEPOSIT_COL_SOURCE_FILE], fields[BusinessLogic.DEPOSIT_COL_SHEET_NAME]);
    }

    // 待ち行列の末尾（最後に展開されたカード）から処理する
    private void StartNextDetail()
    {
        if (detailQueue.Count == 0) return;
        if (detailWorker == null)
        {
            detailWorker = new BackgroundWorker();
            detailWorker.DoWork += delegate(object s, DoWorkEventArgs e)
            {
                var request = (DetailRequest)e.Argument;
                try { ProcessDetailRequest(request); }
                catch (Exception ex) { request.Error = ex; }
                e.Result = request;
            };
            detailWorker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs e)
            {
                detailRunningKey = null;
                OnDetailCompleted((DetailRequest)e.Result);
                StartNextDetail();
            };
        }
        if (detailWorker.IsBusy) return;
        var next = detailQueue[detailQueue.Count - 1];
        detailQueue.RemoveAt(detailQueue.Count - 1);
        detailRunningKey = next.RowKey;
        next.TypeMapping = detailTypeMapping;   // 始める時点の口座種別の補正表（待っている間に読み直しても新しい表で読む）
        detailWorker.RunWorkerAsync(next);
    }

    // ワーカー: 照会ファイルの特定→骨格→シートのセル→口座の明細→計算。UI には触らない
    private void ProcessDetailRequest(DetailRequest request)
    {
        var path = request.ChosenPath;
        if (path == null && !request.ResolvedPaths.TryGetValue(request.ResolveKey, out path))
        {
            var result = new DetailResult();
            path = ResolveDetailFileSilent(request, result);
            if (path == null)
            {
                request.Result = result;   // 特定できず（候補を持ち帰る）
                return;
            }
            request.ResolvedPath = path;
        }

        // 骨格とセル: 更新日時が変わっていなければキャッシュの写しを使う
        request.Stamp = File.GetLastWriteTimeUtc(path);
        InquiryWorkbook workbook;
        DateTime stamp;
        SheetCells cells = null;
        if (request.Workbooks.TryGetValue(path, out workbook) && request.Stamps.TryGetValue(path, out stamp) && stamp == request.Stamp)
        {
            if (request.ExpectedSheet != null) request.Sheets.TryGetValue(path + "|" + request.ExpectedSheet, out cells);
        }
        else workbook = BusinessLogic.OpenInquiryWorkbook(path);
        request.Workbook = workbook;

        var data = new DetailAccountData { SourcePath = path, Status = DetailStatus.NoDetailPage };
        var sheet = request.ExpectedSheet != null ? workbook.FindSheet(request.ExpectedSheet) : null;
        if (sheet != null)
        {
            if (cells == null) cells = BusinessLogic.ReadSheetCells(workbook, sheet.Value.Value, BusinessLogic.GetDetailWantedColumns(config.Detail));
            request.Cells = cells;
            data = BusinessLogic.ExtractDetailAccount(cells,
                request.Fields[BusinessLogic.DEPOSIT_COL_ACCOUNT_NUM], request.Fields[BusinessLogic.DEPOSIT_COL_ACCOUNT_TYPE],
                request.Fields[BusinessLogic.DEPOSIT_COL_INSTITUTION], request.TypeMapping,
                config.Detail, workbook.Date1904, eraMapping);
            data.SourcePath = path;
            data.SheetName = sheet.Value.Key;
        }
        BusinessLogic.AnalyzeDetailAccount(data, BusinessLogic.BuildDetailDetections(request.Fields));
        request.Result = new DetailResult { Data = data };
    }

    // 対話なしのファイル解決: 登録日の年度フォルダで候補を集め、0件なら全年度に
    // 広げる。期待シートを持つ候補がちょうど1件ならそのパス、それ以外は null（候補を result に残す）
    private string ResolveDetailFileSilent(DetailRequest request, DetailResult result)
    {
        var fields = request.Fields;
        var yearKey = BusinessLogic.GetSearchYearKey(fields[BusinessLogic.DEPOSIT_COL_REG_DATE], eraMapping, config.DepositSearchFolders);

        var addressNumber = fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM];
        var candidates = CollectInquiryFileCandidates(BuildLabeledFolders(config.DepositSearchFolders, yearKey), addressNumber, request.ExpectedSheet);
        if (candidates.Count == 0 && yearKey != null)
        {
            yearKey = null;
            candidates = CollectInquiryFileCandidates(BuildLabeledFolders(config.DepositSearchFolders, null), addressNumber, request.ExpectedSheet);
        }
        var matched = candidates.Where(c => c.HasExpectedSheet).ToList();
        if (matched.Count == 1) return matched[0].FullPath;

        result.Candidates = candidates;
        result.YearKey = yearKey;
        return null;
    }

    // ワーカー完了（UI スレッド）: キャッシュを更新し、当該行のカードがまだ画面にあれば描く
    private void OnDetailCompleted(DetailRequest request)
    {
        if (request.Error != null)
        {
            WriteErrorLog("明細", request.Error);
            foreach (var expansion in FindExpansionsByRowKey(request.RowKey))
                ShowDetailNoteOnly(expansion, "明細を読み取れませんでした");
            return;
        }
        if (request.ResolvedPath != null) detailResolvedPaths[request.ResolveKey] = request.ResolvedPath;
        if (request.Workbook != null)
        {
            var key = request.Workbook.Path;
            detailWorkbookCache[key] = request.Workbook;
            detailWorkbookStamps[key] = request.Stamp;
            if (request.Cells != null && request.ExpectedSheet != null) detailSheetCache[key + "|" + request.ExpectedSheet] = request.Cells;
            TouchDetailCache(key);
        }
        // 要求のあとで口座種別の補正表を読み直していたら、古い表での結果は使わず今の表で読み直す
        //（特定したファイル・骨格・セルは表に依存しないので取り込んでおき、読み直しはそれを使う）
        if (request.TypeMapping != detailTypeMapping)
        {
            foreach (var expansion in FindExpansionsByRowKey(request.RowKey))
                if (expansion.Requested) RequestDetail(expansion);
            return;
        }
        detailResultCache[request.RowKey] = request.Result;
        foreach (var expansion in FindExpansionsByRowKey(request.RowKey))
            RenderDetailSection(expansion, request.Result);
    }

    // 行キーに対応する、いま画面にあるカードの展開部（通常は0または1件）
    private List<DepositCardExpansion> FindExpansionsByRowKey(string rowKey)
    {
        var result = new List<DepositCardExpansion>();
        foreach (var expansion in depositCardExpansions.Values)
            if (expansion.Section != null && BusinessLogic.GetDepositLedgerRowKey(expansion.Fields, eraMapping) == rowKey) result.Add(expansion);
        return result;
    }

    // 特定結果のキャッシュの鍵（年度|宛名番号|期待シート。年度は登録日から・決まらなければ空）
    private string GetDetailResolveKey(string[] fields, string expectedSheet)
    {
        var fiscalYear = BusinessLogic.GetRegistrationFiscalYear(fields[BusinessLogic.DEPOSIT_COL_REG_DATE], eraMapping);
        return (fiscalYear == null ? "" : fiscalYear.Value.ToString()) + "|" +
            fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM] + "|" + expectedSheet;
    }

    // 明細の読み取りでセッション内に特定済みの照会ファイル（職員が候補から選んだもの・ワーカーが期待シートの
    // 一致 1 件で特定したもの）を候補の形で返す（明細を開いていない行・特定できず候補のままの行は null）。
    // まだ在るかは呼び出し側のワーカーが確かめる。期待シートの有無は、選んだものは明細がそのシートを読めたか、
    // 特定したものは一致そのもの
    private InquiryFileCandidate FindRememberedInquiryFile(string[] fields)
    {
        var expectedSheet = GetExpectedSheetName(fields);
        var rowKey = BusinessLogic.GetDepositLedgerRowKey(fields, eraMapping);
        string path;
        if (detailChosenPaths.TryGetValue(rowKey, out path))
        {
            DetailResult cached;
            bool sheetRead = detailResultCache.TryGetValue(rowKey, out cached) && cached.Data != null && cached.Data.SheetName != null;
            return new InquiryFileCandidate { FullPath = path, HasExpectedSheet = sheetRead };
        }
        if (detailResolvedPaths.TryGetValue(GetDetailResolveKey(fields, expectedSheet), out path))
            return new InquiryFileCandidate { FullPath = path, HasExpectedSheet = true };
        return null;
    }

    // 候補ダイアログで選ばれたファイルで読み直す（選択は行キーに対してセッション内で記憶する）
    private void ChooseDetailFile(DepositCardExpansion expansion, DetailResult result)
    {
        var fields = expansion.Fields;
        var rowKey = BusinessLogic.GetDepositLedgerRowKey(fields, eraMapping);
        var chosen = ShowCandidateDialog(config.DepositSearchFolders, result.Candidates, result.YearKey, fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM], GetExpectedSheetName(fields), null,
            null, "照会ファイルの候補", "明細を読む", "明細を読むファイルを選んでください");
        if (chosen == null) return;
        detailChosenPaths[rowKey] = chosen.FullPath;
        detailResultCache.Remove(rowKey);
        RequestDetail(expansion);
    }

    // --- セクションの描画 ---

    private void ShowDetailLoading(DepositCardExpansion expansion)
    {
        var section = expansion.Section;
        StopDetailSpinner(expansion);
        ShowCategoryLines(expansion, true);
        SetDetailOpenButton(expansion, null);
        section.Children.Clear();

        var line = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 2, 0, 0) };
        line.Children.Add(MakeSmallSpinner(out expansion.Spinner, new Thickness(0, 1, 6, 0)));
        line.Children.Add(new TextBlock { Text = "照会ファイルを読み込み中...", FontSize = 11, Foreground = BrushTextMuted });
        section.Children.Add(line);
    }

    private void ShowDetailNoteOnly(DepositCardExpansion expansion, string note)
    {
        StopDetailSpinner(expansion);
        ShowCategoryLines(expansion, true);
        SetDetailOpenButton(expansion, null);
        expansion.Section.Children.Clear();
        expansion.Section.Children.Add(MakeDetailNote(note));
    }

    private static void StopDetailSpinner(DepositCardExpansion expansion)
    {
        if (expansion.Spinner == null) return;
        StopSpinner(expansion.Spinner);
        expansion.Spinner = null;
    }

    // パネル再構築時: 読み込み中のカードが捨てられる前に回転アニメーションを止める
    private void StopDetailSpinners()
    {
        foreach (var expansion in depositCardExpansions.Values) StopDetailSpinner(expansion);
    }

    // 結果を明細セクションに描く。読み取れた場合は推移グラフ+集約表、それ以外は1行の退行表示
    private void RenderDetailSection(DepositCardExpansion expansion, DetailResult result)
    {
        if (result.Data == null)
        {
            if (result.Candidates.Count == 0)
            {
                ShowDetailNoteOnly(expansion, "照会ファイルが見つかりませんでした");
                return;
            }
            // 候補が複数（期待シートを持つものが1件に絞れない）: 退行表示の行をクリックすると候補ダイアログから選べる
            StopDetailSpinner(expansion);
            ShowCategoryLines(expansion, true);
            SetDetailOpenButton(expansion, null);
            var section = expansion.Section;
            section.Children.Clear();
            var note = MakeDetailNote("照会ファイルを特定できませんでした（候補 " + result.Candidates.Count + " 件・クリックして選ぶ）");
            note.Foreground = BrushAccent;
            note.Cursor = Cursors.Hand;
            Action choose = delegate { ChooseDetailFile(expansion, result); };
            note.MouseLeftButtonUp += delegate(object s, MouseButtonEventArgs e)
            {
                e.Handled = true;   // カードの開閉に伝えない
                choose();
            };
            MakeKeyActivatable(note, choose);
            section.Children.Add(note);
            return;
        }

        var data = result.Data;
        switch (data.Status)
        {
            case DetailStatus.NoDetailPage:
                ShowDetailNoteOnly(expansion, "明細ページが見つかりません");
                return;
            case DetailStatus.NoTransactions:
                ShowDetailNoteOnly(expansion, "この期間に取引はありません");
                return;
        }

        StopDetailSpinner(expansion);
        ShowCategoryLines(expansion, false);   // 検出行は集約表に吸収する（名称が二度出ないように）
        var body = expansion.Section;
        body.Children.Clear();
        var fields = expansion.Fields;
        SetDetailOpenButton(expansion, delegate { OpenDetailWindow(fields, data); });
        body.Children.Add(BuildDetailChartHost(data, DetailSparklineStyle));
        // 検出の表は行（台帳の検出）があるときだけ置く（見出しだけの表を出さない）。
        // 集約は台帳の検出に当たった取引からしか作られないため、検出がなければ集約もない
        var detections = BusinessLogic.BuildDetailDetections(fields);
        if (detections.Count > 0) body.Children.Add(BuildDetailAggregateTable(data.Aggregates, detections));
    }

    // 取引カテゴリの検出行と、その下の破線の表示切替（明細を読めた口座では集約表が肩代わりする）。枠の余白も決め直す
    private static void ShowCategoryLines(DepositCardExpansion expansion, bool visible)
    {
        if (expansion.CategoryLines != null) expansion.CategoryLines.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        if (expansion.Rule != null)
        {
            // 破線は展開部の行と明細の境。検出行を隠すときは反対債権などカテゴリ以外の行が残るときだけ出す
            bool othersRemain = false;
            foreach (UIElement child in expansion.Detail.Children)
                if (!ReferenceEquals(child, expansion.CategoryLines) && !ReferenceEquals(child, expansion.Frame)) othersRemain = true;
            expansion.Rule.Visibility = visible || othersRemain ? Visibility.Visible : Visibility.Collapsed;
        }
        UpdateDetailFrameSpacing(expansion);
    }

    // --- 残高推移のスパークライン ---

    // グラフの受け皿。幅はカードの内側で決まる（スクロールバーの有無でも変わる）ため、
    // SizeChanged のたびに描き直す
    private UIElement BuildDetailChartHost(DetailAccountData data, DetailChartStyle style)
    {
        var host = new Canvas
        {
            Background = Brushes.Transparent,   // ホバーを拾うため
            ClipToBounds = true
        };
        if (!style.FillHeight) host.Height = style.TotalHeight;
        var tip = new ToolTip
        {
            Style = (Style)window.FindResource("DetectTip"),
            PlacementTarget = host,
            Placement = PlacementMode.Relative
        };
        tip.LayoutTransform = displayScaleTransform;   // 表示倍率に追従
        var points = new List<DetailChartPoint>();
        host.SizeChanged += delegate(object s, SizeChangedEventArgs e) { DrawDetailChart(host, data, style, e.NewSize.Width, e.NewSize.Height, points); };
        host.MouseMove += delegate(object s, MouseEventArgs e) { ShowDetailPointTip(host, tip, points, e.GetPosition(host)); };
        host.MouseLeave += delegate { tip.IsOpen = false; };
        return host;
    }

    // グラフ上の点（ホバーの当たり判定用）
    private class DetailChartPoint
    {
        public double X;
        public double Y;
        public DetailTransaction Tx;
    }

    // 推移グラフを描く。上段: 残高の折れ線（回答の残高=実線・●／推定=破線・白抜き○／
    // 合わない点=コーラルの輪。取引数が多いときは点なし）。下段: 入金=検出色の棒・出金=細線。
    // 点と棒は塗りと枠の組ごとに 1 個の Path にまとめる
    private void DrawDetailChart(Canvas canvas, DetailAccountData data, DetailChartStyle style, double width, double height, List<DetailChartPoint> points)
    {
        canvas.Children.Clear();
        points.Clear();
        if (width < 80 || data.PeriodFrom == null) return;

        const double padR = 8, padT = 8, padB = 16, edge = 4;
        double barH = style.FillHeight ? Math.Round(height * 0.18) : style.BarHeight;
        double chartH = style.FillHeight ? Math.Max(60, height - barH - 6) : style.ChartHeight;
        double barBase = chartH + barH + 4;

        // 値の範囲（回答の残高と推定の両方）。全行に残高がなければ推移は描かず棒だけにする
        long? min = null, max = null;
        foreach (var tx in data.Transactions)
        {
            var v = tx.Balance ?? tx.EstimatedBalance;
            if (v == null) continue;
            if (min == null || v < min) min = v;
            if (max == null || v > max) max = v;
        }
        bool hasLine = min != null;
        double lo = 0, hi = 1;
        if (hasLine)
        {
            double pad = Math.Max((max.Value - min.Value) * 0.08, 1);
            lo = min.Value - pad;
            hi = max.Value + pad;
        }
        Func<long, double> yOf = v => padT + (1 - (v - lo) / (hi - lo)) * (chartH - padT - padB);

        // 縦の目盛: きりのよい刻み（1・2・5×10^n・4本以下）で単位を揃える。左の余白は
        // 目盛ラベルの実測幅だけ取る（固定幅にすると短いラベルのとき左に空白が残る）
        var ticks = new List<KeyValuePair<long, string>>();
        double labelWidth = 0;
        if (hasLine)
        {
            long step = NiceTickStep(hi - lo);
            for (long v = (long)Math.Ceiling(lo / step) * step; v <= hi; v += step)
            {
                var label = FormatBalanceTick(v, step);
                ticks.Add(new KeyValuePair<long, string>(v, label));
                labelWidth = Math.Max(labelWidth, MeasureChartText(label));
            }
        }
        double padL = hasLine ? labelWidth + 8 : edge;

        // 横軸は最初の取引日〜最後の取引日を描画域いっぱいに割り当てる（月初・月末に丸めると
        // 先頭と末尾に空白ができる）。同じ日だけなら前後1日を足して幅を持たせる
        var periodFrom = data.PeriodFrom.Value;
        var periodTo = data.PeriodTo.Value;
        if (periodTo <= periodFrom)
        {
            periodFrom = periodFrom.AddDays(-1);
            periodTo = periodTo.AddDays(1);
        }
        double span = (periodTo - periodFrom).TotalDays;
        double plotW = width - padL - padR - edge * 2;
        Func<DateTime, double> xOf = d => padL + edge + (d - periodFrom).TotalDays / span * plotW;

        // 月境界（範囲内の月初）と月ラベル。先頭の月は月初が範囲外のため左端にラベルだけ置く
        //（最初の月境界が近ければ重なるので省く）
        double labelY = chartH - padB + 4;
        var firstBoundary = new DateTime(periodFrom.Year, periodFrom.Month, 1).AddMonths(1);
        if (firstBoundary > periodTo || xOf(firstBoundary) - xOf(periodFrom) >= 26)
            AddChartText(canvas, periodFrom.Month + "月", xOf(periodFrom), labelY, false);
        for (var month = firstBoundary; month <= periodTo; month = month.AddMonths(1))
        {
            double x = xOf(month);
            AddChartLine(canvas, x, padT, x, barBase, BrushLineWeak, 1, null);
            if (xOf(periodTo) - x >= 26) AddChartText(canvas, month.Month + "月", x + 3, labelY, false);
        }

        foreach (var tick in ticks)
        {
            double y = yOf(tick.Key);
            AddChartLine(canvas, padL, y, width - padR, y, BrushLineWeak, 1, null);
            AddChartText(canvas, tick.Value, padL - 4, y - 6, true);
        }

        // 折れ線: 実線と破線を別のジオメトリにまとめる（点の数だけ図形を作らない）
        if (hasLine)
        {
            var solid = new StreamGeometry();
            var dashed = new StreamGeometry();
            DetailChartPoint previous = null;
            using (var solidCtx = solid.Open())
            using (var dashedCtx = dashed.Open())
            {
                foreach (var tx in data.Transactions)
                {
                    var v = tx.Balance ?? tx.EstimatedBalance;
                    if (tx.Date == null || v == null) continue;
                    var point = new DetailChartPoint { X = xOf(tx.Date.Value), Y = yOf(v.Value), Tx = tx };
                    if (previous != null)
                    {
                        var ctx = (previous.Tx.Balance == null || tx.Balance == null) ? dashedCtx : solidCtx;
                        ctx.BeginFigure(new Point(previous.X, previous.Y), false, false);
                        ctx.LineTo(new Point(point.X, point.Y), true, false);
                    }
                    points.Add(point);
                    previous = point;
                }
            }
            solid.Freeze();
            dashed.Freeze();
            canvas.Children.Add(new System.Windows.Shapes.Path { Data = solid, Stroke = BrushAccent, StrokeThickness = 1.5 });
            canvas.Children.Add(new System.Windows.Shapes.Path
            {
                Data = dashed, Stroke = BrushAccent, StrokeThickness = 1.5,
                StrokeDashArray = new DoubleCollection { 3, 2.5 }
            });

            // 点: 塗りと枠の組ごとに 1 個の Path にまとめる（取引の数だけ図形を作らない。重なった点も抜けないよう Nonzero）
            bool markers = points.Count <= style.MarkerLimit;
            var rings = new StreamGeometry { FillRule = FillRule.Nonzero };
            var solidDots = new StreamGeometry { FillRule = FillRule.Nonzero };
            var hollowDots = new StreamGeometry { FillRule = FillRule.Nonzero };
            using (var ringCtx = rings.Open())
            using (var solidCtx = solidDots.Open())
            using (var hollowCtx = hollowDots.Open())
            {
                foreach (var point in points)
                {
                    if (style.GapRings && point.Tx.Gap != null) AddCircle(ringCtx, point.X, point.Y, 6.5);
                    if (!markers) continue;
                    if (point.Tx.Balance != null) AddCircle(solidCtx, point.X, point.Y, 3.2);
                    else AddCircle(hollowCtx, point.X, point.Y, 2.6);
                }
            }
            AddChartGeometry(canvas, rings, Brushes.Transparent, BrushChipCoralFg, 1.2);
            AddChartGeometry(canvas, solidDots, BrushAccent, null, 0);
            AddChartGeometry(canvas, hollowDots, Brushes.White, BrushAccent, 1.4);
        }

        // 入出金の棒（下段）。高さは最大額に対する比
        long maxAmount = 0;
        foreach (var tx in data.Transactions)
            maxAmount = Math.Max(maxAmount, Math.Max(tx.Deposit ?? 0, tx.Withdrawal ?? 0));
        if (maxAmount > 0)
        {
            // 棒: 色ごとに 1 個の Path にまとめる（入金はカテゴリの色・出金は細線。描く順は入金→出金。
            // 同じ日の棒は丸ごと重なるので Nonzero で塗る）
            double barW = data.Transactions.Count > style.MarkerLimit ? 1.5 : 4;
            var depositBars = new Dictionary<Brush, GeometryGroup>();
            var withdrawBars = new GeometryGroup { FillRule = FillRule.Nonzero };
            foreach (var tx in data.Transactions)
            {
                if (tx.Date == null) continue;
                double x = xOf(tx.Date.Value);
                if (tx.Deposit > 0)
                {
                    double h = Math.Max(1.5, (double)tx.Deposit.Value / maxAmount * barH);
                    var brush = DetailCategoryBrush(tx.Category);
                    GeometryGroup group;
                    if (!depositBars.TryGetValue(brush, out group)) depositBars[brush] = group = new GeometryGroup { FillRule = FillRule.Nonzero };
                    group.Children.Add(new RectangleGeometry(new Rect(x - barW / 2, barBase - h, barW, h), 0.8, 0.8));
                }
                if (tx.Withdrawal > 0)
                {
                    double h = Math.Max(1.5, (double)tx.Withdrawal.Value / maxAmount * barH);
                    withdrawBars.Children.Add(new RectangleGeometry(new Rect(x - 0.6, barBase - h, 1.2, h), 0.8, 0.8));
                }
            }
            foreach (var pair in depositBars) AddChartGeometry(canvas, pair.Value, pair.Key, null, 0);
            AddChartGeometry(canvas, withdrawBars, BrushDetailWithdraw, null, 0);
        }
        points.Sort(delegate(DetailChartPoint a, DetailChartPoint b) { return a.X.CompareTo(b.X); });
    }

    // 目盛の刻み: 1・2・5×10^n のうち、値域に入る目盛が4本以下になる最小のもの（1円未満にはしない）
    private static long NiceTickStep(double range)
    {
        double magnitude = Math.Pow(10, Math.Floor(Math.Log10(Math.Max(range / 4, 1))));
        foreach (var factor in new[] { 1, 2, 5, 10, 20, 50 })
        {
            long step = Math.Max(1, (long)Math.Round(factor * magnitude));
            if (range / step <= 4) return step;
        }
        return Math.Max(1, (long)Math.Round(100 * magnitude));
    }

    // 目盛の表記。刻みが1万円以上なら万円単位（刻みは1万の倍数なので整数で表せる）、
    // それ未満はそのまま。軸の中で単位が混ざらない
    private static string FormatBalanceTick(long value, long step)
    {
        if (value != 0 && step >= 10000) return (value / 10000).ToString("#,0") + "万";
        return value.ToString("#,0");
    }

    // 入金の棒の色（検出カテゴリの文字色。当たらない入金は薄いグレー）
    private static Brush DetailCategoryBrush(string category)
    {
        switch (category)
        {
            case "年金": return BrushChipTealFg;
            case "給与": return BrushChipVioletFg;
            case "保険": return BrushChipAmberFg;
            case "定期入金":
            case "定期出金": return BrushChipNeutralFg;
            default: return BrushDetailOtherDeposit;
        }
    }

    private static void AddChartLine(Canvas canvas, double x1, double y1, double x2, double y2, Brush stroke, double thickness, DoubleCollection dash)
    {
        var line = new System.Windows.Shapes.Line { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Stroke = stroke, StrokeThickness = thickness };
        if (dash != null) line.StrokeDashArray = dash;
        canvas.Children.Add(line);
    }

    private static void AddChartText(Canvas canvas, string text, double x, double y, bool rightAlign)
    {
        var block = new TextBlock { Text = text, FontFamily = FontMono, FontSize = 9.5, Foreground = BrushTextMuted };
        if (rightAlign) x -= MeasureChartText(text);
        Canvas.SetLeft(block, x);
        Canvas.SetTop(block, y);
        canvas.Children.Add(block);
    }

    // 軸ラベルの描画幅（目盛の幅を決めるために先に測る）
    private static double MeasureChartText(string text)
    {
        var block = new TextBlock { Text = text, FontFamily = FontMono, FontSize = 9.5 };
        block.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
        return block.DesiredSize.Width;
    }

    // 円を閉じた図形として足す（上半分と下半分の弧）
    private static void AddCircle(StreamGeometryContext ctx, double x, double y, double radius)
    {
        ctx.BeginFigure(new Point(x - radius, y), true, true);
        ctx.ArcTo(new Point(x + radius, y), new Size(radius, radius), 0, false, SweepDirection.Clockwise, true, false);
        ctx.ArcTo(new Point(x - radius, y), new Size(radius, radius), 0, false, SweepDirection.Clockwise, true, false);
    }

    // ジオメトリを 1 個の Path にして足す（凍結・ヒットテストなし＝ホバーは Canvas が受ける）
    private static void AddChartGeometry(Canvas canvas, Geometry geometry, Brush fill, Brush stroke, double thickness)
    {
        geometry.Freeze();
        canvas.Children.Add(new System.Windows.Shapes.Path
        {
            Data = geometry, Fill = fill, Stroke = stroke, StrokeThickness = thickness, IsHitTestVisible = false
        });
    }

    // ホバー位置に最も近い点（横方向）のツールチップを出す。点が近くになければ閉じる
    private void ShowDetailPointTip(Canvas canvas, ToolTip tip, List<DetailChartPoint> points, Point position)
    {
        DetailChartPoint nearest = null;
        double best = 8;   // 当たりとみなす横距離（px）
        int lowIdx = 0, highIdx = points.Count - 1;
        while (lowIdx <= highIdx)   // X 昇順のため二分探索で近傍を絞る
        {
            int mid = (lowIdx + highIdx) / 2;
            if (points[mid].X < position.X) lowIdx = mid + 1; else highIdx = mid - 1;
        }
        for (int i = Math.Max(0, lowIdx - 1); i <= Math.Min(points.Count - 1, lowIdx); i++)
        {
            double distance = Math.Abs(points[i].X - position.X);
            if (distance < best) { best = distance; nearest = points[i]; }
        }
        if (nearest == null)
        {
            tip.IsOpen = false;
            return;
        }
        tip.Content = BuildDetailPointTipContent(nearest.Tx);
        tip.HorizontalOffset = nearest.X + 12;
        tip.VerticalOffset = nearest.Y - 8;
        tip.IsOpen = true;
    }

    // 点のツールチップ（暗色サーフェス）: 日付・摘要・入出金（検出チップ付き）・残高または推定・差額
    private UIElement BuildDetailPointTipContent(DetailTransaction tx)
    {
        var stack = new StackPanel { MaxWidth = 360 };
        stack.Children.Add(new TextBlock
        {
            Text = tx.Date == null ? "日付なし" : tx.Date.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture),
            FontFamily = FontMono, FontSize = 11, Foreground = BrushTipText
        });
        stack.Children.Add(new TextBlock { Text = tx.Summary, FontSize = 12, Foreground = BrushTipText, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 2, 0, 0) });

        if (tx.Deposit != null || tx.Withdrawal != null)
        {
            var amount = new TextBlock { FontSize = 12, Foreground = BrushTipText, Margin = new Thickness(0, 2, 0, 0) };
            amount.Inlines.Add(new Run(tx.Deposit != null ? "入金 " : "出金 "));
            amount.Inlines.Add(new Run((tx.Deposit ?? tx.Withdrawal).Value.ToString("#,0")) { FontFamily = FontMono });
            if (tx.Category != null) amount.Inlines.Add(new Run("　" + tx.Category) { Foreground = DetailTipCategoryBrush(tx.Category) });
            stack.Children.Add(amount);
        }

        var balance = new TextBlock { FontSize = 12, Margin = new Thickness(0, 2, 0, 0) };
        if (tx.Balance != null)
        {
            balance.Foreground = BrushTipText;
            balance.Inlines.Add(new Run("残高 "));
            balance.Inlines.Add(new Run(tx.Balance.Value.ToString("#,0")) { FontFamily = FontMono });
        }
        else if (tx.EstimatedBalance != null)
        {
            balance.Foreground = BrushTipNeutralFg;
            balance.Inlines.Add(new Run("残高 "));
            balance.Inlines.Add(new Run(tx.EstimatedBalance.Value.ToString("#,0")) { FontFamily = FontMono });
            balance.Inlines.Add(new Run("（推定）"));
        }
        else
        {
            balance.Foreground = BrushTipNeutralFg;
            balance.Text = "残高 不明";
        }
        stack.Children.Add(balance);

        if (tx.Gap != null)
        {
            var gap = new TextBlock { FontSize = 12, Foreground = BrushTipGapText, Margin = new Thickness(0, 2, 0, 0) };
            gap.Inlines.Add(new Run("明細にない差額 "));
            gap.Inlines.Add(new Run((tx.Gap > 0 ? "+" : "") + tx.Gap.Value.ToString("#,0")) { FontFamily = FontMono });
            stack.Children.Add(gap);
        }
        return stack;
    }

    // 暗色ツールチップ上のカテゴリ文字色（検出列のチップ配色と同系）
    private static Brush DetailTipCategoryBrush(string category)
    {
        switch (category)
        {
            case "年金": return BrushTipTealFg;
            case "給与": return BrushTipVioletFg;
            case "保険": return BrushTipAmberFg;
            default: return BrushTipNeutralFg;
        }
    }

    // --- 検出の表（台帳の検出行と明細の集約を1つにまとめる）---

    // 横並びの2要素（先頭＝名称・2つ目＝周期）: 2つ目の幅を先に確保し、残りの幅を先頭に渡す。
    // 先頭が収まらなければ先頭だけが縮み（TextTrimming で省略記号）、2つ目は先頭の直後に全部出る。
    // 文字の大きさが違う2つは下端で揃える（同じフォントなら下端揃えで基線がほぼ揃う）。
    // StackPanel は子を無限幅で測るため、横並びのまま先頭だけを省略させるにはこの配置が要る
    private class ShrinkFirstPanel : Panel
    {
        protected override Size MeasureOverride(Size availableSize)
        {
            if (Children.Count == 0) return new Size();
            double tailWidth = 0, height = 0;
            if (Children.Count > 1)
            {
                Children[1].Measure(new Size(double.PositiveInfinity, availableSize.Height));
                tailWidth = Children[1].DesiredSize.Width;
                height = Children[1].DesiredSize.Height;
            }
            Children[0].Measure(new Size(Math.Max(0, availableSize.Width - tailWidth), availableSize.Height));
            height = Math.Max(height, Children[0].DesiredSize.Height);
            return new Size(Children[0].DesiredSize.Width + tailWidth, height);
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            if (Children.Count == 0) return finalSize;
            double tailWidth = Children.Count > 1 ? Children[1].DesiredSize.Width : 0;
            double headWidth = Math.Min(Children[0].DesiredSize.Width, Math.Max(0, finalSize.Width - tailWidth));
            double headHeight = Children[0].DesiredSize.Height;
            Children[0].Arrange(new Rect(0, finalSize.Height - headHeight, headWidth, headHeight));
            if (Children.Count > 1)
            {
                double tailHeight = Children[1].DesiredSize.Height;
                Children[1].Arrange(new Rect(headWidth, finalSize.Height - tailHeight, tailWidth, tailHeight));
            }
            return finalSize;
        }
    }

    // 行＝台帳の検出。明細で当たったものは合計の大きい順に 回数・1回あたり・直近 を添え、当たらなかった
    // ものは末尾に薄字で名称と周期だけ置く（台帳には検出があるが明細では見つからなかった、という事実）。
    // 列: 検出（チップ＋名称＋周期。名称だけを省略記号で切り周期は残す・全文はツールチップ）・回数（中央）・1回あたり（右）・直近（中央）
    private UIElement BuildDetailAggregateTable(List<DetectionAggregate> aggregates, List<DetailDetection> detections)
    {
        var grid = new Grid { Margin = new Thickness(0, 6, 0, 0) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        // 見出し行（検出は左・数値と日付は中央）
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        var headers = new[] { "検出", "回数", "1回あたり", "直近" };
        for (int c = 0; c < headers.Length; c++)
        {
            var head = new TextBlock
            {
                Text = headers[c],
                FontSize = 10,
                Foreground = BrushTextMuted,
                Margin = new Thickness(c == 0 ? 8 : 4, 0, 4, 2),   // 「検出」はチップ内の文字位置（左パディング 8）に揃える
                TextAlignment = c == 0 ? TextAlignment.Left : TextAlignment.Center
            };
            Grid.SetRow(head, 0);
            Grid.SetColumn(head, c);
            grid.Children.Add(head);
        }
        AddGridRule(grid, 0, headers.Length, BrushLine);

        // 行の並び: 当たった検出（集約の順）→ 当たらなかった検出（台帳の記載順）
        var rows = new List<KeyValuePair<DetailDetection, DetectionAggregate>>();
        foreach (var aggregate in aggregates)
        {
            var detection = detections.Find(d => d.Category == aggregate.Category && d.Name == aggregate.Name);
            rows.Add(new KeyValuePair<DetailDetection, DetectionAggregate>(detection, aggregate));
        }
        foreach (var detection in detections)
        {
            bool matched = false;
            foreach (var pair in rows) if (ReferenceEquals(pair.Key, detection)) matched = true;
            if (!matched) rows.Add(new KeyValuePair<DetailDetection, DetectionAggregate>(detection, null));
        }

        for (int i = 0; i < rows.Count; i++)
        {
            var detection = rows[i].Key;
            var aggregate = rows[i].Value;
            int row = i + 1;
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var category = aggregate != null ? aggregate.Category : detection.Category;
            var name = aggregate != null ? aggregate.Name : detection.Name;
            var bucket = detection != null ? detection.Bucket : null;
            var textBrush = aggregate != null ? BrushTextMain : BrushTextMuted;

            // 検出セル: チップの直後に名称（余白なし・検出行と同じ並び）
            var cell = new Grid { Margin = new Thickness(0, 3, 4, 3) };
            cell.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            cell.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            Brush chipBg, chipFg;
            DetailCategoryChipBrushes(category, out chipBg, out chipFg);
            var chip = new Border
            {
                Background = chipBg,
                CornerRadius = new CornerRadius(9),
                Padding = new Thickness(8, 1, 8, 2),
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Center
            };
            chip.Child = new TextBlock { Text = category, FontSize = 10.5, Foreground = chipFg };
            cell.Children.Add(chip);

            // 名称と周期: 入りきらないときは名称だけを省略し、周期は常に全部見せる
            var nameLine = new ShrinkFirstPanel
            {
                Margin = new Thickness(6, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                ToolTip = bucket != null ? name + "（" + bucket + "）" : name
            };
            nameLine.Children.Add(new TextBlock { Text = name, FontSize = 11, Foreground = textBrush, TextTrimming = TextTrimming.CharacterEllipsis });
            if (bucket != null) nameLine.Children.Add(new TextBlock { Text = "（" + bucket + "）", FontSize = 10, Foreground = BrushTextMuted });
            Grid.SetColumn(nameLine, 1);
            cell.Children.Add(nameLine);
            Grid.SetRow(cell, row);
            grid.Children.Add(cell);

            if (aggregate != null)
            {
                AddAggregateCell(grid, row, 1, aggregate.Count.ToString(), TextAlignment.Center, BrushTextMain);
                AddAggregateCell(grid, row, 2, aggregate.PerAmount.ToString("#,0"), TextAlignment.Right, BrushTextMain);
                AddAggregateCell(grid, row, 3, aggregate.LastDate == null ? "" : aggregate.LastDate.Value.ToString("MM/dd", CultureInfo.InvariantCulture), TextAlignment.Center, BrushTextSub);
            }
            if (i < rows.Count - 1) AddGridRule(grid, row, headers.Length, BrushLineWeak);
        }
        return grid;
    }

    private static void AddAggregateCell(Grid grid, int row, int column, string text, TextAlignment align, Brush foreground)
    {
        var cell = new TextBlock
        {
            Text = text,
            FontFamily = FontMono,
            FontSize = 11,
            Foreground = foreground,
            TextAlignment = align,
            Margin = new Thickness(4, 0, 4, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetRow(cell, row);
        Grid.SetColumn(cell, column);
        grid.Children.Add(cell);
    }

    // 行の下辺に罫線を引く（全列にまたがる Border・内容の下に重ねる）
    private static void AddGridRule(Grid grid, int row, int columnSpan, Brush brush)
    {
        var rule = new Border { BorderBrush = brush, BorderThickness = new Thickness(0, 0, 0, 1), VerticalAlignment = VerticalAlignment.Bottom };
        Grid.SetRow(rule, row);
        Grid.SetColumnSpan(rule, columnSpan);
        grid.Children.Add(rule);
    }

    // 検出カテゴリのチップ配色（一覧・展開部と同一）
    private static void DetailCategoryChipBrushes(string category, out Brush bg, out Brush fg)
    {
        switch (category)
        {
            case "年金": bg = BrushChipTealBg; fg = BrushChipTealFg; break;
            case "給与": bg = BrushChipVioletBg; fg = BrushChipVioletFg; break;
            case "保険": bg = BrushChipAmberBg; fg = BrushChipAmberFg; break;
            default: bg = BrushChipNeutralBg; fg = BrushChipNeutralFg; break;
        }
    }

    // --- 明細ウィンドウ（⧉）---
    //
    // 口座単位の子ウィンドウ（Owner 付き・非モーダル）。パネルと同じ読み取り結果を参照し、
    // 大きい推移グラフ・凡例・絞り込み（検出のみ／入金／出金／すべて＋摘要の検索）・全行の表
    // （仮想化）を出す。同じ口座を再度開くと既存のウィンドウを前面へ出す

    private const string XAML_NS = "xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'";

    // 表の1行（バインディング用）
    public class DetailRowView
    {
        public string Date { get; set; }
        public string Summary { get; set; }
        public string Withdrawal { get; set; }
        public string Deposit { get; set; }
        public string Balance { get; set; }
        public Brush BalanceBrush { get; set; }
        public FontStyle BalanceFontStyle { get; set; }
        public string BalanceTip { get; set; }
        public Visibility GapVisibility { get; set; }
        public string GapTip { get; set; }
        public string Category { get; set; }
        public Brush ChipBg { get; set; }
        public Brush ChipFg { get; set; }
        public Visibility ChipVisibility { get; set; }
        public Brush RowBackground { get; set; }
    }

    // 人物ファイルを「最後に使った」順の末尾へ移し、上限を超えた古いファイルの骨格・更新日時・シートのセルと、
    // そのファイルから読んだ結果を捨てる（次に開くと骨格から読み直す）
    private void TouchDetailCache(string path)
    {
        if (string.IsNullOrEmpty(path)) return;
        detailCacheOrder.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
        detailCacheOrder.Add(path);
        while (detailCacheOrder.Count > DETAIL_CACHE_MAX_FILES)
        {
            var victim = detailCacheOrder[0];
            detailCacheOrder.RemoveAt(0);
            detailWorkbookCache.Remove(victim);
            detailWorkbookStamps.Remove(victim);
            foreach (var sheetKey in detailSheetCache.Keys.Where(k => k.StartsWith(victim + "|", StringComparison.OrdinalIgnoreCase)).ToList())
                detailSheetCache.Remove(sheetKey);
            foreach (var rowKey in detailResultCache.Where(p => p.Value.Data != null && string.Equals(p.Value.Data.SourcePath, victim, StringComparison.OrdinalIgnoreCase)).Select(p => p.Key).ToList())
                detailResultCache.Remove(rowKey);
        }
    }

    // 台帳を読み直すとき: 行キー→結果の対応は台帳の検出に依存するため捨て、開いている
    // 明細ウィンドウも閉じる（照会ファイルの骨格・セルのキャッシュは更新日時で管理しているので残す）
    private void ResetDetailResults()
    {
        detailResultCache.Clear();
        detailChosenPaths.Clear();
        foreach (var detailWindow in new List<Window>(detailWindows.Values)) detailWindow.Close();
    }

    private void OpenDetailWindow(string[] fields, DetailAccountData data)
    {
        var rowKey = BusinessLogic.GetDepositLedgerRowKey(fields, eraMapping);
        Window existing;
        if (detailWindows.TryGetValue(rowKey, out existing))
        {
            existing.Activate();
            return;
        }

        var detailWindow = new Window
        {
            Title = "明細",
            Owner = window,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Width = lastDetailWindowSize.Width,
            Height = lastDetailWindowSize.Height,
            MinWidth = 720,
            MinHeight = 480,
            Background = Brushes.White,
            FontFamily = FontUi,
            ShowInTaskbar = false,
            Icon = window.Icon
        };
        var root = new Grid { LayoutTransform = displayScaleTransform };   // 表示倍率に追従
        for (int i = 0; i < 6; i++)   // グラフ（1）と表（4）はウィンドウの高さを 1:1 で分ける
            root.RowDefinitions.Add(new RowDefinition { Height = i == 1 || i == 4 ? new GridLength(1, GridUnitType.Star) : GridLength.Auto });

        // 0: タイトル行
        var titleBar = new Border { Background = BrushIrisPanelBg, BorderBrush = BrushLine, BorderThickness = new Thickness(0, 0, 0, 1), Padding = new Thickness(14, 9, 14, 9) };
        var titleGrid = new Grid();
        titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        // 誰の・どの口座・いつの回答かをウィンドウ単体で読めるようにする（本体で別の人物を選んでも紐付けが分かる）。
        // 1行の見出しなので強弱は付けず、番号と日付だけ等幅にする
        var titleText = new TextBlock { FontSize = 12, Foreground = BrushTextMain, VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis };
        titleText.Inlines.Add(new Run(fields[BusinessLogic.DEPOSIT_COL_NAME] + " "));
        titleText.Inlines.Add(new Run(fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM]) { FontFamily = FontMono });
        titleText.Inlines.Add(new Run("　" + fields[BusinessLogic.DEPOSIT_COL_INSTITUTION] + " " + fields[BusinessLogic.DEPOSIT_COL_BRANCH_NAME] + " " + fields[BusinessLogic.DEPOSIT_COL_ACCOUNT_TYPE] + " "));
        titleText.Inlines.Add(new Run(fields[BusinessLogic.DEPOSIT_COL_ACCOUNT_NUM]) { FontFamily = FontMono });
        titleText.Inlines.Add(new Run("　"));
        titleText.Inlines.Add(new Run(BusinessLogic.FormatLedgerDate(fields[BusinessLogic.DEPOSIT_COL_BASE_DATE], eraMapping)) { FontFamily = FontMono });
        titleGrid.Children.Add(titleText);
        var sourceText = new TextBlock
        {
            Text = System.IO.Path.GetFileName(data.SourcePath ?? "") + " ／ " + (data.SheetName ?? ""),
            FontSize = 10.5, Foreground = BrushTextMuted, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(16, 0, 0, 0)
        };
        Grid.SetColumn(sourceText, 1);
        titleGrid.Children.Add(sourceText);
        titleBar.Child = titleGrid;
        root.Children.Add(titleBar);

        // 1: 大グラフ（スパークラインと同じ描画関数・行の高さいっぱいに描く・常に点を打つ）
        var chart = BuildDetailChartHost(data, DetailLargeChartStyle);
        ((FrameworkElement)chart).Margin = new Thickness(14, 10, 14, 4);
        ((FrameworkElement)chart).MinHeight = 120;
        Grid.SetRow(chart, 1);
        root.Children.Add(chart);

        // 2: 凡例
        var legend = BuildDetailLegend();
        Grid.SetRow(legend, 2);
        root.Children.Add(legend);

        // 3: 絞り込み行（セグメント＋検索欄＋件数）
        var filterBar = new Border { Background = BrushIrisPanelBg, BorderBrush = BrushLine, BorderThickness = new Thickness(0, 1, 0, 1), Padding = new Thickness(14, 6, 14, 2) };
        var filterGrid = new Grid();
        filterGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        filterGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        filterGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        filterGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var segments = new StackPanel { Orientation = Orientation.Horizontal };
        var toggleStyle = (Style)window.FindResource("AdvToggleChip");
        var modes = new[]
        {
            new KeyValuePair<string, BusinessLogic.DetailFilterMode>("検出のみ", BusinessLogic.DetailFilterMode.Hits),
            new KeyValuePair<string, BusinessLogic.DetailFilterMode>("入金", BusinessLogic.DetailFilterMode.Deposits),
            new KeyValuePair<string, BusinessLogic.DetailFilterMode>("出金", BusinessLogic.DetailFilterMode.Withdrawals),
            new KeyValuePair<string, BusinessLogic.DetailFilterMode>("すべて", BusinessLogic.DetailFilterMode.All)
        };
        // 既定は「検出のみ」。検出に当たった取引が1件もない口座は「すべて」で開く（0件の表で開かないように）
        var initialMode = data.Aggregates.Count > 0 ? BusinessLogic.DetailFilterMode.Hits : BusinessLogic.DetailFilterMode.All;
        var toggles = new List<ToggleButton>();
        foreach (var mode in modes)
        {
            var toggle = new ToggleButton { Style = toggleStyle, Content = mode.Key, Tag = mode.Value, IsChecked = mode.Value == initialMode };
            toggles.Add(toggle);
            segments.Children.Add(toggle);
        }
        filterGrid.Children.Add(segments);

        // 検索欄は本体の検索欄と同じ構成（角丸の枠・🔍・枠なしの入力欄・プレースホルダ）
        var searchFrame = new Border
        {
            Width = 220, Height = 26, Margin = new Thickness(10, 0, 0, 4),
            Background = Brushes.White, BorderBrush = BrushLine, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(7)
        };
        var searchGrid = new Grid();
        searchGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        searchGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        searchGrid.Children.Add(new TextBlock { Text = "🔍", FontSize = 11, Foreground = BrushTextMuted, Margin = new Thickness(9, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center });
        var searchBox = new TextBox
        {
            BorderThickness = new Thickness(0), Background = Brushes.Transparent, FontSize = 11.5, Foreground = BrushTextMain,
            Padding = new Thickness(6, 3, 6, 3), VerticalContentAlignment = VerticalAlignment.Center
        };
        Grid.SetColumn(searchBox, 1);
        searchGrid.Children.Add(searchBox);
        var searchPlaceholder = new TextBlock
        {
            Text = "摘要で絞り込み", FontSize = 11.5, Foreground = BrushTextMuted,
            Margin = new Thickness(9, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center, IsHitTestVisible = false
        };
        Grid.SetColumn(searchPlaceholder, 1);
        searchGrid.Children.Add(searchPlaceholder);
        WirePlaceholder(searchBox, searchPlaceholder);
        searchFrame.Child = searchGrid;
        Grid.SetColumn(searchFrame, 1);
        filterGrid.Children.Add(searchFrame);

        var countText = new TextBlock { FontFamily = FontMono, FontSize = 11, Foreground = BrushTextMuted, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 0, 4) };
        Grid.SetColumn(countText, 3);
        filterGrid.Children.Add(countText);
        filterBar.Child = filterGrid;
        Grid.SetRow(filterBar, 3);
        root.Children.Add(filterBar);

        // 4: 全行の表（仮想化・行高32）
        var list = BuildDetailListView();
        Grid.SetRow(list, 4);
        root.Children.Add(list);

        // 5: 注記（件数・伏字・差額。パネルには置かない）
        var note = new TextBlock { FontSize = 10.5, Foreground = BrushTextMuted, Margin = new Thickness(14, 6, 14, 8), TextWrapping = TextWrapping.Wrap };
        int depositCount = 0, withdrawalCount = 0;
        foreach (var tx in data.Transactions)
        {
            if (tx.Deposit > 0) depositCount++;
            if (tx.Withdrawal > 0) withdrawalCount++;
        }
        note.Text = "取引 " + data.Transactions.Count + " 件（入金 " + depositCount + "・出金 " + withdrawalCount + "）。" +
            (data.MaskedCount > 0 ? "残高が伏せられた " + data.MaskedCount + " 行は入出金から推定した値です。" : "") +
            (data.GapCount > 0 ? "回答の残高と明細の積み上げが合わない箇所が " + data.GapCount + " 件あります（明細に載っていない取引の分・印の行）。" : "");
        Grid.SetRow(note, 5);
        root.Children.Add(note);

        // 絞り込みの適用（検索はデバウンス）
        var filterTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
        Action apply = delegate
        {
            var mode = BusinessLogic.DetailFilterMode.Hits;
            foreach (var toggle in toggles)
                if (toggle.IsChecked == true) mode = (BusinessLogic.DetailFilterMode)toggle.Tag;
            var filtered = BusinessLogic.FilterDetailTransactions(data.Transactions, mode, searchBox.Text);
            list.ItemsSource = BuildDetailRowViews(filtered);
            countText.Text = filtered.Count + " / " + data.Transactions.Count + " 件";
        };
        foreach (var toggle in toggles)
        {
            var self = toggle;
            self.Checked += delegate
            {
                foreach (var other in toggles)
                    if (!ReferenceEquals(other, self)) other.IsChecked = false;
                apply();
            };
            self.Unchecked += delegate
            {
                // 単一選択: 押し直しで全解除にはしない
                bool any = false;
                foreach (var other in toggles) any |= other.IsChecked == true;
                if (!any) self.IsChecked = true;
            };
        }
        searchBox.TextChanged += delegate { filterTimer.Stop(); filterTimer.Start(); };
        filterTimer.Tick += delegate { filterTimer.Stop(); apply(); };
        apply();

        // キー: Esc=閉じる・Ctrl+F=摘要の検索欄へ・Alt+1〜4=絞り込みの切替（左から順に）
        detailWindow.PreviewKeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                detailWindow.Close();
                e.Handled = true;
            }
            else if (e.Key == Key.F && Keyboard.Modifiers == ModifierKeys.Control)
            {
                searchBox.Focus();
                searchBox.SelectAll();
                e.Handled = true;
            }
            else if (e.Key == Key.System && Keyboard.Modifiers == ModifierKeys.Alt &&
                     e.SystemKey >= Key.D1 && e.SystemKey < Key.D1 + toggles.Count)
            {
                toggles[e.SystemKey - Key.D1].IsChecked = true;
                e.Handled = true;
            }
        };

        detailWindow.Content = root;
        detailWindows[rowKey] = detailWindow;
        detailWindow.Closed += delegate
        {
            filterTimer.Stop();
            detailWindows.Remove(rowKey);
            if (detailWindow.WindowState == WindowState.Normal) lastDetailWindowSize = new Size(detailWindow.Width, detailWindow.Height);
        };
        detailWindow.Show();
    }

    // 凡例（回答の残高／推定／入金／出金）
    private static UIElement BuildDetailLegend()
    {
        var legend = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(14, 0, 14, 6) };
        legend.Children.Add(MakeLegendItem(new System.Windows.Shapes.Ellipse { Width = 8, Height = 8, Fill = BrushAccent }, "回答の残高"));

        // 推定の見本は「短い破線の上に白抜き○」（グラフの描き方そのもの。文字で「破線」と言わない）
        var estimateMark = new Canvas { Width = 14, Height = 12 };
        estimateMark.Children.Add(new System.Windows.Shapes.Line
        {
            X1 = 0, X2 = 14, Y1 = 6, Y2 = 6, Stroke = BrushAccent, StrokeThickness = 1.5,
            StrokeDashArray = new DoubleCollection { 3, 2.5 }
        });
        var estimateDot = new System.Windows.Shapes.Ellipse { Width = 7, Height = 7, Fill = Brushes.White, Stroke = BrushAccent, StrokeThickness = 1.5 };
        Canvas.SetLeft(estimateDot, 3.5);
        Canvas.SetTop(estimateDot, 2.5);
        estimateMark.Children.Add(estimateDot);
        legend.Children.Add(MakeLegendItem(estimateMark, "入出金から推定"));
        legend.Children.Add(MakeLegendItem(new System.Windows.Shapes.Ellipse { Width = 8, Height = 8, Fill = Brushes.Transparent, Stroke = BrushChipCoralFg, StrokeThickness = 1.2 }, "明細にない差額"));

        // 入金の棒は当たった検出のカテゴリ色（年金・給与・保険・定期入金）・当たらない入金は薄灰。見本は5色を並べる
        var depositMarks = new StackPanel { Orientation = Orientation.Horizontal };
        foreach (var brush in new[] { BrushChipTealFg, BrushChipVioletFg, BrushChipAmberFg, BrushChipNeutralFg, BrushDetailOtherDeposit })
            depositMarks.Children.Add(new System.Windows.Shapes.Rectangle { Width = 4, Height = 10, Fill = brush, RadiusX = 0.8, RadiusY = 0.8, Margin = new Thickness(0, 0, 1.5, 0) });
        legend.Children.Add(MakeLegendItem(depositMarks, "入金"));
        legend.Children.Add(MakeLegendItem(new System.Windows.Shapes.Rectangle { Width = 2, Height = 10, Fill = BrushDetailWithdraw }, "出金"));
        return legend;
    }

    private static UIElement MakeLegendItem(FrameworkElement mark, string text)
    {
        var item = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 20, 0) };
        mark.VerticalAlignment = VerticalAlignment.Center;
        mark.Margin = new Thickness(0, 0, 4, 0);
        item.Children.Add(mark);
        item.Children.Add(new TextBlock { Text = text, FontSize = 10.5, Foreground = BrushTextMuted, VerticalAlignment = VerticalAlignment.Center });
        return item;
    }

    // 全行の表: 一覧と同じ GridView（フラットな列見出し・行高32・仮想化）。行の背景は検出に当たった行を淡アイリス面にする
    private ListView BuildDetailListView()
    {
        var list = new ListView
        {
            BorderThickness = new Thickness(0),
            Background = Brushes.White,
            FontSize = 12,
            SelectionMode = SelectionMode.Single,
            ItemContainerStyle = (Style)XamlReader.Parse(
                "<Style TargetType='ListViewItem' " + XAML_NS + " xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>" +
                "<Setter Property='Foreground' Value='#1A1523'/><Setter Property='HorizontalContentAlignment' Value='Stretch'/>" +
                "<Setter Property='Padding' Value='0'/><Setter Property='Margin' Value='0'/><Setter Property='Focusable' Value='False'/>" +
                "<Setter Property='Template'><Setter.Value><ControlTemplate TargetType='ListViewItem'>" +
                "<Border x:Name='rowBd' Height='32' Background='{Binding RowBackground}' BorderBrush='#F1F0F4' BorderThickness='0,0,0,1'>" +
                "<GridViewRowPresenter VerticalAlignment='Center'/></Border>" +
                "<ControlTemplate.Triggers><Trigger Property='IsMouseOver' Value='True'>" +
                "<Setter TargetName='rowBd' Property='Background' Value='#FAFAFC'/></Trigger></ControlTemplate.Triggers>" +
                "</ControlTemplate></Setter.Value></Setter></Style>")
        };
        VirtualizingStackPanel.SetIsVirtualizing(list, true);
        VirtualizingStackPanel.SetVirtualizationMode(list, VirtualizationMode.Recycling);
        ScrollViewer.SetHorizontalScrollBarVisibility(list, ScrollBarVisibility.Auto);

        var view = new GridView { AllowsColumnReorder = false };
        var headerStyle = (Style)window.FindResource(typeof(GridViewColumnHeader));
        var headerCenter = (Style)window.FindResource("ColHeadCenter");
        view.Columns.Add(MakeDetailColumn("日付", 104, headerStyle,
            "<TextBlock Text='{Binding Date}' FontFamily='Consolas' Foreground='#6F6E77' Margin='12,0,4,0'/>"));
        var summaryColumn = MakeDetailColumn("摘要", 320, headerStyle,
            "<TextBlock Text='{Binding Summary}' TextTrimming='CharacterEllipsis' ToolTip='{Binding Summary}' Margin='12,0,4,0'/>");
        view.Columns.Add(summaryColumn);
        view.Columns.Add(MakeDetailColumn("出金", 112, headerCenter,
            "<TextBlock Text='{Binding Withdrawal}' FontFamily='Consolas' HorizontalAlignment='Right' Margin='4,0,10,0'/>"));
        view.Columns.Add(MakeDetailColumn("入金", 112, headerCenter,
            "<TextBlock Text='{Binding Deposit}' FontFamily='Consolas' HorizontalAlignment='Right' Margin='4,0,10,0'/>"));
        view.Columns.Add(MakeDetailColumn("残高", 136, headerCenter,
            "<StackPanel Orientation='Horizontal' HorizontalAlignment='Right' Margin='4,0,10,0'>" +
            "<TextBlock Text='◦' Foreground='#993C1D' Margin='0,0,4,0' ToolTip='{Binding GapTip}' Visibility='{Binding GapVisibility}'/>" +
            "<TextBlock Text='{Binding Balance}' FontFamily='Consolas' Foreground='{Binding BalanceBrush}' FontStyle='{Binding BalanceFontStyle}' ToolTip='{Binding BalanceTip}'/></StackPanel>"));
        view.Columns.Add(MakeDetailColumn("検出", 96, headerStyle,
            "<Border Background='{Binding ChipBg}' CornerRadius='9' Padding='8,1,8,2' HorizontalAlignment='Left' Margin='12,0,4,0' Visibility='{Binding ChipVisibility}'>" +
            "<TextBlock Text='{Binding Category}' FontSize='10.5' Foreground='{Binding ChipFg}'/></Border>"));
        list.View = view;

        // 摘要の列は残りの幅を使う（他列の合計と縦スクロールバーの分を除く）
        list.SizeChanged += delegate(object s, SizeChangedEventArgs e)
        {
            double others = 0;
            foreach (var column in view.Columns)
                if (!ReferenceEquals(column, summaryColumn)) others += column.Width;
            summaryColumn.Width = Math.Max(160, e.NewSize.Width - others - 22);
        };
        return list;
    }

    private static GridViewColumn MakeDetailColumn(string header, double width, Style headerStyle, string cellXaml)
    {
        return new GridViewColumn
        {
            Header = header,
            Width = width,
            HeaderContainerStyle = headerStyle,
            CellTemplate = (DataTemplate)XamlReader.Parse("<DataTemplate " + XAML_NS + ">" + cellXaml + "</DataTemplate>")
        };
    }

    // 表示用の行を組み立てる（伏字の残高は推定値を薄い斜体・差額のある行は残高の左に印＝数字の右端を揃える）
    private static List<DetailRowView> BuildDetailRowViews(List<DetailTransaction> transactions)
    {
        var rows = new List<DetailRowView>(transactions.Count);
        foreach (var tx in transactions)
        {
            var row = new DetailRowView
            {
                Date = tx.Date == null ? "" : tx.Date.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture),
                Summary = tx.Summary,
                Withdrawal = tx.Withdrawal == null ? "" : tx.Withdrawal.Value.ToString("#,0"),
                Deposit = tx.Deposit == null ? "" : tx.Deposit.Value.ToString("#,0"),
                BalanceBrush = BrushTextMain,
                BalanceFontStyle = FontStyles.Normal,
                GapVisibility = Visibility.Collapsed,
                ChipVisibility = Visibility.Collapsed,
                RowBackground = tx.Category != null ? BrushIrisPanelBg : Brushes.White
            };
            if (tx.Balance != null) row.Balance = tx.Balance.Value.ToString("#,0");
            else if (tx.EstimatedBalance != null)
            {
                row.Balance = tx.EstimatedBalance.Value.ToString("#,0");
                row.BalanceBrush = BrushTextMuted;
                row.BalanceFontStyle = FontStyles.Italic;
                row.BalanceTip = tx.BalanceMasked ? "回答では伏せられているため、入出金から推定した値" : "回答に残高がないため、入出金から推定した値";
            }
            else row.Balance = "";
            if (tx.Gap != null)
            {
                row.GapVisibility = Visibility.Visible;
                row.GapTip = "回答の残高と明細の積み上げが合わない（明細にない差額 " + (tx.Gap > 0 ? "+" : "") + tx.Gap.Value.ToString("#,0") + "）";
            }
            if (tx.Category != null)
            {
                Brush bg, fg;
                DetailCategoryChipBrushes(tx.Category, out bg, out fg);
                row.Category = tx.Category;
                row.ChipBg = bg;
                row.ChipFg = fg;
                row.ChipVisibility = Visibility.Visible;
            }
            rows.Add(row);
        }
        return rows;
    }
}