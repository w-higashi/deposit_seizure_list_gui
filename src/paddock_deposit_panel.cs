// ==============================================================================
// paddock_deposit_panel.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 預金タブ: 詳細パネル（年度横断の人物ビュー・照会履歴のタイムライン・口座カードの
// 選択と展開・人物ヘッダー）。
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
    // 詳細パネル（年度横断の人物ビュー）
    // ==============================================================
    //
    // 一覧の行を選択すると右側に開き、当該人物の照会履歴を全年度の台帳から
    // 横断収集して表示する。選択中年度の行はメモリ上の depositAllRecords から即時に、
    // 他年度の行はファイルを非同期に走査して合流させる（人物単位の都度読み込み）。

    // 詳細パネルの回次1件（照会履歴タイムラインの1ノード）
    private class DepositPersonRound
    {
        public string Key;                                                   // 回次キー（回答基準日、口座なしのみの回次は登録日。正準形 yyyy/MM/dd）
        public List<string[]> AccountRows = new List<string[]>();            // 口座あり行（生の全列・収集順）
        public List<string> NoAccountInstitutions = new List<string>();      // 口座なし行の金融機関名
    }

    // パネルを開いて指定行の人物を表示する（一覧の選択行に対応するカードを選択状態にする）。
    // 生命保険のパネルが開いていれば先に閉じる（同時に開くのは一方だけ）
    private void OpenDepositDetailPanel(DepositLedgerRecord record)
    {
        if (insuranceTab.isInsurancePanelOpen) insuranceTab.CloseInsuranceDetailPanel(false);
        isDepositPanelOpen = true;
        SwitchDepositPanelPerson(record.AddressNumber, record.RawFields);
        AnimateDetailPanel(depositDetailPanel, DETAIL_PANEL_WIDTH);
    }

    // パネルを閉じる（表示中の人物情報は保持し、同一人物での再オープン時に再利用する。
    // 他年度の走査が進行中でも中断しない — 結果は非表示のパネルに無害に反映される）
    // returnFocus=false は検索条件の変更で閉じる場合（入力中の欄からフォーカスを奪わない）
    private void CloseDepositDetailPanel(bool returnFocus = true)
    {
        isDepositPanelOpen = false;
        AnimateDetailPanel(depositDetailPanel, 0);

        if (returnFocus) ReturnFocusFromDepositPanel();
    }

    // パネルを開いた元の一覧へフォーカスを戻す。台帳一覧では選択中の行コンテナへ
    //（ListView 本体へのフォーカスだと ↑↓ が行移動ではなくキーボードスクロールとして
    // 処理されてしまう）、最近見た人では選択中の履歴行へ
    private void ReturnFocusFromDepositPanel()
    {
        if (depositLedgerView.Visibility == Visibility.Visible)
        {
            if (!FocusDepositLedgerSelection()) depositLedgerList.Focus();
        }
        else if (recentView.Visibility == Visibility.Visible && recentSelectedBorder != null)
            recentSelectedBorder.Focus();
    }

    // 一覧の選択行の行コンテナへフォーカスする（選択がない・行が生成されていなければ false）
    private bool FocusDepositLedgerSelection()
    {
        var item = depositLedgerList.ItemContainerGenerator.ContainerFromItem(depositLedgerList.SelectedItem) as ListViewItem;
        if (item == null) return false;
        item.Focus();
        return true;
    }

    // パネルの表示対象を指定の人物に切り替える
    // focusFields は選択するカードの行（一覧の選択行と連動・口座なし行なら選択なし）。
    // 同一人物の別口座行を選択した場合はカード選択のみ更新する（アコーディオンの展開
    // 状態を保つ。人物ヘッダーは選択行連動のため SelectDepositCard 側が更新する）
    private void SwitchDepositPanelPerson(string addressNumber, string[] focusFields)
    {
        depositPendingCardFields = focusFields;   // 再構築時の選択復元用（未成立の選択の意図を保持する）
        depositPendingExpandFields = focusFields; // 選択の成立と同時に当該カードを自動展開する
        if (addressNumber == depositPanelAddressNumberShown)
        {
            ApplyDepositCardSelection(focusFields);
            return;
        }
        depositSelectedCardFields = focusFields;
        depositPanelScroll.ScrollToTop();   // 前の人物のスクロール位置を持ち越さない

        // 「すべての年度」（selectedYear=null）は全年度の行が depositAllRecords に載っているため他年度なし
        var selectedYear = GetSelectedYear(depositYearCombo);
        var otherYears = new List<int>();
        if (selectedYear != null)
            foreach (var year in depositLedgerFiles.Keys)
                if (year != selectedYear.Value) otherYears.Add(year);

        // 他年度がない、または全ての他年度がこの人物についてキャッシュ済みなら即時に合成表示
        bool allCached = true;
        foreach (var year in otherYears)
        {
            Dictionary<string, List<string[]>> yearCache;
            if (!depositOtherYearCache.TryGetValue(year, out yearCache) || !yearCache.ContainsKey(addressNumber))
            {
                allCached = false;
                break;
            }
        }
        if (otherYears.Count == 0 || allCached)
        {
            MergeAndRenderDepositPanel(addressNumber);
            return;
        }

        // まず選択中年度の分を即時表示し、他年度分はデバウンス後に非同期で合流させる
        //（走査中はヘッダーのミニスピナーで示す）
        RenderDepositPersonPanel(addressNumber, CollectDepositSelectedYearRows(addressNumber), null);
        SetDepositPanelScanSpinner(true);
        depositCrossYearTimer.Stop();
        depositCrossYearTimer.Start();
    }

    // 表示中の台帳（メモリ上の depositAllRecords）から指定人物の行を集める
    private List<string[]> CollectDepositSelectedYearRows(string addressNumber)
    {
        var rows = new List<string[]>();
        foreach (var record in depositAllRecords)
            if (record.AddressNumber == addressNumber) rows.Add(record.RawFields);
        return rows;
    }

    // 他年度台帳から表示中人物の行を走査する（キャッシュにない年度のみ）
    private void StartDepositCrossYearLoad()
    {
        if (depositPanelAddressNumberShown == null) return;
        if (depositPersonWorker != null && depositPersonWorker.IsBusy)
        {
            depositCrossYearPending = true;   // 走査完了後に改めて現在の人物で再走査する
            return;
        }

        var addressNumber = depositPanelAddressNumberShown;
        var selectedYear = GetSelectedYear(depositYearCombo);

        // 走査対象 = 選択中年度以外で、この人物の結果が未キャッシュの年度（「すべての年度」は走査なし）
        var targets = new List<KeyValuePair<int, string>>();
        foreach (var pair in depositLedgerFiles)
        {
            if (selectedYear == null || pair.Key == selectedYear.Value) continue;
            Dictionary<string, List<string[]>> yearCache;
            if (depositOtherYearCache.TryGetValue(pair.Key, out yearCache) && yearCache.ContainsKey(addressNumber)) continue;
            targets.Add(pair);
        }
        if (targets.Count == 0)
        {
            MergeAndRenderDepositPanel(addressNumber);
            return;
        }
        SetDepositPanelScanSpinner(true);   // 走査の間はヘッダーのミニスピナーで示す（入口によらずここで点ける）

        var epoch = depositLedgerEpoch;   // 走査を始めた世代（F5・台帳フォルダの変更で進む。古い世代の結果はキャッシュに入れない）
        depositPersonWorker = new BackgroundWorker();
        depositPersonWorker.DoWork += delegate(object s, DoWorkEventArgs e)
        {
            // 年度ファイルを丸ごと読み、この人物の行だけを拾って残りは捨てる
            // （台帳全体をメモリに保持しないための人物単位読み込み）
            var found = new Dictionary<int, List<string[]>>();
            var pool = new Dictionary<string, string>();   // 同じ値の文字列を 1 本に共有する（この走査の間だけ）
            foreach (var target in targets)
            {
                try
                {
                    var text = ReadLedgerText(target.Value);
                    var rawRecords = BusinessLogic.ParseCsvText(text);
                    var header = rawRecords.Count > 0 ? rawRecords[0] : new string[0];
                    List<string> missing;
                    var map = BusinessLogic.BuildLedgerColumnMap(header, BusinessLogic.DEPOSIT_LEDGER_COLUMNS, BusinessLogic.DEPOSIT_REQUIRED_COLUMNS, out missing);
                    if (map == null) continue;   // 見出しがない・必須の列が欠ける年度は読めなかった年度として扱う
                    bool identity = BusinessLogic.IsIdentityColumnMap(map, header.Length, BusinessLogic.DEPOSIT_LEDGER_COLUMN_COUNT);
                    var rows = new List<string[]>();
                    for (int i = 1; i < rawRecords.Count; i++)
                    {
                        var fields = rawRecords[i];
                        // 列数の合わない行は読み飛ばす（並びどおりなら27列未満だけ・組み直すときは見出しと違う列数の行）
                        if (identity ? fields.Length < BusinessLogic.DEPOSIT_LEDGER_COLUMN_COUNT : fields.Length != header.Length) continue;
                        if (!identity) fields = BusinessLogic.ReorderLedgerFields(fields, map);
                        if (fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM] != addressNumber) continue;
                        BusinessLogic.InternFields(fields, pool);
                        rows.Add(fields);
                    }
                    found[target.Key] = rows;
                }
                catch
                {
                    // 読めなかった年度はキャッシュに入れない（次回の切替時に再試行される）
                    // 失敗した事実は合成表示時に脚注として示す
                }
            }
            e.Result = new object[] { addressNumber, found };
        };
        depositPersonWorker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs e)
        {
            // DoWork の外で起きた例外（e.Result は null）は走査の結果なしとして同じ流れに進む
            //（キャッシュに入らないため次回の切替時に再試行される）
            if (e.Error != null) WriteErrorLog("年度横断の走査（預金）", e.Error);
            var parts = e.Error == null ? (object[])e.Result : new object[] { addressNumber, new Dictionary<int, List<string[]>>() };
            var resultAddressNumber = (string)parts[0];
            var found = (Dictionary<int, List<string[]>>)parts[1];

            if (epoch == depositLedgerEpoch)   // 捨てる操作のあとの完了は合流しない（別フォルダ・F5 前の行をキャッシュに入れない）
            {
                foreach (var pair in found)
                {
                    if (!depositOtherYearCache.ContainsKey(pair.Key))
                        depositOtherYearCache[pair.Key] = new Dictionary<string, List<string[]>>();
                    depositOtherYearCache[pair.Key][resultAddressNumber] = pair.Value;
                }
            }

            // 走査中に人物が切り替わっていた場合は現在の人物で改めて走査する
            if (depositCrossYearPending)
            {
                depositCrossYearPending = false;
                StartDepositCrossYearLoad();
                return;
            }
            if (epoch == depositLedgerEpoch && depositPanelAddressNumberShown == resultAddressNumber)
                MergeAndRenderDepositPanel(resultAddressNumber);
        };
        depositPersonWorker.RunWorkerAsync();
    }

    // 選択中年度の行とキャッシュ済みの他年度の行を合成してパネルを描画する
    private void MergeAndRenderDepositPanel(string addressNumber)
    {
        var rows = CollectDepositSelectedYearRows(addressNumber);
        var selectedYear = GetSelectedYear(depositYearCombo);
        var failedYears = new List<int>();

        // 「すべての年度」は depositAllRecords が全年度分（読めなかった年度は読み込み時の控えを引き継ぐ）
        if (selectedYear == null)
            failedYears.AddRange(depositLoadFailedYears);
        else
            foreach (var year in depositLedgerFiles.Keys)
            {
                if (year == selectedYear.Value) continue;
                Dictionary<string, List<string[]>> yearCache;
                List<string[]> personRows = null;
                if (depositOtherYearCache.TryGetValue(year, out yearCache))
                    yearCache.TryGetValue(addressNumber, out personRows);
                if (personRows != null) rows.AddRange(personRows);
                else failedYears.Add(year);
            }

        string footnote = null;
        if (failedYears.Count > 0)
            footnote = FormatYearLabels(failedYears) + " の台帳を読み込めませんでした";

        // 全年度を突き合わせても行が1件もない場合はその旨を示す（台帳の差し替え等で
        // 最近見た人の履歴の人物が台帳から消えた場合に起こりうる。落ちない・隠さない）
        if (rows.Count == 0 && footnote == null)
            footnote = "現在の台帳にこの人物の行が見つかりません";

        SetDepositPanelScanSpinner(false);   // 走査完了（合流結果の描画に切り替わる）
        RenderDepositPersonPanel(addressNumber, rows, footnote);
    }

    // 年度横断走査中のミニスピナーの表示切替（人物ヘッダー右端・走査中のみ回転させる）
    private void SetDepositPanelScanSpinner(bool scanning)
    {
        depositPanelMiniSpinner.Visibility = scanning ? Visibility.Visible : Visibility.Collapsed;
        if (scanning) StartSpinner(depositPanelSpinnerRotation);
        else StopSpinner(depositPanelSpinnerRotation);
    }

    // 収集済みの行（選択中年度+他年度・口座あり/なしの両方を含む）からパネルの内容を組み立てる
    private void RenderDepositPersonPanel(string addressNumber, List<string[]> rows, string footnote)
    {
        // メモは人物に属し行データに依存しないため、表示人物が変わったときのみ再構築する
        //（他年度分の合流による再描画で「さらに表示」の展開状態を巻き戻さない）
        if (addressNumber != depositPanelAddressNumberShown)
        {
            depositNotes.Expanded = false;
            BuildNotesSection(depositNotes, addressNumber);
        }
        depositPanelAddressNumberShown = addressNumber;

        // --- 人物ヘッダーの退行値: 登録日が最新の非空値（口座なし行は氏名等が空のため）。
        //     表示は選択カードの行の値を優先する（選択行連動——末尾の ApplyDepositCardSelection
        //     経由で SelectDepositCard → UpdateDepositPersonHeader が必ず描画する）---
        depositPanelFallbackName = null;
        depositPanelFallbackKana = null;
        depositPanelFallbackAddress = null;
        foreach (var fields in rows.OrderByDescending(
            f => BusinessLogic.GetRegistrationSortKey(f[BusinessLogic.DEPOSIT_COL_REG_DATE], eraMapping)))
        {
            if (depositPanelFallbackName == null && fields[BusinessLogic.DEPOSIT_COL_NAME] != "") depositPanelFallbackName = fields[BusinessLogic.DEPOSIT_COL_NAME];
            if (depositPanelFallbackKana == null && fields[BusinessLogic.DEPOSIT_COL_KANA] != "") depositPanelFallbackKana = fields[BusinessLogic.DEPOSIT_COL_KANA];
            if (depositPanelFallbackAddress == null && fields[BusinessLogic.DEPOSIT_COL_ADDRESS] != "") depositPanelFallbackAddress = fields[BusinessLogic.DEPOSIT_COL_ADDRESS];
        }
        depositPanelAddressNumber.Text = addressNumber;

        // --- 回次グルーピング+ファクト統計（事実の集計のみ・評価値は置かない）---
        var rounds = GroupDepositPersonRounds(rows);
        int accountCount = 0;
        foreach (var fields in rows)
            if (fields[BusinessLogic.DEPOSIT_COL_HAS_ACCOUNT] != "0") accountCount++;

        depositPanelStatRounds.Text = rounds.Count.ToString();
        depositPanelStatAccounts.Text = accountCount.ToString();
        depositPanelStatFirst.Text = rounds.Count > 0
            ? rounds[rounds.Count - 1].Key   // 回次キーは正準形（yyyy/MM/dd）のためそのまま表示できる
            : "－";

        // --- 照会履歴タイムライン（新しい回次が上）---
        bool cardFocused = depositTimelinePanel.IsKeyboardFocusWithin;
        depositTimelinePanel.Children.Clear();
        depositPanelCards.Clear();
        StopDetailSpinners();
        depositCardExpansions.Clear();
        for (int i = 0; i < rounds.Count; i++)
            depositTimelinePanel.Children.Add(BuildDepositRoundNode(rounds[i], i, i == rounds.Count - 1));

        // カード選択の復元。選択が成立していない場合（他年度の行の合流前に
        // 当該カードがまだ存在しない等）は、直近に指定された行での選択を試みる
        ApplyDepositCardSelection(depositSelectedCardFields != null ? depositSelectedCardFields : depositPendingCardFields);
        if (cardFocused) FocusLater(delegate { if (depositSelectedCardBorder != null) depositSelectedCardBorder.Focus(); });

        depositPanelFootnote.Text = footnote != null ? footnote : "";
        depositPanelFootnote.Visibility = footnote != null ? Visibility.Visible : Visibility.Collapsed;
    }

    // 行を回次（回答基準日、口座なしのみの回次は登録日）でグルーピングし、新しい順に並べる。
    // キーは正準形（yyyy/MM/dd）のため、日付形式が違っても同じ日付は同一回次に
    // まとまり、Ordinal 比較がそのまま時系列順になる
    private List<DepositPersonRound> GroupDepositPersonRounds(List<string[]> rows)
    {
        var map = new Dictionary<string, DepositPersonRound>();
        foreach (var fields in rows)
        {
            var baseDate = fields[BusinessLogic.DEPOSIT_COL_BASE_DATE];
            var key = BusinessLogic.FormatLedgerDate(
                baseDate != "" ? baseDate : fields[BusinessLogic.DEPOSIT_COL_REG_DATE], eraMapping);
            DepositPersonRound round;
            if (!map.TryGetValue(key, out round))
            {
                round = new DepositPersonRound { Key = key };
                map[key] = round;
            }
            if (fields[BusinessLogic.DEPOSIT_COL_HAS_ACCOUNT] == "0")
                round.NoAccountInstitutions.Add(fields[BusinessLogic.DEPOSIT_COL_INSTITUTION]);
            else
                round.AccountRows.Add(fields);
        }
        return map.Values.OrderByDescending(r => r.Key, StringComparer.Ordinal).ToList();
    }

    // 回次1件のタイムラインノード（レール+内容）を組み立てる
    // index はタイムライン上の位置（0=最新）。レールのドットの濃淡は時系列位置の表現であり評価ではない
    private UIElement BuildDepositRoundNode(DepositPersonRound round, int index, bool isLast)
    {
        var node = new Grid();
        node.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        node.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        // --- レール（ドット+次の回次への縦線。最終ノードは線なし）---
        var rail = new Grid { Width = 9, Margin = new Thickness(0, 4, 0, 0) };
        rail.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        rail.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        var dot = new System.Windows.Shapes.Ellipse
        {
            Width = 9,
            Height = 9,
            Fill = BrushAccent,
            Opacity = Math.Max(0.35, 1.0 - index * 0.13)
        };
        rail.Children.Add(dot);
        if (!isLast)
        {
            var line = new Border
            {
                Width = 1.5,
                Background = BrushRailLine,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 3, 0, 0)
            };
            Grid.SetRow(line, 1);
            rail.Children.Add(line);
        }
        node.Children.Add(rail);

        // --- 回次の内容（日付ヘッダー+口座行+口座なし行）---
        var content = new StackPanel { Margin = new Thickness(11, 0, 0, isLast ? 2 : 16) };
        Grid.SetColumn(content, 1);

        var head = new Grid { Margin = new Thickness(0, 0, 0, 7) };
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var dateText = new TextBlock
        {
            Text = round.Key,   // 回次キーは正準形（yyyy/MM/dd）のためそのまま表示できる
            FontFamily = FontMono,
            FontSize = 12.5,
            Foreground = BrushTextMain
        };
        head.Children.Add(dateText);
        if (round.AccountRows.Count > 0)
        {
            var countText = new TextBlock
            {
                Text = "口座 " + round.AccountRows.Count + "件",
                FontSize = 10.5,
                Foreground = BrushTextMuted,
                VerticalAlignment = VerticalAlignment.Bottom
            };
            Grid.SetColumn(countText, 1);
            head.Children.Add(countText);
        }
        content.Children.Add(head);

        foreach (var fields in round.AccountRows)
            content.Children.Add(BuildDepositAccountRow(fields));

        if (round.NoAccountInstitutions.Count > 0)
            content.Children.Add(BuildNoAccountLine("口座なし", round.NoAccountInstitutions));

        node.Children.Add(content);
        return node;
    }

    // 口座行1件（1行表示+クリックで選択と展開）を組み立てる
    // 全てのカードがクリックで選択でき、ダブルクリックでその回次の照会ファイルを開く。
    // 展開部は検出内容の行と明細セクション（照会ファイルを読める行のみ）。
    // 最終取引日は全回次で表示する（明細セクションのある口座はその先頭の ⧉ の行、ない口座は独立した1行）
    private UIElement BuildDepositAccountRow(string[] fields)
    {
        // 展開部の内容を先に組み立て、空なら展開なし（選択のみ）の表示にする
        bool detailEligible = IsDetailEligible(fields);
        StackPanel categoryLines;
        var detailLines = BuildDepositAccountDetailLines(fields, !detailEligible, out categoryLines);
        bool expandable = detailLines.Count > 0 || detailEligible;

        var card = new Border
        {
            Style = (Style)window.FindResource("AccountCard"),
            Margin = new Thickness(0, 0, 0, 6),
            Cursor = Cursors.Hand
        };
        depositPanelCards.Add(new KeyValuePair<string[], Border>(fields, card));

        var outer = new StackPanel();
        card.Child = outer;

        // --- 1行表示: 金融機関・支店 / 種別・口座番号 / 残高 ---
        var head = new Grid();
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var left = new StackPanel();
        var institution = fields[BusinessLogic.DEPOSIT_COL_INSTITUTION];
        var branch = fields[BusinessLogic.DEPOSIT_COL_BRANCH_NAME];
        left.Children.Add(new TextBlock
        {
            Text = branch != "" ? institution + " " + branch : institution,
            FontSize = 12,
            Foreground = BrushTextMain,
            TextTrimming = TextTrimming.CharacterEllipsis
        });
        var accountLine = new TextBlock
        {
            FontSize = 11,
            Foreground = BrushTextSub,
            Margin = new Thickness(0, 2, 0, 0)
        };
        var accountType = fields[BusinessLogic.DEPOSIT_COL_ACCOUNT_TYPE];
        if (accountType != "") accountLine.Inlines.Add(new Run(accountType + "・"));
        accountLine.Inlines.Add(new Run(fields[BusinessLogic.DEPOSIT_COL_ACCOUNT_NUM]) { FontFamily = FontMono });
        left.Children.Add(accountLine);
        head.Children.Add(left);

        // 右側は縦積み（残高の右端を全カードで揃え、シェブロンはその下に置く）
        var right = new StackPanel
        {
            VerticalAlignment = VerticalAlignment.Top,
            Margin = new Thickness(8, 0, 0, 0)
        };
        var balance = fields[BusinessLogic.DEPOSIT_COL_BALANCE];
        var balanceText = new TextBlock
        {
            FontFamily = FontMono,
            FontSize = 12,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        if (BusinessLogic.IsSentinelBalance(balance))
        {
            balanceText.Text = "対象外";
            balanceText.Foreground = BrushTextMuted;
            balanceText.ToolTip = "回答コード: " + balance;
        }
        else
        {
            balanceText.Text = balance;
            balanceText.Foreground = BrushTextMain;
        }
        right.Children.Add(balanceText);

        TextBlock chevron = null;
        if (expandable)
        {
            chevron = new TextBlock
            {
                Text = "▾",
                FontSize = 11.5,
                Foreground = BrushTextMuted,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 1, 0, 0)
            };
            right.Children.Add(chevron);
        }
        Grid.SetColumn(right, 1);
        head.Children.Add(right);
        outer.Children.Add(head);

        // --- 展開部（クリックで開閉・同時に開くのは1枚）---
        DepositCardExpansion expansion = null;
        if (expandable)
        {
            var detail = new StackPanel { Visibility = Visibility.Collapsed, Margin = new Thickness(0, 6, 0, 0) };
            foreach (var line in detailLines) detail.Children.Add(line);
            expansion = new DepositCardExpansion { Card = card, Fields = fields, Detail = detail, Chevron = chevron, CategoryLines = categoryLines };
            if (detailEligible)
            {
                var lastTx = fields[BusinessLogic.DEPOSIT_COL_LAST_TX];
                Decorator openSlot;
                expansion.Section = new StackPanel();
                expansion.Rule = detailLines.Count > 0 ? MakeDashedRule() : null;
                expansion.Toolbar = BuildDetailToolbar(lastTx, out openSlot);
                expansion.OpenSlot = openSlot;
                expansion.HasLastTx = lastTx != "";
                expansion.Frame = BuildDetailSectionFrame(expansion);
                UpdateDetailFrameSpacing(expansion);
                detail.Children.Add(expansion.Frame);
            }
            depositCardExpansions[card] = expansion;
            outer.Children.Add(detail);
        }

        // --- クリック／Space=選択+展開切替・ダブルクリック／Enter=照会ファイルを開く・↑↓=カードの移動 ---
        Action selectAndToggle = delegate
        {
            SelectDepositCard(fields, card);
            depositPendingExpandFields = null;   // 職員が自分で選んだカードが優先（保留中の自動展開は捨てる）
            if (expansion != null)
            {
                if (ReferenceEquals(depositExpandedCard, expansion)) CollapseDepositCard();
                else ExpandDepositCard(expansion, false);
            }
        };
        Action selectAndOpen = delegate
        {
            SelectDepositCard(fields, card);
            OpenDepositInquiryFileForRow(fields);
        };
        int lastCardClickCount = 0;
        card.MouseLeftButtonDown += delegate(object s, MouseButtonEventArgs e)
        {
            lastCardClickCount = e.ClickCount;
            if (e.ClickCount >= 2) selectAndOpen();
        };
        card.MouseLeftButtonUp += delegate
        {
            if (lastCardClickCount >= 2) return;   // ダブルクリック時は展開を二重に切り替えない
            selectAndToggle();
            if (depositTimelinePanel.IsKeyboardFocusWithin) card.Focus();   // キーで辿っている途中なら枠も移す
        };
        card.Focusable = true;
        card.FocusVisualStyle = keyFocusStyle;
        card.KeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.OriginalSource != card || Keyboard.Modifiers != ModifierKeys.None) return;
            if (e.Key == Key.Space) selectAndToggle();
            else if (e.Key == Key.Enter) selectAndOpen();
            else if (e.Key == Key.Up || e.Key == Key.Down)
            {
                // 端のカードでは何もしない（パネルのスクロールに任せる）
                if (!MoveDepositPanelCardFocus(card, e.Key == Key.Down ? 1 : -1)) return;
            }
            else return;
            e.Handled = true;
        };
        return card;
    }

    // 描画順（新しい回次が上）で隣のカードへフォーカスと選択を移す。移れなければ false
    private bool MoveDepositPanelCardFocus(Border from, int delta)
    {
        int index = depositPanelCards.FindIndex(delegate(KeyValuePair<string[], Border> pair) { return ReferenceEquals(pair.Value, from); });
        int next = index + delta;
        if (index < 0 || next < 0 || next >= depositPanelCards.Count) return false;
        SelectDepositCard(depositPanelCards[next].Key, depositPanelCards[next].Value);
        depositPendingExpandFields = null;   // 職員が自分で選んだカードが優先（クリックと同じ）
        depositPanelCards[next].Value.Focus();
        return true;
    }

    // 選択中のカード（なければ先頭のカードを選んで）へフォーカスを移す
    private void FocusDepositPanelCard()
    {
        if (depositSelectedCardBorder == null && depositPanelCards.Count > 0) SelectDepositCard(depositPanelCards[0].Key, depositPanelCards[0].Value);
        if (depositSelectedCardBorder != null) depositSelectedCardBorder.Focus();
    }

    // 明細セクションの対象か: 電子回答の口座あり行で、照会結果フォルダが設定されていること
    //（紙照会行は明細を持たない・口座なし行は回答シート自体がない）
    private bool IsDetailEligible(string[] fields)
    {
        return config.DepositSearchFolders != null &&
            fields[BusinessLogic.DEPOSIT_COL_HAS_ACCOUNT] != "0" && !IsPaperRow(fields);
    }

    // 展開中のカードを閉じる
    private void CollapseDepositCard()
    {
        if (depositExpandedCard == null) return;
        depositExpandedCard.Detail.Visibility = Visibility.Collapsed;
        depositExpandedCard.Chevron.Text = "▾";
        depositExpandedCard = null;
        depositExpandedCardFields = null;
    }

    // カードを展開する（開いていた別のカードは閉じる）。明細セクションを持つカードは
    // 初回の展開で照会ファイルの読み取りを要求する。scroll=true は展開部が見える位置まで送る
    private void ExpandDepositCard(DepositCardExpansion expansion, bool scroll)
    {
        if (!ReferenceEquals(depositExpandedCard, expansion)) CollapseDepositCard();
        depositExpandedCard = expansion;
        depositExpandedCardFields = expansion.Fields;
        expansion.Detail.Visibility = Visibility.Visible;
        expansion.Chevron.Text = "▴";
        if (expansion.Section != null && !expansion.Requested)
        {
            expansion.Requested = true;
            RequestDetail(expansion);
        }
        if (scroll)
        {
            // 展開部の高さが確定してから送る（Visible にした直後は 0 のまま）
            Dispatcher.BeginInvoke(new Action(delegate { expansion.Card.BringIntoView(); }), DispatcherPriority.Loaded);
        }
    }

    // パネル再構築後の展開状態の復元: 選択と同時に自動展開する行（一覧の行クリック・人物切替）
    // があればそのカードを開いて見える位置へ送り、なければ直前まで開いていたカードを開き直す
    private void RestoreDepositCardExpansion()
    {
        // カードが作り直されていれば、開いていたカードの参照はもう画面にない
        if (depositExpandedCard != null && !depositCardExpansions.ContainsKey(depositExpandedCard.Card)) depositExpandedCard = null;

        string[] foundFields;
        Border card = depositPendingExpandFields != null ? FindDepositPanelCard(depositPendingExpandFields, out foundFields) : null;
        if (card != null)
        {
            depositPendingExpandFields = null;
            DepositCardExpansion expansion;
            if (depositCardExpansions.TryGetValue(card, out expansion)) ExpandDepositCard(expansion, true);
            else CollapseDepositCard();   // 展開部のないカード（検出なし・照会結果フォルダ未設定など）
            return;
        }
        if (depositExpandedCardFields == null) return;
        card = FindDepositPanelCard(depositExpandedCardFields, out foundFields);
        DepositCardExpansion previous;
        if (card != null && depositCardExpansions.TryGetValue(card, out previous)) ExpandDepositCard(previous, false);
    }

    // 指定の行に対応するカードを選択状態にする（対応するカードがなければ選択なし）
    // 口座なし行はカードを持たないため、自然に「選択なし=開くボタン無効」へ落ちる
    private void ApplyDepositCardSelection(string[] fields)
    {
        string[] foundFields;
        var found = FindDepositPanelCard(fields, out foundFields);
        SelectDepositCard(foundFields, found);
        RestoreDepositCardExpansion();
    }

    // 行に対応する描画済みカードを探す（なければ null）。参照が一致しない場合は行キーで
    // 照合する（最近見た人の履歴は保存時のスナップショット配列を持つため、台帳から
    // 作られたカードと参照が異なる）
    private Border FindDepositPanelCard(string[] fields, out string[] foundFields)
    {
        foundFields = null;
        if (fields == null) return null;
        foreach (var pair in depositPanelCards)
        {
            if (ReferenceEquals(pair.Key, fields))
            {
                foundFields = pair.Key;
                return pair.Value;
            }
        }
        var key = BusinessLogic.GetDepositLedgerRowKey(fields, eraMapping);
        foreach (var pair in depositPanelCards)
        {
            if (BusinessLogic.GetDepositLedgerRowKey(pair.Key, eraMapping) == key)
            {
                foundFields = pair.Key;
                return pair.Value;
            }
        }
        return null;
    }

    // カードの選択状態を切り替える（選択中カード=「照会ファイルを開く」の対象）
    private void SelectDepositCard(string[] fields, Border card)
    {
        if (depositSelectedCardBorder != null)
        {
            // ローカル値を消してスタイル（ホバー含む）の表示へ戻す
            depositSelectedCardBorder.ClearValue(Border.BackgroundProperty);
            depositSelectedCardBorder.ClearValue(Border.BorderBrushProperty);
        }
        depositSelectedCardFields = fields;
        depositSelectedCardBorder = card;
        if (card != null)
        {
            card.Background = BrushSelectedCardBg;
            card.BorderBrush = BrushAccent;
        }
        UpdateDepositOpenFileButton();
        UpdateDepositPersonHeader();
    }

    // 人物ヘッダー（氏名・カナ・届出住所）を現在のカード選択に合わせて更新する。
    // 同一人物でも回次により届出住所が異なることがあり、差押リスト連携時に別口座の
    // 届出住所のまま取り違えないように、選択カードの行の値を優先して表示する。
    // 選択がない・選択行の項目が空のときは退行値（登録日が最新の非空値）で補う
    //（表示から情報を隠さない）
    private void UpdateDepositPersonHeader()
    {
        var name = depositPanelFallbackName;
        var kana = depositPanelFallbackKana;
        var address = depositPanelFallbackAddress;
        if (depositSelectedCardFields != null)
        {
            if (depositSelectedCardFields[BusinessLogic.DEPOSIT_COL_NAME] != "") name = depositSelectedCardFields[BusinessLogic.DEPOSIT_COL_NAME];
            if (depositSelectedCardFields[BusinessLogic.DEPOSIT_COL_KANA] != "") kana = depositSelectedCardFields[BusinessLogic.DEPOSIT_COL_KANA];
            if (depositSelectedCardFields[BusinessLogic.DEPOSIT_COL_ADDRESS] != "") address = depositSelectedCardFields[BusinessLogic.DEPOSIT_COL_ADDRESS];
        }
        depositPanelName.Text = name != null ? name : NO_NAME_TEXT;
        depositPanelAvatarText.Text = string.IsNullOrEmpty(name) ? "？" : StringInfo.GetNextTextElement(name);   // 頭文字（サロゲートペア・結合文字も 1 文字として）
        depositPanelKana.Text = kana != null ? kana : "";
        depositPanelAddress.Text = address != null ? address : "";
        depositPanelAddress.Visibility = address != null ? Visibility.Visible : Visibility.Collapsed;
    }

    // フッター「照会ファイルを開く」の有効/無効とツールチップを現在の状態に合わせる
    private void UpdateDepositOpenFileButton()
    {
        if (config.DepositSearchFolders == null)
        {
            depositPanelOpenFileButton.IsEnabled = false;
            depositPanelOpenFileButton.ToolTip = "config の depositSearchFolders を設定すると有効になります";
        }
        else if (depositSelectedCardBorder == null)
        {
            depositPanelOpenFileButton.IsEnabled = false;
            depositPanelOpenFileButton.ToolTip = "照会シートのある口座行を選択すると開けます";
        }
        else
        {
            depositPanelOpenFileButton.IsEnabled = true;
            depositPanelOpenFileButton.ToolTip = null;
        }
        UpdateDepositSeizureButton();   // 「差押リストに追加」も同じ選択状態に追従させる
    }

    // フッター「差押リストに追加」の有効/無効とツールチップを現在の状態に合わせる
    // 有効条件は「照会ファイルを開く」と同一（ファイル解決を共有するため）。
    // 連携先未設定時はボタン自体が非表示（ApplySeizureIntegrationVisibility）
    private void UpdateDepositSeizureButton()
    {
        if (depositSeizureScriptPath == null) return;
        if (config.DepositSearchFolders == null)
        {
            depositPanelSeizureButton.IsEnabled = false;
            depositPanelSeizureButton.ToolTip = "config の depositSearchFolders を設定すると有効になります";
        }
        else if (depositSelectedCardBorder == null)
        {
            depositPanelSeizureButton.IsEnabled = false;
            depositPanelSeizureButton.ToolTip = "照会シートのある口座行を選択すると追加できます";
        }
        else if (!IsSeizureTargetKey(BusinessLogic.NormalizeForSearch(
            depositSelectedCardFields[BusinessLogic.DEPOSIT_COL_INSTITUTION]), depositSeizureTargetPatterns, depositSeizureTargetMatchCache))
        {
            // 対象金融機関の制御: 無効にするだけでツールチップは出さない
            depositPanelSeizureButton.IsEnabled = false;
            depositPanelSeizureButton.ToolTip = null;
        }
        else
        {
            depositPanelSeizureButton.IsEnabled = true;
            depositPanelSeizureButton.ToolTip = null;
        }
    }

    // depositSeizureListScript／insuranceSeizureListScript が未設定または不存在の場合、その種別の
    // パネルの連携ボタンを非表示にする（グレースフルデグラデーション。depositSearchFolders の
    // 「無効+ツールチップ」とは扱いが異なり、連携を使わない環境では機能ごと見せない）。
    // ナビ「差押登録」は預金・生命保険のどちらの連携先もないときだけ非表示にする
    private void ApplySeizureIntegrationVisibility()
    {
        if (depositSeizureScriptPath == null)
        {
            depositPanelSeizureButton.Visibility = Visibility.Collapsed;
            Grid.SetColumnSpan(depositPanelOpenFileButton, 3);   // 「照会ファイルを開く」をフッター全幅に広げる
        }
        insuranceTab.ApplyInsuranceSeizureIntegrationVisibility();
        if (depositSeizureScriptPath != null || insuranceTab.insuranceSeizureScriptPath != null) return;
        navSeizure.Visibility = Visibility.Collapsed;
        railSeizure.Visibility = Visibility.Collapsed;
    }

    // 口座行の展開部の行（反対債権の種目別チップ+金額、取引カテゴリのチップ+原文、
    // 最終取引日の行）を組み立てる。表示する内容がなければ空リストを返す。
    // includeLastTxLine=false は最終取引日を明細セクションの ⧉ の行に置く口座
    private List<UIElement> BuildDepositAccountDetailLines(string[] fields, bool includeLastTxLine, out StackPanel categoryLines)
    {
        var lines = new List<UIElement>();
        categoryLines = null;

        // 反対債権（値ありの種目のみ・金額は Consolas）
        var counterClaims = new[]
        {
            new { Name = "証書貸付", Value = fields[BusinessLogic.DEPOSIT_COL_LOAN_DEED] },
            new { Name = "手形貸付", Value = fields[BusinessLogic.DEPOSIT_COL_LOAN_BILL] },
            new { Name = "割引手形", Value = fields[BusinessLogic.DEPOSIT_COL_LOAN_DISCOUNT] },
            new { Name = "当座貸越", Value = fields[BusinessLogic.DEPOSIT_COL_OVERDRAFT] }
        };
        foreach (var claim in counterClaims)
        {
            if (!BusinessLogic.HasCounterClaimValue(claim.Value)) continue;
            lines.Add(BuildDepositPanelDetectionLine(claim.Name, BrushChipCoralBg, BrushChipCoralFg, claim.Value.Trim(), true));
        }

        // 取引カテゴリ（値ありのみ・原文をそのまま全掲する。配色は一覧のチップと同一）
        var categories = new[]
        {
            new { Name = "年金",     Value = fields[BusinessLogic.DEPOSIT_COL_PENSION],      Bg = BrushChipTealBg,    Fg = BrushChipTealFg },
            new { Name = "給与",     Value = fields[BusinessLogic.DEPOSIT_COL_SALARY],       Bg = BrushChipVioletBg,  Fg = BrushChipVioletFg },
            new { Name = "保険",     Value = fields[BusinessLogic.DEPOSIT_COL_INSURANCE],    Bg = BrushChipAmberBg,   Fg = BrushChipAmberFg },
            new { Name = "定期入金", Value = fields[BusinessLogic.DEPOSIT_COL_REG_DEPOSIT],  Bg = BrushChipNeutralBg, Fg = BrushChipNeutralFg },
            new { Name = "定期出金", Value = fields[BusinessLogic.DEPOSIT_COL_REG_WITHDRAW], Bg = BrushChipNeutralBg, Fg = BrushChipNeutralFg }
        };
        // 5カテゴリは1つにまとめる（明細を読めた口座では集約表に置き換えて非表示にする）
        foreach (var category in categories)
        {
            if (string.IsNullOrWhiteSpace(category.Value)) continue;
            if (categoryLines == null)
            {
                categoryLines = new StackPanel();
                lines.Add(categoryLines);
            }
            categoryLines.Children.Add(BuildDepositPanelDetectionLine(category.Name, category.Bg, category.Fg, category.Value.Trim(), false));
        }

        // 当該回次時点の最終取引日
        if (includeLastTxLine && fields[BusinessLogic.DEPOSIT_COL_LAST_TX] != "")
        {
            var meta = MakeDepositLastTxText(fields[BusinessLogic.DEPOSIT_COL_LAST_TX]);
            meta.Margin = new Thickness(0, 5, 0, 0);
            lines.Add(meta);
        }
        return lines;
    }

    // 「最終取引日 yyyy/MM/dd」の文字（展開部の独立した行と明細セクションの ⧉ の行で共用）
    private static TextBlock MakeDepositLastTxText(string lastTx)
    {
        var text = new TextBlock { FontSize = 10.5, Foreground = BrushTextMuted };
        text.Inlines.Add(new Run("最終取引日 "));
        text.Inlines.Add(new Run(lastTx) { FontFamily = FontMono });
        return text;
    }

    // 展開部の1行（チップ+内容）を組み立てる。金額行のみ内容を Consolas で表示する
    private UIElement BuildDepositPanelDetectionLine(string chipText, Brush chipBg, Brush chipFg, string content, bool mono)
    {
        var grid = new Grid { Margin = new Thickness(0, 3, 0, 0) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        // 余白は検出列の暗色ツールチップと同じ比率（チップ上1/下2・本文上1）で、本文がチップより下がって見えないようにする。
        // チップの上下を非対称にするのは、Meiryo UI の字面が行ボックスの下寄りで、対称だと文字が下がって見えるため
        var chip = new Border
        {
            Background = chipBg,
            CornerRadius = new CornerRadius(9),
            Padding = new Thickness(8, 1, 8, 2),
            VerticalAlignment = VerticalAlignment.Top
        };
        chip.Child = new TextBlock { Text = chipText, FontSize = 10.5, Foreground = chipFg };
        grid.Children.Add(chip);

        var text = new TextBlock
        {
            Text = content,
            FontSize = 11,
            Foreground = BrushTextMain,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(8, 1, 0, 0)
        };
        if (mono) text.FontFamily = FontMono;
        Grid.SetColumn(text, 1);
        grid.Children.Add(text);
        return grid;
    }

    // クリックイベントの発生元から ListViewItem のデータ行を取得する（余白・ヘッダーは null）
    private static DepositLedgerRecord GetDepositRecordFromSource(object source)
    {
        var element = source as DependencyObject;
        while (element != null && !(element is ListViewItem))
        {
            // Run などの非 Visual 要素は VisualTreeHelper で辿れないため論理ツリーへ切り替える
            if (element is Visual || element is System.Windows.Media.Media3D.Visual3D)
                element = VisualTreeHelper.GetParent(element);
            else
                element = LogicalTreeHelper.GetParent(element);
        }
        var item = element as ListViewItem;
        return item == null ? null : item.DataContext as DepositLedgerRecord;
    }
}