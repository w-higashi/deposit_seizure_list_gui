// ==============================================================================
// paddock_deposit_search.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 預金タブ: 検索と絞り込み・詳細検索（条件の読み取り・条件チップ・担当のトグル）。
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
    // 検索・フィルタ
    // ==============================================================

    // 現在の検索条件・詳細検索の条件を全行に適用して一覧を更新する。
    // 比較は読み込み時に事前計算した正規化キー・数値・日付に対して行う
    private void ApplyDepositFilter()
    {
        // 検索文字列も行側と同じ正規化を通す（半角/全角・空白の揺れを両側で片寄せする）
        var searchText = BusinessLogic.NormalizeForSearch(depositSearchBox.Text);
        var keyIndex = depositSearchKeyCombo.SelectedIndex;   // 0=宛名番号 1=氏名 2=カナ氏名 3=金融機関名
        var f = depositAdvFilter;

        var filtered = new List<DepositLedgerRecord>(depositAllRecords.Count);
        foreach (var record in depositAllRecords)
        {
            // 口座なし行は「行の分類」の表示スイッチが入口（既定OFF=非表示）
            if (record.IsNoAccount && !f.ShowNoAccount) continue;
            if (MatchesDepositConditions(record, f, searchText, keyIndex)) filtered.Add(record);
        }

        // 0 件のときは、表示スイッチを入れれば並ぶ口座なし行を同じ条件で数える（照会した事実の案内に
        // 使う。口座なし行は上の走査で先に除いているため、数え直すのはこのときだけ）
        int hiddenNoAccount = 0;
        if (filtered.Count == 0 && !f.ShowNoAccount)
        {
            foreach (var record in depositAllRecords)
                if (record.IsNoAccount && MatchesDepositConditions(record, f, searchText, keyIndex)) hiddenNoAccount++;
        }

        depositLedgerList.ItemsSource = filtered;

        // フッター: 表示件数 / 全件数（詳細検索の条件数・読み飛ばしがあれば併記）
        var notes = new List<string>();
        if (depositAdvConditionCount > 0)
            notes.Add("詳細検索 " + depositAdvConditionCount + " 条件");
        if (depositSkippedRowCount > 0)
            notes.Add("列数不正 " + depositSkippedRowCount.ToString("N0") + " 行を読み飛ばし");
        if (depositLoadFailedYears.Count > 0)
            notes.Add(FormatYearLabels(depositLoadFailedYears) + " の台帳を読み込めませんでした");
        var status = "表示 " + filtered.Count.ToString("N0") + " 行 / 全 " + depositAllRecords.Count.ToString("N0") + " 行";
        if (notes.Count > 0) status += "（" + string.Join("・", notes.ToArray()) + "）";
        SetStatusText(depositStatusLeft, status);

        // 空状態: 並ぶ行があれば消す（読み込みが終わる前の絞り込みで出たものも、行が入ったら消す）。
        // 並ぶ行がないときは、台帳に行があれば「該当する行はありません」と口座なし行の案内、台帳が 0 行なら
        //「台帳に行がありません」。検出中・読み込み中は新たに出さない（検出と読み込みの前の 0 行で出さないため。
        // 読み込みの完了時の絞り込みで出る）。台帳が見つからない・読み込めなかったときの文言は出したままにする
        if (filtered.Count > 0)
            depositEmptyStateText.Visibility = Visibility.Collapsed;
        else if (!(depositDetectWorker != null && depositDetectWorker.IsBusy) && !(depositLoadWorker != null && depositLoadWorker.IsBusy))
        {
            if (depositAllRecords.Count > 0)
                ShowDepositFilterEmptyState(hiddenNoAccount);
            else if (depositEmptyStateText.Visibility != Visibility.Visible)
                ShowEmptyState(depositEmptyStateText, "台帳に行がありません。");
        }
    }

    // 行 1 件が検索条件・詳細検索の条件（クイック検索・項目検索・残高・日付・検出・行の分類・担当）に
    // 合うか。口座なし行の表示スイッチはここでは見ない（走査する側が先に除く。0 件のときの
    // 口座なし行の数え直しでも同じ条件で判定するため、ここに置く）
    private static bool MatchesDepositConditions(DepositLedgerRecord record, DepositAdvancedFilter f, string searchText, int keyIndex)
    {
        bool classSelected = f.ClassRepresentative || f.ClassSpecialOnly;

        // クイック検索（選択中のキーで部分一致）
        if (searchText.Length > 0)
        {
            string target;
            switch (keyIndex)
            {
                case 1:  target = record.SearchName; break;
                case 2:  target = record.SearchKana; break;
                case 3:  target = record.SearchInstitution; break;
                default: target = record.SearchAddressNumber; break;
            }
            if (target.IndexOf(searchText, StringComparison.Ordinal) < 0) return false;
        }

        // 項目検索4欄（それぞれ部分一致・すべて AND）
        if (f.AddressNumber.Length > 0 && record.SearchAddressNumber.IndexOf(f.AddressNumber, StringComparison.Ordinal) < 0) return false;
        if (f.Name.Length > 0          && record.SearchName.IndexOf(f.Name, StringComparison.Ordinal) < 0) return false;
        if (f.Kana.Length > 0          && record.SearchKana.IndexOf(f.Kana, StringComparison.Ordinal) < 0) return false;
        if (f.Institution.Length > 0   && record.SearchInstitution.IndexOf(f.Institution, StringComparison.Ordinal) < 0) return false;

        // 残高範囲（値を持たない行=センチネル・非数・口座なしは条件設定時に対象外）
        if (f.HasBalance)
        {
            if (record.BalanceValue == null) return false;
            if (f.BalanceMin != null && record.BalanceValue.Value < f.BalanceMin.Value) return false;
            if (f.BalanceMax != null && record.BalanceValue.Value > f.BalanceMax.Value) return false;
        }

        // 日付範囲（パース不能・空欄の行は条件設定時に対象外）
        if (f.HasBaseDate)
        {
            if (record.BaseDateValue == null) return false;
            if (f.BaseDateFrom != null && record.BaseDateValue.Value < f.BaseDateFrom.Value) return false;
            if (f.BaseDateTo != null && record.BaseDateValue.Value > f.BaseDateTo.Value) return false;
        }
        if (f.HasLastTx)
        {
            if (record.LastTxValue == null) return false;
            if (f.LastTxFrom != null && record.LastTxValue.Value < f.LastTxFrom.Value) return false;
            if (f.LastTxTo != null && record.LastTxValue.Value > f.LastTxTo.Value) return false;
        }

        // 検出（選択したものをすべて含む行。口座なし行は検出を持たないため設定時に落ちる）
        if (f.HasDetection)
        {
            bool lacks = false;
            for (int i = 0; i < f.Detections.Length; i++)
            {
                if (f.Detections[i] && !record.DetectionFlags[i]) { lacks = true; break; }
            }
            if (lacks) return false;
        }

        // 行の分類（代表口座/特殊口座のみ の OR。口座あり行のみ判定し、口座なし行は素通り）
        if (classSelected && !record.IsNoAccount)
        {
            bool match = (f.ClassRepresentative && record.Classification == "1")
                      || (f.ClassSpecialOnly && record.Classification == "2");
            if (!match) return false;
        }

        // 担当（選択したもののいずれかに該当。表記ゆれ統合後のクラスタ単位で照合する。
        // 担当が空欄の行=口座なし行・旧27列の行は条件設定時に対象外）
        if (f.HasStaff && (record.StaffKey == null || !f.StaffKeys.Contains(record.StaffKey))) return false;

        return true;
    }

    // 絞り込みが 0 件のときの空状態。行の分類「口座なし」を ON にすれば並ぶ口座なし行があれば、
    // その件数の文をリンク（押すと ON になる）で添える。ほかの空状態と同じ TextBlock に Inline で組む
    //（ほかの文言は Text で入れるため、そのときこの Inline は消える）
    private void ShowDepositFilterEmptyState(int hiddenNoAccount)
    {
        var inlines = depositEmptyStateText.Inlines;
        inlines.Clear();
        inlines.Add(new Run("該当する行はありません。"));
        if (hiddenNoAccount > 0)
        {
            inlines.Add(new LineBreak());
            inlines.Add(new InlineUIContainer(MakeDepositNoAccountLink(hiddenNoAccount)) { BaselineAlignment = BaselineAlignment.Center });
        }
        depositEmptyStateText.Visibility = Visibility.Visible;
    }

    // 「口座なしの行が N 件あります ›」のリンク（条件チップと同じ形の Border。クリック・Enter・Space で押せる）
    private Border MakeDepositNoAccountLink(int count)
    {
        var text = new TextBlock
        {
            Text = "口座なしの行が " + count.ToString("N0") + " 件あります ›",
            FontSize = 12,
            Foreground = BrushAccent
        };
        var link = new Border
        {
            Background = Brushes.Transparent,
            CornerRadius = new CornerRadius(9),
            Padding = new Thickness(8, 1, 8, 2),
            Cursor = Cursors.Hand,
            ToolTip = "クリックで口座なしの行を表示",
            Child = text
        };
        link.MouseEnter += delegate { text.TextDecorations = TextDecorations.Underline; };
        link.MouseLeave += delegate { text.TextDecorations = null; };
        link.MouseLeftButtonUp += delegate { ShowDepositNoAccountRows(); };
        MakeKeyActivatable(link, ShowDepositNoAccountRows, keyFocusPillStyle);
        return link;
    }

    // 行の分類「口座なし」を ON にして口座なし行を並べ、一覧の先頭行へ移る（検索欄の ↓ と同じ）。
    // 絞り込み・条件チップ・詳細検索ボタンの件数は Checked の既存の経路で更新される
    private void ShowDepositNoAccountRows()
    {
        depositAdvClassNoAccount.IsChecked = true;
        if (depositLedgerList.Items.Count == 0) return;
        depositLedgerList.SelectedIndex = 0;
        depositLedgerList.ScrollIntoView(depositLedgerList.Items[0]);
        FocusLater(delegate { FocusDepositLedgerSelection(); });
    }

    // ==============================================================
    // 詳細検索
    // ==============================================================
    //
    // 検索バー右端の「詳細検索」トグルで条件パネルを開閉する。条件は入力と同時に
    // 一覧へ反映され、フィルタチップ行に着脱可能な条件チップとして現れる
    //（パネルを閉じても条件は持続する。チップのクリック=そのグループの解除）。
    // グループ間は AND・検出は「すべて含む」・行の分類は OR。
    // 「口座なし」トグルだけは分類の OR ではなく口座なし行の表示スイッチ（既定OFF）。

    // 詳細検索の条件一式。テキストは正規化済み・数値/日付はパース済みで保持し、
    // ApplyDepositFilter は比較だけを行う
    private class DepositAdvancedFilter
    {
        public string AddressNumber = "";             // 項目検索（正規化済み。空=条件なし）
        public string Name = "";
        public string Kana = "";
        public string Institution = "";
        public long? BalanceMin;                      // 残高範囲（null=その側の指定なし）
        public long? BalanceMax;
        public DateTime? BaseDateFrom;                // 回答基準日の範囲
        public DateTime? BaseDateTo;
        public DateTime? LastTxFrom;                  // 最終取引日の範囲
        public DateTime? LastTxTo;
        public bool[] Detections = new bool[6];       // 検出（DETECTION_NAMES と同順・選択=すべて含む）
        public bool ClassRepresentative;              // 行の分類「代表口座」（登録推奨=1）
        public bool ClassSpecialOnly;                 // 行の分類「特殊口座のみ」（登録推奨=2）
        public bool ShowNoAccount;                    // 行の分類「口座なし」= 口座なし行の表示スイッチ
        public HashSet<string> StaffKeys = new HashSet<string>();   // 担当（選択クラスタのメンバーキーの合併。空=条件なし）
        public List<string> StaffNames = new List<string>();        // 〃 チップ表示用の代表名（トグルの並び順）

        public bool HasBalance   { get { return BalanceMin != null || BalanceMax != null; } }
        public bool HasBaseDate  { get { return BaseDateFrom != null || BaseDateTo != null; } }
        public bool HasLastTx    { get { return LastTxFrom != null || LastTxTo != null; } }
        public bool HasDetection { get { return Detections.Any(d => d); } }
        public bool HasClass     { get { return ClassRepresentative || ClassSpecialOnly || ShowNoAccount; } }
        public bool HasStaff     { get { return StaffKeys.Count > 0; } }
    }

    // 文字入力による絞り込みの再実行を、入力が止まってから 250ms 後の1回にまとめる
    //（打鍵のたびに一覧を組み直さないための性能対策）。トグルの切替・
    // 条件チップの✕・「条件をすべてクリア」は OnDepositAdvancedFilterChanged 直呼びの即時反映
    private void RestartDepositFilterDebounce()
    {
        if (suppressDepositAdvChange) return;
        if (depositFilterDebounceTimer == null)
        {
            depositFilterDebounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
            depositFilterDebounceTimer.Tick += delegate { OnDepositAdvancedFilterChanged(); };
        }
        depositFilterDebounceTimer.Stop();
        depositFilterDebounceTimer.Start();
    }

    // 詳細検索の入力・選択の変化を一覧へ反映する（読み取り→チップ再構築→トグル表示→絞り込み）
    private void OnDepositAdvancedFilterChanged()
    {
        if (suppressDepositAdvChange) return;
        if (depositFilterDebounceTimer != null) depositFilterDebounceTimer.Stop();   // 保留中の遅延反映はこの実行が兼ねる

        // 条件が変わると一覧が入れ替わるため、開いている詳細パネルは閉じる
        //（年度切替・ビュー切替と同じ扱い。入力中の欄からフォーカスは奪わない）
        if (isDepositPanelOpen) CloseDepositDetailPanel(false);
        ReadDepositAdvancedInputs();
        RebuildDepositFilterChips();
        UpdateDepositAdvToggleVisual();
        ApplyDepositFilter();
    }

    // パネルの入力値を条件セットへ読み取る
    private void ReadDepositAdvancedInputs()
    {
        var f = depositAdvFilter;
        f.AddressNumber = BusinessLogic.NormalizeForSearch(depositAdvAddressBox.Text);
        f.Name          = BusinessLogic.NormalizeForSearch(depositAdvNameBox.Text);
        f.Kana          = BusinessLogic.NormalizeForSearch(depositAdvKanaBox.Text);
        f.Institution   = BusinessLogic.NormalizeForSearch(depositAdvInstitutionBox.Text);

        f.BalanceMin = ReadAmountInput(depositAdvBalanceMinBox, depositAdvBalanceMinBd);
        f.BalanceMax = ReadAmountInput(depositAdvBalanceMaxBox, depositAdvBalanceMaxBd);
        f.BaseDateFrom = ReadDateInput(depositAdvBaseFromBox, depositAdvBaseFromBd);
        f.BaseDateTo   = ReadDateInput(depositAdvBaseToBox, depositAdvBaseToBd);
        f.LastTxFrom   = ReadDateInput(depositAdvTxFromBox, depositAdvTxFromBd);
        f.LastTxTo     = ReadDateInput(depositAdvTxToBox, depositAdvTxToBd);

        for (int i = 0; i < depositAdvDetectChips.Length; i++)
            f.Detections[i] = depositAdvDetectChips[i].IsChecked == true;
        f.ClassRepresentative = depositAdvClassRep.IsChecked == true;
        f.ClassSpecialOnly    = depositAdvClassSpecial.IsChecked == true;
        f.ShowNoAccount       = depositAdvClassNoAccount.IsChecked == true;

        f.StaffKeys.Clear();
        f.StaffNames.Clear();
        for (int i = 0; i < depositAdvStaffToggles.Count; i++)
        {
            if (depositAdvStaffToggles[i].IsChecked != true) continue;
            f.StaffNames.Add(depositStaffClusters[i].Representative);
            foreach (var key in depositStaffClusters[i].MemberKeys) f.StaffKeys.Add(key);
        }
    }

    // 条件チップ行を現在の条件から組み直す。チップの数はバッジとフッター表記が共有する
    private void RebuildDepositFilterChips()
    {
        depositFilterChipsRow.Children.Clear();

        AddDepositTextConditionChip("宛名番号", depositAdvAddressBox);
        AddDepositTextConditionChip("氏名", depositAdvNameBox);
        AddDepositTextConditionChip("カナ氏名", depositAdvKanaBox);
        AddDepositTextConditionChip("金融機関名", depositAdvInstitutionBox);

        var f = depositAdvFilter;
        if (f.HasBalance)
        {
            var text = "残高 " +
                (f.BalanceMin != null ? f.BalanceMin.Value.ToString("N0") + "円" : "") + "〜" +
                (f.BalanceMax != null ? f.BalanceMax.Value.ToString("N0") + "円" : "");
            AddConditionChip(depositFilterChipsRow, depositSearchBox, text, delegate
            {
                suppressDepositAdvChange = true;
                depositAdvBalanceMinBox.Text = "";
                depositAdvBalanceMaxBox.Text = "";
                suppressDepositAdvChange = false;
                OnDepositAdvancedFilterChanged();
            });
        }
        if (f.HasBaseDate)
        {
            AddConditionChip(depositFilterChipsRow, depositSearchBox, FormatDateRangeChip("回答基準日", f.BaseDateFrom, f.BaseDateTo), delegate
            {
                suppressDepositAdvChange = true;
                depositAdvBaseFromBox.Text = "";
                depositAdvBaseToBox.Text = "";
                suppressDepositAdvChange = false;
                OnDepositAdvancedFilterChanged();
            });
        }
        if (f.HasLastTx)
        {
            AddConditionChip(depositFilterChipsRow, depositSearchBox, FormatDateRangeChip("最終取引日", f.LastTxFrom, f.LastTxTo), delegate
            {
                suppressDepositAdvChange = true;
                depositAdvTxFromBox.Text = "";
                depositAdvTxToBox.Text = "";
                suppressDepositAdvChange = false;
                OnDepositAdvancedFilterChanged();
            });
        }
        if (f.HasDetection)
        {
            var names = new List<string>();
            for (int i = 0; i < f.Detections.Length; i++)
                if (f.Detections[i]) names.Add(BusinessLogic.DETECTION_NAMES[i]);
            AddConditionChip(depositFilterChipsRow, depositSearchBox, "検出: " + string.Join("・", names.ToArray()), delegate
            {
                suppressDepositAdvChange = true;
                foreach (var chip in depositAdvDetectChips) chip.IsChecked = false;
                suppressDepositAdvChange = false;
                OnDepositAdvancedFilterChanged();
            });
        }
        if (f.HasClass)
        {
            var names = new List<string>();
            if (f.ClassRepresentative) names.Add("代表口座");
            if (f.ClassSpecialOnly) names.Add("特殊口座のみ");
            if (f.ShowNoAccount) names.Add("口座なし");
            AddConditionChip(depositFilterChipsRow, depositSearchBox, "行の分類: " + string.Join("・", names.ToArray()), delegate
            {
                suppressDepositAdvChange = true;
                depositAdvClassRep.IsChecked = false;
                depositAdvClassSpecial.IsChecked = false;
                depositAdvClassNoAccount.IsChecked = false;
                suppressDepositAdvChange = false;
                OnDepositAdvancedFilterChanged();
            });
        }
        if (f.HasStaff)
        {
            AddConditionChip(depositFilterChipsRow, depositSearchBox, "担当: " + string.Join("・", f.StaffNames.ToArray()), delegate
            {
                suppressDepositAdvChange = true;
                foreach (var toggle in depositAdvStaffToggles) toggle.IsChecked = false;
                suppressDepositAdvChange = false;
                OnDepositAdvancedFilterChanged();
            });
        }

        depositAdvConditionCount = depositFilterChipsRow.Children.Count;
        depositFilterChipsRow.Visibility = depositAdvConditionCount > 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    // 項目検索欄1つ分の条件チップ（値があるときのみ。表示は入力の生値・解除で欄をクリア）
    private void AddDepositTextConditionChip(string label, TextBox box)
    {
        var value = box.Text.Trim();
        if (value == "") return;
        AddConditionChip(depositFilterChipsRow, depositSearchBox, label + ": " + value, delegate
        {
            suppressDepositAdvChange = true;
            box.Text = "";
            suppressDepositAdvChange = false;
            OnDepositAdvancedFilterChanged();
        });
    }

    // 詳細検索パネルを開閉する（高さアニメーション・条件は閉じても持続する）
    private void ToggleDepositAdvancedPanel()
    {
        depositAdvPanelOpen = !depositAdvPanelOpen;
        if (depositAdvPanelOpen) AnimateExpand(depositAdvPanel);
        else AnimateCollapse(depositAdvPanel);
        UpdateDepositAdvToggleVisual();
    }

    // 「詳細検索」トグルの見た目を更新する（条件数バッジと開閉キャレット。
    // トグル自体は常時 選択面+アイリス文字の固定表示）
    private void UpdateDepositAdvToggleVisual()
    {
        depositAdvCountBadge.Visibility = depositAdvConditionCount > 0 ? Visibility.Visible : Visibility.Collapsed;
        depositAdvCountText.Text = depositAdvConditionCount.ToString();
        depositAdvToggleCaret.Text = depositAdvPanelOpen ? "▲" : "▼";
    }

    // パネルの「条件をすべてクリア」（クイック検索は対象外）
    private void ClearDepositAdvancedFilter()
    {
        suppressDepositAdvChange = true;
        depositAdvAddressBox.Text = "";
        depositAdvNameBox.Text = "";
        depositAdvKanaBox.Text = "";
        depositAdvInstitutionBox.Text = "";
        depositAdvBalanceMinBox.Text = "";
        depositAdvBalanceMaxBox.Text = "";
        depositAdvBaseFromBox.Text = "";
        depositAdvBaseToBox.Text = "";
        depositAdvTxFromBox.Text = "";
        depositAdvTxToBox.Text = "";
        foreach (var chip in depositAdvDetectChips) chip.IsChecked = false;
        depositAdvClassRep.IsChecked = false;
        depositAdvClassSpecial.IsChecked = false;
        depositAdvClassNoAccount.IsChecked = false;
        foreach (var toggle in depositAdvStaffToggles) toggle.IsChecked = false;
        suppressDepositAdvChange = false;
        OnDepositAdvancedFilterChanged();
    }

    // 担当グループのトグル群を選択中年度の台帳から組み直す。
    // 年度切替・staffFuzzyChars の変更で呼ばれ、選択中の条件は同一クラスタ
    //（代表の正規化キーの一致）が残れば維持・なければ解除される
    private void RebuildDepositStaffToggles()
    {
        // 復元用に現在の選択を代表の正規化キーで控える
        var selectedKeys = new HashSet<string>();
        for (int i = 0; i < depositAdvStaffToggles.Count; i++)
        {
            if (depositAdvStaffToggles[i].IsChecked == true)
                selectedKeys.Add(BusinessLogic.NormalizeForSearch(depositStaffClusters[i].Representative));
        }

        depositStaffClusters = BusinessLogic.BuildStaffClusters(
            depositStaffRawValues, new HashSet<char>(config.StaffFuzzyChars));

        depositAdvStaffToggles.Clear();
        depositAdvStaffPanel.Children.Clear();
        suppressDepositAdvChange = true;   // IsChecked の復元で Checked イベントを発火させない
        foreach (var cluster in depositStaffClusters)
        {
            var toggle = new ToggleButton
            {
                Style = (Style)window.FindResource("AdvToggleChip"),
                Content = cluster.Representative,
                IsChecked = selectedKeys.Contains(BusinessLogic.NormalizeForSearch(cluster.Representative))
            };
            // 2値以上を統合したトグルは、吸収した生値を暗色ツールチップで全掲する
            if (cluster.RawValues.Count > 1)
            {
                var tip = new ToolTip { Style = (Style)window.FindResource("DetectTip") };
                tip.LayoutTransform = displayScaleTransform;   // 表示倍率に追従
                tip.Content = BuildStaffMergeTip(cluster);
                toggle.ToolTip = tip;
                ToolTipService.SetInitialShowDelay(toggle, 350);
                ToolTipService.SetShowDuration(toggle, 60000);
            }
            toggle.Checked += delegate { OnDepositAdvancedFilterChanged(); };
            toggle.Unchecked += delegate { OnDepositAdvancedFilterChanged(); };
            depositAdvStaffPanel.Children.Add(toggle);
            depositAdvStaffToggles.Add(toggle);
        }
        suppressDepositAdvChange = false;

        depositAdvStaffPanel.Visibility = depositStaffClusters.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        depositAdvStaffEmptyText.Visibility = depositStaffClusters.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
    }
}