// ==============================================================================
// paddock_insurance_search.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 生命保険タブ: 検索と絞り込み・詳細検索（条件の読み取り・条件チップ・担当のトグル）・
// 絞り込みが 0 件のときの案内。
// PaddockApp の入れ子のクラス InsuranceTab の一部（partial class）。ツールの説明とビルドの方法は
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
    private partial class InsuranceTab
    {
        // ==============================================================
        // 検索・フィルタ
        // ==============================================================

        // 現在の検索条件・詳細検索の条件を全行に適用して一覧を更新する。
        // 比較は読み込み時に事前計算した正規化キー・数値・日付に対して行う
        private void ApplyInsuranceFilter()
        {
            // 検索文字列も行側と同じ正規化を通す（半角/全角・空白の揺れを両側で片寄せする）
            var searchText = BusinessLogic.NormalizeForSearch(insuranceSearchBox.Text);
            var keyIndex = insuranceSearchKeyCombo.SelectedIndex;   // 0=宛名番号 1=氏名 2=カナ氏名 3=保険会社名 4=証券番号
            var f = insuranceAdvFilter;

            var filtered = new List<InsuranceLedgerRecord>(insuranceAllRecords.Count);
            foreach (var record in insuranceAllRecords)
            {
                // 契約なし行・電子回答不可行は「行の分類」の表示スイッチが入口（既定OFF=非表示）
                if (record.IsNoContract && !f.ShowNoContract) continue;
                if (record.IsUnanswered && !f.ShowUnanswered) continue;
                if (MatchesInsuranceConditions(record, f, searchText, keyIndex)) filtered.Add(record);
            }

            // 0 件のときは、表示スイッチを入れれば並ぶ契約なし行・電子回答不可行を同じ条件で分類ごとに数える
            //（照会した事実の案内に使う。これらの行は上の走査で先に除いているため、数え直すのはこのときだけ。
            // ON の分類は 1 度目の走査に含まれていて 0 件なので数えない）
            int hiddenNoContract = 0;
            int hiddenUnanswered = 0;
            if (filtered.Count == 0 && (!f.ShowNoContract || !f.ShowUnanswered))
            {
                foreach (var record in insuranceAllRecords)
                {
                    if (record.IsNoContract && !f.ShowNoContract && MatchesInsuranceConditions(record, f, searchText, keyIndex)) hiddenNoContract++;
                    else if (record.IsUnanswered && !f.ShowUnanswered && MatchesInsuranceConditions(record, f, searchText, keyIndex)) hiddenUnanswered++;
                }
            }

            insuranceLedgerList.ItemsSource = filtered;

            // フッター: 表示件数 / 全件数（詳細検索の条件数・読み飛ばしがあれば併記）
            var notes = new List<string>();
            if (insuranceAdvConditionCount > 0)
                notes.Add("詳細検索 " + insuranceAdvConditionCount + " 条件");
            if (insuranceSkippedRowCount > 0)
                notes.Add("列数不正 " + insuranceSkippedRowCount.ToString("N0") + " 行を読み飛ばし");
            if (insuranceLoadFailedYears.Count > 0)
                notes.Add(FormatYearLabels(insuranceLoadFailedYears) + " の台帳を読み込めませんでした");
            var status = "表示 " + filtered.Count.ToString("N0") + " 行 / 全 " + insuranceAllRecords.Count.ToString("N0") + " 行";
            if (notes.Count > 0) status += "（" + string.Join("・", notes.ToArray()) + "）";
            app.SetStatusText(insuranceStatusLeft, status);

            // 空状態: 並ぶ行があれば消す（読み込みが終わる前の絞り込みで出たものも、行が入ったら消す）。
            // 並ぶ行がないときは、台帳に行があれば「該当する行はありません」と隠れている行の案内、台帳が 0 行なら
            //「台帳に行がありません」。検出中・読み込み中は新たに出さない（検出と読み込みの前の 0 行で出さないため。
            // 読み込みの完了時の絞り込みで出る）。台帳が見つからない・読み込めなかったときの文言は出したままにする
            if (filtered.Count > 0)
                insuranceEmptyStateText.Visibility = Visibility.Collapsed;
            else if (!(insuranceDetectWorker != null && insuranceDetectWorker.IsBusy) && !(insuranceLoadWorker != null && insuranceLoadWorker.IsBusy))
            {
                if (insuranceAllRecords.Count > 0)
                    ShowInsuranceFilterEmptyState(hiddenNoContract, hiddenUnanswered);
                else if (insuranceEmptyStateText.Visibility != Visibility.Visible)
                    ShowEmptyState(insuranceEmptyStateText, "台帳に行がありません。");
            }
        }

        // 行 1 件が検索条件・詳細検索の条件（クイック検索・項目検索・解約返戻金・日付・注記・担当）に
        // 合うか。契約なし行・電子回答不可行の表示スイッチはここでは見ない（走査する側が先に除く。0 件のときの
        // 数え直しでも同じ条件で判定するため、ここに置く）
        private static bool MatchesInsuranceConditions(InsuranceLedgerRecord record, InsuranceAdvancedFilter f, string searchText, int keyIndex)
        {
            // クイック検索（選択中のキーで部分一致）
            if (searchText.Length > 0)
            {
                string target;
                switch (keyIndex)
                {
                    case 1:  target = record.SearchName; break;
                    case 2:  target = record.SearchKana; break;
                    case 3:  target = record.SearchCompany; break;
                    case 4:  target = record.SearchPolicyNumber; break;
                    default: target = record.SearchAddressNumber; break;
                }
                if (target.IndexOf(searchText, StringComparison.Ordinal) < 0) return false;
            }

            // 項目検索5欄（それぞれ部分一致・すべて AND）
            if (f.AddressNumber.Length > 0 && record.SearchAddressNumber.IndexOf(f.AddressNumber, StringComparison.Ordinal) < 0) return false;
            if (f.Name.Length > 0          && record.SearchName.IndexOf(f.Name, StringComparison.Ordinal) < 0) return false;
            if (f.Kana.Length > 0          && record.SearchKana.IndexOf(f.Kana, StringComparison.Ordinal) < 0) return false;
            if (f.Company.Length > 0       && record.SearchCompany.IndexOf(f.Company, StringComparison.Ordinal) < 0) return false;
            if (f.PolicyNumber.Length > 0  && record.SearchPolicyNumber.IndexOf(f.PolicyNumber, StringComparison.Ordinal) < 0) return false;

            // 解約返戻金の範囲（値を持たない行=読めない生値・空欄は条件設定時に対象外）
            if (f.HasSurrender)
            {
                if (record.SurrenderValue == null) return false;
                if (f.SurrenderMin != null && record.SurrenderValue.Value < f.SurrenderMin.Value) return false;
                if (f.SurrenderMax != null && record.SurrenderValue.Value > f.SurrenderMax.Value) return false;
            }

            // 日付範囲（パース不能・空欄の行は条件設定時に対象外）
            if (f.HasBaseDate)
            {
                if (record.BaseDateValue == null) return false;
                if (f.BaseDateFrom != null && record.BaseDateValue.Value < f.BaseDateFrom.Value) return false;
                if (f.BaseDateTo != null && record.BaseDateValue.Value > f.BaseDateTo.Value) return false;
            }
            if (f.HasMaturity)
            {
                if (record.MaturityValue == null) return false;
                if (f.MaturityFrom != null && record.MaturityValue.Value < f.MaturityFrom.Value) return false;
                if (f.MaturityTo != null && record.MaturityValue.Value > f.MaturityTo.Value) return false;
            }

            // 注記（選択したものをすべて含む行。契約なし行・電子回答不可行は注記を持たないため設定時に落ちる）
            if (f.HasNote)
            {
                bool lacks = false;
                for (int i = 0; i < f.Notes.Length; i++)
                {
                    if (f.Notes[i] && !record.NoteFlags[i]) { lacks = true; break; }
                }
                if (lacks) return false;
            }

            // 担当（選択したもののいずれかに該当。表記ゆれ統合後のクラスタ単位で照合する。
            // 担当が空欄の行は条件設定時に対象外）
            if (f.HasStaff && (record.StaffKey == null || !f.StaffKeys.Contains(record.StaffKey))) return false;

            return true;
        }

        // 絞り込みが 0 件のときの空状態。行の分類「契約なし」「電子回答不可」を ON にすれば並ぶ行があれば、
        // その件数の文をリンク（押すと ON になる）で添える。ほかの空状態と同じ TextBlock に Inline で組む
        //（ほかの文言は Text で入れるため、そのときこの Inline は消える）
        private void ShowInsuranceFilterEmptyState(int hiddenNoContract, int hiddenUnanswered)
        {
            var inlines = insuranceEmptyStateText.Inlines;
            inlines.Clear();
            inlines.Add(new Run("該当する行はありません。"));
            if (hiddenNoContract + hiddenUnanswered > 0)
            {
                inlines.Add(new LineBreak());
                inlines.Add(new InlineUIContainer(MakeInsuranceHiddenRowsLink(hiddenNoContract, hiddenUnanswered)) { BaselineAlignment = BaselineAlignment.Center });
            }
            insuranceEmptyStateText.Visibility = Visibility.Visible;
        }

        // 「契約なしの行が N 件あります ›」のリンク（条件チップと同じ形の Border。クリック・Enter・Space で押せる）。
        // 隠れている分類が片方なら「契約なし」「電子回答不可」、両方なら「契約なし・電子回答不可」の 1 行で
        // 件数は合計。押すと文言の分類のスイッチだけが ON になる
        private Border MakeInsuranceHiddenRowsLink(int hiddenNoContract, int hiddenUnanswered)
        {
            bool noContract = hiddenNoContract > 0;
            bool unanswered = hiddenUnanswered > 0;
            var label = noContract && unanswered ? "契約なし・電子回答不可" : noContract ? "契約なし" : "電子回答不可";
            var text = new TextBlock
            {
                Text = label + "の行が " + (hiddenNoContract + hiddenUnanswered).ToString("N0") + " 件あります ›",
                FontSize = 12,
                Foreground = BrushAccent
            };
            var link = new Border
            {
                Background = Brushes.Transparent,
                CornerRadius = new CornerRadius(9),
                Padding = new Thickness(8, 1, 8, 2),
                Cursor = Cursors.Hand,
                ToolTip = "クリックで" + label + "の行を表示",
                Child = text
            };
            link.MouseEnter += delegate { text.TextDecorations = TextDecorations.Underline; };
            link.MouseLeave += delegate { text.TextDecorations = null; };
            Action show = delegate { ShowInsuranceHiddenRows(noContract, unanswered); };
            link.MouseLeftButtonUp += delegate { show(); };
            app.MakeKeyActivatable(link, show, app.keyFocusPillStyle);
            return link;
        }

        // 行の分類のスイッチを ON にして隠れていた行を並べ、一覧の先頭行へ移る（検索欄の ↓ と同じ）。
        // 2 つ立てるときも絞り込みは 1 回だけ（変化の通知を止めて立ててから反映する）。
        // 条件チップ・詳細検索ボタンの件数はその反映で更新される
        private void ShowInsuranceHiddenRows(bool noContract, bool unanswered)
        {
            suppressInsuranceAdvChange = true;
            if (noContract) insuranceAdvClassNoContract.IsChecked = true;
            if (unanswered) insuranceAdvClassUnanswered.IsChecked = true;
            suppressInsuranceAdvChange = false;
            OnInsuranceAdvancedFilterChanged();
            if (insuranceLedgerList.Items.Count == 0) return;
            insuranceLedgerList.SelectedIndex = 0;
            insuranceLedgerList.ScrollIntoView(insuranceLedgerList.Items[0]);
            app.FocusLater(delegate { FocusInsuranceLedgerSelection(); });
        }

        // ==============================================================
        // 詳細検索
        // ==============================================================
        //
        // 検索バー右端の「詳細検索」トグルで条件パネルを開閉する。条件は入力と同時に
        // 一覧へ反映され、フィルタチップ行に着脱可能な条件チップとして現れる
        //（パネルを閉じても条件は持続する。チップのクリック=そのグループの解除）。
        // グループ間は AND・注記は「すべて含む」。
        // 行の分類の「契約なし」「電子回答不可」は契約なし行・電子回答不可行の表示スイッチ（既定OFF・互いに独立）。

        // 詳細検索の条件一式。テキストは正規化済み・数値/日付はパース済みで保持し、
        // ApplyInsuranceFilter は比較だけを行う
        private class InsuranceAdvancedFilter
        {
            public string AddressNumber = "";             // 項目検索（正規化済み。空=条件なし）
            public string Name = "";
            public string Kana = "";
            public string Company = "";
            public string PolicyNumber = "";
            public long? SurrenderMin;                    // 解約返戻金の範囲（null=その側の指定なし）
            public long? SurrenderMax;
            public DateTime? BaseDateFrom;                // 回答基準日の範囲
            public DateTime? BaseDateTo;
            public DateTime? MaturityFrom;                // 満期年月日の範囲
            public DateTime? MaturityTo;
            public bool[] Notes = new bool[BusinessLogic.INSURANCE_NOTE_NAMES.Length];   // 注記（INSURANCE_NOTE_NAMES と同順・選択=すべて含む）
            public bool ShowNoContract;                   // 行の分類「契約なし」= 契約なし行の表示スイッチ
            public bool ShowUnanswered;                   // 行の分類「電子回答不可」= 電子回答不可行の表示スイッチ
            public HashSet<string> StaffKeys = new HashSet<string>();   // 担当（選択クラスタのメンバーキーの合併。空=条件なし）
            public List<string> StaffNames = new List<string>();        // 〃 チップ表示用の代表名（トグルの並び順）

            public bool HasSurrender { get { return SurrenderMin != null || SurrenderMax != null; } }
            public bool HasBaseDate  { get { return BaseDateFrom != null || BaseDateTo != null; } }
            public bool HasMaturity  { get { return MaturityFrom != null || MaturityTo != null; } }
            public bool HasNote      { get { return Notes.Any(n => n); } }
            public bool HasClass     { get { return ShowNoContract || ShowUnanswered; } }
            public bool HasStaff     { get { return StaffKeys.Count > 0; } }
        }

        // 文字入力による絞り込みの再実行を、入力が止まってから 250ms 後の1回にまとめる
        //（打鍵のたびに一覧を組み直さないための性能対策）。トグルの切替・
        // 条件チップの✕・「条件をすべてクリア」は OnInsuranceAdvancedFilterChanged 直呼びの即時反映
        private void RestartInsuranceFilterDebounce()
        {
            if (suppressInsuranceAdvChange) return;
            if (insuranceFilterDebounceTimer == null)
            {
                insuranceFilterDebounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
                insuranceFilterDebounceTimer.Tick += delegate { OnInsuranceAdvancedFilterChanged(); };
            }
            insuranceFilterDebounceTimer.Stop();
            insuranceFilterDebounceTimer.Start();
        }

        // 詳細検索の入力・選択の変化を一覧へ反映する（読み取り→チップ再構築→トグル表示→絞り込み）
        public void OnInsuranceAdvancedFilterChanged()
        {
            if (suppressInsuranceAdvChange) return;
            if (insuranceFilterDebounceTimer != null) insuranceFilterDebounceTimer.Stop();   // 保留中の遅延反映はこの実行が兼ねる

            // 条件が変わると一覧が入れ替わるため、開いている詳細パネルは閉じる
            //（年度切替・ビュー切替と同じ扱い。入力中の欄からフォーカスは奪わない）
            if (isInsurancePanelOpen) CloseInsuranceDetailPanel(false);
            ReadInsuranceAdvancedInputs();
            RebuildInsuranceFilterChips();
            UpdateInsuranceAdvToggleVisual();
            ApplyInsuranceFilter();
        }

        // パネルの入力値を条件セットへ読み取る
        private void ReadInsuranceAdvancedInputs()
        {
            var f = insuranceAdvFilter;
            f.AddressNumber = BusinessLogic.NormalizeForSearch(insuranceAdvAddressBox.Text);
            f.Name          = BusinessLogic.NormalizeForSearch(insuranceAdvNameBox.Text);
            f.Kana          = BusinessLogic.NormalizeForSearch(insuranceAdvKanaBox.Text);
            f.Company       = BusinessLogic.NormalizeForSearch(insuranceAdvCompanyBox.Text);
            f.PolicyNumber  = BusinessLogic.NormalizeForSearch(insuranceAdvPolicyBox.Text);

            f.SurrenderMin = ReadAmountInput(insuranceAdvSurrenderMinBox, insuranceAdvSurrenderMinBd);
            f.SurrenderMax = ReadAmountInput(insuranceAdvSurrenderMaxBox, insuranceAdvSurrenderMaxBd);
            f.BaseDateFrom = ReadDateInput(insuranceAdvBaseFromBox, insuranceAdvBaseFromBd);
            f.BaseDateTo   = ReadDateInput(insuranceAdvBaseToBox, insuranceAdvBaseToBd);
            f.MaturityFrom = ReadDateInput(insuranceAdvMaturityFromBox, insuranceAdvMaturityFromBd);
            f.MaturityTo   = ReadDateInput(insuranceAdvMaturityToBox, insuranceAdvMaturityToBd);

            for (int i = 0; i < insuranceAdvNoteChips.Length; i++)
                f.Notes[i] = insuranceAdvNoteChips[i].IsChecked == true;
            f.ShowNoContract = insuranceAdvClassNoContract.IsChecked == true;
            f.ShowUnanswered = insuranceAdvClassUnanswered.IsChecked == true;

            f.StaffKeys.Clear();
            f.StaffNames.Clear();
            for (int i = 0; i < insuranceAdvStaffToggles.Count; i++)
            {
                if (insuranceAdvStaffToggles[i].IsChecked != true) continue;
                f.StaffNames.Add(insuranceStaffClusters[i].Representative);
                foreach (var key in insuranceStaffClusters[i].MemberKeys) f.StaffKeys.Add(key);
            }
        }

        // 条件チップ行を現在の条件から組み直す。チップの数はバッジとフッター表記が共有する
        private void RebuildInsuranceFilterChips()
        {
            insuranceFilterChipsRow.Children.Clear();

            AddInsuranceTextConditionChip("宛名番号", insuranceAdvAddressBox);
            AddInsuranceTextConditionChip("氏名", insuranceAdvNameBox);
            AddInsuranceTextConditionChip("カナ氏名", insuranceAdvKanaBox);
            AddInsuranceTextConditionChip("保険会社名", insuranceAdvCompanyBox);
            AddInsuranceTextConditionChip("証券番号", insuranceAdvPolicyBox);

            var f = insuranceAdvFilter;
            if (f.HasSurrender)
            {
                var text = "解約返戻金 " +
                    (f.SurrenderMin != null ? f.SurrenderMin.Value.ToString("N0") + "円" : "") + "〜" +
                    (f.SurrenderMax != null ? f.SurrenderMax.Value.ToString("N0") + "円" : "");
                app.AddConditionChip(insuranceFilterChipsRow, insuranceSearchBox, text, delegate
                {
                    suppressInsuranceAdvChange = true;
                    insuranceAdvSurrenderMinBox.Text = "";
                    insuranceAdvSurrenderMaxBox.Text = "";
                    suppressInsuranceAdvChange = false;
                    OnInsuranceAdvancedFilterChanged();
                });
            }
            if (f.HasBaseDate)
            {
                app.AddConditionChip(insuranceFilterChipsRow, insuranceSearchBox, FormatDateRangeChip("回答基準日", f.BaseDateFrom, f.BaseDateTo), delegate
                {
                    suppressInsuranceAdvChange = true;
                    insuranceAdvBaseFromBox.Text = "";
                    insuranceAdvBaseToBox.Text = "";
                    suppressInsuranceAdvChange = false;
                    OnInsuranceAdvancedFilterChanged();
                });
            }
            if (f.HasMaturity)
            {
                app.AddConditionChip(insuranceFilterChipsRow, insuranceSearchBox, FormatDateRangeChip("満期年月日", f.MaturityFrom, f.MaturityTo), delegate
                {
                    suppressInsuranceAdvChange = true;
                    insuranceAdvMaturityFromBox.Text = "";
                    insuranceAdvMaturityToBox.Text = "";
                    suppressInsuranceAdvChange = false;
                    OnInsuranceAdvancedFilterChanged();
                });
            }
            if (f.HasNote)
            {
                var names = new List<string>();
                for (int i = 0; i < f.Notes.Length; i++)
                    if (f.Notes[i]) names.Add(BusinessLogic.INSURANCE_NOTE_NAMES[i]);
                app.AddConditionChip(insuranceFilterChipsRow, insuranceSearchBox, "注記: " + string.Join("・", names.ToArray()), delegate
                {
                    suppressInsuranceAdvChange = true;
                    foreach (var chip in insuranceAdvNoteChips) chip.IsChecked = false;
                    suppressInsuranceAdvChange = false;
                    OnInsuranceAdvancedFilterChanged();
                });
            }
            if (f.HasClass)
            {
                var names = new List<string>();
                if (f.ShowNoContract) names.Add("契約なし");
                if (f.ShowUnanswered) names.Add("電子回答不可");
                app.AddConditionChip(insuranceFilterChipsRow, insuranceSearchBox, "行の分類: " + string.Join("・", names.ToArray()), delegate
                {
                    suppressInsuranceAdvChange = true;
                    insuranceAdvClassNoContract.IsChecked = false;
                    insuranceAdvClassUnanswered.IsChecked = false;
                    suppressInsuranceAdvChange = false;
                    OnInsuranceAdvancedFilterChanged();
                });
            }
            if (f.HasStaff)
            {
                app.AddConditionChip(insuranceFilterChipsRow, insuranceSearchBox, "担当: " + string.Join("・", f.StaffNames.ToArray()), delegate
                {
                    suppressInsuranceAdvChange = true;
                    foreach (var toggle in insuranceAdvStaffToggles) toggle.IsChecked = false;
                    suppressInsuranceAdvChange = false;
                    OnInsuranceAdvancedFilterChanged();
                });
            }

            insuranceAdvConditionCount = insuranceFilterChipsRow.Children.Count;
            insuranceFilterChipsRow.Visibility = insuranceAdvConditionCount > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        // 項目検索欄1つ分の条件チップ（値があるときのみ。表示は入力の生値・解除で欄をクリア）
        private void AddInsuranceTextConditionChip(string label, TextBox box)
        {
            var value = box.Text.Trim();
            if (value == "") return;
            app.AddConditionChip(insuranceFilterChipsRow, insuranceSearchBox, label + ": " + value, delegate
            {
                suppressInsuranceAdvChange = true;
                box.Text = "";
                suppressInsuranceAdvChange = false;
                OnInsuranceAdvancedFilterChanged();
            });
        }

        // 詳細検索パネルを開閉する（高さアニメーション・条件は閉じても持続する）
        private void ToggleInsuranceAdvancedPanel()
        {
            insuranceAdvPanelOpen = !insuranceAdvPanelOpen;
            if (insuranceAdvPanelOpen) AnimateExpand(insuranceAdvPanel);
            else AnimateCollapse(insuranceAdvPanel);
            UpdateInsuranceAdvToggleVisual();
        }

        // キーで詳細検索を開閉する（Ctrl+Shift+F）。開いたら最初の欄へ、閉じたら検索欄へ
        public void ToggleInsuranceAdvancedPanelByKey()
        {
            ToggleInsuranceAdvancedPanel();
            if (insuranceAdvPanelOpen) app.FocusLater(delegate { insuranceAdvAddressBox.Focus(); });
            else app.FocusLater(delegate { insuranceSearchBox.Focus(); });
        }

        // キーで検索欄へ移る（Ctrl+F。全選択して打ち直せるように）
        public void FocusInsuranceSearchByKey()
        {
            app.FocusLater(delegate { insuranceSearchBox.Focus(); insuranceSearchBox.SelectAll(); });
        }

        // 「詳細検索」トグルの見た目を更新する（条件数バッジと開閉キャレット。
        // トグル自体は常時 選択面+アイリス文字の固定表示）
        private void UpdateInsuranceAdvToggleVisual()
        {
            insuranceAdvCountBadge.Visibility = insuranceAdvConditionCount > 0 ? Visibility.Visible : Visibility.Collapsed;
            insuranceAdvCountText.Text = insuranceAdvConditionCount.ToString();
            insuranceAdvToggleCaret.Text = insuranceAdvPanelOpen ? "▲" : "▼";
        }

        // パネルの「条件をすべてクリア」（クイック検索は対象外）
        private void ClearInsuranceAdvancedFilter()
        {
            suppressInsuranceAdvChange = true;
            insuranceAdvAddressBox.Text = "";
            insuranceAdvNameBox.Text = "";
            insuranceAdvKanaBox.Text = "";
            insuranceAdvCompanyBox.Text = "";
            insuranceAdvPolicyBox.Text = "";
            insuranceAdvSurrenderMinBox.Text = "";
            insuranceAdvSurrenderMaxBox.Text = "";
            insuranceAdvBaseFromBox.Text = "";
            insuranceAdvBaseToBox.Text = "";
            insuranceAdvMaturityFromBox.Text = "";
            insuranceAdvMaturityToBox.Text = "";
            foreach (var chip in insuranceAdvNoteChips) chip.IsChecked = false;
            insuranceAdvClassNoContract.IsChecked = false;
            insuranceAdvClassUnanswered.IsChecked = false;
            foreach (var toggle in insuranceAdvStaffToggles) toggle.IsChecked = false;
            suppressInsuranceAdvChange = false;
            OnInsuranceAdvancedFilterChanged();
        }

        // 担当グループのトグル群を選択中年度の台帳から組み直す。
        // 年度切替・staffFuzzyChars の変更で呼ばれ、選択中の条件は同一クラスタ
        //（代表の正規化キーの一致）が残れば維持・なければ解除される
        public void RebuildInsuranceStaffToggles()
        {
            // 復元用に現在の選択を代表の正規化キーで控える
            var selectedKeys = new HashSet<string>();
            for (int i = 0; i < insuranceAdvStaffToggles.Count; i++)
            {
                if (insuranceAdvStaffToggles[i].IsChecked == true)
                    selectedKeys.Add(BusinessLogic.NormalizeForSearch(insuranceStaffClusters[i].Representative));
            }

            insuranceStaffClusters = BusinessLogic.BuildStaffClusters(
                insuranceStaffRawValues, new HashSet<char>(app.config.StaffFuzzyChars));

            insuranceAdvStaffToggles.Clear();
            insuranceAdvStaffPanel.Children.Clear();
            suppressInsuranceAdvChange = true;   // IsChecked の復元で Checked イベントを発火させない
            foreach (var cluster in insuranceStaffClusters)
            {
                var toggle = new ToggleButton
                {
                    Style = (Style)app.window.FindResource("AdvToggleChip"),
                    Content = cluster.Representative,
                    IsChecked = selectedKeys.Contains(BusinessLogic.NormalizeForSearch(cluster.Representative))
                };
                // 2値以上を統合したトグルは、吸収した生値を暗色ツールチップで全掲する
                if (cluster.RawValues.Count > 1)
                {
                    var tip = new ToolTip { Style = (Style)app.window.FindResource("DetectTip") };
                    tip.LayoutTransform = app.displayScaleTransform;   // 表示倍率に追従
                    tip.Content = app.BuildStaffMergeTip(cluster);
                    toggle.ToolTip = tip;
                    ToolTipService.SetInitialShowDelay(toggle, 350);
                    ToolTipService.SetShowDuration(toggle, 60000);
                }
                toggle.Checked += delegate { OnInsuranceAdvancedFilterChanged(); };
                toggle.Unchecked += delegate { OnInsuranceAdvancedFilterChanged(); };
                insuranceAdvStaffPanel.Children.Add(toggle);
                insuranceAdvStaffToggles.Add(toggle);
            }
            suppressInsuranceAdvChange = false;

            insuranceAdvStaffPanel.Visibility = insuranceStaffClusters.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            insuranceAdvStaffEmptyText.Visibility = insuranceStaffClusters.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
        }
    }
}