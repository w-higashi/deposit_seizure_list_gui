// ==============================================================================
// paddock_insurance_panel.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 生命保険タブ: 詳細パネル（年度横断の人物ビュー・照会履歴のタイムライン・契約カードの
// 選択と展開・人物ヘッダー）と、一覧の行の操作・照会ファイルを開く・差押リスト連携の入口。
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
        // 詳細パネル（年度横断の人物ビュー）
        // ==============================================================
        //
        // 一覧の行を選択すると右側に開き、当該人物の照会履歴を全年度の台帳から
        // 横断収集して表示する。選択中年度の行はメモリ上の insuranceAllRecords から即時に、
        // 他年度の行はファイルを非同期に走査して合流させる（人物単位の都度読み込み）。
        // 預金のパネルと同じ列に置き、同時に開くのは一方だけ

        // --- 詳細パネルの状態 ---

        public bool isInsurancePanelOpen;                        // パネルの開閉状態
        public string insurancePanelAddressNumberShown;          // パネルに表示中の宛名番号（非同期結果の適用判定・メモの欄の組み直しにも使用）
        private DispatcherTimer insuranceClickTimer;             // シングル/ダブルクリックの判別用
        private InsuranceLedgerRecord insurancePendingClickRecord;   // クリック判別待ちの行
        private int insuranceLastClickCount;                     // 直近の MouseDown のクリック回数
        private int insuranceLastMouseDownTick;                  // 直近の MouseDown の時刻（Environment.TickCount）
        private DispatcherTimer insuranceCrossYearTimer;         // 年度横断読み込みのデバウンス用
        private BackgroundWorker insurancePersonWorker;          // 他年度台帳の走査用
        private bool insuranceCrossYearPending;                  // 走査中に人物が切り替わった場合の再走査フラグ

        // 契約カードの選択状態（選択中カード = 「照会ファイルを開く」の対象）
        // 一覧の行選択と連動し、パネル内のカードクリックでも切り替わる
        private string[] insuranceSelectedCardFields;            // 選択中カードの行（台帳の並びの全列）
        private string[] insurancePendingCardFields;             // 選択が未成立のときに復元を試みる行（SwitchInsurancePanelPerson が受けた focusFields）
        private Border insuranceSelectedCardBorder;              // 選択中カードの Border（表示切替用）

        // 人物ヘッダーの退行値（登録日が最新の非空値・RenderInsurancePersonPanel で計算する）。
        // 表示は選択カードの行の値を優先し、空の項目はこの値へ退行する
        private string insurancePanelFallbackName;
        private string insurancePanelFallbackKana;
        private string insurancePanelFallbackAddress;
        private string insurancePanelFallbackBirthDate;
        private List<KeyValuePair<string[], Border>> insurancePanelCards =
            new List<KeyValuePair<string[], Border>>();          // 描画済みカードの行→Border 対応

        private readonly Dictionary<Border, InsuranceCardExpansion> insuranceCardExpansions = new Dictionary<Border, InsuranceCardExpansion>();   // カード→展開部
        private InsuranceCardExpansion insuranceExpandedCard;    // 展開中のカード（同時に1枚）
        private string[] insuranceExpandedCardFields;            // 展開中のカードの行（パネル再構築時の復元用）
        private string[] insurancePendingExpandFields;           // 選択の成立と同時に自動展開する行

        // 台帳ファイル全体はキャッシュせず、実際に閲覧した人物の行のみを保持する
        // （過去年度ファイルは追記されないため、セッション中のキャッシュは陳腐化しない）
        private Dictionary<int, Dictionary<string, List<string[]>>> insuranceOtherYearCache =
            new Dictionary<int, Dictionary<string, List<string[]>>>();

        public readonly NotesHost insuranceNotes = new NotesHost();   // 生命保険の詳細パネルのメモの欄

        // --- 差押リスト連携 ---

        public string insuranceSeizureScriptPath;                // 生命保険差押予定一覧 作成ツールのパス（解決済み。null=連携UI非表示）
        private List<SeizureTargetPattern> insuranceSeizureTargetPatterns;   // 差押リストの対象保険会社（null=制御なし）
        private Dictionary<string, bool> insuranceSeizureTargetMatchCache;   // 保険会社キー→対象か（年度横断のカードにも効かせるメモ化）

        // --- UI 要素（詳細パネル）---

        private Border insuranceDetailPanel;
        private TextBlock insurancePanelAvatarText;
        private TextBlock insurancePanelName;
        private Run insurancePanelAddressNumber;                 // 宛名番号とカナは1つの TextBlock の Run 2本（ベースラインを揃える）
        private Run insurancePanelKana;
        private TextBlock insurancePanelAddress;
        private TextBlock insurancePanelBirthDate;               // 契約者の生年月日（日付だけ・値がなければ行ごと非表示）
        private Border insurancePanelCopyChip;                   // 「コピーしました」チップ（1.5秒で消える）
        private TextBlock insurancePanelStatRounds;
        private TextBlock insurancePanelStatContracts;
        private TextBlock insurancePanelStatFirst;
        private ScrollViewer insurancePanelScroll;
        private StackPanel insuranceTimelinePanel;
        private TextBlock insurancePanelFootnote;
        private System.Windows.Shapes.Path insurancePanelMiniSpinner;   // 年度横断走査のミニスピナー（System.IO.Path と衝突するため完全修飾）
        private RotateTransform insurancePanelSpinnerRotation;
        private Button insurancePanelCloseButton;
        private Button insurancePanelOpenFileButton;
        private Button insurancePanelSeizureButton;

        // 詳細パネルの回次1件（照会履歴タイムラインの1ノード）
        private class InsurancePersonRound
        {
            public string Key;                                                // 回次キー（回答基準日、回答基準日のない行は登録日。正準形 yyyy/MM/dd）
            public List<string[]> ContractRows = new List<string[]>();        // 契約あり行（台帳の並びの全列・収集順）
            public List<string> UnansweredCompanies = new List<string>();     // 電子回答不可行の保険会社名
            public List<string> NoContractCompanies = new List<string>();     // 契約なし行の保険会社名
        }

        // 契約カードの展開部（開閉の対象）
        private class InsuranceCardExpansion
        {
            public Border Card;
            public string[] Fields;
            public StackPanel Detail;                       // 展開部全体（開閉の対象）
            public TextBlock Chevron;
        }

        // ==============================================================
        // 画面の要素の取得・イベントの配線（詳細パネルと一覧の行の操作）
        // ==============================================================

        private void FindInsurancePanelUiElements()
        {
            var window = app.window;
            insuranceDetailPanel        = (Border)window.FindName("InsuranceDetailPanel");
            insurancePanelAvatarText    = (TextBlock)window.FindName("InsurancePanelAvatarText");
            insurancePanelName          = (TextBlock)window.FindName("InsurancePanelName");
            insurancePanelAddressNumber = (Run)window.FindName("InsurancePanelAddressNumber");
            insurancePanelKana          = (Run)window.FindName("InsurancePanelKana");
            insurancePanelAddress       = (TextBlock)window.FindName("InsurancePanelAddress");
            insurancePanelBirthDate     = (TextBlock)window.FindName("InsurancePanelBirthDate");
            insurancePanelCopyChip      = (Border)window.FindName("InsurancePanelCopyChip");
            insurancePanelStatRounds    = (TextBlock)window.FindName("InsurancePanelStatRounds");
            insurancePanelStatContracts = (TextBlock)window.FindName("InsurancePanelStatContracts");
            insurancePanelStatFirst     = (TextBlock)window.FindName("InsurancePanelStatFirst");
            insurancePanelScroll        = (ScrollViewer)window.FindName("InsurancePanelScroll");
            insuranceNotes.Panel        = (StackPanel)window.FindName("InsuranceNotesPanel");
            insuranceTimelinePanel      = (StackPanel)window.FindName("InsuranceTimelinePanel");
            insurancePanelFootnote      = (TextBlock)window.FindName("InsurancePanelFootnote");
            insurancePanelMiniSpinner   = (System.Windows.Shapes.Path)window.FindName("InsurancePanelMiniSpinner");
            insurancePanelSpinnerRotation = (RotateTransform)insurancePanelMiniSpinner.RenderTransform;
            insurancePanelCloseButton   = (Button)window.FindName("InsurancePanelCloseButton");
            insurancePanelOpenFileButton = (Button)window.FindName("InsurancePanelOpenFileButton");
            insurancePanelSeizureButton = (Button)window.FindName("InsurancePanelSeizureButton");
        }

        // 一覧の行の操作（クリックの判別・ダブルクリック・Enter・Space・Ctrl+O・Ctrl+D・右クリック・選択の変化）と
        // 詳細パネル（フッターのボタン・人物ヘッダーのコピー・✕・キー・年度横断のデバウンス）の配線
        private void SetupInsurancePanelEvents()
        {
            // --- 一覧の行クリック（シングル/ダブルの判別）---
            // パネルが閉じているときのマウスクリックのみ、ダブルクリック判定が確定するまで
            // パネルを開くのを保留する（ダブルクリック時にパネルがチラつくのを防ぐ）。
            // OSのダブルクリック判定は1回目の Down からの経過時間で行われるため、
            // 待機時間は Down 起点で計算し、Up までに消化した分を差し引いて最短化する。
            insuranceClickTimer = new DispatcherTimer();
            insuranceClickTimer.Tick += delegate
            {
                insuranceClickTimer.Stop();
                // 保留中に選択が移動した場合（キーボード操作・再絞り込み等）は破棄する
                if (!isInsurancePanelOpen && insurancePendingClickRecord != null &&
                    ReferenceEquals(insurancePendingClickRecord, insuranceLedgerList.SelectedItem))
                    OpenInsuranceDetailPanel(insurancePendingClickRecord);
                insurancePendingClickRecord = null;
            };

            insuranceLedgerList.PreviewMouseLeftButtonDown += delegate(object s, MouseButtonEventArgs e)
            {
                insuranceLastClickCount = e.ClickCount;
                insuranceLastMouseDownTick = Environment.TickCount;
            };
            insuranceLedgerList.PreviewMouseLeftButtonUp += delegate(object s, MouseButtonEventArgs e)
            {
                if (insuranceLastClickCount >= 2) return;   // ダブルクリック2回目の Up では保留を再開しない
                var record = GetInsuranceRecordFromSource(e.OriginalSource);
                if (record == null) return;               // ヘッダー・余白・スクロールバーのクリックは無視
                if (isInsurancePanelOpen) return;         // 開いているときは SelectionChanged が中身を切替済み

                // 残り待機時間 = ダブルクリック判定時間 − Down からの経過時間
                // （経過済みならダブルクリックはもう成立しないため即座に開く）
                int remaining = (int)GetDoubleClickTime() - (Environment.TickCount - insuranceLastMouseDownTick);
                if (remaining <= 0)
                {
                    OpenInsuranceDetailPanel(record);
                    return;
                }
                insurancePendingClickRecord = record;
                insuranceClickTimer.Stop();
                insuranceClickTimer.Interval = TimeSpan.FromMilliseconds(remaining);
                insuranceClickTimer.Start();
            };
            insuranceLedgerList.MouseDoubleClick += delegate(object s, MouseButtonEventArgs e)
            {
                insuranceClickTimer.Stop();
                insurancePendingClickRecord = null;
                var record = GetInsuranceRecordFromSource(e.OriginalSource);
                if (record != null) OpenInsuranceInquiryFileForRow(record.RawFields);
            };

            // --- Enter: 選択行の照会ファイルを開く（ダブルクリックと同じ動作）---
            insuranceLedgerList.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.Key != Key.Enter) return;
                var record = insuranceLedgerList.SelectedItem as InsuranceLedgerRecord;
                if (record != null)
                {
                    OpenInsuranceInquiryFileForRow(record.RawFields);
                    e.Handled = true;
                }
            };

            // --- Space=詳細パネルを開く（開いていればパネルのカードへ移る）・Ctrl+O=照会ファイルを開く・
            //     Ctrl+D=差押リストに追加（右クリックメニューと同じ条件）。Space は行コンテナが先に
            //     処理するため Preview で受ける ---
            insuranceLedgerList.PreviewKeyDown += delegate(object s, KeyEventArgs e)
            {
                var record = insuranceLedgerList.SelectedItem as InsuranceLedgerRecord;
                if (record == null) return;
                var modifiers = Keyboard.Modifiers;
                if (e.Key == Key.Space && modifiers == ModifierKeys.None)
                {
                    if (!isInsurancePanelOpen) OpenInsuranceDetailPanel(record);
                    else FocusInsurancePanelCard();
                }
                else if (e.Key == Key.O && modifiers == ModifierKeys.Control) OpenInsuranceInquiryFileForRow(record.RawFields);
                else if (e.Key == Key.D && modifiers == ModifierKeys.Control)
                {
                    if (CanAddToInsuranceSeizureList(record)) AddInsuranceRowToSeizureList(record.RawFields);
                }
                else return;
                e.Handled = true;
            };

            // --- 右クリックメニュー（行の右クリックで選択された行が対象）---
            var openMenuItem = new MenuItem { Header = "照会ファイルを開く" };
            openMenuItem.Click += delegate
            {
                var record = insuranceLedgerList.SelectedItem as InsuranceLedgerRecord;
                if (record != null) OpenInsuranceInquiryFileForRow(record.RawFields);
            };
            var listMenu = new ContextMenu();
            listMenu.Items.Add(openMenuItem);
            listMenu.Opened += delegate { openMenuItem.IsEnabled = insuranceLedgerList.SelectedItem != null; };

            // 「差押リストに追加」は連携先設定時のみ生成・追加する（連携を使わない環境では項目ごと出さない）
            if (insuranceSeizureScriptPath != null)
            {
                var seizureMenuItem = new MenuItem { Header = "差押リストに追加" };
                seizureMenuItem.Click += delegate
                {
                    var record = insuranceLedgerList.SelectedItem as InsuranceLedgerRecord;
                    if (record != null) AddInsuranceRowToSeizureList(record.RawFields);
                };
                listMenu.Items.Add(seizureMenuItem);
                listMenu.Opened += delegate
                {
                    seizureMenuItem.IsEnabled = CanAddToInsuranceSeizureList(insuranceLedgerList.SelectedItem as InsuranceLedgerRecord);
                };
            }
            insuranceLedgerList.ContextMenu = listMenu;

            // --- 行選択の変化（クリック・↑↓共通）: パネルが開いていれば中身を切替 ---
            // パネルが閉じているときは ↑↓ だけではパネルを開かない（開くのはクリックか Space）
            insuranceLedgerList.SelectionChanged += delegate
            {
                if (!isInsurancePanelOpen) return;
                var record = insuranceLedgerList.SelectedItem as InsuranceLedgerRecord;
                if (record != null) SwitchInsurancePanelPerson(record.AddressNumber, record.RawFields);
            };

            // --- フッター「照会ファイルを開く」「差押リストに追加」: 選択中カードが対象 ---
            insurancePanelOpenFileButton.Click += delegate { OpenInsuranceInquiryFileForRow(insuranceSelectedCardFields); };
            insurancePanelSeizureButton.Click += delegate { AddInsuranceRowToSeizureList(insuranceSelectedCardFields); };

            // --- 人物ヘッダーの各項目はダブルクリック（または Tab で移って Enter）でその文字列をクリップボードへ ---
            app.WireCopy(insurancePanelName, delegate { return insurancePanelName.Text == NO_NAME_TEXT ? "" : insurancePanelName.Text; }, insurancePanelCopyChip);
            app.WireCopy(insurancePanelAddressNumber, delegate { return insurancePanelAddressNumber.Text; }, insurancePanelCopyChip);
            app.WireCopy(insurancePanelKana, delegate { return insurancePanelKana.Text; }, insurancePanelCopyChip);
            app.WireCopy(insurancePanelAddress, delegate { return insurancePanelAddress.Text; }, insurancePanelCopyChip);
            app.WireCopy(insurancePanelBirthDate, delegate { return insurancePanelBirthDate.Text; }, insurancePanelCopyChip);

            // --- パネルを閉じる（✕。Esc は本体のキー処理 OnWindowPreviewKeyDown）---
            insurancePanelCloseButton.Click += delegate { CloseInsuranceDetailPanel(); };
            insuranceDetailPanel.KeyDown += OnInsuranceDetailPanelKeyDown;

            // --- 年度横断読み込みのデバウンス（行を高速に移動している間は走査を保留する）---
            insuranceCrossYearTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
            insuranceCrossYearTimer.Tick += delegate
            {
                insuranceCrossYearTimer.Stop();
                StartInsuranceCrossYearLoad();
            };
        }

        // 詳細パネルの中のキー（カード・人物ヘッダー・ボタンなど、メモの入力欄以外）。
        // ← で元の一覧へ戻る・Ctrl+D で選択中のカードを差押リストへ（下部の「差押リストに追加」が押せるとき）
        private void OnInsuranceDetailPanelKeyDown(object sender, KeyEventArgs e)
        {
            if (e.OriginalSource is TextBox) return;
            var modifiers = Keyboard.Modifiers;
            if (e.Key == Key.Left && modifiers == ModifierKeys.None)
            {
                ReturnFocusFromInsurancePanel();
            }
            else if (e.Key == Key.D && modifiers == ModifierKeys.Control)
            {
                if (insurancePanelSeizureButton.Visibility == Visibility.Visible && insurancePanelSeizureButton.IsEnabled)
                    AddInsuranceRowToSeizureList(insuranceSelectedCardFields);
            }
            else return;
            e.Handled = true;
        }

        // クリックイベントの発生元から ListViewItem のデータ行を取得する（余白・ヘッダーは null）
        private static InsuranceLedgerRecord GetInsuranceRecordFromSource(object source)
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
            return item == null ? null : item.DataContext as InsuranceLedgerRecord;
        }

        // ==============================================================
        // パネルの開閉・人物の切替・年度横断の読み込み
        // ==============================================================

        // パネルを開いて指定行の人物を表示する（一覧の選択行に対応するカードを選択状態にする）。
        // 預金のパネルが開いていれば先に閉じる（同時に開くのは一方だけ）
        private void OpenInsuranceDetailPanel(InsuranceLedgerRecord record)
        {
            if (app.isDepositPanelOpen) app.CloseDepositDetailPanel(false);
            isInsurancePanelOpen = true;
            SwitchInsurancePanelPerson(record.AddressNumber, record.RawFields);
            AnimateDetailPanel(insuranceDetailPanel, DETAIL_PANEL_WIDTH);
        }

        // 最近見た人の履歴行から開く（記録した行のカードを行キーで照合して選択する）。預金のパネルが
        // 開いていれば先に閉じる。台帳を読む前なら台帳ファイルの検出だけを行い、その人の全年度の行は
        // パネルの年度横断の走査で集める（一覧はこのとき作らない——生命保険タブを開いたときに読む）
        public void OpenInsurancePanelFromRecent(string[] fields)
        {
            if (app.isDepositPanelOpen) app.CloseDepositDetailPanel(false);
            EnsureInsuranceLedgerDetected();
            isInsurancePanelOpen = true;
            SwitchInsurancePanelPerson(fields[BusinessLogic.INSURANCE_COL_ADDRESS_NUM], fields);
            AnimateDetailPanel(insuranceDetailPanel, DETAIL_PANEL_WIDTH);
        }

        // パネルを閉じる（表示中の人物情報は保持し、同一人物での再オープン時に再利用する。
        // 他年度の走査が進行中でも中断しない — 結果は非表示のパネルに無害に反映される）
        // returnFocus=false は検索条件の変更で閉じる場合（入力中の欄からフォーカスを奪わない）
        public void CloseInsuranceDetailPanel(bool returnFocus = true)
        {
            isInsurancePanelOpen = false;
            AnimateDetailPanel(insuranceDetailPanel, 0);

            if (returnFocus) ReturnFocusFromInsurancePanel();
        }

        // パネルを開いた元の一覧へフォーカスを戻す。生命保険の一覧では選択中の行コンテナへ
        //（ListView 本体へのフォーカスだと ↑↓ が行移動ではなくキーボードスクロールとして
        // 処理されてしまう）、最近見た人では選択中の履歴行へ
        private void ReturnFocusFromInsurancePanel()
        {
            if (insuranceLedgerView.Visibility == Visibility.Visible)
            {
                if (!FocusInsuranceLedgerSelection()) insuranceLedgerList.Focus();
            }
            else if (app.recentView.Visibility == Visibility.Visible && app.recentSelectedBorder != null)
                app.recentSelectedBorder.Focus();
        }

        // パネルの表示対象を指定の人物に切り替える
        // focusFields は選択するカードの行（一覧の選択行と連動・契約なし行・電子回答不可行なら選択なし）。
        // 同一人物の別契約行を選択した場合はカード選択のみ更新する（アコーディオンの展開
        // 状態を保つ。人物ヘッダーは選択行連動のため SelectInsuranceCard 側が更新する）
        private void SwitchInsurancePanelPerson(string addressNumber, string[] focusFields)
        {
            insurancePendingCardFields = focusFields;   // 再構築時の選択復元用（未成立の選択の意図を保持する）
            insurancePendingExpandFields = focusFields; // 選択の成立と同時に当該カードを自動展開する
            if (addressNumber == insurancePanelAddressNumberShown)
            {
                ApplyInsuranceCardSelection(focusFields);
                return;
            }
            insuranceSelectedCardFields = focusFields;
            insurancePanelScroll.ScrollToTop();   // 前の人物のスクロール位置を持ち越さない

            // 読み込み済みで「すべての年度」なら全年度の行が insuranceAllRecords に載っているため他年度なし。
            // 台帳を読む前（最近見た人から開いたとき）は全年度が他年度（ファイルの走査で集める）
            var otherYears = new List<int>();
            foreach (var year in insuranceLedgerFiles.Keys)
                if (!IsInsuranceYearInMemory(year)) otherYears.Add(year);

            // 他年度がない、または全ての他年度がこの人物についてキャッシュ済みなら即時に合成表示
            bool allCached = true;
            foreach (var year in otherYears)
            {
                Dictionary<string, List<string[]>> yearCache;
                if (!insuranceOtherYearCache.TryGetValue(year, out yearCache) || !yearCache.ContainsKey(addressNumber))
                {
                    allCached = false;
                    break;
                }
            }
            if (otherYears.Count == 0 || allCached)
            {
                MergeAndRenderInsurancePanel(addressNumber);
                return;
            }

            // まず選択中年度の分を即時表示し、他年度分はデバウンス後に非同期で合流させる
            //（走査中はヘッダーのミニスピナーで示す）
            RenderInsurancePersonPanel(addressNumber, CollectInsuranceSelectedYearRows(addressNumber), null);
            SetInsurancePanelScanSpinner(true);
            insuranceCrossYearTimer.Stop();
            insuranceCrossYearTimer.Start();
        }

        // 表示中の台帳（メモリ上の insuranceAllRecords）から指定人物の行を集める
        private List<string[]> CollectInsuranceSelectedYearRows(string addressNumber)
        {
            var rows = new List<string[]>();
            foreach (var record in insuranceAllRecords)
                if (record.AddressNumber == addressNumber) rows.Add(record.RawFields);
            return rows;
        }

        // 指定の年度の行がメモリ上（insuranceAllRecords）にあるか: 台帳を読み込み済みで、表示中の年度が
        // 「すべての年度」かその年度。読む前（最近見た人から開いたとき）はどの年度もファイルの走査で集める
        private bool IsInsuranceYearInMemory(int year)
        {
            if (!insuranceLedgerLoaded) return false;
            var selectedYear = GetSelectedYear(insuranceYearCombo);
            return selectedYear == null || selectedYear.Value == year;
        }

        // 他年度台帳から表示中人物の行を走査する（キャッシュにない年度のみ）。
        // 各ファイルは見出しの列名で読む（年度により列の並びが違いうるため）
        private void StartInsuranceCrossYearLoad()
        {
            if (insurancePanelAddressNumberShown == null) return;
            if (insurancePersonWorker != null && insurancePersonWorker.IsBusy)
            {
                insuranceCrossYearPending = true;   // 走査完了後に改めて現在の人物で再走査する
                return;
            }

            var addressNumber = insurancePanelAddressNumberShown;

            // 走査対象 = メモリ上にない年度で、この人物の結果が未キャッシュの年度（読み込み済みで「すべての年度」は走査なし）
            var targets = new List<KeyValuePair<int, string>>();
            foreach (var pair in insuranceLedgerFiles)
            {
                if (IsInsuranceYearInMemory(pair.Key)) continue;
                Dictionary<string, List<string[]>> yearCache;
                if (insuranceOtherYearCache.TryGetValue(pair.Key, out yearCache) && yearCache.ContainsKey(addressNumber)) continue;
                targets.Add(pair);
            }
            if (targets.Count == 0)
            {
                MergeAndRenderInsurancePanel(addressNumber);
                return;
            }
            SetInsurancePanelScanSpinner(true);   // 走査の間はヘッダーのミニスピナーで示す（入口によらずここで点ける）

            var epoch = insuranceLedgerEpoch;   // 走査を始めた世代（F5・台帳フォルダの変更で進む。古い世代の結果はキャッシュに入れない）
            insurancePersonWorker = new BackgroundWorker();
            insurancePersonWorker.DoWork += delegate(object s, DoWorkEventArgs e)
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
                        var map = BusinessLogic.BuildLedgerColumnMap(header, BusinessLogic.INSURANCE_LEDGER_COLUMNS, BusinessLogic.INSURANCE_REQUIRED_COLUMNS, out missing);
                        if (map == null) continue;   // 見出しがない・必須の列が欠ける年度は読めなかった年度として扱う
                        bool identity = BusinessLogic.IsIdentityColumnMap(map, header.Length, BusinessLogic.INSURANCE_LEDGER_COLUMN_COUNT);
                        var rows = new List<string[]>();
                        for (int i = 1; i < rawRecords.Count; i++)
                        {
                            var fields = rawRecords[i];
                            if (fields.Length != header.Length) continue;   // 列数が見出しと違う行は読み飛ばす
                            if (!identity) fields = BusinessLogic.ReorderLedgerFields(fields, map);
                            if (fields[BusinessLogic.INSURANCE_COL_ADDRESS_NUM] != addressNumber) continue;
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
            insurancePersonWorker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs e)
            {
                // DoWork の外で起きた例外（e.Result は null）は走査の結果なしとして同じ流れに進む
                //（キャッシュに入らないため次回の切替時に再試行される）
                if (e.Error != null) app.WriteErrorLog("年度横断の走査（生命保険）", e.Error);
                var parts = e.Error == null ? (object[])e.Result : new object[] { addressNumber, new Dictionary<int, List<string[]>>() };
                var resultAddressNumber = (string)parts[0];
                var found = (Dictionary<int, List<string[]>>)parts[1];

                if (epoch == insuranceLedgerEpoch)   // 捨てる操作のあとの完了は合流しない（別フォルダ・F5 前の行をキャッシュに入れない）
                {
                    foreach (var pair in found)
                    {
                        if (!insuranceOtherYearCache.ContainsKey(pair.Key))
                            insuranceOtherYearCache[pair.Key] = new Dictionary<string, List<string[]>>();
                        insuranceOtherYearCache[pair.Key][resultAddressNumber] = pair.Value;
                    }
                }

                // 走査中に人物が切り替わっていた場合は現在の人物で改めて走査する
                if (insuranceCrossYearPending)
                {
                    insuranceCrossYearPending = false;
                    StartInsuranceCrossYearLoad();
                    return;
                }
                if (epoch == insuranceLedgerEpoch && insurancePanelAddressNumberShown == resultAddressNumber)
                    MergeAndRenderInsurancePanel(resultAddressNumber);
            };
            insurancePersonWorker.RunWorkerAsync();
        }

        // 選択中年度の行とキャッシュ済みの他年度の行を合成してパネルを描画する
        private void MergeAndRenderInsurancePanel(string addressNumber)
        {
            var rows = CollectInsuranceSelectedYearRows(addressNumber);
            var selectedYear = GetSelectedYear(insuranceYearCombo);
            var failedYears = new List<int>();

            // 読み込み済みで「すべての年度」は insuranceAllRecords が全年度分（読めなかった年度は読み込み時の控えを引き継ぐ）。
            // 台帳を読む前はどの年度も人物単位のキャッシュから合流する
            if (insuranceLedgerLoaded && selectedYear == null)
                failedYears.AddRange(insuranceLoadFailedYears);
            else
                foreach (var year in insuranceLedgerFiles.Keys)
                {
                    if (IsInsuranceYearInMemory(year)) continue;
                    Dictionary<string, List<string[]>> yearCache;
                    List<string[]> personRows = null;
                    if (insuranceOtherYearCache.TryGetValue(year, out yearCache))
                        yearCache.TryGetValue(addressNumber, out personRows);
                    if (personRows != null) rows.AddRange(personRows);
                    else failedYears.Add(year);
                }

            string footnote = null;
            if (failedYears.Count > 0)
                footnote = FormatYearLabels(failedYears) + " の台帳を読み込めませんでした";

            // 全年度を突き合わせても行が1件もない場合はその旨を示す（一覧の行から開いたときは起こらない。
            // 最近見た人の履歴の人物が台帳から消えた場合に起こりうる。落ちない・隠さない）
            if (rows.Count == 0 && footnote == null)
                footnote = "現在の台帳にこの人物の行が見つかりません";

            SetInsurancePanelScanSpinner(false);   // 走査完了（合流結果の描画に切り替わる）
            RenderInsurancePersonPanel(addressNumber, rows, footnote);
        }

        // 年度横断走査中のミニスピナーの表示切替（人物ヘッダー右端・走査中のみ回転させる）
        private void SetInsurancePanelScanSpinner(bool scanning)
        {
            insurancePanelMiniSpinner.Visibility = scanning ? Visibility.Visible : Visibility.Collapsed;
            if (scanning) StartSpinner(insurancePanelSpinnerRotation);
            else StopSpinner(insurancePanelSpinnerRotation);
        }

        // ==============================================================
        // パネルの描画（人物ヘッダー・統計・タイムライン・契約カード）
        // ==============================================================

        // 収集済みの行（選択中年度+他年度・契約あり/なし・電子回答不可を含む）からパネルの内容を組み立てる
        private void RenderInsurancePersonPanel(string addressNumber, List<string[]> rows, string footnote)
        {
            // メモは人物に属し行データに依存しないため、表示人物が変わったときのみ再構築する
            //（他年度分の合流による再描画で「さらに表示」の展開状態を巻き戻さない）
            if (addressNumber != insurancePanelAddressNumberShown)
            {
                insuranceNotes.Expanded = false;
                app.BuildNotesSection(insuranceNotes, addressNumber);
            }
            insurancePanelAddressNumberShown = addressNumber;

            // --- 人物ヘッダーの退行値: 登録日が最新の非空値（契約なし行・電子回答不可行は氏名等が空のため）。
            //     表示は選択カードの行の値を優先する（選択行連動——末尾の ApplyInsuranceCardSelection
            //     経由で SelectInsuranceCard → UpdateInsurancePersonHeader が必ず描画する）---
            insurancePanelFallbackName = null;
            insurancePanelFallbackKana = null;
            insurancePanelFallbackAddress = null;
            insurancePanelFallbackBirthDate = null;

            // 行が1件もないとき（台帳を読む前の走査中・最近見た人の履歴の人物が台帳から消えたとき）は、
            // 開いた行（履歴の控え）の値で補う。行があるときは同じ人物の行のため結果は変わらない
            var headerRows = rows;
            if (insurancePendingCardFields != null)
                headerRows = rows.Concat(new[] { insurancePendingCardFields }).ToList();
            foreach (var fields in headerRows.OrderByDescending(
                f => BusinessLogic.GetRegistrationSortKey(f[BusinessLogic.INSURANCE_COL_REG_DATE], app.eraMapping)))
            {
                if (insurancePanelFallbackName == null && fields[BusinessLogic.INSURANCE_COL_NAME] != "") insurancePanelFallbackName = fields[BusinessLogic.INSURANCE_COL_NAME];
                if (insurancePanelFallbackKana == null && fields[BusinessLogic.INSURANCE_COL_KANA] != "") insurancePanelFallbackKana = fields[BusinessLogic.INSURANCE_COL_KANA];
                if (insurancePanelFallbackAddress == null && fields[BusinessLogic.INSURANCE_COL_ADDRESS] != "") insurancePanelFallbackAddress = fields[BusinessLogic.INSURANCE_COL_ADDRESS];
                if (insurancePanelFallbackBirthDate == null && fields[BusinessLogic.INSURANCE_COL_BIRTH_DATE] != "") insurancePanelFallbackBirthDate = fields[BusinessLogic.INSURANCE_COL_BIRTH_DATE];
            }
            insurancePanelAddressNumber.Text = addressNumber;

            // --- 回次グルーピング+ファクト統計（事実の集計のみ・評価値は置かない）---
            var rounds = GroupInsurancePersonRounds(rows);
            int contractCount = 0;
            foreach (var fields in rows)
                if (fields[BusinessLogic.INSURANCE_COL_HAS_CONTRACT].Trim() == "1") contractCount++;

            insurancePanelStatRounds.Text = rounds.Count.ToString();
            insurancePanelStatContracts.Text = contractCount.ToString();
            insurancePanelStatFirst.Text = rounds.Count > 0
                ? rounds[rounds.Count - 1].Key   // 回次キーは正準形（yyyy/MM/dd）のためそのまま表示できる
                : "－";

            // --- 照会履歴タイムライン（新しい回次が上）---
            bool cardFocused = insuranceTimelinePanel.IsKeyboardFocusWithin;
            insuranceTimelinePanel.Children.Clear();
            insurancePanelCards.Clear();
            insuranceCardExpansions.Clear();
            for (int i = 0; i < rounds.Count; i++)
                insuranceTimelinePanel.Children.Add(BuildInsuranceRoundNode(rounds[i], i, i == rounds.Count - 1));

            // カード選択の復元。選択が成立していない場合（他年度の行の合流前に
            // 当該カードがまだ存在しない等）は、直近に指定された行での選択を試みる
            ApplyInsuranceCardSelection(insuranceSelectedCardFields != null ? insuranceSelectedCardFields : insurancePendingCardFields);
            if (cardFocused) app.FocusLater(delegate { if (insuranceSelectedCardBorder != null) insuranceSelectedCardBorder.Focus(); });

            insurancePanelFootnote.Text = footnote != null ? footnote : "";
            insurancePanelFootnote.Visibility = footnote != null ? Visibility.Visible : Visibility.Collapsed;
        }

        // 行を回次（回答基準日、回答基準日のない行は登録日）でグルーピングし、新しい順に並べる。
        // キーは正準形（yyyy/MM/dd）のため、日付形式が違っても同じ日付は同一回次に
        // まとまり、Ordinal 比較がそのまま時系列順になる。
        // 契約有無の値で 契約あり行（カード）／電子回答不可行／契約なし行（保険会社名の1行）に振り分ける
        private List<InsurancePersonRound> GroupInsurancePersonRounds(List<string[]> rows)
        {
            var map = new Dictionary<string, InsurancePersonRound>();
            foreach (var fields in rows)
            {
                var baseDate = fields[BusinessLogic.INSURANCE_COL_BASE_DATE];
                var key = BusinessLogic.FormatLedgerDate(
                    baseDate != "" ? baseDate : fields[BusinessLogic.INSURANCE_COL_REG_DATE], app.eraMapping);
                InsurancePersonRound round;
                if (!map.TryGetValue(key, out round))
                {
                    round = new InsurancePersonRound { Key = key };
                    map[key] = round;
                }
                var hasContract = fields[BusinessLogic.INSURANCE_COL_HAS_CONTRACT].Trim();
                if (hasContract == "1")
                    round.ContractRows.Add(fields);
                else if (hasContract == "2")
                    round.UnansweredCompanies.Add(fields[BusinessLogic.INSURANCE_COL_COMPANY]);
                else
                    round.NoContractCompanies.Add(fields[BusinessLogic.INSURANCE_COL_COMPANY]);
            }
            return map.Values.OrderByDescending(r => r.Key, StringComparer.Ordinal).ToList();
        }

        // 回次1件のタイムラインノード（レール+内容）を組み立てる
        // index はタイムライン上の位置（0=最新）。レールのドットの濃淡は時系列位置の表現であり評価ではない
        private UIElement BuildInsuranceRoundNode(InsurancePersonRound round, int index, bool isLast)
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

            // --- 回次の内容（日付ヘッダー+契約カード+電子回答不可行+契約なし行）---
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
            if (round.ContractRows.Count > 0)
            {
                var countText = new TextBlock
                {
                    Text = "契約 " + round.ContractRows.Count + "件",
                    FontSize = 10.5,
                    Foreground = BrushTextMuted,
                    VerticalAlignment = VerticalAlignment.Bottom
                };
                Grid.SetColumn(countText, 1);
                head.Children.Add(countText);
            }
            content.Children.Add(head);

            foreach (var fields in round.ContractRows)
                content.Children.Add(BuildInsuranceContractCard(fields));

            if (round.UnansweredCompanies.Count > 0)
                content.Children.Add(BuildNoAccountLine("電子回答不可", round.UnansweredCompanies));

            if (round.NoContractCompanies.Count > 0)
            {
                var noContractLine = BuildNoAccountLine("契約なし", round.NoContractCompanies);
                // 電子回答不可の行に続くときは行の間を空ける（1行だけなら預金の口座なし行と同じ余白）
                if (round.UnansweredCompanies.Count > 0) ((FrameworkElement)noContractLine).Margin = new Thickness(1, 8, 0, 0);
                content.Children.Add(noContractLine);
            }

            node.Children.Add(content);
            return node;
        }

        // 契約カード1件（1行表示+クリックで選択と展開）を組み立てる
        // 全てのカードがクリックで選択でき、ダブルクリックでその回次の照会ファイルを開く。
        // 展開部は契約・返戻金等・差押等の3ブロック（値のある行だけ。行がなければ展開なし＝選択のみ）
        private UIElement BuildInsuranceContractCard(string[] fields)
        {
            // 展開部の内容を先に組み立て、空なら展開なし（選択のみ）の表示にする
            var blocks = BusinessLogic.BuildInsurancePanelBlocks(fields);
            bool expandable = blocks.Count > 0;

            var card = new Border
            {
                Style = (Style)app.window.FindResource("AccountCard"),
                Margin = new Thickness(0, 0, 0, 6),
                Cursor = Cursors.Hand
            };
            insurancePanelCards.Add(new KeyValuePair<string[], Border>(fields, card));

            var outer = new StackPanel();
            card.Child = outer;

            // --- 1行表示: 保険会社名 / 証券番号・契約種類 / 解約返戻金 ---
            var head = new Grid();
            head.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var left = new StackPanel();
            left.Children.Add(new TextBlock
            {
                Text = fields[BusinessLogic.INSURANCE_COL_COMPANY],
                FontSize = 12,
                Foreground = BrushTextMain,
                TextTrimming = TextTrimming.CharacterEllipsis
            });
            var policyLine = new TextBlock
            {
                FontSize = 11,
                Foreground = BrushTextSub,
                Margin = new Thickness(0, 2, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            policyLine.Inlines.Add(new Run(fields[BusinessLogic.INSURANCE_COL_POLICY_NUM]) { FontFamily = FontMono });
            var contractType = fields[BusinessLogic.INSURANCE_COL_CONTRACT_TYPE];
            if (contractType != "") policyLine.Inlines.Add(new Run("・" + contractType));
            left.Children.Add(policyLine);
            head.Children.Add(left);

            // 右側は縦積み（解約返戻金の右端を全カードで揃え、シェブロンはその下に置く）
            var right = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(8, 0, 0, 0)
            };
            var surrender = fields[BusinessLogic.INSURANCE_COL_SURRENDER];
            right.Children.Add(new TextBlock
            {
                Text = surrender,
                FontFamily = FontMono,
                FontSize = 12,
                HorizontalAlignment = HorizontalAlignment.Right,
                // 数値として読めない生値（「電子回答不可」など）は薄文字色で生値のまま（一覧の列と同じ）
                Foreground = surrender != "" && BusinessLogic.ParseLedgerAmount(surrender) == null ? BrushTextMutedRow : BrushTextMain
            });
            right.Children.Add(new TextBlock
            {
                Text = "解約返戻金",
                FontSize = 10,
                Foreground = BrushTextMuted,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 1, 0, 0)
            });

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

            // --- 注記のチップ（全部出す。暗色ツールチップは一覧の注記列と同じ第2層）---
            List<TooltipLineInfo> tipLines;
            var chipInfos = BuildInsuranceNoteChipInfos(BusinessLogic.BuildInsuranceNoteFlags(fields), fields, out tipLines);
            if (chipInfos.Count > 0)
            {
                var marks = new WrapPanel { Margin = new Thickness(0, 5, 0, 0) };
                foreach (var chip in chipInfos)
                {
                    var pill = new Border
                    {
                        Background = chip.Bg,
                        CornerRadius = new CornerRadius(9),
                        Padding = new Thickness(8, 1, 8, 2),
                        Margin = new Thickness(0, 0, 4, 0)
                    };
                    pill.Child = new TextBlock { Text = chip.Text, FontSize = 10.5, Foreground = chip.Fg };
                    marks.Children.Add(pill);
                }
                var tip = new ToolTip { Style = (Style)app.window.FindResource("DetectTip") };
                tip.LayoutTransform = app.displayScaleTransform;   // 表示倍率に追従
                tip.Content = app.BuildDetectTipContent(tipLines);
                marks.ToolTip = tip;
                ToolTipService.SetInitialShowDelay(marks, 350);
                ToolTipService.SetShowDuration(marks, 60000);
                outer.Children.Add(marks);
            }

            // --- 展開部（クリックで開閉・同時に開くのは1枚）---
            InsuranceCardExpansion expansion = null;
            if (expandable)
            {
                var detail = new StackPanel { Visibility = Visibility.Collapsed, Margin = new Thickness(0, 6, 0, 0) };
                foreach (var line in BuildInsurancePanelBlockRows(blocks)) detail.Children.Add(line);
                expansion = new InsuranceCardExpansion { Card = card, Fields = fields, Detail = detail, Chevron = chevron };
                insuranceCardExpansions[card] = expansion;
                outer.Children.Add(detail);
            }

            // --- クリック／Space=選択+展開切替・ダブルクリック／Enter=照会ファイルを開く・↑↓=カードの移動 ---
            Action selectAndToggle = delegate
            {
                SelectInsuranceCard(fields, card);
                insurancePendingExpandFields = null;   // 職員が自分で選んだカードが優先（保留中の自動展開は捨てる）
                if (expansion != null)
                {
                    if (ReferenceEquals(insuranceExpandedCard, expansion)) CollapseInsuranceCard();
                    else ExpandInsuranceCard(expansion, false);
                }
            };
            Action selectAndOpen = delegate
            {
                SelectInsuranceCard(fields, card);
                OpenInsuranceInquiryFileForRow(fields);
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
                if (insuranceTimelinePanel.IsKeyboardFocusWithin) card.Focus();   // キーで辿っている途中なら枠も移す
            };
            card.Focusable = true;
            card.FocusVisualStyle = app.keyFocusStyle;
            card.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.OriginalSource != card || Keyboard.Modifiers != ModifierKeys.None) return;
                if (e.Key == Key.Space) selectAndToggle();
                else if (e.Key == Key.Enter) selectAndOpen();
                else if (e.Key == Key.Up || e.Key == Key.Down)
                {
                    // 端のカードでは何もしない（パネルのスクロールに任せる）
                    if (!MoveInsurancePanelCardFocus(card, e.Key == Key.Down ? 1 : -1)) return;
                }
                else return;
                e.Handled = true;
            };
            return card;
        }

        // 展開部のブロック（見出し+「ラベル／値」の行）を組み立てる。
        // 値は片を並べた1つの TextBlock（数値・日付は Consolas）。読めない生値は薄文字色
        private List<UIElement> BuildInsurancePanelBlockRows(List<InsurancePanelBlock> blocks)
        {
            var lines = new List<UIElement>();
            for (int b = 0; b < blocks.Count; b++)
            {
                lines.Add(new TextBlock
                {
                    Text = blocks[b].Title,
                    FontSize = 10.5,
                    Foreground = BrushTextMuted,
                    Margin = new Thickness(0, b == 0 ? 0 : 4, 0, 3)
                });
                foreach (var row in blocks[b].Rows)
                {
                    var grid = new Grid { Margin = new Thickness(0, 0, 0, 3) };
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(132) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                    grid.Children.Add(new TextBlock { Text = row.Label, FontSize = 11.5, Foreground = BrushTextSub });

                    var value = new TextBlock
                    {
                        FontSize = 11.5,
                        Foreground = row.IsRawText ? BrushTextMutedRow : BrushTextMain,
                        TextWrapping = TextWrapping.Wrap,
                        Margin = new Thickness(8, 0, 0, 0)
                    };
                    foreach (var part in row.Parts)
                    {
                        var run = new Run(part.Text);
                        if (part.Mono) run.FontFamily = FontMono;
                        value.Inlines.Add(run);
                    }
                    if (row.ToolTip != null) value.ToolTip = row.ToolTip;
                    Grid.SetColumn(value, 1);
                    grid.Children.Add(value);
                    lines.Add(grid);
                }
            }
            return lines;
        }

        // 描画順（新しい回次が上）で隣のカードへフォーカスと選択を移す。移れなければ false
        private bool MoveInsurancePanelCardFocus(Border from, int delta)
        {
            int index = insurancePanelCards.FindIndex(delegate(KeyValuePair<string[], Border> pair) { return ReferenceEquals(pair.Value, from); });
            int next = index + delta;
            if (index < 0 || next < 0 || next >= insurancePanelCards.Count) return false;
            SelectInsuranceCard(insurancePanelCards[next].Key, insurancePanelCards[next].Value);
            insurancePendingExpandFields = null;   // 職員が自分で選んだカードが優先（クリックと同じ）
            insurancePanelCards[next].Value.Focus();
            return true;
        }

        // 選択中のカード（なければ先頭のカードを選んで）へフォーカスを移す
        public void FocusInsurancePanelCard()
        {
            if (insuranceSelectedCardBorder == null && insurancePanelCards.Count > 0) SelectInsuranceCard(insurancePanelCards[0].Key, insurancePanelCards[0].Value);
            if (insuranceSelectedCardBorder != null) insuranceSelectedCardBorder.Focus();
        }

        // 展開中のカードを閉じる
        private void CollapseInsuranceCard()
        {
            if (insuranceExpandedCard == null) return;
            insuranceExpandedCard.Detail.Visibility = Visibility.Collapsed;
            insuranceExpandedCard.Chevron.Text = "▾";
            insuranceExpandedCard = null;
            insuranceExpandedCardFields = null;
        }

        // カードを展開する（開いていた別のカードは閉じる）。scroll=true は展開部が見える位置まで送る
        private void ExpandInsuranceCard(InsuranceCardExpansion expansion, bool scroll)
        {
            if (!ReferenceEquals(insuranceExpandedCard, expansion)) CollapseInsuranceCard();
            insuranceExpandedCard = expansion;
            insuranceExpandedCardFields = expansion.Fields;
            expansion.Detail.Visibility = Visibility.Visible;
            expansion.Chevron.Text = "▴";
            if (scroll)
            {
                // 展開部の高さが確定してから送る（Visible にした直後は 0 のまま）
                expansion.Card.Dispatcher.BeginInvoke(new Action(delegate { expansion.Card.BringIntoView(); }), DispatcherPriority.Loaded);
            }
        }

        // パネル再構築後の展開状態の復元: 選択と同時に自動展開する行（一覧の行クリック・人物切替）
        // があればそのカードを開いて見える位置へ送り、なければ直前まで開いていたカードを開き直す
        private void RestoreInsuranceCardExpansion()
        {
            // カードが作り直されていれば、開いていたカードの参照はもう画面にない
            if (insuranceExpandedCard != null && !insuranceCardExpansions.ContainsKey(insuranceExpandedCard.Card)) insuranceExpandedCard = null;

            string[] foundFields;
            Border card = insurancePendingExpandFields != null ? FindInsurancePanelCard(insurancePendingExpandFields, out foundFields) : null;
            if (card != null)
            {
                insurancePendingExpandFields = null;
                InsuranceCardExpansion expansion;
                if (insuranceCardExpansions.TryGetValue(card, out expansion)) ExpandInsuranceCard(expansion, true);
                else CollapseInsuranceCard();   // 展開部のないカード（展開部の行が1つもない）
                return;
            }
            if (insuranceExpandedCardFields == null) return;
            card = FindInsurancePanelCard(insuranceExpandedCardFields, out foundFields);
            InsuranceCardExpansion previous;
            if (card != null && insuranceCardExpansions.TryGetValue(card, out previous)) ExpandInsuranceCard(previous, false);
        }

        // 指定の行に対応するカードを選択状態にする（対応するカードがなければ選択なし）
        // 契約なし行・電子回答不可行はカードを持たないため、自然に「選択なし=開くボタン無効」へ落ちる
        private void ApplyInsuranceCardSelection(string[] fields)
        {
            string[] foundFields;
            var found = FindInsurancePanelCard(fields, out foundFields);
            SelectInsuranceCard(foundFields, found);
            RestoreInsuranceCardExpansion();
        }

        // 行に対応する描画済みカードを探す（なければ null）。参照が一致しない場合は行キーで
        // 照合する（同じ行でも配列の参照が違うことがあるため）
        private Border FindInsurancePanelCard(string[] fields, out string[] foundFields)
        {
            foundFields = null;
            if (fields == null) return null;
            foreach (var pair in insurancePanelCards)
            {
                if (ReferenceEquals(pair.Key, fields))
                {
                    foundFields = pair.Key;
                    return pair.Value;
                }
            }
            var key = BusinessLogic.GetInsuranceLedgerRowKey(fields, app.eraMapping);
            foreach (var pair in insurancePanelCards)
            {
                if (BusinessLogic.GetInsuranceLedgerRowKey(pair.Key, app.eraMapping) == key)
                {
                    foundFields = pair.Key;
                    return pair.Value;
                }
            }
            return null;
        }

        // カードの選択状態を切り替える（選択中カード=「照会ファイルを開く」の対象）
        private void SelectInsuranceCard(string[] fields, Border card)
        {
            if (insuranceSelectedCardBorder != null)
            {
                // ローカル値を消してスタイル（ホバー含む）の表示へ戻す
                insuranceSelectedCardBorder.ClearValue(Border.BackgroundProperty);
                insuranceSelectedCardBorder.ClearValue(Border.BorderBrushProperty);
            }
            insuranceSelectedCardFields = fields;
            insuranceSelectedCardBorder = card;
            if (card != null)
            {
                card.Background = BrushSelectedCardBg;
                card.BorderBrush = BrushAccent;
            }
            UpdateInsuranceOpenFileButton();
            UpdateInsurancePersonHeader();
        }

        // 人物ヘッダー（氏名・カナ・住所・生年月日）を現在のカード選択に合わせて更新する。
        // 同一人物でも回次により住所が異なることがあり、差押リスト連携時に別契約の
        // 住所のまま取り違えないように、選択カードの行の値を優先して表示する。
        // 選択がない・選択行の項目が空のときは退行値（登録日が最新の非空値）で補う
        //（表示から情報を隠さない）
        private void UpdateInsurancePersonHeader()
        {
            var name = insurancePanelFallbackName;
            var kana = insurancePanelFallbackKana;
            var address = insurancePanelFallbackAddress;
            var birthDate = insurancePanelFallbackBirthDate;
            if (insuranceSelectedCardFields != null)
            {
                if (insuranceSelectedCardFields[BusinessLogic.INSURANCE_COL_NAME] != "") name = insuranceSelectedCardFields[BusinessLogic.INSURANCE_COL_NAME];
                if (insuranceSelectedCardFields[BusinessLogic.INSURANCE_COL_KANA] != "") kana = insuranceSelectedCardFields[BusinessLogic.INSURANCE_COL_KANA];
                if (insuranceSelectedCardFields[BusinessLogic.INSURANCE_COL_ADDRESS] != "") address = insuranceSelectedCardFields[BusinessLogic.INSURANCE_COL_ADDRESS];
                if (insuranceSelectedCardFields[BusinessLogic.INSURANCE_COL_BIRTH_DATE] != "") birthDate = insuranceSelectedCardFields[BusinessLogic.INSURANCE_COL_BIRTH_DATE];
            }
            insurancePanelName.Text = name != null ? name : NO_NAME_TEXT;
            insurancePanelAvatarText.Text = string.IsNullOrEmpty(name) ? "？" : StringInfo.GetNextTextElement(name);   // 頭文字（サロゲートペア・結合文字も 1 文字として）
            insurancePanelKana.Text = kana != null ? kana : "";
            insurancePanelAddress.Text = address != null ? address : "";
            insurancePanelAddress.Visibility = address != null ? Visibility.Visible : Visibility.Collapsed;
            insurancePanelBirthDate.Text = birthDate != null ? birthDate : "";
            insurancePanelBirthDate.Visibility = birthDate != null ? Visibility.Visible : Visibility.Collapsed;
        }

        // フッター「照会ファイルを開く」の有効/無効とツールチップを現在の状態に合わせる
        private void UpdateInsuranceOpenFileButton()
        {
            if (app.config.InsuranceSearchFolders == null)
            {
                insurancePanelOpenFileButton.IsEnabled = false;
                insurancePanelOpenFileButton.ToolTip = "config の insuranceSearchFolders を設定すると有効になります";
            }
            else if (insuranceSelectedCardBorder == null)
            {
                insurancePanelOpenFileButton.IsEnabled = false;
                insurancePanelOpenFileButton.ToolTip = "照会シートのある契約行を選択すると開けます";
            }
            else
            {
                insurancePanelOpenFileButton.IsEnabled = true;
                insurancePanelOpenFileButton.ToolTip = null;
            }
            UpdateInsuranceSeizureButton();   // 「差押リストに追加」も同じ選択状態に追従させる
        }

        // フッター「差押リストに追加」の有効/無効とツールチップを現在の状態に合わせる
        // 有効条件は「照会ファイルを開く」と同一（ファイル解決を共有するため）。
        // 連携先未設定時はボタン自体が非表示（ApplyInsuranceSeizureIntegrationVisibility）
        public void UpdateInsuranceSeizureButton()
        {
            if (insuranceSeizureScriptPath == null) return;
            if (app.config.InsuranceSearchFolders == null)
            {
                insurancePanelSeizureButton.IsEnabled = false;
                insurancePanelSeizureButton.ToolTip = "config の insuranceSearchFolders を設定すると有効になります";
            }
            else if (insuranceSelectedCardBorder == null)
            {
                insurancePanelSeizureButton.IsEnabled = false;
                insurancePanelSeizureButton.ToolTip = "照会シートのある契約行を選択すると追加できます";
            }
            else if (!IsSeizureTargetKey(BusinessLogic.NormalizeForSearch(
                insuranceSelectedCardFields[BusinessLogic.INSURANCE_COL_COMPANY]), insuranceSeizureTargetPatterns, insuranceSeizureTargetMatchCache))
            {
                // 対象保険会社の制御: 無効にするだけでツールチップは出さない
                insurancePanelSeizureButton.IsEnabled = false;
                insurancePanelSeizureButton.ToolTip = null;
            }
            else
            {
                insurancePanelSeizureButton.IsEnabled = true;
                insurancePanelSeizureButton.ToolTip = null;
            }
        }

        // insuranceSeizureListScript が未設定または不存在の場合、生命保険のパネルの連携ボタンを非表示にする
        // （連携を使わない環境では機能ごと見せない。ナビ「差押登録」の表示は預金と合わせて外側で決める）
        public void ApplyInsuranceSeizureIntegrationVisibility()
        {
            if (insuranceSeizureScriptPath != null) return;
            insurancePanelSeizureButton.Visibility = Visibility.Collapsed;
            Grid.SetColumnSpan(insurancePanelOpenFileButton, 3);   // 「照会ファイルを開く」をフッター全幅に広げる
        }

        // ==============================================================
        // 照会ファイルを開く・差押リスト連携（入口。解決の本体と起動は預金と共用）
        // ==============================================================

        // 生命保険の台帳の行から照会ファイルを解決して開く（一覧・パネル・全ての「開く」動線の入口）
        private void OpenInsuranceInquiryFileForRow(string[] fields)
        {
            var resolved = ResolveInsuranceInquiryFileForRow(fields,
                "照会ファイルの候補", "開く", "開くファイルを選んでください");
            if (resolved != null)
            {
                app.RecordRecentInsurancePerson(fields);   // 実際に開く行のみ記録する（キャンセル・0件時は記録しない）
                app.OpenInquiryFileInExcel(resolved.FullPath, resolved.ExpectedSheet);
            }
        }

        // 生命保険の台帳の行から照会ファイルを解決する（契約なし・シートのない電子回答不可・照会結果フォルダ
        // 未設定をここで振り分け、契約シートのある行は ResolveInquiryFile へ渡す）。
        // 電子回答不可の行は基本情報だけの短いシートがあればそれを開く。
        // 候補ダイアログの見出し・プライマリボタン・選択案内の文言は終端アクションに合わせて
        // 呼び出し側が渡す。解決できなかった・職員がキャンセルした場合は null を返す
        private ResolvedInquiryFile ResolveInsuranceInquiryFileForRow(
            string[] fields, string dialogTitle, string primaryButtonText, string selectPrompt)
        {
            if (fields == null) return null;
            var hasContract = fields[BusinessLogic.INSURANCE_COL_HAS_CONTRACT].Trim();
            if (hasContract != "1" && hasContract != "2")
            {
                app.ShowTransientStatus("契約なし回答に照会シートはありません");
                return null;
            }
            if (hasContract == "2" && fields[BusinessLogic.INSURANCE_COL_SHEET_NAME] == "")
            {
                app.ShowTransientStatus("電子回答不可の回答に照会シートはありません");
                return null;
            }
            if (app.config.InsuranceSearchFolders == null)
            {
                app.ShowTransientStatus("config の insuranceSearchFolders が未設定のため、照会ファイルを開く機能は無効です");
                return null;
            }

            return app.ResolveInquiryFile(app.config.InsuranceSearchFolders,
                fields[BusinessLogic.INSURANCE_COL_ADDRESS_NUM], fields[BusinessLogic.INSURANCE_COL_SOURCE_FILE],
                fields[BusinessLogic.INSURANCE_COL_SHEET_NAME], fields[BusinessLogic.INSURANCE_COL_REG_DATE],
                null, dialogTitle, primaryButtonText, selectPrompt);
        }

        // 差押リストの対象保険会社: config の insuranceSeizureTargetCompanies に一致しない
        // 保険会社の行では「差押リストに追加」を無効化する（非表示にしない・無効時の
        // ツールチップも出さない）。照合結果はキー単位でメモ化する——判定は年度横断の
        // カードにも及ぶため、パターンが変わらない限りキャッシュは年度を切り替えても有効
        public void RebuildInsuranceSeizureTargets()
        {
            insuranceSeizureTargetPatterns = BusinessLogic.ParseSeizureTargetBanks(app.config.InsuranceSeizureTargetCompanies);
            insuranceSeizureTargetMatchCache = new Dictionary<string, bool>();
        }

        // 一覧の行を差押リストに追加できるか: 連携先があり、契約あり行で、ファイル解決（insuranceSearchFolders）が
        // 設定済みで、対象保険会社に一致する行（右クリックメニューと Ctrl+D で共用）
        private bool CanAddToInsuranceSeizureList(InsuranceLedgerRecord record)
        {
            return insuranceSeizureScriptPath != null && record != null && !record.IsMutedRow && app.config.InsuranceSearchFolders != null &&
                IsSeizureTargetKey(record.SearchCompany, insuranceSeizureTargetPatterns, insuranceSeizureTargetMatchCache);
        }

        // 台帳の行から照会ファイルを解決し、生命保険差押予定一覧 作成ツールに引き渡す
        // （右クリックメニュー・パネルフッターの「差押リストに追加」の入口）。契約はツールの側で選ぶ
        private void AddInsuranceRowToSeizureList(string[] fields)
        {
            if (insuranceSeizureScriptPath == null) return;   // 連携系UIは非表示のため通常は到達しない

            var resolved = ResolveInsuranceInquiryFileForRow(fields,
                "差押リストに追加 — 照会ファイルの候補", "このファイルで追加", "追加するファイルを選んでください");
            if (resolved != null && app.LaunchSeizureListTool(insuranceSeizureScriptPath, "生命保険差押予定一覧 作成ツール", resolved.FullPath))
                app.RecordRecentInsurancePerson(fields);   // 実際に引き渡した行のみ記録する（起動中断・失敗時は記録しない）
        }
    }
}