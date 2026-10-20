// ==============================================================================
// paddock_insurance_ledger.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 生命保険タブ: 状態とUI要素の入れ物・UI要素の取得と配線・年度の選択・台帳の読み込み・
// 一覧の行と注記のチップの組み立て・注記列の幅。
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
    // 生命保険タブの状態と画面の一式。預金タブの同じ役割の関数と字面で対になる名前を持ち、
    // 預金の処理と共用する部品（読み込み中の覆い・空の状態・条件チップ・入力欄の補正・案内・
    // フォーカス枠など）は外側の app を通して呼ぶ（静的なものはそのまま）。外側から使うメンバー
    // だけ public（private の入れ子のクラスの public は外側からしか見えない）
    private partial class InsuranceTab
    {
        private readonly PaddockApp app;

        // --- 状態 ---

        public string insuranceLedgerFolder;                    // 台帳CSVの検出フォルダ（解決済み）
        private Dictionary<int, string> insuranceLedgerFiles;   // 年度 → 台帳CSVパス（null = まだ検出していない）
        private int insuranceLedgerEpoch;                       // 台帳の状態の世代（初回表示・F5・台帳フォルダの変更で進める。古い世代のワーカーの完了は合流しない）
        private bool insuranceLedgerLoaded;                     // 読み込み済み（生命保険タブを一度表示し、年度セレクタを組んで一覧を読んだ）。
                                                                // 最近見た人から開くときは検出だけ行うため、検出済みと読み込み済みを分けて持つ
        public List<InsuranceLedgerRecord> insuranceAllRecords = new List<InsuranceLedgerRecord>();   // 表示中の台帳の全行（年度指定=1本／すべての年度=全ファイルの結合・ソート済み）
        private int insuranceSkippedRowCount;                   // 列数が見出しと違うため読み飛ばした行数
        private List<int> insuranceLoadFailedYears = new List<int>();   // 「すべての年度」で読めなかった年度（フッターに出す）

        private BackgroundWorker insuranceLoadWorker;           // 台帳読み込みの非同期実行用
        private BackgroundWorker insuranceDetectWorker;         // 台帳ファイルの検出の非同期実行用（検出中は空状態を出さない判定に使う）
        private bool suppressInsuranceYearChange;               // 年度セレクタ初期化中の SelectionChanged 抑止

        // --- UI 要素 ---

        public DockPanel insuranceLedgerView;                   // 中央ペイン: 生命保険の一覧ビュー
        public ComboBox insuranceYearCombo;
        public ComboBox insuranceSearchKeyCombo;
        public TextBox insuranceSearchBox;
        private TextBlock insuranceSearchPlaceholder;
        private Button insuranceClearButton;
        private ListView insuranceLedgerList;
        private GridViewColumn insuranceNoteColumn;             // 注記列（幅を台帳のチップ列に合わせて更新する）

        // --- 詳細検索の UI 要素・状態 ---

        private Border insuranceAdvToggle;                      // 「詳細検索」トグル（常時 選択面・条件ありは件数バッジ）
        private Border insuranceAdvCountBadge;
        private TextBlock insuranceAdvCountText;
        private TextBlock insuranceAdvToggleCaret;              // パネルの開閉を示すキャレット（▼/▲）
        private Border insuranceAdvPanel;                       // 条件パネル（開閉は高さアニメーション・条件は閉じても持続）
        private WrapPanel insuranceFilterChipsRow;              // 条件チップ行（コードから動的生成）
        private TextBox insuranceAdvAddressBox;                 // 項目検索5欄
        private TextBox insuranceAdvNameBox;
        private TextBox insuranceAdvKanaBox;
        private TextBox insuranceAdvCompanyBox;
        private TextBox insuranceAdvPolicyBox;
        private TextBox insuranceAdvSurrenderMinBox;            // 解約返戻金の範囲
        private TextBox insuranceAdvSurrenderMaxBox;
        private Border insuranceAdvSurrenderMinBd;              // 入力枠（パース不能な入力をコーラル枠で提示）
        private Border insuranceAdvSurrenderMaxBd;
        private TextBox insuranceAdvBaseFromBox;                // 回答基準日の範囲
        private TextBox insuranceAdvBaseToBox;
        private Border insuranceAdvBaseFromBd;
        private Border insuranceAdvBaseToBd;
        private TextBox insuranceAdvMaturityFromBox;            // 満期年月日の範囲
        private TextBox insuranceAdvMaturityToBox;
        private Border insuranceAdvMaturityFromBd;
        private Border insuranceAdvMaturityToBd;
        private ToggleButton[] insuranceAdvNoteChips;           // 注記8種（INSURANCE_NOTE_NAMES と同順）
        private ToggleButton insuranceAdvClassNoContract;       // 行の分類「契約なし」= 契約なし行の表示スイッチ（既定OFF）
        private ToggleButton insuranceAdvClassUnanswered;       // 〃「電子回答不可」= 電子回答不可行の表示スイッチ（既定OFF）
        private WrapPanel insuranceAdvStaffPanel;               // 担当グループ（トグルは台帳から動的生成）
        private TextBlock insuranceAdvStaffEmptyText;           // 〃 担当の記録がない年度の案内
        private List<ToggleButton> insuranceAdvStaffToggles = new List<ToggleButton>();   // 生成済みの担当トグル（insuranceStaffClusters と同順）
        private List<StaffCluster> insuranceStaffClusters = new List<StaffCluster>();     // 現在の統合クラスタ
        private List<string> insuranceStaffRawValues = new List<string>();                // 選択中年度の担当職員名（distinct 生値・初出順）
        private InsuranceAdvancedFilter insuranceAdvFilter = new InsuranceAdvancedFilter();   // 現在の条件（正規化・パース済み）
        private int insuranceAdvConditionCount;                 // アクティブな条件チップの数（バッジ・フッター表記が共有）
        private bool suppressInsuranceAdvChange;                // 一括クリア中の変更イベント抑止
        private bool insuranceAdvPanelOpen;                     // パネルの開閉の論理状態（アニメーション中の基準）
        private DispatcherTimer insuranceFilterDebounceTimer;   // 文字入力の絞り込み反映を遅らせるタイマー（250ms）
        private Grid insuranceLoadingOverlay;
        private RotateTransform insuranceSpinnerRotation;       // 読み込みスピナーの回転（表示中のみ回す）
        private TextBlock insuranceEmptyStateText;
        public TextBlock insuranceStatusLeft;                   // フッター左（一時的な案内の表示先にもなる）
        private TextBlock insuranceStatusRight;

        public InsuranceTab(PaddockApp app)
        {
            this.app = app;
        }

        // 台帳を読み込み済み（生命保険タブを一度表示した）か。読み込む前は設定ビューの一致の開示や
        // 担当トグルの組み直しの対象にしない（最近見た人から開くための検出だけでは読み込み済みにしない）
        public bool IsLoaded
        {
            get { return insuranceLedgerLoaded; }
        }

        // ==============================================================
        // 画面の要素の取得・イベントの配線
        // ==============================================================

        public void FindInsuranceUiElements()
        {
            var window = app.window;
            insuranceLedgerView        = (DockPanel)window.FindName("InsuranceLedgerView");
            insuranceYearCombo         = (ComboBox)window.FindName("InsuranceYearCombo");
            insuranceSearchKeyCombo    = (ComboBox)window.FindName("InsuranceSearchKeyCombo");
            insuranceSearchBox         = (TextBox)window.FindName("InsuranceSearchBox");
            insuranceSearchPlaceholder = (TextBlock)window.FindName("InsuranceSearchPlaceholder");
            insuranceClearButton       = (Button)window.FindName("InsuranceClearButton");
            // 詳細検索
            insuranceAdvToggle         = (Border)window.FindName("InsuranceAdvToggle");
            insuranceAdvCountBadge     = (Border)window.FindName("InsuranceAdvCountBadge");
            insuranceAdvCountText      = (TextBlock)window.FindName("InsuranceAdvCountText");
            insuranceAdvToggleCaret    = (TextBlock)window.FindName("InsuranceAdvToggleCaret");
            insuranceAdvPanel          = (Border)window.FindName("InsuranceAdvPanel");
            insuranceFilterChipsRow    = (WrapPanel)window.FindName("InsuranceFilterChipsRow");
            insuranceAdvAddressBox     = (TextBox)window.FindName("InsuranceAdvAddressBox");
            insuranceAdvNameBox        = (TextBox)window.FindName("InsuranceAdvNameBox");
            insuranceAdvKanaBox        = (TextBox)window.FindName("InsuranceAdvKanaBox");
            insuranceAdvCompanyBox     = (TextBox)window.FindName("InsuranceAdvCompanyBox");
            insuranceAdvPolicyBox      = (TextBox)window.FindName("InsuranceAdvPolicyBox");
            insuranceAdvSurrenderMinBox = (TextBox)window.FindName("InsuranceAdvSurrenderMinBox");
            insuranceAdvSurrenderMaxBox = (TextBox)window.FindName("InsuranceAdvSurrenderMaxBox");
            insuranceAdvSurrenderMinBd = (Border)window.FindName("InsuranceAdvSurrenderMinBd");
            insuranceAdvSurrenderMaxBd = (Border)window.FindName("InsuranceAdvSurrenderMaxBd");
            insuranceAdvBaseFromBox    = (TextBox)window.FindName("InsuranceAdvBaseFromBox");
            insuranceAdvBaseToBox      = (TextBox)window.FindName("InsuranceAdvBaseToBox");
            insuranceAdvBaseFromBd     = (Border)window.FindName("InsuranceAdvBaseFromBd");
            insuranceAdvBaseToBd       = (Border)window.FindName("InsuranceAdvBaseToBd");
            insuranceAdvMaturityFromBox = (TextBox)window.FindName("InsuranceAdvMaturityFromBox");
            insuranceAdvMaturityToBox  = (TextBox)window.FindName("InsuranceAdvMaturityToBox");
            insuranceAdvMaturityFromBd = (Border)window.FindName("InsuranceAdvMaturityFromBd");
            insuranceAdvMaturityToBd   = (Border)window.FindName("InsuranceAdvMaturityToBd");
            insuranceAdvNoteChips = new ToggleButton[BusinessLogic.INSURANCE_NOTE_NAMES.Length];   // INSURANCE_NOTE_NAMES と同順
            for (int i = 0; i < insuranceAdvNoteChips.Length; i++)
                insuranceAdvNoteChips[i] = (ToggleButton)window.FindName("InsuranceAdvNote" + i);
            insuranceAdvClassNoContract = (ToggleButton)window.FindName("InsuranceAdvClassNoContract");
            insuranceAdvClassUnanswered = (ToggleButton)window.FindName("InsuranceAdvClassUnanswered");
            insuranceAdvStaffPanel     = (WrapPanel)window.FindName("InsuranceAdvStaffPanel");
            insuranceAdvStaffEmptyText = (TextBlock)window.FindName("InsuranceAdvStaffEmptyText");
            insuranceLedgerList        = (ListView)window.FindName("InsuranceLedgerList");
            insuranceNoteColumn        = (GridViewColumn)window.FindName("InsuranceNoteColumn");
            insuranceLoadingOverlay    = (Grid)window.FindName("InsuranceLoadingOverlay");
            insuranceSpinnerRotation   = (RotateTransform)((System.Windows.Shapes.Path)window.FindName("InsuranceSpinnerPath")).RenderTransform;
            insuranceEmptyStateText    = (TextBlock)window.FindName("InsuranceEmptyStateText");
            insuranceStatusLeft        = (TextBlock)window.FindName("InsuranceStatusLeft");
            insuranceStatusRight       = (TextBlock)window.FindName("InsuranceStatusRight");
            FindInsurancePanelUiElements();
        }

        // 生命保険タブの配線（年度・検索欄・詳細検索）。一覧の行の操作と詳細パネルの配線は
        // SetupInsurancePanelEvents が持つ
        public void SetupInsuranceEvents()
        {
            var window = app.window;

            // --- 年度セレクタ ---
            insuranceYearCombo.SelectionChanged += delegate
            {
                app.UpdateRailYearLabel();   // レールの年度表示は抑止中（初期化）も追従させる
                if (suppressInsuranceYearChange) return;
                if (insuranceYearCombo.SelectedItem == null) return;

                // 年度切替でパネルを閉じる（一覧が総入れ替えになるため、前の年度で
                // 開いた人物の文脈を右側に残さない。ビュー切替時と同じ扱い）
                if (isInsurancePanelOpen) CloseInsuranceDetailPanel();
                insurancePanelAddressNumberShown = null;   // 開き直したら組み直す（捨てる前に組んだパネルを使い回さない）
                LoadInsuranceLedger(GetSelectedYear(insuranceYearCombo));
            };

            // --- 検索ボックス ---
            // プレースホルダ・クリアボタンの表示切替は即時、絞り込みの反映はデバウンス
            insuranceSearchBox.TextChanged += delegate
            {
                bool hasText = insuranceSearchBox.Text.Length > 0;
                insuranceSearchPlaceholder.Visibility = hasText ? Visibility.Collapsed : Visibility.Visible;
                insuranceClearButton.Visibility = hasText ? Visibility.Visible : Visibility.Collapsed;
                RestartInsuranceFilterDebounce();
            };

            // ↓ で一覧の先頭行へ移る（保留中の絞り込みは先に反映する）
            insuranceSearchBox.PreviewKeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.Key != Key.Down || Keyboard.Modifiers != ModifierKeys.None) return;
                if (insuranceFilterDebounceTimer != null && insuranceFilterDebounceTimer.IsEnabled) OnInsuranceAdvancedFilterChanged();
                if (insuranceLedgerList.Items.Count == 0) return;
                insuranceLedgerList.SelectedIndex = 0;
                insuranceLedgerList.ScrollIntoView(insuranceLedgerList.Items[0]);
                app.FocusLater(delegate { FocusInsuranceLedgerSelection(); });
                e.Handled = true;
            };

            // --- クリアボタン（検索ボックス内 ×。解除は待たずに即時反映する）---
            insuranceClearButton.Click += delegate { insuranceSearchBox.Text = ""; insuranceSearchBox.Focus(); OnInsuranceAdvancedFilterChanged(); };

            // --- 検索キーの切替（入力中のみ再絞り込み）---
            insuranceSearchKeyCombo.SelectionChanged += delegate
            {
                if (insuranceSearchBox.Text.Trim().Length > 0) ApplyInsuranceFilter();
            };

            // --- 詳細検索 ---
            insuranceAdvToggle.MouseLeftButtonUp += delegate { ToggleInsuranceAdvancedPanel(); };
            app.MakeKeyActivatable(insuranceAdvToggle, ToggleInsuranceAdvancedPanel);
            // ホバーは条件チップと同じ色遣い（トグルは常時 選択面のため）
            insuranceAdvToggle.MouseEnter += delegate { insuranceAdvToggle.Background = BrushChipVioletBg; };
            insuranceAdvToggle.MouseLeave += delegate { insuranceAdvToggle.Background = BrushSelectedCardBg; };

            // 条件の変化は OnInsuranceAdvancedFilterChanged に集約する（文字入力はデバウンス経由・トグルは即時）
            insuranceAdvAddressBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            insuranceAdvNameBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            insuranceAdvKanaBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            insuranceAdvCompanyBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            insuranceAdvPolicyBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            insuranceAdvSurrenderMinBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            insuranceAdvSurrenderMaxBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            insuranceAdvBaseFromBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            insuranceAdvBaseToBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            insuranceAdvMaturityFromBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            insuranceAdvMaturityToBox.TextChanged += delegate { RestartInsuranceFilterDebounce(); };
            foreach (var chip in insuranceAdvNoteChips)
            {
                chip.Checked += delegate { OnInsuranceAdvancedFilterChanged(); };
                chip.Unchecked += delegate { OnInsuranceAdvancedFilterChanged(); };
            }
            insuranceAdvClassNoContract.Checked += delegate { OnInsuranceAdvancedFilterChanged(); };
            insuranceAdvClassNoContract.Unchecked += delegate { OnInsuranceAdvancedFilterChanged(); };
            insuranceAdvClassUnanswered.Checked += delegate { OnInsuranceAdvancedFilterChanged(); };
            insuranceAdvClassUnanswered.Unchecked += delegate { OnInsuranceAdvancedFilterChanged(); };

            // 「条件をすべてクリア」とプレースホルダ（配線後は参照しないためローカルで引く）
            var insuranceAdvClearLink = (TextBlock)window.FindName("InsuranceAdvClearLink");
            insuranceAdvClearLink.MouseLeftButtonUp += delegate { ClearInsuranceAdvancedFilter(); };
            app.MakeKeyActivatable(insuranceAdvClearLink, ClearInsuranceAdvancedFilter);
            insuranceAdvClearLink.MouseEnter += delegate { insuranceAdvClearLink.Foreground = BrushTextMain; };
            insuranceAdvClearLink.MouseLeave += delegate { insuranceAdvClearLink.Foreground = BrushTextSub; };
            WirePlaceholder(insuranceAdvAddressBox, (TextBlock)window.FindName("InsuranceAdvAddressPh"));
            WirePlaceholder(insuranceAdvNameBox, (TextBlock)window.FindName("InsuranceAdvNamePh"));
            WirePlaceholder(insuranceAdvKanaBox, (TextBlock)window.FindName("InsuranceAdvKanaPh"));
            WirePlaceholder(insuranceAdvCompanyBox, (TextBlock)window.FindName("InsuranceAdvCompanyPh"));
            WirePlaceholder(insuranceAdvPolicyBox, (TextBlock)window.FindName("InsuranceAdvPolicyPh"));
            WirePlaceholder(insuranceAdvSurrenderMinBox, (TextBlock)window.FindName("InsuranceAdvSurrenderMinPh"));
            WirePlaceholder(insuranceAdvSurrenderMaxBox, (TextBlock)window.FindName("InsuranceAdvSurrenderMaxPh"));
            WirePlaceholder(insuranceAdvBaseFromBox, (TextBlock)window.FindName("InsuranceAdvBaseFromPh"));
            WirePlaceholder(insuranceAdvBaseToBox, (TextBlock)window.FindName("InsuranceAdvBaseToPh"));
            WirePlaceholder(insuranceAdvMaturityFromBox, (TextBlock)window.FindName("InsuranceAdvMaturityFromPh"));
            WirePlaceholder(insuranceAdvMaturityToBox, (TextBlock)window.FindName("InsuranceAdvMaturityToPh"));

            // 解約返戻金欄の3桁区切り整形と、解約返戻金・日付欄のフォーカス喪失時の補正
            WireAmountFormatting(insuranceAdvSurrenderMinBox, OnInsuranceAdvancedFilterChanged);
            WireAmountFormatting(insuranceAdvSurrenderMaxBox, OnInsuranceAdvancedFilterChanged);
            WireDateNormalize(insuranceAdvBaseFromBox, OnInsuranceAdvancedFilterChanged);
            WireDateNormalize(insuranceAdvBaseToBox, OnInsuranceAdvancedFilterChanged);
            WireDateNormalize(insuranceAdvMaturityFromBox, OnInsuranceAdvancedFilterChanged);
            WireDateNormalize(insuranceAdvMaturityToBox, OnInsuranceAdvancedFilterChanged);

            // --- 日付欄のカレンダー（ウィンドウ直下の共有ポップアップ）---
            foreach (var dateBox in new[] { insuranceAdvBaseFromBox, insuranceAdvBaseToBox, insuranceAdvMaturityFromBox, insuranceAdvMaturityToBox })
                app.WireAdvDateBox(dateBox, OnInsuranceAdvancedFilterChanged);

            SetupInsurancePanelEvents();
        }

        // 一覧の選択行の行コンテナへフォーカスする（選択がない・行が生成されていなければ false）
        public bool FocusInsuranceLedgerSelection()
        {
            var item = insuranceLedgerList.ItemContainerGenerator.ContainerFromItem(insuranceLedgerList.SelectedItem) as ListViewItem;
            if (item == null) return false;
            item.Focus();
            return true;
        }

        // ==============================================================
        // 台帳ファイルの検出・年度セレクタ
        // ==============================================================

        // 生命保険タブを初めて表示したときに台帳を検出して読む（起動時は検出もしない。
        // 2度目以降の表示では何もしない——年度・条件・一覧はタブが保持している。
        // 最近見た人から開くための検出だけが済んでいても、ここで検出からやり直して読む）
        public void EnsureInsuranceLedgerLoaded()
        {
            if (insuranceLedgerLoaded) return;
            insuranceLedgerLoaded = true;
            DetectInsuranceLedgerFilesThenPopulate(null);
        }

        // 最近見た人から生命保険の詳細パネルを開くとき、台帳を読む前なら台帳ファイルの検出だけを行う
        //（年度セレクタも一覧も作らない。その人の全年度の行はパネルの年度横断の走査で集める）
        public void EnsureInsuranceLedgerDetected()
        {
            if (insuranceLedgerFiles == null)
                insuranceLedgerFiles = DetectLedgerFiles(insuranceLedgerFolder, INSURANCE_LEDGER_PATTERN, InsuranceLedgerYearRegex);
        }

        // 設定の保存で台帳フォルダが変わったとき。読み込み済みなら検出からやり直して年度セレクタを組み直す
        //（選択は「すべての年度」に戻る）。読み込む前は、検出だけ済んでいれば新しいフォルダで検出し直す（一覧は読まない）
        public void OnInsuranceLedgerFolderChanged(string folder)
        {
            insuranceLedgerFolder = folder;
            insuranceOtherYearCache.Clear();   // 人物キャッシュは別フォルダの内容と混ざらないよう捨てる
            if (isInsurancePanelOpen) CloseInsuranceDetailPanel();
            insurancePanelAddressNumberShown = null;   // 開き直したら組み直す（捨てる前に組んだパネルを使い回さない）
            if (!insuranceLedgerLoaded)
            {
                insuranceLedgerEpoch++;   // 走っていた走査の完了は合流しない（別フォルダの行を混ぜない）
                // 検出だけ済んでいれば新しいフォルダで検出し直す（走査中の合流が台帳ファイルの一覧を引くため null には戻さない）
                if (insuranceLedgerFiles != null)
                    insuranceLedgerFiles = DetectLedgerFiles(insuranceLedgerFolder, INSURANCE_LEDGER_PATTERN, InsuranceLedgerYearRegex);
                return;
            }
            DetectInsuranceLedgerFilesThenPopulate(null);
        }

        // 台帳ファイルの検出をワーカーで行い、終わってから年度セレクタを組んで読む（初回表示・F5・台帳フォルダの
        // 変更の入口。共有フォルダが遅い・切断中でもウィンドウは止まらない）。入口で世代を進めるため、走っていた
        // 古い検出・読み込み・走査の完了は合流しない。検出に失敗したら台帳 0 本として整えてから文言を差し替える
        private void DetectInsuranceLedgerFilesThenPopulate(string preferredItem)
        {
            var epoch = ++insuranceLedgerEpoch;
            var folder = insuranceLedgerFolder;
            insuranceLedgerFiles = new Dictionary<int, string>();   // 検出が終わるまで台帳 0 本（検出中に開いた人物の走査が旧フォルダのファイルを読まない）
            insuranceYearCombo.IsEnabled = false;
            insuranceEmptyStateText.Visibility = Visibility.Collapsed;
            app.ShowLoadingOverlay(insuranceLoadingOverlay, insuranceSpinnerRotation);

            insuranceDetectWorker = new BackgroundWorker();
            insuranceDetectWorker.DoWork += delegate(object s, DoWorkEventArgs e)
            {
                e.Result = DetectLedgerFiles(folder, INSURANCE_LEDGER_PATTERN, InsuranceLedgerYearRegex);
            };
            insuranceDetectWorker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs e)
            {
                if (epoch != insuranceLedgerEpoch) return;   // そのあとに F5・台帳フォルダの変更があった古い検出は合流しない
                if (e.Error != null)
                {
                    app.WriteErrorLog("台帳の検出（生命保険）", e.Error);
                    PopulateInsuranceYearSelector(null);
                    ShowEmptyState(insuranceEmptyStateText, "台帳フォルダを読めませんでした:\n" + e.Error.Message);
                    app.SetStatusText(insuranceStatusLeft, "読み込みエラー");
                    return;
                }
                insuranceLedgerFiles = (Dictionary<int, string>)e.Result;
                PopulateInsuranceYearSelector(preferredItem);
                if (isInsurancePanelOpen)   // 検出の前に開いた人物の他年度を改めて走査する
                {
                    insuranceCrossYearTimer.Stop();
                    insuranceCrossYearTimer.Start();
                }
            };
            insuranceDetectWorker.RunWorkerAsync();
        }

        // 年度セレクタを構築して読み込む。既定は「すべての年度」
        //（年度の帰属は登録日基準のため年度替わり直後は当年度の台帳がほぼ空になる。
        //  既定を全年度にして、探すたびに前年度へ切り替えずに済むようにする）。
        // preferredItem に選択中だった項目を渡すと、残っていればそれを選び直す（再読み込み用）
        private void PopulateInsuranceYearSelector(string preferredItem)
        {
            suppressInsuranceYearChange = true;
            insuranceYearCombo.Items.Clear();

            var years = insuranceLedgerFiles.Keys.OrderByDescending(y => y).ToList();
            if (years.Count > 0) insuranceYearCombo.Items.Add(ALL_YEARS_LABEL);
            foreach (var year in years)
                insuranceYearCombo.Items.Add(BusinessLogic.FormatYearKey(year));

            if (years.Count == 0)
            {
                suppressInsuranceYearChange = false;
                insuranceYearCombo.IsEnabled = false;
                app.UpdateRailYearLabel();     // 台帳なしの表示（—）に更新する
                ResetInsuranceLedgerToEmpty(); // 再読み込みで台帳が消えた場合に前の行を残さない
                ShowEmptyState(insuranceEmptyStateText, "生命保険の台帳CSVが見つかりません。\nフォルダ: " + insuranceLedgerFolder +
                               "\n\nJUDGE（judge.ps1）の生命保険のプロファイルの ledgerFolder 設定と\n" +
                               "paddock_config.json の insuranceLedgerFolder が同じフォルダを指しているか確認してください。");
                app.SetStatusText(insuranceStatusLeft, "台帳なし");
                app.HideLoadingOverlay(insuranceLoadingOverlay, insuranceSpinnerRotation);   // 検出のワーカーが出したオーバーレイを消す
                return;
            }

            insuranceYearCombo.IsEnabled = true;
            var index = preferredItem != null ? insuranceYearCombo.Items.IndexOf(preferredItem) : -1;
            insuranceYearCombo.SelectedIndex = index >= 0 ? index : 0;
            suppressInsuranceYearChange = false;

            LoadInsuranceLedger(GetSelectedYear(insuranceYearCombo));
        }

        // 台帳を読み直す（F5）。ファイルの検出からやり直して年度セレクタを組み直し、選んで
        // いた項目は残っていれば維持する（消えていれば「すべての年度」）。検索欄・詳細検索の
        // 条件は入力欄が持っているため保持され、読み込み完了時の絞り込みで再適用される。
        // 読み込み中に押しても検出からやり直す（古い読み込みの完了は世代で捨てる）
        public void ReloadInsuranceLedger()
        {
            var current = insuranceYearCombo.SelectedItem != null ? insuranceYearCombo.SelectedItem.ToString() : null;
            insuranceOtherYearCache.Clear();   // 他年度の追記も拾えるよう人物単位のキャッシュは捨てる
            if (isInsurancePanelOpen) CloseInsuranceDetailPanel();
            insurancePanelAddressNumberShown = null;   // 開き直したら組み直す（捨てる前に組んだパネルを使い回さない）
            DetectInsuranceLedgerFilesThenPopulate(current);
        }

        // 一覧の状態を「台帳0行」へ戻す（読み込みエラー時・台帳ファイルが1本もないとき）。
        // 担当トグルを空にし、解除された担当条件をチップ・バッジへ反映する
        //（一覧の表示はこの後の空状態・エラー文言が担うため絞り込みは行わない）
        private void ResetInsuranceLedgerToEmpty()
        {
            insuranceAllRecords = new List<InsuranceLedgerRecord>();
            insuranceSkippedRowCount = 0;
            insuranceLoadFailedYears = new List<int>();
            UpdateInsuranceNoteColumnWidth();   // 台帳なし=注記列を下限幅へ戻す
            insuranceStaffRawValues = new List<string>();
            RebuildInsuranceStaffToggles();
            if (app.insuranceSettingsRows.TargetBox != null) app.UpdateSeizureTargetNote(app.insuranceSettingsRows);   // 台帳0行としての一致開示に更新
            ReadInsuranceAdvancedInputs();
            RebuildInsuranceFilterChips();
            UpdateInsuranceAdvToggleVisual();
            insuranceLedgerList.ItemsSource = null;
        }

        // ==============================================================
        // 台帳の読み込み（BackgroundWorker）
        // ==============================================================

        // 読み込み結果（DoWork → RunWorkerCompleted の受け渡し用）
        private class LoadInsuranceResult
        {
            public List<InsuranceLedgerRecord> Records;
            public int SkippedRows;
            public string Label;                  // フッター右の表示（年度指定=ファイル名／すべての年度=範囲と本数）
            public string Error;
            public List<string> StaffRawValues;   // 担当職員名の distinct 生値（初出順・担当トグルの列挙用）
            public List<int> FailedYears;         // 読めなかった年度（すべての年度で一部のファイルが読めなかった場合）
        }

        // 台帳CSVを非同期で読み込む。year=null は「すべての年度」（台帳フォルダの全ファイルを結合）。
        // 各ファイルは1行目を見出しとして列名で読み、台帳の並びへ組み直して持つ
        private void LoadInsuranceLedger(int? year)
        {
            // 読む対象を年度の降順に並べる（結合後は安定ソートのため、同一基準日内の並びは
            // 新しい年度→古い年度・各ファイル内は追記順になる）
            var targets = new List<KeyValuePair<int, string>>();
            if (year == null)
            {
                foreach (var y in insuranceLedgerFiles.Keys.OrderByDescending(k => k))
                    targets.Add(new KeyValuePair<int, string>(y, insuranceLedgerFiles[y]));
            }
            else
            {
                string path;
                if (!insuranceLedgerFiles.TryGetValue(year.Value, out path)) return;
                targets.Add(new KeyValuePair<int, string>(year.Value, path));
            }
            if (targets.Count == 0) return;
            var epoch = insuranceLedgerEpoch;   // この読み込みを始めた世代（F5・台帳フォルダの変更で進む。古い世代の完了は合流しない）

            insuranceYearCombo.IsEnabled = false;
            insuranceEmptyStateText.Visibility = Visibility.Collapsed;
            app.ShowLoadingOverlay(insuranceLoadingOverlay, insuranceSpinnerRotation);

            insuranceLoadWorker = new BackgroundWorker();
            insuranceLoadWorker.DoWork += delegate(object s, DoWorkEventArgs e)
            {
                var result = new LoadInsuranceResult { FailedYears = new List<int>() };
                result.Label = year != null
                    ? System.IO.Path.GetFileName(targets[0].Value)
                    : "生命保険調査台帳（" + (targets.Count == 1
                        ? BusinessLogic.FormatYearKey(targets[0].Key)
                        : targets[targets.Count - 1].Key + "〜" + BusinessLogic.FormatYearKey(targets[0].Key)) +
                      "・" + targets.Count + "本）";

                var records = new List<InsuranceLedgerRecord>();
                int skipped = 0;
                var staffSeen = new HashSet<string>();
                var staffRaws = new List<string>();
                var pool = new Dictionary<string, string>();   // 同じ値の文字列を 1 本に共有する（この読み込みの間だけ）

                foreach (var target in targets)
                {
                    try
                    {
                        // ファイル全体を読み込んで1回走査でパースする（途中で失敗したファイルの行は載せない）
                        var text = ReadLedgerText(target.Value);
                        var rawRecords = BusinessLogic.ParseCsvText(text);
                        var fileRecords = new List<InsuranceLedgerRecord>(rawRecords.Count);
                        int fileSkipped = 0;

                        // 1行目の見出しから台帳の並びへの対応表を作る。見出しがない・必須の列が欠ける
                        // ファイルはその年度を読めない扱いにする。並びが台帳の並びそのものなら組み直さない
                        List<string> missing;
                        var header = rawRecords.Count > 0 ? rawRecords[0] : new string[0];
                        var map = BusinessLogic.BuildLedgerColumnMap(header, BusinessLogic.INSURANCE_LEDGER_COLUMNS, BusinessLogic.INSURANCE_REQUIRED_COLUMNS, out missing);
                        if (map == null)
                        {
                            throw new InvalidDataException(missing == null
                                ? "見出しがありません"
                                : "見出しに必須の列がありません: " + string.Join("・", missing.ToArray()));
                        }
                        bool identity = BusinessLogic.IsIdentityColumnMap(map, header.Length, BusinessLogic.INSURANCE_LEDGER_COLUMN_COUNT);

                        for (int i = 1; i < rawRecords.Count; i++)
                        {
                            var fields = rawRecords[i];

                            // 列数が見出しと違う行は読み飛ばす（件数のみ記録して落とさない）
                            if (fields.Length != header.Length) { fileSkipped++; continue; }
                            if (!identity) fields = BusinessLogic.ReorderLedgerFields(fields, map);

                            BusinessLogic.InternFields(fields, pool);
                            fileRecords.Add(BuildInsuranceRecord(fields));

                            // 担当トグルの列挙用に生値の distinct を集める（空欄は対象外）
                            var staffRaw = fields[BusinessLogic.INSURANCE_COL_STAFF].Trim();
                            if (staffRaw != "" && staffSeen.Add(staffRaw)) staffRaws.Add(staffRaw);
                        }
                        records.AddRange(fileRecords);
                        skipped += fileSkipped;
                    }
                    catch (Exception ex)
                    {
                        // 年度単位で控えて続行する。読めた分だけ表示し、全て読めなければ読み込みエラー
                        result.FailedYears.Add(target.Key);
                        result.Error = ex.Message;
                    }
                }

                if (result.FailedYears.Count < targets.Count)
                {
                    result.Error = null;

                    // 既定ソート: 回答基準日の降順（回答基準日が空の行は登録日で代替）
                    // SortKey は正準形（yyyy/MM/dd）のため文字列比較で時系列順になる
                    // ※ 安定ソート（OrderBy）を使い、同一基準日内はファイルの追記順
                    //   （＝同一人物の契約が連続する並び）を保持する
                    result.Records = records.OrderByDescending(r => r.SortKey, StringComparer.Ordinal).ToList();
                    result.SkippedRows = skipped;
                    result.StaffRawValues = staffRaws;
                }
                e.Result = result;
            };
            insuranceLoadWorker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs e)
            {
                if (epoch != insuranceLedgerEpoch) return;   // 新しい読み込みが始まっている（オーバーレイと年度セレクタはそちらが面倒を見る）
                insuranceYearCombo.IsEnabled = true;

                // DoWork の外で起きた例外（e.Result は null）は読み込みの失敗と同じ経路に流す
                if (e.Error != null) app.WriteErrorLog("台帳の読み込み（生命保険）", e.Error);
                var result = e.Error == null ? (LoadInsuranceResult)e.Result : new LoadInsuranceResult { Error = e.Error.Message };
                if (result.Error != null)
                {
                    ResetInsuranceLedgerToEmpty();
                    ShowEmptyState(insuranceEmptyStateText, "台帳の読み込みに失敗しました:\n" + result.Error);
                    app.SetStatusText(insuranceStatusLeft, "読み込みエラー");
                    insuranceStatusRight.Text = result.Label;
                    app.HideLoadingOverlay(insuranceLoadingOverlay, insuranceSpinnerRotation);
                    return;
                }

                insuranceAllRecords = result.Records;
                insuranceSkippedRowCount = result.SkippedRows;
                insuranceLoadFailedYears = result.FailedYears;
                insuranceStatusRight.Text = result.Label;

                // 注記列の幅を新しい台帳の最長チップ列に合わせる
                UpdateInsuranceNoteColumnWidth();

                // 担当トグルを新しい年度の台帳から組み直してから絞り込む（選択中の
                // 担当が新年度に存在しない場合の条件解除をチップ・バッジへ反映する）
                insuranceStaffRawValues = result.StaffRawValues;
                RebuildInsuranceStaffToggles();
                OnInsuranceAdvancedFilterChanged();
                if (app.insuranceSettingsRows.TargetBox != null) app.UpdateSeizureTargetNote(app.insuranceSettingsRows);   // 設定の一致開示を新年度で組み直す

                // 一覧への反映が終わってから消灯し、反映後の一覧をフェードイン
                //（反映が終わるまでオーバーレイが覆い続ける）
                app.HideLoadingOverlay(insuranceLoadingOverlay, insuranceSpinnerRotation);
                insuranceLedgerList.BeginAnimation(UIElement.OpacityProperty,
                    new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(100))));
            };
            insuranceLoadWorker.RunWorkerAsync();
        }

        // ==============================================================
        // 行データの構築（台帳の並びの行 → 表示用レコード）
        // ==============================================================

        public InsuranceLedgerRecord BuildInsuranceRecord(string[] fields)
        {
            var record = new InsuranceLedgerRecord();
            record.RawFields = fields;

            var baseDate = fields[BusinessLogic.INSURANCE_COL_BASE_DATE];

            // 行の分類: 契約有無の値で 1=契約あり／2=電子回答不可／それ以外（0・空欄・未知の値）=契約なし
            var hasContract = fields[BusinessLogic.INSURANCE_COL_HAS_CONTRACT].Trim();
            record.IsUnanswered = hasContract == "2";
            record.IsNoContract = !record.IsUnanswered && hasContract != "1";
            record.MutedPillText = record.IsUnanswered ? "電子回答不可" : record.IsNoContract ? "契約なし" : null;

            // 回答基準日は正準形（yyyy/MM/dd）に変換して表示する。台帳の値は7桁和暦・
            // 西暦のどちらでもよい（形式非依存）。契約なし行は空欄
            record.DisplayBaseDate = BusinessLogic.FormatLedgerDate(baseDate, app.eraMapping);

            // ソートキー: 回答基準日が空（契約なし行・短いシートのない電子回答不可行）は登録日で代替。
            // 正準形の Ordinal 比較が時系列順になるため、台帳内で日付形式が混在しても並びが崩れない
            //（変換不能な値は生値のままで、並び位置は保証しない）
            record.SortKey = BusinessLogic.FormatLedgerDate(
                baseDate != "" ? baseDate : fields[BusinessLogic.INSURANCE_COL_REG_DATE], app.eraMapping);

            // 文字列セルは空のとき null で保持する（空文字列を ToolTip にバインドすると
            // 空のツールチップが表示されるため。null なら非表示になる）
            record.AddressNumber   = fields[BusinessLogic.INSURANCE_COL_ADDRESS_NUM];
            record.Name            = NullIfEmpty(fields[BusinessLogic.INSURANCE_COL_NAME]);
            record.Kana            = NullIfEmpty(fields[BusinessLogic.INSURANCE_COL_KANA]);
            record.CompanyName     = NullIfEmpty(fields[BusinessLogic.INSURANCE_COL_COMPANY]);
            record.ContractType    = NullIfEmpty(fields[BusinessLogic.INSURANCE_COL_CONTRACT_TYPE]);
            record.PolicyNumber    = NullIfEmpty(fields[BusinessLogic.INSURANCE_COL_POLICY_NUM]);
            record.ContractStatus  = NullIfEmpty(fields[BusinessLogic.INSURANCE_COL_CONTRACT_STATUS]);
            record.MaturityDate    = fields[BusinessLogic.INSURANCE_COL_MATURITY];

            // 検索用の正規化キー（半角/全角・空白の揺れを吸収した部分一致用）
            record.SearchAddressNumber = BusinessLogic.NormalizeForSearch(fields[BusinessLogic.INSURANCE_COL_ADDRESS_NUM]);
            record.SearchName          = BusinessLogic.NormalizeForSearch(fields[BusinessLogic.INSURANCE_COL_NAME]);
            record.SearchKana          = BusinessLogic.NormalizeForSearch(fields[BusinessLogic.INSURANCE_COL_KANA]);
            record.SearchCompany       = BusinessLogic.NormalizeForSearch(fields[BusinessLogic.INSURANCE_COL_COMPANY]);
            record.SearchPolicyNumber  = BusinessLogic.NormalizeForSearch(fields[BusinessLogic.INSURANCE_COL_POLICY_NUM]);

            // 解約返戻金: 数値として読めれば台帳の値（カンマ付き）をそのまま、読めない生値は薄文字色で出す
            var surrender = fields[BusinessLogic.INSURANCE_COL_SURRENDER];
            record.DisplaySurrenderValue = surrender;
            record.SurrenderValue = BusinessLogic.ParseLedgerAmount(surrender);
            record.IsSurrenderValueText = surrender != "" && record.SurrenderValue == null;

            // 詳細検索用の事前計算。回答基準日は正準形（yyyy/MM/dd）に変換済みの表示値を、
            // 満期年月日は台帳の値を日付化する（パース不能・空欄は null=当該条件の設定時に対象外）
            record.BaseDateValue = BusinessLogic.ParseDisplayDate(record.DisplayBaseDate);
            record.MaturityValue = BusinessLogic.ParseInsuranceDate(record.MaturityDate);
            record.NoteFlags = BusinessLogic.BuildInsuranceNoteFlags(fields);

            // 担当職員名。空欄は null（=担当条件の対象外）
            record.StaffKey = NullIfEmpty(BusinessLogic.NormalizeForSearch(fields[BusinessLogic.INSURANCE_COL_STAFF]));

            BuildInsuranceNoteChips(record, fields);

            // 人物メモの有無（起動時に読み込んだ全ユーザー分の件数表）
            int noteCount;
            if (app.noteCountByAddress.TryGetValue(fields[BusinessLogic.INSURANCE_COL_ADDRESS_NUM], out noteCount) && noteCount > 0)
            {
                record.HasNotes = true;
                record.NotesToolTip = "メモ " + noteCount + "件";
            }
            return record;
        }

        // 注記ごとのチップの配色（INSURANCE_NOTE_NAMES と同順）: 明色面の背景/文字と暗色ツールチップ用の背景/文字
        private static readonly Brush[] InsuranceNoteChipBg = { BrushChipCoralBg, BrushChipCoralBg, BrushChipAmberBg, BrushChipNeutralBg, BrushChipTealBg, BrushChipTealBg, BrushChipNeutralBg, BrushChipNeutralBg };
        private static readonly Brush[] InsuranceNoteChipFg = { BrushChipCoralFg, BrushChipCoralFg, BrushChipAmberFg, BrushChipNeutralFg, BrushChipTealFg, BrushChipTealFg, BrushChipNeutralFg, BrushChipNeutralFg };
        private static readonly Brush[] InsuranceNoteTipBg  = { BrushTipCoralBg, BrushTipCoralBg, BrushTipAmberBg, BrushTipNeutralBg, BrushTipTealBg, BrushTipTealBg, BrushTipNeutralBg, BrushTipNeutralBg };
        private static readonly Brush[] InsuranceNoteTipFg  = { BrushTipCoralFg, BrushTipCoralFg, BrushTipAmberFg, BrushTipNeutralFg, BrushTipTealFg, BrushTipTealFg, BrushTipNeutralFg, BrushTipNeutralFg };

        // 注記チップ（第1層・立っている注記の全部）とツールチップ行（第2層）を構築する。
        // 一覧の行と詳細パネルの契約カードで共用する。
        // 走査順は固定: 差押 → 質権 → 貸付金 → 未払保険料 → 配当 → 年金 → 解約済 → 保険金支払済
        private static List<ChipInfo> BuildInsuranceNoteChipInfos(bool[] flags, string[] fields, out List<TooltipLineInfo> lines)
        {
            var chips = new List<ChipInfo>();
            lines = new List<TooltipLineInfo>();
            for (int i = 0; i < flags.Length; i++)
            {
                if (!flags[i]) continue;
                var name = BusinessLogic.INSURANCE_NOTE_NAMES[i];
                chips.Add(new ChipInfo { Text = name, Bg = InsuranceNoteChipBg[i], Fg = InsuranceNoteChipFg[i] });
                lines.Add(new TooltipLineInfo
                {
                    ChipText = name, ChipBg = InsuranceNoteTipBg[i], ChipFg = InsuranceNoteTipFg[i],
                    Content = BusinessLogic.BuildInsuranceNoteDetail(fields, i), ContentFont = FontUi   // 原文は UI フォント（項目名と値の並び）
                });
            }
            return chips;
        }

        // 一覧の行の注記チップとツールチップ行を組み立てて行に持たせる
        private static void BuildInsuranceNoteChips(InsuranceLedgerRecord record, string[] fields)
        {
            List<TooltipLineInfo> lines;
            var chips = BuildInsuranceNoteChipInfos(record.NoteFlags, fields, out lines);

            // --- 第1層はチップ最大3個+超過件数「+N」---
            if (chips.Count > 3)
            {
                var overflow = chips.Count - 3;
                chips = chips.Take(3).ToList();
                chips.Add(new ChipInfo { Text = "+" + overflow, Bg = BrushChipNeutralBg, Fg = BrushChipNeutralFg });
            }

            record.Chips = chips;
            record.TooltipLines = lines;
            record.HasNote = lines.Count > 0;
        }

        // 注記列の幅を、読み込み済みの全行のチップ列が収まる幅に更新する（預金の検出列と同じ方式。
        // 年度の読み込みごとに最長行を実測して合わせ、絞り込みでは変えない）
        private void UpdateInsuranceNoteColumnWidth()
        {
            if (insuranceNoteColumn == null) return;

            // 加算値は XAML の寸法を写した定数（チップの左右の Padding と右の Margin・セルの左右の Margin・描画誤差の余裕）
            double maxRow = 0;
            foreach (var record in insuranceAllRecords)
            {
                double row = 0;
                if (record.IsMutedRow)
                    row = app.MeasureChipText(record.MutedPillText) + CHIP_PADDING_H;
                else if (record.Chips != null)
                    foreach (var chip in record.Chips)
                        row += app.MeasureChipText(chip.Text) + CHIP_PADDING_H + CHIP_MARGIN_R;
                if (row > maxRow) maxRow = row;
            }

            var width = maxRow + CELL_MARGIN_H + CHIP_COLUMN_SLACK;
            insuranceNoteColumn.Width = width > 260 ? width : 260;   // 下限は 260（チップの短い年度でも列幅を狭めない）
        }

        // メモの追加・削除を一覧の印に反映する（同じ宛名番号の行の印だけを変更通知で描き直す。一覧のコンテナは
        // 作り直さないのでスクロール位置は動かない）
        public void RefreshInsuranceNoteMarks(string addressNumber, int noteCount)
        {
            foreach (var record in insuranceAllRecords)
            {
                if (record.AddressNumber != addressNumber) continue;
                record.HasNotes = noteCount > 0;
                record.NotesToolTip = noteCount > 0 ? "メモ " + noteCount + "件" : null;
            }
        }
    }
}