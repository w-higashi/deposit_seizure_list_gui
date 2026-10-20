// ==============================================================================
// paddock_app.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// アプリケーション本体（定数・ブラシ・フィールド・入れ子の型）と、起動（検証用の起動を
// 含む）・エラーログ・ユーザー履歴の配置・ユーザー設定・設定ファイルと元号と口座種別の補正表の読み込み。
// PaddockApp の静的フィールドはこのファイルにまとめている（初期化子がほかの静的
// フィールドを読むことがあり、partial class の別のファイルに分けると初期化の順序が
// 定まらないため）。
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

// ==============================================================
// アプリケーション本体
// ==============================================================

public partial class PaddockApp : Application
{
    // --- 定数 ---

    private const string CONFIG_FILE      = "paddock_config.json";
    private const string ERA_MAPPING_FILE = "era_mapping.json";
    private const string NOTICE_FILE      = "update_notice.json";
    private const string DEPOSIT_LEDGER_PATTERN   = "財産調査台帳_*.csv";
    private const string INSURANCE_LEDGER_PATTERN = "生命保険調査台帳_*.csv";
    private const string ALL_YEARS_LABEL  = "すべての年度";   // 年度セレクタの先頭項目（全年度の結合・既定）
    private const string ALL_YEARS_RAIL   = "ALL";            // 同・レールの短縮表示
    private const string KIND_INSURANCE   = "insurance";      // 履歴ファイル・差押登録ログの種別の値（生命保険。預金の行は種別を持たない）
    private const string NO_NAME_TEXT     = "（氏名情報なし）"; // 人物ヘッダーの氏名の退行表示
    private const string APP_VERSION      = "1.7.2";   // 設定ビューの表示用（告知の version とは独立）

    private const double DETAIL_PANEL_WIDTH = 440;   // 詳細パネルの展開幅（px。明細セクションの表が収まる幅）
    private const double SIDEBAR_RAIL_WIDTH = 48;     // サイドバーのレール幅（px）
    private const double SIDEBAR_FLYOUT_WIDTH = 208;  // サイドバーの展開幅（px）
    private const int NOTES_VISIBLE_MAX = 3;           // メモの折りたたみ時の表示件数
    private const int NOTE_UNDO_SECONDS = 10;          // メモの削除を「元に戻す」で戻せる秒数
    private const long ERROR_LOG_MAX_BYTES = 1024 * 1024;   // エラーログの世代交代の上限（超えていたら .1 へ改名してから書く）

    // 一覧の XAML の寸法を写した定数（列幅・氏名セルの計算に使う。XAML の値と合うことを検査で固定する）
    private const double CHIP_PADDING_H     = 16;    // チップの Padding='8,1,8,2' の左右
    private const double CHIP_MARGIN_R      = 4;     // チップの Margin='0,0,4,0' の右
    private const double CELL_MARGIN_H      = 16;    // セルの Grid の Margin='12,0,4,0' の左右
    private const double CHIP_COLUMN_SLACK  = 6;     // 検出列・注記列の幅に足す描画誤差の余裕
    private const double NAME_COLUMN_WIDTH  = 150;   // 氏名列の幅（一覧・最近見た人・差押登録で同じ）
    private const double NAME_CELL_PADDING_R = 8;    // コードで組む氏名セルの右余白
    private const double NOTE_MARK_WIDTH    = 19;    // 氏名の右の印 ✎ の幅（Margin 5＋文字 14）

    // 台帳ファイル名から年度を抽出する（預金・生命保険）
    private static readonly Regex DepositLedgerYearRegex   = new Regex(@"財産調査台帳_(\d{4})年度\.csv$", RegexOptions.Compiled);
    private static readonly Regex InsuranceLedgerYearRegex = new Regex(@"生命保険調査台帳_(\d{4})年度\.csv$", RegexOptions.Compiled);

    // ダブルクリック判定時間（OSのユーザー設定値・ミリ秒）を取得する
    // シングルクリックでのパネル表示をこの時間だけ保留し、ダブルクリックと判別する
    [DllImport("user32.dll")]
    private static extern uint GetDoubleClickTime();

    // 照会ファイルを開いた後に Excel のウィンドウを前面へ出す
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    // --- デザイントークン（PADDOCK デザイン言語・全て Freeze 済み）---

    private static SolidColorBrush Frozen(string hex)
    {
        var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        brush.Freeze();
        return brush;
    }

    private static SolidColorBrush FrozenArgb(byte a, byte r, byte g, byte b)
    {
        var brush = new SolidColorBrush(Color.FromArgb(a, r, g, b));
        brush.Freeze();
        return brush;
    }

    // チップ配色（明色面・カテゴリ別）: 背景/文字
    // ※ チップはコードから動的生成するためブラシを保持する。固定UIの色は XAML 側にリテラルで持つ
    public static readonly Brush BrushChipCoralBg   = Frozen("#FAECE7");   // 反対債権
    public static readonly Brush BrushChipCoralFg   = Frozen("#993C1D");
    public static readonly Brush BrushChipVioletBg  = Frozen("#EEEDFB");   // 給与
    public static readonly Brush BrushChipVioletFg  = Frozen("#4C4AC0");
    public static readonly Brush BrushChipTealBg    = Frozen("#E1F5EE");   // 年金
    public static readonly Brush BrushChipTealFg    = Frozen("#0F6E56");
    public static readonly Brush BrushChipAmberBg   = Frozen("#FAEEDA");   // 保険
    public static readonly Brush BrushChipAmberFg   = Frozen("#854F0B");
    public static readonly Brush BrushChipNeutralBg = Frozen("#F4F4F6");   // 定期入金・定期出金・超過件数
    public static readonly Brush BrushChipNeutralFg = Frozen("#6F6E77");

    // チップ配色（ツールチップの暗色サーフェス用・カテゴリ別）: 背景は各色の半透明トーン
    public static readonly Brush BrushTipCoralBg    = FrozenArgb(41, 240, 153, 123);
    public static readonly Brush BrushTipCoralFg    = Frozen("#F0997B");
    public static readonly Brush BrushTipVioletBg   = FrozenArgb(41, 175, 169, 236);
    public static readonly Brush BrushTipVioletFg   = Frozen("#AFA9EC");
    public static readonly Brush BrushTipTealBg     = FrozenArgb(41, 93, 202, 165);
    public static readonly Brush BrushTipTealFg     = Frozen("#5DCAA5");
    public static readonly Brush BrushTipAmberBg    = FrozenArgb(41, 250, 199, 117);
    public static readonly Brush BrushTipAmberFg    = Frozen("#FAC775");
    public static readonly Brush BrushTipNeutralBg  = FrozenArgb(31, 244, 244, 246);
    public static readonly Brush BrushTipNeutralFg  = Frozen("#B9B8C0");

    // 詳細パネル用の基本色（パネルの中身はコードから動的生成するためブラシを保持する）
    public static readonly Brush BrushAccent    = Frozen("#5B5BD6");   // アクセント（アイリス）・レールのドット
    public static readonly Brush BrushRailLine  = Frozen("#E9E8F2");   // タイムラインのレール線
    public static readonly Brush BrushTextMain  = Frozen("#1A1523");   // テキスト（主）
    public static readonly Brush BrushTextSub   = Frozen("#6F6E77");   // テキスト（副）
    public static readonly Brush BrushTextMuted = Frozen("#9B9AA3");   // テキスト（淡）

    // ミュート行の文字と数値として読めない生値（一覧の XAML の DataTrigger と同じ値）
    public static readonly Brush BrushTextMutedRow = Frozen("#B9B8C0");

    // 口座行カード・候補ダイアログの選択状態と面の色
    public static readonly Brush BrushSelectedCardBg = Frozen("#F6F5FE");   // 選択中の背景（一覧の選択行と同色）
    public static readonly Brush BrushSelectedLine   = Frozen("#DCDAF6");   // 選択面の枠線（条件チップ・詳細検索トグル）
    public static readonly Brush BrushIrisPanelBg    = Frozen("#F9F8FD");   // 淡アイリス面（詳細検索パネルと同色・メモの入力エリア）
    public static readonly Brush BrushSurfaceAlt     = Frozen("#FCFCFD");   // フッター帯・セクションヘッダー
    public static readonly Brush BrushLine           = Frozen("#ECEBF1");   // 罫線（強）
    public static readonly Brush BrushLineWeak       = Frozen("#F1F0F4");   // 罫線（弱・行間）
    public static readonly Brush BrushRowHover       = Frozen("#FAFAFC");   // 行ホバー背景（コード生成の行用）
    public static readonly Brush BrushTipText        = Frozen("#F4F3F8");   // 暗色ツールチップの原文テキスト

    // 「明細を見る」の推移グラフ
    public static readonly Brush BrushDetailWithdraw     = Frozen("#D97A5A");   // 出金の細線
    public static readonly Brush BrushDetailOtherDeposit = Frozen("#C9C8D0");   // 検出に当たらない入金の棒
    public static readonly Brush BrushTipGapText         = Frozen("#F2B9A6");   // ツールチップの差額

    private static readonly FontFamily FontMono = new FontFamily("Consolas");
    private static readonly FontFamily FontUi   = new FontFamily("Meiryo UI");
    private static readonly Typeface ChipTypeface = new Typeface(FontUi,
        FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);   // 検出チップの幅実測用

    // --- フィールド ---

    private Window window;
    private string exeDir;
    private AppConfig config;
    private Dictionary<int, EraEntry> eraMapping;         // 元号マッピング

    private string depositLedgerFolder;                   // 預金の台帳CSVの検出フォルダ（解決済み）
    private string depositSeizureScriptPath;              // 預金差押予定一覧 作成ツールのパス（解決済み。null=連携UI非表示）
    private List<SeizureTargetPattern> depositSeizureTargetPatterns;   // 差押リストの対象金融機関（null=制御なし）
    private Dictionary<string, bool> depositSeizureTargetMatchCache;   // 金融機関キー→対象か（年度横断のカードにも効かせるメモ化）
    private Dictionary<int, string> depositLedgerFiles = new Dictionary<int, string>();   // 年度 → 台帳CSVパス（検出が終わるまで空）
    private int depositLedgerEpoch;                       // 台帳の状態の世代（起動・F5・台帳フォルダの変更で進める。古い世代のワーカーの完了は合流しない）
    private List<DepositLedgerRecord> depositAllRecords = new List<DepositLedgerRecord>();   // 表示中の台帳の全行（年度指定=1本／すべての年度=全ファイルの結合・ソート済み）
    private int depositSkippedRowCount;                   // 列数不足で読み飛ばした行数
    private List<int> depositLoadFailedYears = new List<int>();   // 「すべての年度」で読めなかった年度（フッターと脚注に出す）

    private BackgroundWorker depositLoadWorker;           // 台帳読み込みの非同期実行用
    private BackgroundWorker depositDetectWorker;         // 台帳ファイルの検出の非同期実行用（検出中は空状態を出さない判定に使う）
    private bool suppressDepositYearChange;               // 年度セレクタ初期化中の SelectionChanged 抑止

    // 生命保険タブ（状態と画面の一式。起動時に作り、台帳は生命保険タブを初めて表示したときに読む）
    private InsuranceTab insuranceTab;
    private bool isInsuranceTabCurrent;                   // 表示中（ほかのビューでは直前に表示していた）台帳のタブが生命保険か（起動直後は預金）

    // --- UI 要素 ---

    private Border depositNavLedger;                  // ナビ「預金」
    private TextBlock depositNavLedgerText;
    private TextBlock depositNavLedgerIcon;
    private Border insuranceNavLedger;                // ナビ「生命保険」
    private TextBlock insuranceNavLedgerText;
    private TextBlock insuranceNavLedgerIcon;
    private Border navRecent;                         // ナビ「最近見た人」
    private TextBlock navRecentText;
    private TextBlock navRecentIcon;
    private Border navSeizure;                        // ナビ「差押登録」
    private TextBlock navSeizureText;
    private TextBlock navSeizureIcon;
    private Border navSettings;                       // ナビ「設定」
    private TextBlock navSettingsText;
    private TextBlock navSettingsIcon;

    // --- サイドバー折りたたみの UI 要素・状態 ---
    private Border sidebarRail;                       // アイコンレール（既定表示・48px）
    private Border sidebarFlyout;                     // 展開パネル（レールに被さるオーバーレイ・208px）
    private TextBlock railKindText;                   // レール下端の種別の表示（「預金」「生保」。年度の短縮表示の上）
    private TextBlock railYearText;                   // レール下端の年度短縮表示（'26 形式）
    private Border depositRailLedger;                 // レールのナビ5項目（展開パネル側と活性表示を同期）
    private Border insuranceRailLedger;
    private Border railRecent;
    private Border railSeizure;
    private Border railSettings;
    private TextBlock depositRailLedgerIcon;
    private TextBlock insuranceRailLedgerIcon;
    private TextBlock railRecentIcon;
    private TextBlock railSeizureIcon;
    private TextBlock railSettingsIcon;
    private DispatcherTimer sidebarExpandTimer;       // レールのホバー0.2秒で展開
    private DispatcherTimer sidebarCollapseTimer;     // 展開パネルからの離脱0.2秒で収納
    private bool sidebarExpanded;                     // 展開の論理状態（アニメーション中の基準）
    private DockPanel depositLedgerView;              // 中央ペイン: 預金の一覧ビュー（生命保険のビューは insuranceTab が持つ）
    private DockPanel recentView;                     // 中央ペイン: 最近見た人ビュー
    private DockPanel seizureView;                    // 中央ペイン: 差押登録ビュー
    private DockPanel settingsView;                   // 中央ペイン: 設定ビュー

    private StackPanel depositYearPanel;              // 展開パネルの年度の欄（預金用・生命保険用。表示中のタブの分だけ見せる）
    private StackPanel insuranceYearPanel;
    private ComboBox depositYearCombo;
    private ComboBox depositSearchKeyCombo;
    private TextBox depositSearchBox;
    private TextBlock depositSearchPlaceholder;
    private Button depositClearButton;
    private ListView depositLedgerList;
    private GridViewColumn depositDetectColumn;       // 検出列（幅を台帳のチップ列に合わせて更新する）

    // --- 詳細検索の UI 要素・状態 ---

    private Border depositAdvToggle;                      // 「詳細検索」トグル（常時 選択面・条件ありは件数バッジ）
    private Border depositAdvCountBadge;
    private TextBlock depositAdvCountText;
    private TextBlock depositAdvToggleCaret;              // パネルの開閉を示すキャレット（▼/▲）
    private Border depositAdvPanel;                       // 条件パネル（開閉は高さアニメーション・条件は閉じても持続）
    private WrapPanel depositFilterChipsRow;              // 条件チップ行（コードから動的生成）
    private TextBox depositAdvAddressBox;                 // 項目検索4欄
    private TextBox depositAdvNameBox;
    private TextBox depositAdvKanaBox;
    private TextBox depositAdvInstitutionBox;
    private TextBox depositAdvBalanceMinBox;              // 残高範囲
    private TextBox depositAdvBalanceMaxBox;
    private Border depositAdvBalanceMinBd;                // 入力枠（パース不能な入力をコーラル枠で提示）
    private Border depositAdvBalanceMaxBd;
    private TextBox depositAdvBaseFromBox;                // 回答基準日の範囲
    private TextBox depositAdvBaseToBox;
    private Border depositAdvBaseFromBd;
    private Border depositAdvBaseToBd;
    private TextBox depositAdvTxFromBox;                  // 最終取引日の範囲
    private TextBox depositAdvTxToBox;
    private Border depositAdvTxFromBd;
    private Border depositAdvTxToBd;
    private Popup advDatePopup;                           // 日付欄の共有カレンダー（すべての日付欄で使い回す）
    private System.Windows.Controls.Calendar advDateCalendar;   // System.Globalization.Calendar と衝突するため完全修飾
    private TextBlock advDateTodayLink;                   // フッター「今日」リンク
    private Run advDateTodayValue;                        // 「今日」リンクの日付部分（開くたびに更新）
    private TextBox advDateTarget;                        // カレンダーの対象欄（開いた欄）
    private Action advDateApply;                          // 日付を確定したときの反映（開いた欄の絞り込み）
    private bool advDateSuppress;                         // プログラム側の選択設定でイベントを発生させない
    private ToggleButton[] depositAdvDetectChips;         // 検出6種（DETECTION_NAMES と同順）
    private ToggleButton depositAdvClassRep;              // 行の分類「代表口座」（登録推奨=1）
    private ToggleButton depositAdvClassSpecial;          // 〃「特殊口座のみ」（登録推奨=2）
    private ToggleButton depositAdvClassNoAccount;        // 〃「口座なし」= 口座なし行の表示スイッチ（既定OFF）
    private WrapPanel depositAdvStaffPanel;               // 担当グループ（トグルは台帳から動的生成）
    private TextBlock depositAdvStaffEmptyText;           // 〃 担当の記録がない年度の案内
    private List<ToggleButton> depositAdvStaffToggles = new List<ToggleButton>();   // 生成済みの担当トグル（depositStaffClusters と同順）
    private List<StaffCluster> depositStaffClusters = new List<StaffCluster>();     // 現在の統合クラスタ
    private List<string> depositStaffRawValues = new List<string>();                // 選択中年度の担当職員名（distinct 生値・初出順）
    private DepositAdvancedFilter depositAdvFilter = new DepositAdvancedFilter();   // 現在の条件（正規化・パース済み）
    private int depositAdvConditionCount;                 // アクティブな条件チップの数（バッジ・フッター表記が共有）
    private bool suppressDepositAdvChange;                // 一括クリア中の変更イベント抑止
    private bool depositAdvPanelOpen;                     // パネルの開閉の論理状態（アニメーション中の基準）
    private DispatcherTimer depositFilterDebounceTimer;   // 文字入力の絞り込み反映を遅らせるタイマー（250ms）
    private Grid depositLoadingOverlay;
    private RotateTransform depositSpinnerRotation;       // 読み込みスピナーの回転（表示中のみ回す）
    private TextBlock depositEmptyStateText;
    private TextBlock depositStatusLeft;
    private TextBlock depositStatusRight;

    // --- 詳細パネルの UI 要素 ---

    private Border depositDetailPanel;
    private TextBlock depositPanelAvatarText;
    private TextBlock depositPanelName;
    private Run depositPanelAddressNumber;                   // 宛名番号とカナは1つの TextBlock の Run 2本（ベースラインを揃える）
    private Run depositPanelKana;
    private TextBlock depositPanelAddress;
    private Border depositPanelCopyChip;                     // 「コピーしました」チップ（1.5秒で消える）
    private TextBlock depositPanelStatRounds;
    private TextBlock depositPanelStatAccounts;
    private TextBlock depositPanelStatFirst;
    private ScrollViewer depositPanelScroll;
    private readonly NotesHost depositNotes = new NotesHost();   // 預金の詳細パネルのメモの欄
    private StackPanel depositTimelinePanel;
    private TextBlock depositPanelFootnote;
    private System.Windows.Shapes.Path depositPanelMiniSpinner;   // 年度横断走査のミニスピナー（System.IO.Path と衝突するため完全修飾）
    private RotateTransform depositPanelSpinnerRotation;
    private Button depositPanelCloseButton;
    private Button depositPanelOpenFileButton;
    private Button depositPanelSeizureButton;

    // --- 差押登録ビューの UI 要素・状態 ---

    private Border seizureCard;                       // 年度セクションを収める白カード
    private StackPanel seizureSectionsPanel;          // 年度セクションの並び（コードから動的生成）
    private TextBlock seizureEmptyText;               // ログ0件時の案内
    private TextBlock seizureFooterText;              // フッター（件数・ログパス）
    private List<SeizureLogEntry> seizureLog = new List<SeizureLogEntry>();   // 読み込み済みログ（全件）

    // --- 最近見た人ビュー ---
    private Border recentCard;                        // 年度セクションを収める白カード
    private StackPanel recentSectionsPanel;           // 年度セクションの並び（コードから動的生成）
    private TextBlock recentSubtitleText;             // 見出し下の説明（config の recentMax を反映）
    private TextBlock recentEmptyText;                // 履歴0件時の案内
    private TextBlock recentFooterText;               // フッター（件数・履歴ファイルのパス）
    private List<RecentPersonEntry> recentPeople = new List<RecentPersonEntry>();   // 読み込み済み履歴（全件）
    private bool persistRecent = true;                // 履歴の永続化が有効か（読み書き失敗時に false に切り替わる）
    private int recentSkippedCount;                   // 起動時に形式不正で読み飛ばした履歴エントリの件数（フッターに出す）
    private Border recentSelectedBorder;              // 選択中の履歴行の Border（表示切替用）

    // --- 設定ビューの UI 要素・状態 ---
    // 中身（6枚のカード）はビューへ入るたびにコードから再構築する。ただし未保存の
    // 変更がある間は再構築しない（ナビで離れて戻っても編集状態と保存バーが残る）

    private StackPanel settingsPanel;                     // 6枚のカードを収める本体
    private Border settingsSaveBar;                       // 保存バー（変更あり時のみ表示）
    private TextBlock settingsChangeSummary;              // 保存バーの変更内容の要約
    private Button settingsSaveButton;
    private Button settingsDiscardButton;
    private TextBlock settingsFooterLeft;                 // フッター（設定ファイルのパス / 保存結果の一時表示）
    private TextBlock settingsFooterRight;                // フッター（バージョン表示）
    private TextBox settingsRecentMaxBox;                 // 最近見た人の保持件数の入力欄
    private TextBox settingsFuzzyBox;                     // 担当の化け置換文字の入力欄
    private Border settingsFuzzyChip;                     // 〃 採用文字数のチップ
    private TextBox settingsTypeMappingBox;               // 口座種別の補正表のパスの入力欄
    private Border settingsTypeMappingChip;               // 〃 読めるかどうかのチップ
    private string settingsTypeMappingChecked;            // 〃 チップを最後に確かめた入力値（同じ値なら読み直さない）
    private bool settingsDirty;                           // 未保存の変更あり
    private bool settingsSaveBarShown;                    // 保存バーの論理表示状態（アニメーション中の基準）
    private bool suppressSettingsDirty;                   // 構築中の TextChanged で変更扱いにしない
    private DispatcherTimer settingsChipTimer;            // 存在チェックのデバウンス用（250ms・ネットワークドライブ対策）
    private DispatcherTimer settingsFooterTimer;          // 保存メッセージの一時表示の復元用
    private string settingsFooterNormal = "";             // 通常時のフッター左側の文言
    private AppConfig settingsShownConfig;                // ビューが示している値（設定ファイルの今の値。読めなければメモリ上の config）
    private string settingsConfigText;                    // ビューを組んだ時点のファイルの本文（保存前の照合用。null = 読めなかった＝照合しない）

    // 保存済みの共通の設定値のスナップショット（変更検出と「変更を破棄」の戻し先。種別ごとの分は入れ物が持つ）
    private string settingsSnapRecentMax = "";
    private string settingsSnapFuzzy = "";
    private string settingsSnapTypeMapping = "";

    // 照会結果フォルダの編集行1つ分の UI 要素の対応
    private class SettingsFolderRow
    {
        public TextBox YearBox;                           // 年度キー
        public TextBox PathBox;                           // フォルダパス
        public Border Chip;                               // 存在チェックのチップ
    }

    // 設定ビューの種別（預金・生命保険）ごとの4行（台帳フォルダ／照会結果フォルダ／差押リスト連携／
    // 差押リストの対象）の部品と控えの入れ物。カードの組み立て・存在チェック・変更検出・保存を
    // 同じ関数で両方の種別に当てるため、行の見出しや文言の種別で変わる語もここから引く
    private class SettingsKindRows
    {
        public string Kind;                               // 種別の名前（「預金」「生命保険」。カードの題・変更の要約）
        public string LedgerKey;                          // config のキー名（行の見出しの小字・保存）
        public string FoldersKey;
        public string ScriptKey;
        public string TargetKey;
        public string LedgerPattern;                      // 台帳の検出パターン（注記の台帳ファイル数）
        public string ToolName;                           // 連携先のツールの呼び名（参照ダイアログの題）
        public string TargetNoun;                         // 対象の語（「金融機関」「保険会社」。行の見出し・開示の文）
        public TextBox LedgerBox;                         // 台帳フォルダの入力欄
        public TextBlock LedgerPlaceholder;               // 〃 未設定時のプレースホルダ
        public TextBlock LedgerNote;                      // 〃 現在の検出先の注記
        public StackPanel FolderRowsPanel;                // 照会結果フォルダの行の入れ物
        public List<SettingsFolderRow> FolderRows = new List<SettingsFolderRow>();
        public TextBox ScriptBox;                         // 差押リスト連携先の入力欄
        public Border ScriptChip;                         // 〃 存在チェック/再起動待ちのチップ
        public TextBox TargetBox;                         // 差押リストの対象の入力欄
        public Border TargetChip;                         // 〃 一致種類数のチップ
        public Border TargetDisclosure;                   // 〃 一致の開示（淡アイリス面）
        public TextBlock TargetMatchText;                 // 〃 一致した名前の1行列挙
        public TextBlock TargetMissText;                  // 〃 効いていない設定の1行
        public bool ScriptRestartPending;                 // 連携先を保存済み・再起動待ち（次回起動まで表示）

        // 保存済みの設定値のスナップショット（変更検出と「変更を破棄」の戻し先）
        public string SnapLedger = "";
        public string SnapScript = "";
        public string SnapTarget = "";
        public List<string[]> SnapFolders = new List<string[]>();   // {年度, パス} の組
    }
    private readonly SettingsKindRows depositSettingsRows = new SettingsKindRows
    {
        Kind = "預金", LedgerKey = "depositLedgerFolder", FoldersKey = "depositSearchFolders",
        ScriptKey = "depositSeizureListScript", TargetKey = "depositSeizureTargetBanks",
        LedgerPattern = DEPOSIT_LEDGER_PATTERN, ToolName = "預金差押予定一覧 作成ツール", TargetNoun = "金融機関"
    };
    private readonly SettingsKindRows insuranceSettingsRows = new SettingsKindRows
    {
        Kind = "生命保険", LedgerKey = "insuranceLedgerFolder", FoldersKey = "insuranceSearchFolders",
        ScriptKey = "insuranceSeizureListScript", TargetKey = "insuranceSeizureTargetCompanies",
        LedgerPattern = INSURANCE_LEDGER_PATTERN, ToolName = "生命保険差押予定一覧 作成ツール", TargetNoun = "保険会社"
    };

    // --- アップデート告知の UI 要素・状態 ---

    private Grid noticeOverlay;                       // 告知オーバーレイ（暗転+カード）
    private StackPanel noticeContent;                 // 告知内容の動的描画エリア
    private Button noticeCloseButton;
    private string lastSeenVersion = "";              // ユーザーが最後に確認した告知バージョン
    private string pendingNoticeVersion;              // 表示中の告知バージョン（閉じたときに既読として保存）

    // --- 詳細パネルの状態 ---

    private bool isDepositPanelOpen;                         // パネルの開閉状態
    private string depositPanelAddressNumberShown;           // パネルに表示中の宛名番号（非同期結果の適用判定・メモの欄の組み直しにも使用）
    private DispatcherTimer depositClickTimer;               // シングル/ダブルクリックの判別用
    private DepositLedgerRecord depositPendingClickRecord;   // クリック判別待ちの行
    private int depositLastClickCount;                       // 直近の MouseDown のクリック回数
    private int depositLastMouseDownTick;                    // 直近の MouseDown の時刻（Environment.TickCount）
    private DispatcherTimer depositCrossYearTimer;           // 年度横断読み込みのデバウンス用
    private BackgroundWorker depositPersonWorker;            // 他年度台帳の走査用
    private bool depositCrossYearPending;                    // 走査中に人物が切り替わった場合の再走査フラグ

    // 口座行カードの選択状態（選択中カード = 「照会ファイルを開く」の対象）
    // 一覧の行選択と連動し、パネル内のカードクリックでも切り替わる
    private string[] depositSelectedCardFields;              // 選択中カードの行（生の全列）
    private string[] depositPendingCardFields;               // 選択が未成立のときに復元を試みる行（SwitchDepositPanelPerson が受けた focusFields）
    private Border depositSelectedCardBorder;                // 選択中カードの Border（表示切替用）

    // 人物ヘッダーの退行値（登録日が最新の非空値・RenderDepositPersonPanel で計算する）。
    // 表示は選択カードの行の値を優先し、空の項目はこの値へ退行する
    private string depositPanelFallbackName;
    private string depositPanelFallbackKana;
    private string depositPanelFallbackAddress;
    private List<KeyValuePair<string[], Border>> depositPanelCards =
        new List<KeyValuePair<string[], Border>>();          // 描画済みカードの行→Border 対応

    // --- 明細を見る ---

    private const int DETAIL_CACHE_MAX_FILES = 30;   // 骨格・セルをセッション内に持つ人物ファイルの上限（古いものから捨てる）

    // 推移グラフの寸法。パネルのスパークラインと明細ウィンドウの大グラフは同じ描画関数を
    // この寸法だけ変えて使う
    private class DetailChartStyle
    {
        public double ChartHeight;                     // 折れ線の領域の高さ（月ラベル込み）
        public double BarHeight;                       // 入出金の棒の領域の高さ
        public bool FillHeight;                        // true: 受け皿の高さに合わせて上の2つを決める（棒は 18%）
        public bool GapRings;                          // 回答の残高と積み上げが合わない点にコーラルの輪を描く（小さなグラフでは点と重なって読めないため描かない）
        public int MarkerLimit;                        // 取引数がこれを超えると点を打たない
        public double TotalHeight { get { return ChartHeight + BarHeight + 6; } }
    }
    private static readonly DetailChartStyle DetailSparklineStyle = new DetailChartStyle { ChartHeight = 104, BarHeight = 28, MarkerLimit = 60 };
    private static readonly DetailChartStyle DetailLargeChartStyle = new DetailChartStyle { FillHeight = true, GapRings = true, MarkerLimit = int.MaxValue };

    private readonly Dictionary<string, Window> detailWindows = new Dictionary<string, Window>();   // 行キー→開いている明細ウィンドウ
    private Size lastDetailWindowSize = new Size(960, 640);                                          // 最後に閉じたウィンドウの大きさ（セッション内）

    private readonly Dictionary<Border, DepositCardExpansion> depositCardExpansions = new Dictionary<Border, DepositCardExpansion>();   // カード→展開部
    private DepositCardExpansion depositExpandedCard;   // 展開中のカード（同時に1枚）
    private string[] depositExpandedCardFields;         // 展開中のカードの行（パネル再構築時の復元用）
    private string[] depositPendingExpandFields;        // 選択の成立と同時に自動展開する行

    private BackgroundWorker detailWorker;              // 照会ファイルの読み取り（1本・後着優先）
    private readonly List<DetailRequest> detailQueue = new List<DetailRequest>();
    private string detailRunningKey;                    // ワーカーが処理中の行キー
    private AccountTypeMapping detailTypeMapping;       // 口座種別の補正表（読めなければ null＝補正しない）

    // セッション内キャッシュ（UI スレッドだけが書く。ワーカーには要求時の写しを渡す。人物ファイルは DETAIL_CACHE_MAX_FILES を超えると古いものから捨てる）
    private readonly Dictionary<string, DetailResult> detailResultCache = new Dictionary<string, DetailResult>();        // 行キー→結果
    private readonly Dictionary<string, string> detailResolvedPaths = new Dictionary<string, string>();                // 年度|宛名番号|期待シート→特定できたファイル
    private readonly Dictionary<string, InquiryWorkbook> detailWorkbookCache = new Dictionary<string, InquiryWorkbook>(StringComparer.OrdinalIgnoreCase);   // 人物ファイル→骨格
    private readonly Dictionary<string, DateTime> detailWorkbookStamps = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);         // 同・更新日時
    private readonly Dictionary<string, SheetCells> detailSheetCache = new Dictionary<string, SheetCells>(StringComparer.OrdinalIgnoreCase);         // 人物ファイル|シート→セル
    private readonly Dictionary<string, string> detailChosenPaths = new Dictionary<string, string>();                  // 行キー→候補から選んだファイル
    private readonly List<string> detailCacheOrder = new List<string>();                                                   // 人物ファイルのパス・最後に使った順（末尾が最新）

    // 口座カードの展開部（検出行+明細セクション）
    private class DepositCardExpansion
    {
        public Border Card;
        public string[] Fields;
        public StackPanel Detail;                       // 展開部全体（開閉の対象）
        public TextBlock Chevron;
        public StackPanel Section;                      // 明細セクションの中身（明細の対象外なら null）
        public StackPanel CategoryLines;                // 取引カテゴリの検出行（明細を読めたら集約表に置き換える。なければ null）
        public UIElement Rule;                          // 展開部の行（検出・反対債権）と明細セクションの間の破線（それらの行がなければ null）
        public StackPanel Frame;                        // 明細セクションの枠（破線＋⧉ の行＋中身）
        public Grid Toolbar;                            // ⧉ の行（左に最終取引日・右に ⧉）
        public Decorator OpenSlot;                      // ⧉ の受け皿（明細を読めたときだけ ⧉ を入れる）
        public Action OpenWindow;                       // ⧉ の処理（⧉ が出ている間だけ。Ctrl+M も同じ処理を呼ぶ）
        public bool HasLastTx;                          // ⧉ の行に最終取引日を出しているか
        public bool Requested;                          // 明細の読み取りを要求済み
        public RotateTransform Spinner;                 // 読み込み中表示の回転（差し替え時に止める）
    }

    // 明細の読み取り要求（UI スレッドで作り、ワーカーが埋めて返す）
    private class DetailRequest
    {
        public string[] Fields;
        public string RowKey;
        public string ExpectedSheet;
        public string ResolveKey;                      // 特定結果のキャッシュキー（年度|宛名番号|期待シート）
        public string ChosenPath;                      // 候補ダイアログで選ばれたファイル（なければ null）
        public AccountTypeMapping TypeMapping;         // 要求時の口座種別の補正表（結果が古い表のものか見分ける）

        // UI 側キャッシュの写し（ワーカーは骨格の更新日時を確かめてから使う）
        public Dictionary<string, string> ResolvedPaths;
        public Dictionary<string, InquiryWorkbook> Workbooks;
        public Dictionary<string, DateTime> Stamps;
        public Dictionary<string, SheetCells> Sheets;

        // ワーカーの結果（UI 側キャッシュの更新に使う）
        public DetailResult Result;
        public string ResolvedPath;                    // 特定できたファイル（候補ダイアログで選んだものは含めない）
        public InquiryWorkbook Workbook;
        public DateTime Stamp;
        public SheetCells Cells;
        public Exception Error;
    }

    // 口座1件の明細の結果（読み取れた・特定できなかった のどちらか）
    private class DetailResult
    {
        public DetailAccountData Data;                 // 読み取れた場合
        public List<InquiryFileCandidate> Candidates;  // 特定できなかった場合の候補（0件・複数）
        public string YearKey;                         // 候補を集めた年度（全年度なら null。候補ダイアログの「全年度から探す」の有無に使う）
    }

    private DispatcherTimer statusTimer;               // 一時ステータス表示の復元用
    private TextBlock statusTransientTarget;           // 一時表示中の欄（表示中のビューのフッター左）
    private string statusRestoreText;                  // 一時表示の前の文言（一時表示を終えたら戻す）

    // 他年度の走査結果キャッシュ: 年度 → 宛名番号 → その人物の行（生の全列）
    // 台帳ファイル全体はキャッシュせず、実際に閲覧した人物の行のみを保持する
    // （過去年度ファイルは追記されないため、セッション中のキャッシュは陳腐化しない）
    private Dictionary<int, Dictionary<string, List<string[]>>> depositOtherYearCache =
        new Dictionary<int, Dictionary<string, List<string[]>>>();

    // 人物メモ: 宛名番号 → メモの一覧。起動時に
    // 全ユーザーの person_notes（history/{ユーザー名}/）をまとめて読み込んでキャッシュする
    //（パネルを開くたびのフォルダ走査を避ける — depositOtherYearCache と同じ思想。
    //  他ユーザーの追記はセッション中には反映されず、次回起動で最新になる）
    private Dictionary<string, List<PersonNoteEntry>> myPersonNotes;      // 自分のメモ（null = 未読み込み）
    private Dictionary<string, List<PersonNoteEntry>> otherPersonNotes;   // 他ユーザーのメモ（読み取り専用）
    private Dictionary<string, int> noteCountByAddress = new Dictionary<string, int>();   // 宛名番号→メモ件数（全ユーザー分。一覧の印に使う）
    private bool persistNotes = true;                  // 自分のメモファイルが読めない・書けない場合は false（保存を止め追加・編集も無効化）

    // メモの削除を元に戻す控え（全体で 1 つ。削除した 1 件を NOTE_UNDO_SECONDS の間だけメモリに持ち、
    // 追加・書き直し・別の削除・同じ欄での別の人物の表示・秒切れで捨てる。パネルを閉じるだけでは捨てない）
    private NotesHost noteUndoHost;                    // 「元に戻す」の行を出している欄（null = 控えなし）
    private string noteUndoAddress;                    // 削除したメモの宛名番号
    private PersonNoteEntry noteUndoNote;              // 削除したメモ（同じ作成日時のまま戻す）
    private int noteUndoIndex;                         // 削除前の並び（作成日時の降順）での位置＝行を出す位置
    private int noteUndoSeconds;                       // 残り秒
    private Border noteUndoRow;                        // 「元に戻す」の行（一覧を描き直すたびに差し替わる。控えを捨てるときに外す）
    private TextBlock noteUndoSecondsText;             // 残り秒の表示（同上）
    private DispatcherTimer noteUndoTimer;             // 1 秒ごとに残り秒を減らす

    // ==============================================================
    // エントリポイント
    // ==============================================================

    [STAThread]
    public static void Main()
    {
        var app = new PaddockApp();
        app.Startup += app.OnAppStartup;
        app.Run();
    }

    private void OnAppStartup(object sender, StartupEventArgs e)
    {
        exeDir = AppDomain.CurrentDomain.BaseDirectory;
        InitUserHistoryPaths();   // ユーザー履歴の配置（history/{ユーザー名}/）— エラーログの置き場所もここ

        // 検証用の起動: --dump-detail <照会ファイル> <シート名> <口座番号> <口座種別> <出力JSON> [検出] [金融機関名 補正表]
        // 明細の読み取りと計算の結果を JSON に書き出して終了する（画面は出さない。独立オラクルとの
        // 突合用）。検出は台帳の 年金・給与・保険・定期入金・定期出金 の5列を "||" で連結した文字列。
        // 金融機関名と口座種別の補正表のパスを渡すと口座種別を補正しても照合する（渡さなければ補正しない）
        if ((e.Args.Length == 6 || e.Args.Length == 7 || e.Args.Length == 9) && e.Args[0] == "--dump-detail")
        {
            DumpDetailForVerification(e.Args[1], e.Args[2], e.Args[3], e.Args[4], e.Args[5], e.Args.Length >= 7 ? e.Args[6] : "",
                e.Args.Length == 9 ? e.Args[7] : "", e.Args.Length == 9 ? e.Args[8] : null);
            Shutdown();
            return;
        }

        // 起動後の例外はログに残して続行する（黙って落ちない。原因はログで追う）
        DispatcherUnhandledException += delegate(object sender2, DispatcherUnhandledExceptionEventArgs ex)
        {
            WriteErrorLog("実行中", ex.Exception);
            ex.Handled = true;
        };
        try
        {
            StartupCore();
        }
        catch (Exception ex)
        {
            // 起動中の例外は知らせて終了する（ウィンドウのないプロセスを残さない）
            var logPath = WriteErrorLog("起動", ex);
            MessageBox.Show("PADDOCK の起動に失敗しました。\n\n" + ex.GetType().Name + ": " + ex.Message +
                (logPath == null ? "" : "\n\n詳細はログに記録しました。\n" + logPath),
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
        }
    }

    // 例外を history\{ユーザー名}\paddock_error.log に追記する（書けなければ null）
    private string WriteErrorLog(string phase, Exception ex)
    {
        return WriteErrorLog(phase, ex.ToString());
    }

    // 知らせの本文（例外の全文、または読み飛ばしの件数のような 1 行）を history\{ユーザー名}\paddock_error.log に
    // 追記する（書けなければ null）。上限を超えていたら paddock_error.log.1 へ改名してから書く（1 世代だけ残す。
    // 改名できないときはその回だけ見送り、追記は行う）
    private string WriteErrorLog(string phase, string detail)
    {
        try
        {
            if (!Directory.Exists(userHistDir)) Directory.CreateDirectory(userHistDir);
            var path = GetErrorLogPath();
            try
            {
                if (File.Exists(path) && new FileInfo(path).Length > ERROR_LOG_MAX_BYTES)
                {
                    var previous = path + ".1";
                    if (File.Exists(previous)) File.Delete(previous);
                    File.Move(path, previous);
                }
            }
            catch { }   // 世代交代に失敗しても追記は続ける
            File.AppendAllText(path,
                "[" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture) + "] PADDOCK " + APP_VERSION + " " + phase + "\r\n" +
                detail + "\r\n\r\n", Encoding.UTF8);
            return path;
        }
        catch
        {
            return null;
        }
    }

    // エラーログのパス（history/{ユーザー名}/paddock_error.log）
    private string GetErrorLogPath()
    {
        return System.IO.Path.Combine(userHistDir, "paddock_error.log");
    }

    // 明細の読み取りと計算の結果を JSON ファイルに書き出す（検証用起動・--dump-detail）。
    // 設定と元号マッピングは通常起動と同じファイルから読み、失敗は error キーで返す。
    // 口座種別の補正表は引数で渡したものだけを使う（渡さなければ補正しない）
    private void DumpDetailForVerification(string xlsmPath, string sheetName, string accountNumber, string accountType, string outputPath, string detectionsRaw,
        string institutionName, string typeMappingPath)
    {
        var sb = new StringBuilder();
        try
        {
            var detailConfig = LoadConfig(System.IO.Path.Combine(exeDir, CONFIG_FILE)).Detail;
            var eraMap = LoadEraMapping(System.IO.Path.Combine(exeDir, ERA_MAPPING_FILE));
            AccountTypeMapping typeMapping = null;
            if (typeMappingPath != null)
            {
                typeMapping = LoadAccountTypeMapping(typeMappingPath);
                if (typeMapping == null) throw new InvalidDataException("口座種別の補正表を読めません: " + typeMappingPath);
            }
            var workbook = BusinessLogic.OpenInquiryWorkbook(xlsmPath);
            var data = BusinessLogic.ReadDetailAccount(workbook, sheetName, accountNumber, accountType, institutionName, typeMapping, detailConfig, eraMap);

            // 検出は台帳の行の形（5列分だけ埋めた配列）に戻してから取り出す
            var fields = new string[BusinessLogic.DEPOSIT_LEDGER_COLUMN_COUNT];
            for (int i = 0; i < fields.Length; i++) fields[i] = "";
            var rawParts = detectionsRaw.Split(new[] { "||" }, StringSplitOptions.None);
            int[] columns = { BusinessLogic.DEPOSIT_COL_PENSION, BusinessLogic.DEPOSIT_COL_SALARY, BusinessLogic.DEPOSIT_COL_INSURANCE, BusinessLogic.DEPOSIT_COL_REG_DEPOSIT, BusinessLogic.DEPOSIT_COL_REG_WITHDRAW };
            for (int i = 0; i < columns.Length && i < rawParts.Length; i++) fields[columns[i]] = rawParts[i];
            BusinessLogic.AnalyzeDetailAccount(data, BusinessLogic.BuildDetailDetections(fields));

            sb.Append("{\"status\": \"" + data.Status + "\", \"sheetName\": " + JsonValue(data.SheetName) +
                ", \"date1904\": " + (workbook.Date1904 ? "true" : "false") +
                ", \"maskedCount\": " + data.MaskedCount + ", \"gapCount\": " + data.GapCount +
                ", \"periodFrom\": " + JsonValue(data.PeriodFrom) + ", \"periodTo\": " + JsonValue(data.PeriodTo) +
                ", \"aggregates\": [");
            for (int i = 0; i < data.Aggregates.Count; i++)
            {
                var a = data.Aggregates[i];
                if (i > 0) sb.Append(",");
                sb.Append("\r\n  {\"category\": " + JsonValue(a.Category) + ", \"name\": " + JsonValue(a.Name) +
                    ", \"count\": " + a.Count + ", \"perAmount\": " + a.PerAmount + ", \"lastDate\": " + JsonValue(a.LastDate) + "}");
            }
            sb.Append("], \"transactions\": [");
            for (int i = 0; i < data.Transactions.Count; i++)
            {
                var tx = data.Transactions[i];
                if (i > 0) sb.Append(",");
                sb.Append("\r\n  {\"row\": " + tx.SheetRow +
                    ", \"date\": " + JsonValue(tx.Date) +
                    ", \"summary\": " + JsonValue(tx.Summary) +
                    ", \"withdrawal\": " + JsonValue(tx.Withdrawal) +
                    ", \"deposit\": " + JsonValue(tx.Deposit) +
                    ", \"balance\": " + JsonValue(tx.Balance) +
                    ", \"masked\": " + (tx.BalanceMasked ? "true" : "false") +
                    ", \"estimated\": " + JsonValue(tx.EstimatedBalance) +
                    ", \"gap\": " + JsonValue(tx.Gap) +
                    ", \"category\": " + JsonValue(tx.Category) +
                    ", \"detection\": " + JsonValue(tx.DetectionName) + "}");
            }
            sb.Append("]}");
        }
        catch (Exception ex)
        {
            sb.Length = 0;
            sb.Append("{\"error\": " + JsonValue(ex.GetType().Name + ": " + ex.Message) + "}");
        }
        File.WriteAllText(outputPath, sb.ToString(), new UTF8Encoding(false));
    }

    private static string JsonValue(string value)
    {
        return value == null ? "null" : "\"" + EscapeJson(value) + "\"";
    }

    private static string JsonValue(long? value)
    {
        return value == null ? "null" : value.Value.ToString(CultureInfo.InvariantCulture);
    }

    private static string JsonValue(DateTime? value)
    {
        return value == null ? "null" : "\"" + value.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + "\"";
    }

    // 起動処理の本体（設定・履歴の読込 → ウィンドウ構築 → 台帳の検出と読込 → 表示）
    private void StartupCore()
    {
        LoadUserPrefs();          // ユーザー設定（画面の表示倍率など）

        config = LoadConfig(System.IO.Path.Combine(exeDir, CONFIG_FILE));
        try
        {
            eraMapping = LoadEraMapping(System.IO.Path.Combine(exeDir, ERA_MAPPING_FILE));
        }
        catch (Exception ex)
        {
            // 読めない元号マッピングはファイルがないときと同じ扱い（7桁元号の日付は生値のまま）で起動を続ける
            eraMapping = new Dictionary<int, EraEntry>();
            WriteErrorLog("元号マッピングの読み込み", ex);
        }
        detailTypeMapping = LoadAccountTypeMapping(config.AccountTypeMappingFile);   // 明細の口座種別の照合に使う補正表
        LoadRecentPeople();   // 最近見た人の履歴（読めない場合はセッション内のみに退行）
        LoadNoticeState();    // アップデート告知の既読バージョン（読めない場合は未読扱い）
        EnsureNotesLoaded();  // 人物メモ（全ユーザー分。一覧の印に使うため台帳の読み込みより先）
        RebuildNoteCounts();

        // 台帳フォルダの解決（config 未設定時は exe と同一フォルダ）。生命保険の台帳の検出と読み込みは
        // 生命保険タブを初めて表示したときに行う（フォルダの解決だけ先に済ませる）
        depositLedgerFolder = string.IsNullOrWhiteSpace(config.DepositLedgerFolder) ? exeDir : config.DepositLedgerFolder.Trim();
        insuranceTab = new InsuranceTab(this);
        insuranceTab.insuranceLedgerFolder = string.IsNullOrWhiteSpace(config.InsuranceLedgerFolder) ? exeDir : config.InsuranceLedgerFolder.Trim();

        // 差押リスト連携先の解決（預金・生命保険）。未設定・ファイル不存在なら null のままとし、
        // その種別の連携系UI（右クリック項目・パネルのボタン）を非表示にする。
        // ナビ「差押登録」はどちらかの連携先があれば表示する
        if (!string.IsNullOrWhiteSpace(config.DepositSeizureListScript) &&
            File.Exists(config.DepositSeizureListScript.Trim()))
            depositSeizureScriptPath = config.DepositSeizureListScript.Trim();
        if (!string.IsNullOrWhiteSpace(config.InsuranceSeizureListScript) &&
            File.Exists(config.InsuranceSeizureListScript.Trim()))
            insuranceTab.insuranceSeizureScriptPath = config.InsuranceSeizureListScript.Trim();

        RebuildDepositSeizureTargets();   // 差押リストの対象金融機関の照合パターン
        insuranceTab.RebuildInsuranceSeizureTargets();   // 差押リストの対象保険会社の照合パターン

        window = BuildWindow();
        FindUiElements();
        displayScaleTransform = (ScaleTransform)window.FindResource("DisplayScale");
        keyFocusStyle = (Style)window.FindResource("KeyFocus");
        keyFocusPillStyle = (Style)window.FindResource("KeyFocusPill");
        keyFocusInsetStyle = (Style)window.FindResource("KeyFocusInset");
        ApplyDisplayScale();   // 保存済みの表示倍率を復元する
        SetupEvents();
        ApplySeizureIntegrationVisibility();
        UpdateRecentSubtitle();

        // 白カードの中身を角丸に沿って切り抜く（先頭セクションのヘッダー背景が
        // 四隅にはみ出さないように。設定ビューのカードは MakeSettingsCard 側で適用する）
        ApplyRoundedClip(recentCard, 10);
        ApplyRoundedClip(seizureCard, 10);

        ShowDepositLedgerView();
        UpdateDepositOpenFileButton();

        // 起動時のウィンドウサイズ: 作業領域（タスクバー除く）が 1920×1080 以下なら最大化、
        // それより大きいモニタでは 1920×1080 の中央表示で開く（起動後は自由に可変）
        var workArea = SystemParameters.WorkArea;
        if (workArea.Width <= 1920 || workArea.Height <= 1080)
            window.WindowState = WindowState.Maximized;
        else
        {
            window.Width = 1920;
            window.Height = 1080;
        }

        DetectDepositLedgerFilesThenPopulate(null);

        window.Show();

        // アップデート告知の表示チェック（ウィンドウの表示後。既読なら何もしない）
        CheckUpdateNotice();
    }

    // 最近見た人ビューの見出し下の説明（config の recentMax を反映。設定の保存時にも更新する）
    private void UpdateRecentSubtitle()
    {
        recentSubtitleText.Text = "照会ファイルを開いた・差押リストに追加した行の履歴です" +
            "（このユーザーの分のみ・年度ごとに最新" + config.RecentMax + "件）";
    }

    // ==============================================================
    // ユーザー履歴の配置
    // ==============================================================
    //
    // ユーザー単位ファイル5種（recent_people / seizure_log / notice_state /
    // person_notes / user_prefs）は history/{Windowsユーザー名}/ 直下に保管する。
    // 1ユーザーの持ち物が1箇所にまとまり、整理・バックアップが人単位で完結する。
    // フォルダは書き込み側が必要時に作成し、読み込みはファイル不在なら退行する。

    private string userHistDir;                 // history/{ユーザー名}
    private string userPrefsPath;               // user_prefs.json のパス

    private void InitUserHistoryPaths()
    {
        var histDir = System.IO.Path.Combine(exeDir, "history");
        userHistDir = System.IO.Path.Combine(histDir, Environment.UserName);
        userPrefsPath = System.IO.Path.Combine(userHistDir, "user_prefs.json");
    }

    // ユーザー単位ファイルを原子的に書く（BOM つき UTF-8）。一時ファイル（<path>.tmp）に書いてから
    // 宛先へ移す（宛先がなければ Move・あれば Replace。Replace が使えないボリュームでは上書きコピーに退行）ため、
    // 書いている途中で止まっても宛先のファイルは元のまま残る。失敗は一時ファイルを片づけて例外のまま返す
    //（呼び出し元が以降の保存を止める規約は従来どおり）
    private static void WriteUserFileAtomic(string path, string text)
    {
        var temp = path + ".tmp";
        try
        {
            File.WriteAllText(temp, text, Encoding.UTF8);
            if (!File.Exists(path))
                File.Move(temp, path);
            else
            {
                try { File.Replace(temp, path, null); }
                catch
                {
                    File.Copy(temp, path, true);
                    try { File.Delete(temp); }
                    catch { }   // 宛先は書けている。一時ファイルの掃除失敗で保存を失敗にしない（次の書き込みで上書きされる）
                }
            }
        }
        catch
        {
            try { if (File.Exists(temp)) File.Delete(temp); }
            catch { }   // 一時ファイルの掃除失敗は元の例外を優先する
            throw;
        }
    }

    // ==============================================================
    // ユーザー設定（user_prefs.json）
    // ==============================================================
    //
    // ユーザー単位の設定の受け皿。中身は当面 displayScale（画面の表示倍率 %）
    // のみ。ユーザー単位ファイルのため排他は不要。読めない・不正な値は既定で起動し、
    // 書き込みの失敗はエラー表示せず次回の変更時に再試行する。
    //（表示倍率の反映は ApplyDisplayScale・設定ビューのステッパーは「このユーザーの設定」カード）

    private int displayScale = 100;             // 画面の表示倍率（%）
    private static readonly int[] DisplayScaleSteps = new int[] { 80, 85, 90, 95, 100, 105, 110, 115, 120 };
    private ScaleTransform displayScaleTransform;   // XAML リソース DisplayScale の共有インスタンス

    // 表示倍率を共有 ScaleTransform に反映する（ルート Grid・カレンダーポップアップ・
    // 検出ツールチップが同一インスタンスを参照しているため一括で追従する）
    private void ApplyDisplayScale()
    {
        if (displayScaleTransform == null || displayScaleTransform.IsFrozen) return;   // Freeze 済みでも落ちない
        displayScaleTransform.ScaleX = displayScale / 100.0;
        displayScaleTransform.ScaleY = displayScale / 100.0;
    }

    private void LoadUserPrefs()
    {
        displayScale = 100;
        try
        {
            if (!File.Exists(userPrefsPath)) return;
            var json = File.ReadAllText(userPrefsPath, Encoding.UTF8);
            var value = JsonHelper.GetInt(json, "displayScale", 100);
            if (Array.IndexOf(DisplayScaleSteps, value) >= 0) displayScale = value;
        }
        catch (Exception ex)
        {
            displayScale = 100;   // パース失敗は既定で起動する（落ちない）
            WriteErrorLog("ユーザー設定の読み込み", ex);
        }
    }

    // user_prefs.json を全文再生成で書き込む（recent_people 等と同じ都度保存。
    // 失敗しても動作に影響しない — 次回の変更時に再試行される）
    private void WriteUserPrefs()
    {
        try
        {
            if (!Directory.Exists(userHistDir)) Directory.CreateDirectory(userHistDir);
            WriteUserFileAtomic(userPrefsPath,
                "{\n  \"displayScale\": " + displayScale + "\n}\n");
        }
        catch
        {
            // 書き込み失敗はエラー表示しない
        }
    }

    // ==============================================================
    // 設定・元号マッピング・口座種別の補正表の読み込み
    // ==============================================================

    // paddock_config.json を読み込む（存在しない・解析不能な場合は既定値）
    private AppConfig LoadConfig(string path)
    {
        string text;
        return LoadConfig(path, out text);
    }

    // 同上。text に読めて解釈できた本文を返す（存在しない・読めない・解釈に失敗したときは null）。
    // 設定ビューがファイルの今の値から組み、保存の前に本文を照合するために使う
    private AppConfig LoadConfig(string path, out string text)
    {
        var result = new AppConfig();
        text = null;
        if (!File.Exists(path)) return result;

        try
        {
            var json = File.ReadAllText(path, Encoding.UTF8);
            result.DepositLedgerFolder = JsonHelper.GetString(json, "depositLedgerFolder");
            result.InsuranceLedgerFolder = JsonHelper.GetString(json, "insuranceLedgerFolder");
            result.DepositSearchFolders = ParseSearchFolders(json, "depositSearchFolders");
            result.InsuranceSearchFolders = ParseSearchFolders(json, "insuranceSearchFolders");
            result.DepositSeizureListScript = JsonHelper.GetString(json, "depositSeizureListScript");
            result.InsuranceSeizureListScript = JsonHelper.GetString(json, "insuranceSeizureListScript");
            result.RecentMax = JsonHelper.GetInt(json, "recentMax", 100);
            if (result.RecentMax < 1) result.RecentMax = 100;   // 不正値は既定に戻す

            // キー未設定（null）は内蔵既定「◆」のまま・空文字列は空リスト（統合の無効化）
            var fuzzyRaw = JsonHelper.GetString(json, "staffFuzzyChars");
            if (fuzzyRaw != null) result.StaffFuzzyChars = ParseStaffFuzzyChars(fuzzyRaw);

            result.DepositSeizureTargetBanks = JsonHelper.GetString(json, "depositSeizureTargetBanks");
            result.InsuranceSeizureTargetCompanies = JsonHelper.GetString(json, "insuranceSeizureTargetCompanies");
            result.AccountTypeMappingFile = JsonHelper.GetString(json, "accountTypeMappingFile");
            result.Detail = ParseDetailConfig(json);
            text = json;
        }
        catch (Exception ex)
        {
            // 設定が読めなくても既定値（exe と同一フォルダ）で動作を続ける
            WriteErrorLog("設定の読み込み", ex);
        }
        return result;
    }

    // detail オブジェクト（「明細を見る」の読み取り位置）を解析する。キー単位で
    // 省略でき、不正な値はそのキーだけ既定に戻す（recentMax と同じ規約）
    private static DetailConfig ParseDetailConfig(string json)
    {
        var result = DetailConfig.CreateDefault();
        var keyIdx = json.IndexOf("\"detail\"");
        if (keyIdx < 0) return result;
        var colonIdx = json.IndexOf(':', keyIdx + 8);
        if (colonIdx < 0) return result;
        var objStart = colonIdx + 1;
        while (objStart < json.Length && char.IsWhiteSpace(json[objStart])) objStart++;
        if (objStart >= json.Length || json[objStart] != '{') return result;
        var objEnd = JsonHelper.FindMatchingBracket(json, objStart, '{', '}');
        if (objEnd < 0) return result;
        var objJson = json.Substring(objStart, objEnd - objStart + 1);

        var stopValues = ParseJsonStringArray(objJson, "stopValues");
        if (stopValues != null && stopValues.Count > 0) result.StopValues = stopValues;

        int col, offset;
        var numberCell = JsonHelper.GetString(objJson, "accountNumberCell");
        if (BusinessLogic.TryParseCellRef(numberCell, out col, out offset)) result.AccountNumberCell = numberCell.Trim();
        var typeCell = JsonHelper.GetString(objJson, "accountTypeCell");
        if (BusinessLogic.TryParseCellRef(typeCell, out col, out offset)) result.AccountTypeCell = typeCell.Trim();

        var startRow = JsonHelper.GetInt(objJson, "dataStartRow", result.DataStartRow);
        if (startRow >= 1) result.DataStartRow = startRow;

        result.DateCol = ReadColumnLetters(objJson, "dateCol", result.DateCol);
        result.SummaryCol = ReadColumnLetters(objJson, "summaryCol", result.SummaryCol);
        result.WithdrawalCol = ReadColumnLetters(objJson, "withdrawalCol", result.WithdrawalCol);
        result.DepositCol = ReadColumnLetters(objJson, "depositCol", result.DepositCol);
        result.BalanceCol = ReadColumnLetters(objJson, "balanceCol", result.BalanceCol);
        result.ResolveIndexes();
        return result;
    }

    // 列文字の設定値を読む（未設定・不正値は defaultValue）。大文字に揃えて保持する
    private static string ReadColumnLetters(string json, string key, string defaultValue)
    {
        int index;
        var raw = JsonHelper.GetString(json, key);
        return BusinessLogic.TryParseColumnLetters(raw, out index) ? raw.Trim().ToUpperInvariant() : defaultValue;
    }

    // 文字列配列の設定値を読む（キーなし・配列でない場合は null。空要素は読み飛ばす）
    private static List<string> ParseJsonStringArray(string json, string key)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return null;
        var arrStart = json.IndexOf('[', keyIdx);
        if (arrStart < 0) return null;
        var arrEnd = JsonHelper.FindMatchingBracket(json, arrStart, '[', ']');
        if (arrEnd < 0) return null;
        var result = new List<string>();
        foreach (Match m in Regex.Matches(json.Substring(arrStart, arrEnd - arrStart + 1), "\"((?:[^\"\\\\]|\\\\.)*)\""))
        {
            var value = JsonHelper.UnescapeJsonString(m.Groups[1].Value).Trim();
            if (value != "") result.Add(value);
        }
        return result;
    }

    // staffFuzzyChars（カンマ区切りの化け置換文字リスト）を解析する。Trim 後に
    // 1文字のトークンのみ採用し、それ以外は読み飛ばす
    private static List<char> ParseStaffFuzzyChars(string raw)
    {
        var result = new List<char>();
        foreach (var token in raw.Split(','))
        {
            var trimmed = token.Trim();
            if (trimmed.Length == 1 && !result.Contains(trimmed[0])) result.Add(trimmed[0]);
        }
        return result;
    }

    // staffFuzzyChars の表示・保存用文字列（カンマ区切り。空リストは空文字列）
    private static string FormatStaffFuzzyChars(List<char> fuzzyChars)
    {
        var parts = new string[fuzzyChars.Count];
        for (int i = 0; i < fuzzyChars.Count; i++) parts[i] = fuzzyChars[i].ToString();
        return string.Join(",", parts);
    }

    // 照会結果フォルダのオブジェクト（年度キー → 文字列または文字列配列）を解析する（key は
    // depositSearchFolders・insuranceSearchFolders）。解析できない場合は null を返す（ドリルダウン無効として扱われる）
    private static Dictionary<string, List<string>> ParseSearchFolders(string json, string key)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return null;
        var objStart = json.IndexOf('{', keyIdx);
        if (objStart < 0) return null;
        var objEnd = JsonHelper.FindMatchingBracket(json, objStart, '{', '}');
        if (objEnd < 0) return null;
        var objJson = json.Substring(objStart + 1, objEnd - objStart - 1);

        // "キー": "値" / "キー": ["値", ...] の並びを走査する
        var result = new Dictionary<string, List<string>>();
        var pairRegex = new Regex("\"((?:[^\"\\\\]|\\\\.)*)\"\\s*:\\s*(\"(?:[^\"\\\\]|\\\\.)*\"|\\[[^\\]]*\\])");
        foreach (Match m in pairRegex.Matches(objJson))
        {
            var year = JsonHelper.UnescapeJsonString(m.Groups[1].Value);
            var value = m.Groups[2].Value;
            var folders = new List<string>();
            if (value.StartsWith("["))
            {
                foreach (Match sm in Regex.Matches(value, "\"((?:[^\"\\\\]|\\\\.)*)\""))
                {
                    var folder = JsonHelper.UnescapeJsonString(sm.Groups[1].Value).Trim();
                    if (folder != "") folders.Add(folder);
                }
            }
            else
            {
                var folder = JsonHelper.UnescapeJsonString(value.Trim('"')).Trim();
                if (folder != "") folders.Add(folder);
            }
            if (folders.Count > 0) result[year] = folders;
        }
        return result.Count > 0 ? result : null;
    }

    // era_mapping.json を読み込む（startDate を持たないファイルもそのまま読める）
    private Dictionary<int, EraEntry> LoadEraMapping(string path)
    {
        var map = new Dictionary<int, EraEntry>();
        if (!File.Exists(path)) return map;

        var json = File.ReadAllText(path, Encoding.UTF8);
        for (int code = 1; code <= 9; code++)
        {
            var key = code.ToString();
            var keyIdx = json.IndexOf("\"" + key + "\"");
            if (keyIdx < 0) continue;
            var objStart = json.IndexOf('{', keyIdx);
            if (objStart < 0) continue;
            var objEnd = JsonHelper.FindMatchingBracket(json, objStart, '{', '}');
            if (objEnd < 0) continue;
            var objJson = json.Substring(objStart, objEnd - objStart + 1);

            var name = JsonHelper.GetString(objJson, "name");
            var startYear = JsonHelper.GetInt(objJson, "startYear");
            if (name == null || startYear <= 0) continue;

            var entry = new EraEntry { Name = name, StartYear = startYear };
            entry.StartDate = BusinessLogic.ParseDisplayDate(JsonHelper.GetString(objJson, "startDate"));
            map[code] = entry;
        }
        return map;
    }

    // 口座種別の補正表（JUDGE の account_type_mapping.json）を読み込む。未設定・ファイルがない・
    // 読めない場合は null（明細ページの口座種別を補正せずに照合する。知らせは設定ビューのチップ）
    private static AccountTypeMapping LoadAccountTypeMapping(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return null;
        try
        {
            path = path.Trim();
            if (!File.Exists(path)) return null;
            return BusinessLogic.ParseAccountTypeMapping(File.ReadAllText(path, Encoding.UTF8));
        }
        catch
        {
            return null;
        }
    }
}