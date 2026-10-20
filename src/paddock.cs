// ==============================================================================
// paddock.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 【使用方法】
// paddock_build.bat を実行して paddock.exe を生成し、ダブルクリックで起動する。
// JUDGE（judge.ps1）が蓄積する台帳CSV（財産調査台帳_YYYY年度.csv）を
// 読み込み、宛名番号・氏名・カナ・金融機関名で検索できる一覧を表示する
//（半角/全角・空白の揺れは正規化で吸収する）。「詳細検索」では残高・日付の範囲、
// 検出カテゴリ、行の分類、担当（台帳から自動列挙・表記ゆれは統合する）に
// よる絞り込みを組み合わせられ、条件は着脱可能なチップとして
// 検索バー直下に表示される。口座なし行は既定で非表示（行の分類「口座なし」で表示）。
// 検出列には取引明細分析の結果（反対債権・年金・給与・保険・定期入出金）を
// チップとして圧縮表示し、セルホバーで原文をツールチップ表示する。
// 行を選択すると右側に詳細パネルが開き、当該人物の照会履歴を全年度ファイルから
// 横断収集した年度横断ビュー（人物ヘッダー・ファクト統計・メモ・タイムライン）を
// 表示する。メモ（history/{ユーザー名}/person_notes.json）は人物単位の自由テキストで、
// 全ユーザーの記入分を集約して表示する。追加・編集・削除は自分のファイルにのみ
// 書き込む（他ユーザーのメモは読み取り専用）。
// 一覧・パネルのダブルクリック等から、仕分け後の照会ファイルを Excel で開ける
//（宛名番号でファイルを候補化し、期待シート名の照合で特定・シートジャンプする）。
// 差押リスト連携（config の depositSeizureListScript 設定時のみ有効）では、
// 同じファイル解決を経て預金差押予定一覧 作成ツールに照会ファイルを引き渡し、
// 登録履歴（差押登録ログ）をナビ「差押登録」で年度別に一覧表示する。
// ナビ「最近見た人」は、照会ファイルを開いた・差押リストに追加した行の履歴を
// ユーザー単位で永続化し、年度（行の登録日基準）別に一覧表示する。履歴のクリックで
// 当該人物の詳細パネル（年度横断ビュー）を再表示し、記録した行のカードを選択状態にする。
// ナビ「設定」では paddock_config.json の内容と保管ファイルの状態を確認でき、
// 設定値の編集・保存もできる（設定ファイルへの書き込みはこの保存操作のみ）。
// 画面の表示倍率（80〜120% の9段階）はユーザー単位（user_prefs.json）で
// 保存され、次回起動時も適用される。
// JUDGE の生命保険の台帳（生命保険調査台帳_YYYY年度.csv）はナビ「生命保険」で閲覧する。
// 預金とは別の一覧（回答基準日・契約者・保険会社名・契約種類・証券番号・契約の状態・
// 解約返戻金・満期年月日・注記）を持ち、年度・検索欄・詳細検索（項目検索・注記・
// 行の分類・担当・解約返戻金と日付の範囲）はタブごとに保持される。台帳は生命保険タブを
// 初めて表示したときに読み込む。契約なしの行と電子回答不可の行は既定で非表示
//（行の分類「契約なし」「電子回答不可」で表示）。一覧の行から詳細パネル（契約カードと
// 展開部・差引見込額（参考）・メモ）を開き、照会ファイルを開く・生命保険差押予定一覧
// 作成ツール（entry_insurance.exe）へ渡すことができる（預金と同じ動線）。開いた・渡した
// 行は預金と同じ履歴（最近見た人）に種別つきで載り、履歴から生命保険の詳細パネルを
// 開ける（台帳を読む前なら全年度のファイルからその人の行を探す）。
// 起動時には update_notice.json による告知をユーザーごとに初回のみオーバーレイ表示する
//（既読情報は history/{ユーザー名}/notice_state.json に保管する）。
//
// 設計原則: 判定しない。判断材料を網羅的に提示する。
// （残高と反対債権の差引・履歴の増減判定など、評価を含む加工値は表示しない。
//   照会ファイルの候補が特定できないときも職員の選択に委ねる。
//   差押リスト連携は職員の能動操作の起点であり、ツールによる推奨ではない）
//
// 台帳CSVへの書き込みは JUDGE（judge.ps1・およびバックフィル）のみが
// 行い、本ツールは読み取り専用。照会ファイルも読み取り目的で開くのみ。
// 差押登録ログの書き込みは entry_deposit 側が行い、本ツールは読むのみ。
//
// 【ビルド方法】
// paddock_build.bat を実行（.NET Framework 4.0 の csc.exe を使用）
// ソースは src フォルダの複数のファイルに分かれている（PaddockApp と BusinessLogic は
// partial class）。コンパイルするファイルとその順序は paddock_build.bat に列挙している。
// ※ フォルダ参照ダイアログに System.Windows.Forms.dll が必要（bat に設定済み）
//
// 【必要ファイル（同じフォルダに配置）】
// ＜必須＞
// - src\                （ソースコード一式。このファイルを含む）
// - paddock.ico         （アプリケーションアイコン）
// - paddock_build.bat   （ビルドスクリプト）
// - era_mapping.json    （元号マッピング。STEWARD と同一ファイル）
// ＜任意＞
// - paddock_config.json （設定ファイル。未配置時は既定値で動作）
// - update_notice.json  （アップデート告知。未配置時は告知なしで動作）
// - 財産調査台帳_YYYY年度.csv （預金の台帳。既定では exe と同一フォルダから検出）
// - 生命保険調査台帳_YYYY年度.csv （生命保険の台帳。同上）
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

// アセンブリの版（exe のプロパティに出る。画面に出す版の定数と同じ値に揃える）
[assembly: System.Reflection.AssemblyTitle("PADDOCK")]
[assembly: System.Reflection.AssemblyProduct("PADDOCK — 財産調査台帳")]
[assembly: System.Reflection.AssemblyVersion("1.7.2.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.7.2.0")]

// ==============================================================
// 設定
// ==============================================================

public class AppConfig
{
    // 預金・生命保険の設定は対で持つ（台帳の検出フォルダ・照会結果フォルダ・差押リスト連携先・差押リストの対象）。
    // 台帳CSVの検出フォルダ（未設定時は exe と同一フォルダ）
    public string DepositLedgerFolder { get; set; }
    public string InsuranceLedgerFolder { get; set; }

    // 年度→照会結果フォルダのマッピング（例: "2026年度" → ["D:\...\2026\照会結果"]）
    // 未設定時はドリルダウン（照会ファイルを開く）を無効化する
    public Dictionary<string, List<string>> DepositSearchFolders { get; set; }
    public Dictionary<string, List<string>> InsuranceSearchFolders { get; set; }

    // 預金差押予定一覧 作成ツール（entry_deposit.exe）・生命保険差押予定一覧 作成ツール（entry_insurance.exe）の絶対パス
    // 未設定または該当ファイルが存在しない場合、差押リスト連携のUIを非表示にする
    public string DepositSeizureListScript { get; set; }
    public string InsuranceSeizureListScript { get; set; }

    // 最近見た人の年度ごとの保持件数
    public int RecentMax { get; set; }

    // 担当の表記ゆれ統合に用いる化け置換文字（config はカンマ区切りで指定・
    // キー未設定なら内蔵既定「◆」・空欄ならワイルドカード統合を無効化する）
    public List<char> StaffFuzzyChars { get; set; }

    // 差押リストの対象金融機関・対象保険会社（カンマ区切り・ワイルドカード可・生値のまま保持）。
    // キー未設定・空欄は制御なし（全行で「差押リストに追加」が可能）
    public string DepositSeizureTargetBanks { get; set; }
    public string InsuranceSeizureTargetCompanies { get; set; }

    // 口座種別の補正表（JUDGE の account_type_mapping.json）のパス。未設定・読めないときは
    // 明細ページの口座種別を補正せずに台帳と照合する
    public string AccountTypeMappingFile { get; set; }

    // 「明細を見る」が照会シートを読む位置。キー未設定・不正値は既定（JUDGE の
    // judge_config.json と同値）に戻す
    public DetailConfig Detail { get; set; }

    public AppConfig()
    {
        RecentMax = 100;
        StaffFuzzyChars = new List<char> { '◆' };
        Detail = DetailConfig.CreateDefault();
    }
}

// 「明細を見る」の読み取り位置。明細ページ先頭の目印セル（A列）を1行目とした
// 相対指定で、JUDGE の judge_config.json（detail*）と同じ位置を指す。列文字・セルは
// 解析済みのインデックスも併せて保持する（読み取りのたびに文字列を解かない）
public class DetailConfig
{
    public List<string> StopValues { get; set; }     // 明細ページ先頭の目印文字列（A列・完全一致）
    public string AccountNumberCell { get; set; }    // 口座番号セル（例: "M3"）
    public string AccountTypeCell { get; set; }      // 口座種別セル（例: "F3"）
    public int DataStartRow { get; set; }            // 明細の開始行（目印セルを1行目とする）
    public string DateCol { get; set; }              // 取引日の列
    public string SummaryCol { get; set; }           // 摘要の列
    public string WithdrawalCol { get; set; }        // 出金（払戻）の列
    public string DepositCol { get; set; }           // 入金（預入）の列
    public string BalanceCol { get; set; }           // 残高の列

    // 解析済み（1始まりの列番号・目印行からの行オフセット）
    public int AccountNumberColIndex;
    public int AccountNumberRowOffset;
    public int AccountTypeColIndex;
    public int AccountTypeRowOffset;
    public int DateColIndex;
    public int SummaryColIndex;
    public int WithdrawalColIndex;
    public int DepositColIndex;
    public int BalanceColIndex;

    public static DetailConfig CreateDefault()
    {
        var result = new DetailConfig
        {
            StopValues = new List<string> { "回答書（取引明細表）", "回答書（担保・保護預）" },
            AccountNumberCell = "M3",
            AccountTypeCell = "F3",
            DataStartRow = 7,
            DateCol = "B",
            SummaryCol = "G",
            WithdrawalCol = "Q",
            DepositCol = "U",
            BalanceCol = "Y"
        };
        result.ResolveIndexes();
        return result;
    }

    // 文字列指定から列番号・行オフセットを確定する（値は事前に検証済みであること）
    public void ResolveIndexes()
    {
        BusinessLogic.TryParseCellRef(AccountNumberCell, out AccountNumberColIndex, out AccountNumberRowOffset);
        BusinessLogic.TryParseCellRef(AccountTypeCell, out AccountTypeColIndex, out AccountTypeRowOffset);
        BusinessLogic.TryParseColumnLetters(DateCol, out DateColIndex);
        BusinessLogic.TryParseColumnLetters(SummaryCol, out SummaryColIndex);
        BusinessLogic.TryParseColumnLetters(WithdrawalCol, out WithdrawalColIndex);
        BusinessLogic.TryParseColumnLetters(DepositCol, out DepositColIndex);
        BusinessLogic.TryParseColumnLetters(BalanceCol, out BalanceColIndex);
    }
}

// 口座種別の補正表（JUDGE の account_type_mapping.json）。金融機関ごとの表と global の表を
// ファイルの並びのまま持つ（補正は表の並びの順に見て、最初に当たったキーの値を採るため）
public class AccountTypeMapping
{
    public List<AccountTypeRuleTable> Institutions = new List<AccountTypeRuleTable>();
    public List<AccountTypeRule> Global;             // global の表（キーがなければ null）
}

// 金融機関1つ分の表
public class AccountTypeRuleTable
{
    public string InstitutionKey;                    // 金融機関名（全角化済み・照合用）
    public List<AccountTypeRule> Rules;
}

// 表の1項目（キー→補正後の表記）
public class AccountTypeRule
{
    public string Key;                               // 表のキー（完全一致に使う）
    public Regex Prefix;                             // 末尾が * のキーの前方一致（それ以外は null）
    public string Value;                             // 補正後の表記（値が文字列でなければ null＝補正しない）
}

// 照会シートの1セル（「明細を見る」用）。数値セルは XML の生値（OADate・金額）を文字列のまま持ち、
// 解釈は使う側で行う（Excel COM の Value2 と同じ区別）
public class SheetCell
{
    public string Text;                              // 文字列セルは本文・数値セルは生値
    public bool IsNumeric;                           // 数値セル（型属性なし・"n"）
}

// 照会シート1枚分のセル（必要な列だけを疎に保持する）
public class SheetCells
{
    public Dictionary<int, Dictionary<int, SheetCell>> Rows = new Dictionary<int, Dictionary<int, SheetCell>>();
    public int MaxRow;                               // セルを持つ最終行（UsedRange の末尾に相当）

    public SheetCell Get(int row, int col)
    {
        Dictionary<int, SheetCell> cols;
        SheetCell cell;
        if (Rows.TryGetValue(row, out cols) && cols.TryGetValue(col, out cell)) return cell;
        return null;
    }

    public string GetText(int row, int col)
    {
        var cell = Get(row, col);
        return cell == null ? "" : cell.Text;
    }
}

// 取引明細の1行（「明細を見る」用）
public class DetailTransaction
{
    public DateTime? Date;                           // 取引日（解釈できなければ null）
    public string Summary;                           // 摘要（原本の生値）
    public long? Withdrawal;                         // 出金（空・解釈不能は null）
    public long? Deposit;                            // 入金（同上）
    public long? Balance;                            // 残高（伏字・空・解釈不能は null）
    public bool BalanceMasked;                       // 残高が伏字（"*" 始まり）
    public int SheetRow;                             // 原本シート上の行番号

    // 以下は読み取り後の計算値（BusinessLogic.AnalyzeDetailAccount）
    public long? EstimatedBalance;                   // 残高が無い行の推定残高（入出金の積み上げ。推定できなければ null）
    public long? Gap;                                // 回答の残高 − 積み上げ（一致しない点にのみ値が入る）
    public string Category;                          // 当たった検出のカテゴリ（当たらなければ null）
    public string DetectionName;                     // 当たった検出の名称
}

// 台帳の検出列から取り出した検出1件（カテゴリ・名称・入金側か）
public class DetailDetection
{
    public string Category;
    public string Name;                              // 台帳に記録された名称（表示用）
    public string Bucket;                            // 台帳に記録された周期（例: "25日頃"・"不定期"。なければ null）
    public string NormalizedName;                    // 突合用（NormalizeForSearch 済み）
    public bool IsDeposit;                           // 年金・給与・定期入金=入金側／保険・定期出金=出金側
}

// 検出ごとの集約（回数・1回あたり・直近）
public class DetectionAggregate
{
    public string Category;
    public string Name;
    public int Count;
    public long Total;                               // 合計（並び順の決定に使う。表示しない）
    public long PerAmount;                           // 1回あたり（合計÷回数・整数丸め。出金は負）
    public DateTime? LastDate;                       // 直近の取引日
}

// 口座1件の明細の読み取り結果
public enum DetailStatus
{
    Loaded,                                          // 明細あり
    NoDetailPage,                                    // 期待シートに当該口座の明細ページがない
    NoTransactions                                   // 明細ページはあるが取引が0行
}

public class DetailAccountData
{
    public DetailStatus Status;
    public List<DetailTransaction> Transactions = new List<DetailTransaction>();
    public int MaskedCount;                          // 残高が伏字の行数
    public string SourcePath;                        // 読んだ照会ファイル
    public string SheetName;                         // 読んだシート名

    // 以下は計算値（BusinessLogic.AnalyzeDetailAccount）
    public List<DetectionAggregate> Aggregates = new List<DetectionAggregate>();
    public int GapCount;                             // 回答の残高と積み上げが合わない点の数
    public DateTime? PeriodFrom;                     // 取引日の範囲（最初の取引日。横軸の左端）
    public DateTime? PeriodTo;                       // 同（最後の取引日。横軸の右端）
}

// 照会ファイル1冊の骨格（シート名とパートの対応・共有文字列・1904年始まりか）。
// 開いた Package は握り続けず、シート本体を読むたびに開き直す（共有フォルダ上のファイルを
// ロックしない。共有文字列はファイル単位で1回読めばよい）
public class InquiryWorkbook
{
    public string Path;
    public bool Date1904;
    public List<KeyValuePair<string, Uri>> Sheets = new List<KeyValuePair<string, Uri>>();   // シート名→パート
    public string[] SharedStrings = new string[0];

    // 期待シート名（またはその重複連番形）に一致する最初のシートを返す（なければ null）
    public KeyValuePair<string, Uri>? FindSheet(string expectedSheet)
    {
        foreach (var sheet in Sheets)
            if (BusinessLogic.SheetNameMatchesExpected(sheet.Key, expectedSheet)) return sheet;
        return null;
    }
}

// 元号マッピング1件
public class EraEntry
{
    public string Name { get; set; }                 // 元号名（例: "令和"）
    public int StartYear { get; set; }               // 元号元年の西暦（例: 2019）
    public DateTime? StartDate { get; set; }         // 改元日（西暦→和暦変換の境界判定用。未設定は元年1月1日扱い）
}

// 検出チップ1個（一覧セル第1層のバインディング用）
public class ChipInfo
{
    public string Text { get; set; }                 // チップ表示文字列（例: "給与 25日頃 ×2"）
    public Brush Bg { get; set; }                    // 背景ブラシ（カテゴリ別・Freeze済み）
    public Brush Fg { get; set; }                    // 文字ブラシ（カテゴリ別・Freeze済み）
}

// 検出ツールチップ1行（第2層のバインディング用）
public class TooltipLineInfo
{
    public string ChipText { get; set; }             // 行頭チップの文字列（種目名 or カテゴリ名）
    public Brush ChipBg { get; set; }                // チップ背景（暗色サーフェス用トーン）
    public Brush ChipFg { get; set; }                // チップ文字色
    public string Content { get; set; }              // 原文（生文字列のまま）
    public FontFamily ContentFont { get; set; }      // 原文のフォント（金額行のみ Consolas）
}

// 台帳データ行1件（テーブルバインディング用）
// 表示用に整形した値と、詳細表示・最近見た人・差押リスト連携で使う生フィールドの両方を保持する
public class DepositLedgerRecord : INotifyPropertyChanged
{
    public string DisplayBaseDate { get; set; }      // 表示用回答基準日（yyyy/MM/dd。口座なし行は空欄）
    public string AddressNumber { get; set; }        // 宛名番号
    public string Name { get; set; }                 // 名義人氏名
    public string Kana { get; set; }                 // カナ氏名
    public string InstitutionName { get; set; }      // 金融機関名
    public string BranchName { get; set; }           // 支店名
    public string AccountType { get; set; }          // 口座種別
    public string AccountNumber { get; set; }        // 口座番号
    public string DisplayBalance { get; set; }       // 表示用残高（センチネル値は「対象外」に置換）
    public string BalanceToolTip { get; set; }       // 残高セルのツールチップ（センチネル時のみ生値。null=非表示）
    public string DisplayLastTransaction { get; set; } // 最終取引日（元からyyyy/MM/dd形式）
    public bool IsNoAccount { get; set; }            // 口座なし行（取引有無 == "0"）
    public bool IsMutedRow { get { return IsNoAccount; } }   // 一覧の行を薄文字色で描く（生命保険の行と共用するスタイルの条件）
    public bool HasDetection { get; set; }           // 検出データあり（ツールチップの有効化に使用）
    public bool HasNotes { get { return hasNotes; } set { hasNotes = value; RaisePropertyChanged("HasNotes"); } }   // 人物メモあり（氏名の右の印の表示に使用）
    public string NotesToolTip { get { return notesToolTip; } set { notesToolTip = value; RaisePropertyChanged("NotesToolTip"); } }   // 〃 印のツールチップ「メモ N件」（なしは null=非表示）
    public List<ChipInfo> Chips { get; set; }        // 第1層チップ（最大3個+超過チップ）
    public List<TooltipLineInfo> TooltipLines { get; set; } // 第2層ツールチップの行

    public string SortKey { get; set; }              // ソートキー（回答基準日、空なら登録日。正準形 yyyy/MM/dd で比較）
    public string[] RawFields { get; set; }          // パース済みの全列（旧27列・現行28列。詳細パネル・原本特定用）

    // 検索用の正規化キー（NormalizeForSearch 済み。読み込み時に事前計算する）
    public string SearchAddressNumber { get; set; }
    public string SearchName { get; set; }
    public string SearchKana { get; set; }
    public string SearchInstitution { get; set; }

    // 詳細検索用の事前計算値（読み込み時に確定。キー入力のたびの変換を避け、比較だけで済ませる）
    public string Classification { get; set; }       // 登録推奨列の生値（"1"=代表口座 / "2"=特殊口座のみ）
    public long? BalanceValue { get; set; }          // 残高の数値（センチネル・非数・空は null=金額条件の対象外）
    public DateTime? BaseDateValue { get; set; }     // 回答基準日（表示用 yyyy/MM/dd の日付化。不能は null）
    public DateTime? LastTxValue { get; set; }       // 最終取引日（同上）
    public bool[] DetectionFlags { get; set; }       // 検出の有無（BusinessLogic.DETECTION_NAMES と同順）
    public string StaffKey { get; set; }             // 担当職員名の正規化キー（空欄・旧27列の行は null=担当条件の対象外）

    // 人物メモの印はメモの追加・削除で後から変わるため、この 2 つだけ変更通知を上げる（ほかの値は読み込み時に確定）。
    // 印の付け外しで一覧のコンテナを作り直さずに済む
    private bool hasNotes;
    private string notesToolTip;
    public event PropertyChangedEventHandler PropertyChanged;
    private void RaisePropertyChanged(string name)
    {
        var handler = PropertyChanged;
        if (handler != null) handler(this, new PropertyChangedEventArgs(name));
    }
}

// 生命保険の台帳データ行1件（テーブルバインディング用）。預金の行と同じく、
// 表示用に整形した値と、詳細表示・最近見た人・差押リスト連携で使う生フィールドの両方を保持する
public class InsuranceLedgerRecord : INotifyPropertyChanged
{
    public string DisplayBaseDate { get; set; }      // 表示用回答基準日（yyyy/MM/dd。契約なし行は空欄）
    public string AddressNumber { get; set; }        // 宛名番号
    public string Name { get; set; }                 // 契約者漢字氏名（契約なし行・電子回答不可行は台帳のとおり空）
    public string Kana { get; set; }                 // 契約者カナ氏名
    public string CompanyName { get; set; }          // 保険会社名（補正後）
    public string ContractType { get; set; }         // 契約種類
    public string PolicyNumber { get; set; }         // 証券番号
    public string ContractStatus { get; set; }       // 契約の状態
    public string DisplaySurrenderValue { get; set; } // 解約返戻金（台帳の値のまま）
    public bool IsSurrenderValueText { get; set; }   // 解約返戻金が数値として読めない生値（薄文字色で出す）
    public string MaturityDate { get; set; }         // 満期年月日（台帳の値のまま。終身の 99999999 もそのまま）
    public bool IsNoContract { get; set; }           // 契約なし行（契約有無が 1・2 以外）
    public bool IsUnanswered { get; set; }           // 電子回答不可行（契約有無 == "2"）
    public bool IsMutedRow { get { return IsNoContract || IsUnanswered; } }   // 一覧の行を薄文字色で描く（預金の行と共用するスタイルの条件）
    public string MutedPillText { get; set; }        // 注記列のピルの文言（「契約なし」「電子回答不可」。契約あり行は null=非表示）
    public bool HasNote { get; set; }                // 注記あり（ツールチップの有効化に使用。人物メモの HasNotes とは別）
    public bool HasNotes { get { return hasNotes; } set { hasNotes = value; RaisePropertyChanged("HasNotes"); } }   // 人物メモあり（氏名の右の印の表示に使用）
    public string NotesToolTip { get { return notesToolTip; } set { notesToolTip = value; RaisePropertyChanged("NotesToolTip"); } }   // 〃 印のツールチップ「メモ N件」（なしは null=非表示）
    public List<ChipInfo> Chips { get; set; }        // 第1層チップ（最大3個+超過チップ）
    public List<TooltipLineInfo> TooltipLines { get; set; } // 第2層ツールチップの行

    public string SortKey { get; set; }              // ソートキー（回答基準日、空なら登録日。正準形 yyyy/MM/dd で比較）
    public string[] RawFields { get; set; }          // 台帳の並び（INSURANCE_LEDGER_COLUMNS）の全列（詳細パネル・原本特定用）

    // 検索用の正規化キー（NormalizeForSearch 済み。読み込み時に事前計算する）
    public string SearchAddressNumber { get; set; }
    public string SearchName { get; set; }
    public string SearchKana { get; set; }
    public string SearchCompany { get; set; }
    public string SearchPolicyNumber { get; set; }

    // 詳細検索用の事前計算値（読み込み時に確定。キー入力のたびの変換を避け、比較だけで済ませる）
    public long? SurrenderValue { get; set; }        // 解約返戻金の数値（読めない生値・空は null=金額条件の対象外）
    public DateTime? BaseDateValue { get; set; }     // 回答基準日（表示用 yyyy/MM/dd の日付化。不能は null）
    public DateTime? MaturityValue { get; set; }     // 満期年月日（yyyy/MM/dd・年 1900〜2099 で読めたもの。不能は null）
    public bool[] NoteFlags { get; set; }            // 注記の有無（BusinessLogic.INSURANCE_NOTE_NAMES と同順）
    public string StaffKey { get; set; }             // 担当職員名の正規化キー（空欄は null=担当条件の対象外）

    // 人物メモの印はメモの追加・削除で後から変わるため、この 2 つだけ変更通知を上げる（ほかの値は読み込み時に確定）。
    // 印の付け外しで一覧のコンテナを作り直さずに済む
    private bool hasNotes;
    private string notesToolTip;
    public event PropertyChangedEventHandler PropertyChanged;
    private void RaisePropertyChanged(string name)
    {
        var handler = PropertyChanged;
        if (handler != null) handler(this, new PropertyChangedEventArgs(name));
    }
}

// 生命保険の契約カードの展開部の値の片1つ（数値・日付は Consolas、文字は既定の字で描く）
public class InsurancePanelPart
{
    public string Text { get; set; }                 // 片の文字列（台帳の値のまま。配当の日付は括弧つき）
    public bool Mono { get; set; }                   // Consolas で描く（数値・日付）
}

// 生命保険の契約カードの展開部の1行（ラベル／値）
public class InsurancePanelRow
{
    public string Label { get; set; }                // 台帳の列名（合成行は組の名・差引見込額（参考））
    public List<InsurancePanelPart> Parts { get; set; }   // 値の片（並べて1つの TextBlock に描く）
    public bool IsRawText { get; set; }              // 数値として読めない生値（薄文字色で出す）
    public string ToolTip { get; set; }              // 差引見込額（参考）の式・計算できない理由（ほかの行は null）
}

// 〃 ブロック（契約／返戻金・配当・貸付金等／差押・質権・介入権）。行のないブロックは作らない
public class InsurancePanelBlock
{
    public string Title { get; set; }                // ブロックの見出し
    public List<InsurancePanelRow> Rows { get; set; }
}

// 担当の表記ゆれ統合の1クラスタ。担当フィルターの列挙（トグル）と
// 行の照合に使う。統合は列挙・照合のための同値関係で、台帳の値・表示には関与しない
public class StaffCluster
{
    public string Representative;                    // 代表名（トグルの表示名。化け文字を含まない最長の生値）
    public HashSet<string> MemberKeys;               // メンバーの正規化キー（行の StaffKey との照合用）
    public List<string> RawValues;                   // 吸収した生値（統合トグルのツールチップ用・初出順）
}

// 差押リストの対象金融機関の照合パターン1件。
// JUDGE（judge.ps1）の target_institutions.json と同一の書式:
// ワイルドカードを含まない項目は正規化キーの完全一致、「*」（任意の0文字以上）
// 「?」（任意の1文字）を含む項目はグロブ照合。全角の＊？は正規化（NFKC）が
// 半角へ片寄せするためどちらでも書ける
public class SeizureTargetPattern
{
    public string ExactKey;                          // ワイルドカードなし（完全一致キー。Glob と排他）
    public Regex Glob;                               // ワイルドカードあり（^…$ の全体一致に変換済み）
}

// 差押登録ログの1件（history/{ユーザー名}/seizure_log.json のエントリ）
// 書き込みは entry_deposit 側（WriteSeizureLog）が行い、本ツールは読むのみ
public class SeizureLogEntry
{
    public bool IsInsurance { get; set; }            // 生命保険の登録か（kind が insurance。kind のないエントリは預金）
    public string AddressNumber { get; set; }        // 宛名番号
    public string Name { get; set; }                 // 名義人氏名
    public string InstitutionName { get; set; }      // 金融機関名（生命保険は保険会社名）
    public string ExecutionDate { get; set; }        // 執行日（yyyy-MM-dd で記録される）
    public string DocumentNumber { get; set; }       // 文書番号
    public DateTime RegisteredAt { get; set; }       // 登録日時（パース不能時は DateTime.MinValue）
    public string RegisteredAtRaw { get; set; }      // 登録日時の生文字列（パース不能時の表示用）
}

// 最近見た人の履歴エントリ1件（history/{ユーザー名}/recent_people.json）
public class RecentPersonEntry
{
    public DateTime ViewedAt { get; set; }           // 閲覧日時（照会ファイルを開いた日時。パース不能時は DateTime.MinValue）
    public string ViewedAtRaw { get; set; }          // 閲覧日時の生文字列（パース不能時の表示用）
    public int Year { get; set; }                    // 行の登録日から導出した年度（導出不能時は -1 =「年度不明」）
    public bool IsInsurance { get; set; }            // 生命保険の行か（kind が insurance。kind のないエントリは預金）
    public string[] Fields { get; set; }             // 記録時点の台帳行（種別の並び。預金 27〜28 列・生保 122 列。履歴に残す列だけ値を持ちほかは空。台帳は追記専用のため記録後も正確なまま）
}

// 人物メモ1件（記入者は history/{ユーザー名}/ のフォルダ名から取得する）
public class PersonNoteEntry
{
    public string Author { get; set; }               // 記入者（Windows ユーザー名 = 職員番号）
    public string Text { get; set; }                 // メモ本文（自由テキスト・複数行可）
    public DateTime CreatedAt { get; set; }          // 作成日時（パース不能時は DateTime.MinValue）
    public string CreatedAtRaw { get; set; }         // 作成日時の生文字列（パース不能時の表示用）
}

// アップデート告知（update_notice.json のデシリアライズ用）
public class UpdateNotice
{
    public string Version { get; set; }
    public List<NoticeFeature> Features { get; set; }
    public UpdateNotice() { Features = new List<NoticeFeature>(); }
}
public class NoticeFeature
{
    public string Title { get; set; }
    public string Description { get; set; }
    public List<NoticeSection> Sections { get; set; }
    public NoticeFeature() { Sections = new List<NoticeSection>(); }
}
public class NoticeSection
{
    public string Heading { get; set; }
    public List<string> Items { get; set; }
    public NoticeSection() { Items = new List<string>(); }
}

// ==============================================================
// JSON パーサー（手書き・外部ライブラリ不要）
// LGWAN 環境では NuGet パッケージが使えないため手動パース
// ==============================================================

public static class JsonHelper
{
    // JSON文字列から指定キーの文字列値を取得（値が文字列でなければ null。エスケープは UnescapeJsonString で解く）
    public static string GetString(string json, string key)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return null;
        var colonIdx = json.IndexOf(':', keyIdx + key.Length + 2);
        if (colonIdx < 0) return null;

        var rest = json.Substring(colonIdx + 1).TrimStart();
        if (rest.Length == 0 || rest[0] != '"') return null;

        // 文字列リテラルの終わり（エスケープされていない "）までを生のまま切り出す
        var sb = new StringBuilder();
        bool escaped = false;
        for (int i = 1; i < rest.Length; i++)
        {
            if (escaped) { sb.Append(rest[i]); escaped = false; continue; }
            if (rest[i] == '\\') { sb.Append(rest[i]); escaped = true; continue; }
            if (rest[i] == '"') break;
            sb.Append(rest[i]);
        }
        return UnescapeJsonString(sb.ToString());
    }

    // JSON文字列から指定キーの整数値を取得（値の先頭が数字・- でなければ既定値。null や文字列の値のとき
    // 後方にある別の数字を拾わない）
    public static int GetInt(string json, string key, int defaultValue = 0)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return defaultValue;
        var colonIdx = json.IndexOf(':', keyIdx + key.Length + 2);
        if (colonIdx < 0) return defaultValue;

        var rest = json.Substring(colonIdx + 1).TrimStart();
        if (rest.Length == 0 || !(char.IsDigit(rest[0]) || rest[0] == '-')) return defaultValue;
        var numStr = new StringBuilder();
        foreach (var c in rest)
        {
            if (char.IsDigit(c) || c == '-') numStr.Append(c);
            else break;
        }
        int result;
        return int.TryParse(numStr.ToString(), out result) ? result : defaultValue;
    }

    // 対応する閉じ括弧の位置を返す
    public static int FindMatchingBracket(string json, int openIdx, char open, char close)
    {
        int depth = 0;
        bool inString = false;
        bool escaped = false;
        for (int i = openIdx; i < json.Length; i++)
        {
            if (escaped) { escaped = false; continue; }
            if (json[i] == '\\' && inString) { escaped = true; continue; }
            if (json[i] == '"') { inString = !inString; continue; }
            if (inString) continue;
            if (json[i] == open) depth++;
            else if (json[i] == close) { depth--; if (depth == 0) return i; }
        }
        return -1;
    }

    // JSON 文字列値のエスケープを解除する（\n \t \r は制御文字・それ以外は次の 1 文字をそのまま）
    public static string UnescapeJsonString(string value)
    {
        var sb = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++)
        {
            if (value[i] == '\\' && i + 1 < value.Length)
            {
                i++;
                switch (value[i])
                {
                    case 'n':  sb.Append('\n'); break;
                    case 't':  sb.Append('\t'); break;
                    case 'r':  sb.Append('\r'); break;
                    default:   sb.Append(value[i]); break;   // \\ \" \/ はそのまま
                }
            }
            else sb.Append(value[i]);
        }
        return sb.ToString();
    }
}