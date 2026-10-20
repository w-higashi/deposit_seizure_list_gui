// ==============================================================================
// paddock_insurance_logic.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 画面を持たない処理のうち、生命保険の台帳の列の並びと値の解釈（見出しの対応づけ・
// 金額と日付の読み・注記の判定・詳細パネルの展開部の行と差引見込額（参考）・行の一意キー）。
// BusinessLogic の一部（partial class）。ツールの説明とビルドの方法は
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

public static partial class BusinessLogic
{
    // ---------- 生命保険の台帳: 列の並び ----------

    // 生命保険の台帳の列の並び（JUDGE が書く122列・judge.ps1 の生命保険の台帳の列と同一順序）。
    // 台帳は見出しの列名で読み（列が多く、位置で読むと列の追加のたびに保守が要るため）、
    // この並びへ組み直して持つ。列の追加は末尾だけのため、途中の位置は変わらない
    public static readonly string[] INSURANCE_LEDGER_COLUMNS =
    {
        "記録日時", "登録日", "回答基準日", "宛名番号",
        "契約者カナ氏名", "契約者漢字氏名", "契約者英字名", "契約者生年月日", "契約者郵便番号", "契約者住所",
        "契約者電話番号1", "契約者電話番号2", "契約者勤務先名", "契約者勤務先電話番号",
        "保険会社名", "保険会社コード", "契約有無", "契約有無（回答書）", "顧客番号", "法人番号",
        "証券番号", "契約の状態", "保険料", "支払方法", "払込の別", "契約種類", "契約年月日", "契約期間",
        "払込満了年月日", "据置金の有無", "満期年月日",
        "解約返戻金", "解約済_解約返戻金", "解約済_解約年月日", "解約済_支払方法", "解約済_振込先金融機関コード",
        "解約済_振込先金融機関名", "解約済_振込先支店番号", "解約済_振込先支店名", "解約済_振込先口座番号",
        "解約済_振込先口座種別", "解約済_振込先口座名義人",
        "配当の有無", "配当金額", "配当年月日", "配当利息",
        "貸付金の有無", "貸付金の金額", "未払い保険料の有無", "未払い保険料の金額", "未払い利息の有無",
        "未払い利息の金額", "前払い保険料の有無", "前払い保険料の金額", "据置金額",
        "差押権利者の有無", "差押年月日", "差押権利者名", "介入権行使の可否", "介入権者の明細",
        "質権者の有無", "質権設定年月日", "質権者名",
        "被保険者1_番号", "被保険者1_漢字氏名", "被保険者1_カナ氏名", "被保険者1_英字氏名", "被保険者1_生年月日",
        "被保険者2_番号", "被保険者2_漢字氏名", "被保険者2_カナ氏名", "被保険者2_英字氏名", "被保険者2_生年月日",
        "受取人1_種類", "受取人1_漢字氏名", "受取人1_カナ氏名", "受取人1_英字氏名", "受取人1_続柄",
        "受取人2_種類", "受取人2_漢字氏名", "受取人2_カナ氏名", "受取人2_英字氏名", "受取人2_続柄",
        "指定代理請求人",
        "振込元_金融機関コード", "振込元_金融機関名", "振込元_支店番号", "振込元_支店名", "振込元_口座種別",
        "振込元_口座番号", "振込元_口座名義人",
        "給付金1_説明", "給付金1_金額", "給付金2_説明", "給付金2_金額", "給付金3_説明", "給付金3_金額",
        "給付金4_説明", "給付金4_金額",
        "年金開始日", "年金満了日", "年金支払期間", "年金一括受取試算額",
        "保険金の支払事由", "保険事故発生年月日", "保険金支払年月日", "支払保険金", "保険金支払方法",
        "保険金_振込先金融機関コード", "保険金_振込先金融機関名", "保険金_振込先支店番号", "保険金_振込先支店名",
        "保険金_振込先口座番号", "保険金_振込先口座種別", "保険金_振込先口座名義人",
        "登録推奨", "照会結果ファイル名", "シート名", "受付番号", "保険会社名称（回答書）", "担当職員名", "取組区分"
    };

    public const int INSURANCE_LEDGER_COLUMN_COUNT = 122;

    // この並びの中の位置（一覧・絞り込み・詳細パネルで使う列だけ。配列との対応は検査で確かめる）
    public const int INSURANCE_COL_REG_DATE           = 1;    // 登録日
    public const int INSURANCE_COL_BASE_DATE          = 2;    // 回答基準日
    public const int INSURANCE_COL_ADDRESS_NUM        = 3;    // 宛名番号
    public const int INSURANCE_COL_KANA               = 4;    // 契約者カナ氏名
    public const int INSURANCE_COL_NAME               = 5;    // 契約者漢字氏名
    public const int INSURANCE_COL_BIRTH_DATE         = 7;    // 契約者生年月日
    public const int INSURANCE_COL_ADDRESS            = 9;    // 契約者住所
    public const int INSURANCE_COL_COMPANY            = 14;   // 保険会社名（補正後）
    public const int INSURANCE_COL_HAS_CONTRACT       = 16;   // 契約有無（1=契約あり／2=電子回答不可／それ以外=契約なし）
    public const int INSURANCE_COL_POLICY_NUM         = 20;   // 証券番号
    public const int INSURANCE_COL_CONTRACT_STATUS    = 21;   // 契約の状態
    public const int INSURANCE_COL_CONTRACT_TYPE      = 25;   // 契約種類
    public const int INSURANCE_COL_CONTRACT_DATE      = 26;   // 契約年月日
    public const int INSURANCE_COL_PAID_UP_DATE       = 28;   // 払込満了年月日
    public const int INSURANCE_COL_MATURITY           = 30;   // 満期年月日
    public const int INSURANCE_COL_SURRENDER          = 31;   // 解約返戻金
    public const int INSURANCE_COL_CANCELED_SURRENDER = 32;   // 解約済_解約返戻金
    public const int INSURANCE_COL_CANCELED_DATE      = 33;   // 解約済_解約年月日
    public const int INSURANCE_COL_CANCELED_METHOD    = 34;   // 解約済_支払方法
    public const int INSURANCE_COL_DIVIDEND_FLAG      = 42;   // 配当の有無
    public const int INSURANCE_COL_DIVIDEND_AMOUNT    = 43;   // 配当金額
    public const int INSURANCE_COL_DIVIDEND_DATE      = 44;   // 配当年月日
    public const int INSURANCE_COL_DIVIDEND_INTEREST  = 45;   // 配当利息
    public const int INSURANCE_COL_LOAN_FLAG          = 46;   // 貸付金の有無
    public const int INSURANCE_COL_LOAN_AMOUNT        = 47;   // 貸付金の金額
    public const int INSURANCE_COL_UNPAID_FLAG        = 48;   // 未払い保険料の有無
    public const int INSURANCE_COL_UNPAID_AMOUNT      = 49;   // 未払い保険料の金額
    public const int INSURANCE_COL_UNPAID_INTEREST_FLAG   = 50;   // 未払い利息の有無
    public const int INSURANCE_COL_UNPAID_INTEREST_AMOUNT = 51;   // 未払い利息の金額
    public const int INSURANCE_COL_PREPAID_FLAG       = 52;   // 前払い保険料の有無
    public const int INSURANCE_COL_PREPAID_AMOUNT     = 53;   // 前払い保険料の金額
    public const int INSURANCE_COL_DEFERRED_AMOUNT    = 54;   // 据置金額
    public const int INSURANCE_COL_SEIZURE_FLAG       = 55;   // 差押権利者の有無
    public const int INSURANCE_COL_SEIZURE_DATE       = 56;   // 差押年月日
    public const int INSURANCE_COL_SEIZURE_HOLDER     = 57;   // 差押権利者名
    public const int INSURANCE_COL_INTERVENTION       = 58;   // 介入権行使の可否
    public const int INSURANCE_COL_INTERVENTION_DETAIL = 59;  // 介入権者の明細
    public const int INSURANCE_COL_PLEDGE_FLAG        = 60;   // 質権者の有無
    public const int INSURANCE_COL_PLEDGE_DATE        = 61;   // 質権設定年月日
    public const int INSURANCE_COL_PLEDGE_HOLDER      = 62;   // 質権者名
    public const int INSURANCE_COL_ANNUITY_START      = 99;   // 年金開始日
    public const int INSURANCE_COL_ANNUITY_END        = 100;  // 年金満了日
    public const int INSURANCE_COL_ANNUITY_PERIOD     = 101;  // 年金支払期間
    public const int INSURANCE_COL_ANNUITY_LUMP       = 102;  // 年金一括受取試算額
    public const int INSURANCE_COL_CLAIM_REASON       = 103;  // 保険金の支払事由
    public const int INSURANCE_COL_CLAIM_PAID_DATE    = 105;  // 保険金支払年月日
    public const int INSURANCE_COL_CLAIM_PAID         = 106;  // 支払保険金
    public const int INSURANCE_COL_SOURCE_FILE        = 116;  // 照会結果ファイル名
    public const int INSURANCE_COL_SHEET_NAME         = 117;  // シート名
    public const int INSURANCE_COL_STAFF              = 120;  // 担当職員名

    // 最近見た人の履歴に残す列（一覧の行を組む BuildInsuranceRecord・注記の旗・注記のツールチップが読む列・行キーの列・登録日）。
    // ほかの列（契約者の生年月日・住所・被保険者・受取人・振込元・保険料・契約期間・照会結果ファイル名・シート名など）は履歴に書かない
    public static readonly int[] RECENT_KEPT_INSURANCE_COLUMNS =
    {
        INSURANCE_COL_REG_DATE, INSURANCE_COL_BASE_DATE, INSURANCE_COL_ADDRESS_NUM, INSURANCE_COL_KANA, INSURANCE_COL_NAME,
        INSURANCE_COL_COMPANY, INSURANCE_COL_HAS_CONTRACT, INSURANCE_COL_POLICY_NUM, INSURANCE_COL_CONTRACT_STATUS,
        INSURANCE_COL_CONTRACT_TYPE, INSURANCE_COL_MATURITY, INSURANCE_COL_SURRENDER, INSURANCE_COL_CANCELED_SURRENDER,
        INSURANCE_COL_CANCELED_DATE, INSURANCE_COL_CANCELED_METHOD, INSURANCE_COL_DIVIDEND_FLAG, INSURANCE_COL_DIVIDEND_AMOUNT,
        INSURANCE_COL_DIVIDEND_DATE, INSURANCE_COL_DIVIDEND_INTEREST, INSURANCE_COL_LOAN_FLAG, INSURANCE_COL_LOAN_AMOUNT,
        INSURANCE_COL_UNPAID_FLAG, INSURANCE_COL_UNPAID_AMOUNT, INSURANCE_COL_SEIZURE_FLAG, INSURANCE_COL_SEIZURE_DATE,
        INSURANCE_COL_SEIZURE_HOLDER, INSURANCE_COL_PLEDGE_FLAG, INSURANCE_COL_PLEDGE_DATE, INSURANCE_COL_PLEDGE_HOLDER,
        INSURANCE_COL_ANNUITY_START, INSURANCE_COL_ANNUITY_END, INSURANCE_COL_ANNUITY_PERIOD, INSURANCE_COL_ANNUITY_LUMP,
        INSURANCE_COL_CLAIM_REASON, INSURANCE_COL_CLAIM_PAID_DATE, INSURANCE_COL_CLAIM_PAID, INSURANCE_COL_STAFF
    };

    // 見出しに要る列（欠ければその年度は読めない）
    public static readonly string[] INSURANCE_REQUIRED_COLUMNS =
        { "登録日", "回答基準日", "宛名番号", "保険会社名", "契約有無", "証券番号", "照会結果ファイル名", "シート名" };

    // ---------- 生命保険の台帳: 値の読み ----------

    // 台帳の日付の列（yyyy/MM/dd・年 1900〜2099）を日付にする。読めなければ null。
    // 登録日・回答基準日は形式非依存の ParseLedgerDate で読み、こちらは JUDGE が読めた日付を
    // yyyy/MM/dd に整えて書くそのほかの日付の列（契約年月日・満期年月日など）に当てる。
    // 読めなかった生値（終身の 99999999・「終身払込」など）は日付として扱わない
    public static DateTime? ParseInsuranceDate(string text)
    {
        var value = ParseDisplayDate(text);
        if (value == null || value.Value.Year < MIN_YEAR || value.Value.Year > MAX_YEAR) return null;
        return value;
    }

    // 有無の欄（「有」「無」「電子回答不可」など）が「有」を含むか（値の文字列そのもので見る）
    public static bool HasInsuranceYes(string value)
    {
        return !string.IsNullOrEmpty(value) && value.IndexOf('有') >= 0;
    }

    // 金額の列が数値として読めて 0 でないか（解約していない契約でも 0 が書かれることがあるため、0 は値なし）
    private static bool HasInsuranceNonZeroAmount(string text)
    {
        var value = ParseLedgerAmount(text);
        return value != null && value.Value != 0;
    }

    // ---------- 生命保険の台帳: 注記 ----------

    // 注記の固定順（チップの走査順・詳細検索の注記の条件と共有）
    public static readonly string[] INSURANCE_NOTE_NAMES =
        { "差押", "質権", "貸付金", "未払保険料", "配当", "年金", "解約済", "保険金支払済" };

    // 注記ごとの暗色ツールチップ（第2層）に並べる項目（項目名と台帳の列）。INSURANCE_NOTE_NAMES と同順
    private static readonly string[][] InsuranceNoteDetailLabels =
    {
        new[] { "差押年月日", "差押権利者名" },
        new[] { "質権設定年月日", "質権者名" },
        new[] { "貸付金の金額" },
        new[] { "未払い保険料の金額" },
        new[] { "配当金額", "配当年月日", "配当利息" },
        new[] { "年金開始日", "年金満了日", "年金支払期間", "年金一括受取試算額" },
        new[] { "解約年月日", "解約返戻金", "支払方法" },
        new[] { "保険金の支払事由", "保険金支払年月日", "支払保険金" }
    };
    private static readonly int[][] InsuranceNoteDetailColumns =
    {
        new[] { INSURANCE_COL_SEIZURE_DATE, INSURANCE_COL_SEIZURE_HOLDER },
        new[] { INSURANCE_COL_PLEDGE_DATE, INSURANCE_COL_PLEDGE_HOLDER },
        new[] { INSURANCE_COL_LOAN_AMOUNT },
        new[] { INSURANCE_COL_UNPAID_AMOUNT },
        new[] { INSURANCE_COL_DIVIDEND_AMOUNT, INSURANCE_COL_DIVIDEND_DATE, INSURANCE_COL_DIVIDEND_INTEREST },
        new[] { INSURANCE_COL_ANNUITY_START, INSURANCE_COL_ANNUITY_END, INSURANCE_COL_ANNUITY_PERIOD, INSURANCE_COL_ANNUITY_LUMP },
        new[] { INSURANCE_COL_CANCELED_DATE, INSURANCE_COL_CANCELED_SURRENDER, INSURANCE_COL_CANCELED_METHOD },
        new[] { INSURANCE_COL_CLAIM_REASON, INSURANCE_COL_CLAIM_PAID_DATE, INSURANCE_COL_CLAIM_PAID }
    };

    // 注記8種の立つ・立たない（台帳の並びの行から。INSURANCE_NOTE_NAMES と同順）。
    // 有無の列は「有」を含むとき、日付・金額の列はどちらかが「日付として読める」「数値として読めて 0 でない」
    // ときに立つ。記録された値の圧縮表示であり評価ではない（差引などの計算はしない）
    public static bool[] BuildInsuranceNoteFlags(string[] fields)
    {
        var flags = new bool[INSURANCE_NOTE_NAMES.Length];
        flags[0] = HasInsuranceYes(fields[INSURANCE_COL_SEIZURE_FLAG]);
        flags[1] = HasInsuranceYes(fields[INSURANCE_COL_PLEDGE_FLAG]);
        flags[2] = HasInsuranceYes(fields[INSURANCE_COL_LOAN_FLAG]);
        flags[3] = HasInsuranceYes(fields[INSURANCE_COL_UNPAID_FLAG]);
        flags[4] = HasInsuranceYes(fields[INSURANCE_COL_DIVIDEND_FLAG]);
        flags[5] = ParseInsuranceDate(fields[INSURANCE_COL_ANNUITY_START]) != null || HasInsuranceNonZeroAmount(fields[INSURANCE_COL_ANNUITY_LUMP]);
        flags[6] = ParseInsuranceDate(fields[INSURANCE_COL_CANCELED_DATE]) != null || HasInsuranceNonZeroAmount(fields[INSURANCE_COL_CANCELED_SURRENDER]);
        flags[7] = ParseInsuranceDate(fields[INSURANCE_COL_CLAIM_PAID_DATE]) != null || HasInsuranceNonZeroAmount(fields[INSURANCE_COL_CLAIM_PAID]);
        return flags;
    }

    // 注記1種の暗色ツールチップの原文: 項目の「項目名 値」を全角空白でつなぐ（値のない項目は省く・
    // すべて空なら「—」）。値は台帳のまま
    public static string BuildInsuranceNoteDetail(string[] fields, int noteIndex)
    {
        var labels = InsuranceNoteDetailLabels[noteIndex];
        var columns = InsuranceNoteDetailColumns[noteIndex];
        var parts = new List<string>();
        for (int i = 0; i < columns.Length; i++)
        {
            var value = fields[columns[i]].Trim();
            if (value != "") parts.Add(labels[i] + " " + value);
        }
        return parts.Count > 0 ? string.Join("　", parts.ToArray()) : "—";
    }

    // ---------- 生命保険の台帳: 詳細パネルの展開部 ----------

    // 展開部の行の種類。文字＝列の値そのまま／金額＝数値として読めて 0 でない値か読めない生値／
    // 金額と日付＝配当（金額（日付））／日付と名＝差押・質権（日付　名）／差引見込額（参考）
    private enum InsurancePanelRowKind { Text, Amount, AmountDate, DateName, Net }

    // 展開部の行の仕様（ラベル・種類・列・有無の列）。有無の列は、組の金額・日付・名がどれも出ないのに
    // 「有」を含むときだけ「有」の字を値として出すために見る（注記のチップと展開部が食い違わないように）
    private class InsurancePanelRowSpec
    {
        public string Label;
        public InsurancePanelRowKind Kind;
        public int Column;        // 文字・金額の列。金額と日付は金額の列・日付と名は日付の列
        public int SecondColumn;  // 金額と日付は日付の列・日付と名は名の列（ほかは -1）
        public int FlagColumn;    // 有無の列（見ない行は -1）
        public bool Mono;         // 文字の行を Consolas で描く（日付の列）
    }

    private static InsurancePanelRowSpec PanelText(int column, bool mono)
    {
        return new InsurancePanelRowSpec { Label = INSURANCE_LEDGER_COLUMNS[column], Kind = InsurancePanelRowKind.Text, Column = column, SecondColumn = -1, FlagColumn = -1, Mono = mono };
    }

    private static InsurancePanelRowSpec PanelAmount(int column, int flagColumn)
    {
        return new InsurancePanelRowSpec { Label = INSURANCE_LEDGER_COLUMNS[column], Kind = InsurancePanelRowKind.Amount, Column = column, SecondColumn = -1, FlagColumn = flagColumn, Mono = true };
    }

    private static InsurancePanelRowSpec PanelAmountDate(string label, int amountColumn, int dateColumn, int flagColumn)
    {
        return new InsurancePanelRowSpec { Label = label, Kind = InsurancePanelRowKind.AmountDate, Column = amountColumn, SecondColumn = dateColumn, FlagColumn = flagColumn, Mono = true };
    }

    private static InsurancePanelRowSpec PanelDateName(string label, int dateColumn, int nameColumn, int flagColumn)
    {
        return new InsurancePanelRowSpec { Label = label, Kind = InsurancePanelRowKind.DateName, Column = dateColumn, SecondColumn = nameColumn, FlagColumn = flagColumn, Mono = true };
    }

    // 差引見込額（参考）: 生命保険差押予定一覧 作成ツール（entry_insurance.exe）と同じ式。
    // 項は台帳の金額の列（空欄は 0）で、数値として読めない生値が1つでもあれば計算しない
    public const string INSURANCE_NET_LABEL = "差引見込額（参考）";
    public const string INSURANCE_NET_FORMULA = "解約返戻金 ＋ 配当金額 − 貸付金の金額 − 未払い保険料の金額 − 未払い利息の金額 ＋ 前払い保険料の金額";
    private static readonly int[] InsuranceNetColumns =
        { INSURANCE_COL_SURRENDER, INSURANCE_COL_DIVIDEND_AMOUNT, INSURANCE_COL_LOAN_AMOUNT, INSURANCE_COL_UNPAID_AMOUNT, INSURANCE_COL_UNPAID_INTEREST_AMOUNT, INSURANCE_COL_PREPAID_AMOUNT };
    private static readonly int[] InsuranceNetSigns = { 1, 1, -1, -1, -1, 1 };

    // 展開部のブロックの見出しと行の仕様（並びは固定）。解約返戻金はカードの右上に出すため置かない。
    // 差引見込額（参考）はそのブロックにほかの行があるときだけ出す（なければ差引は解約返戻金そのもの）
    private static readonly string[] INSURANCE_PANEL_BLOCK_TITLES = { "契約", "返戻金・配当・貸付金等", "差押・質権・介入権" };
    private static readonly InsurancePanelRowSpec[][] INSURANCE_PANEL_BLOCKS =
    {
        new[]
        {
            PanelText(INSURANCE_COL_CONTRACT_STATUS, false),
            PanelText(INSURANCE_COL_CONTRACT_DATE, true),
            PanelText(INSURANCE_COL_PAID_UP_DATE, true),
            PanelText(INSURANCE_COL_MATURITY, true)
        },
        new[]
        {
            PanelAmountDate("配当", INSURANCE_COL_DIVIDEND_AMOUNT, INSURANCE_COL_DIVIDEND_DATE, INSURANCE_COL_DIVIDEND_FLAG),
            PanelAmount(INSURANCE_COL_DIVIDEND_INTEREST, -1),
            PanelAmount(INSURANCE_COL_LOAN_AMOUNT, INSURANCE_COL_LOAN_FLAG),
            PanelAmount(INSURANCE_COL_UNPAID_AMOUNT, INSURANCE_COL_UNPAID_FLAG),
            PanelAmount(INSURANCE_COL_UNPAID_INTEREST_AMOUNT, INSURANCE_COL_UNPAID_INTEREST_FLAG),
            PanelAmount(INSURANCE_COL_PREPAID_AMOUNT, INSURANCE_COL_PREPAID_FLAG),
            PanelAmount(INSURANCE_COL_DEFERRED_AMOUNT, -1),
            new InsurancePanelRowSpec { Label = INSURANCE_NET_LABEL, Kind = InsurancePanelRowKind.Net, Column = -1, SecondColumn = -1, FlagColumn = -1, Mono = true }
        },
        new[]
        {
            PanelDateName("差押", INSURANCE_COL_SEIZURE_DATE, INSURANCE_COL_SEIZURE_HOLDER, INSURANCE_COL_SEIZURE_FLAG),
            PanelDateName("質権", INSURANCE_COL_PLEDGE_DATE, INSURANCE_COL_PLEDGE_HOLDER, INSURANCE_COL_PLEDGE_FLAG),
            PanelText(INSURANCE_COL_INTERVENTION, false),
            PanelText(INSURANCE_COL_INTERVENTION_DETAIL, false)
        }
    };

    // 差引見込額（参考）を計算する。unreadable には数値として読めない生値の列名を集め、
    // 1つでもあれば null（空欄は 0 として足す）
    public static long? ComputeInsuranceNetValue(string[] fields, out List<string> unreadable)
    {
        unreadable = new List<string>();
        long total = 0;
        for (int i = 0; i < InsuranceNetColumns.Length; i++)
        {
            var text = fields[InsuranceNetColumns[i]];
            if (string.IsNullOrWhiteSpace(text)) continue;
            var value = ParseLedgerAmount(text);
            if (value == null) { unreadable.Add(INSURANCE_LEDGER_COLUMNS[InsuranceNetColumns[i]]); continue; }
            total += InsuranceNetSigns[i] * value.Value;
        }
        return unreadable.Count > 0 ? (long?)null : total;
    }

    // 金額の列の片を作る。数値として読めて 0 でなければ台帳の値のまま、読めない生値は薄文字色の印つきで
    // 返す。0・空欄は null（行に出さない）
    private static InsurancePanelPart AmountPart(string text, out bool rawText)
    {
        rawText = false;
        var trimmed = text.Trim();
        if (trimmed == "") return null;
        var value = ParseLedgerAmount(trimmed);
        if (value == null) { rawText = true; return new InsurancePanelPart { Text = trimmed, Mono = true }; }
        if (value.Value == 0) return null;
        return new InsurancePanelPart { Text = trimmed, Mono = true };
    }

    private static InsurancePanelRow MakePanelRow(string label, bool rawText, params InsurancePanelPart[] parts)
    {
        return new InsurancePanelRow { Label = label, Parts = new List<InsurancePanelPart>(parts), IsRawText = rawText };
    }

    // 行の仕様1つから展開部の行を作る（出さない行は null）
    private static InsurancePanelRow BuildInsurancePanelRow(InsurancePanelRowSpec spec, string[] fields)
    {
        bool rawText;
        switch (spec.Kind)
        {
            case InsurancePanelRowKind.Text:
            {
                var text = fields[spec.Column].Trim();
                return text == "" ? null : MakePanelRow(spec.Label, false, new InsurancePanelPart { Text = text, Mono = spec.Mono });
            }
            case InsurancePanelRowKind.Amount:
            {
                var part = AmountPart(fields[spec.Column], out rawText);
                if (part != null) return MakePanelRow(spec.Label, rawText, part);
                break;
            }
            case InsurancePanelRowKind.AmountDate:
            {
                var amount = AmountPart(fields[spec.Column], out rawText);
                var date = fields[spec.SecondColumn].Trim();
                if (amount != null && date != "") return MakePanelRow(spec.Label, rawText, amount, new InsurancePanelPart { Text = "（" + date + "）", Mono = true });
                if (amount != null) return MakePanelRow(spec.Label, rawText, amount);
                if (date != "") return MakePanelRow(spec.Label, false, new InsurancePanelPart { Text = date, Mono = true });
                break;
            }
            case InsurancePanelRowKind.DateName:
            {
                var date = fields[spec.Column].Trim();
                var name = fields[spec.SecondColumn].Trim();
                if (date != "" && name != "") return MakePanelRow(spec.Label, false, new InsurancePanelPart { Text = date, Mono = true }, new InsurancePanelPart { Text = "　" + name, Mono = false });
                if (date != "") return MakePanelRow(spec.Label, false, new InsurancePanelPart { Text = date, Mono = true });
                if (name != "") return MakePanelRow(spec.Label, false, new InsurancePanelPart { Text = name, Mono = false });
                break;
            }
        }
        // 組の値が出ないのに有無の列が「有」を含む行は「有」の字を値として出す
        if (spec.FlagColumn >= 0 && HasInsuranceYes(fields[spec.FlagColumn]))
            return MakePanelRow(spec.Label, false, new InsurancePanelPart { Text = "有", Mono = false });
        return null;
    }

    // 契約カードの展開部のブロックと行を組み立てる（値のある行だけ・行のないブロックは返さない）
    public static List<InsurancePanelBlock> BuildInsurancePanelBlocks(string[] fields)
    {
        var blocks = new List<InsurancePanelBlock>();
        for (int b = 0; b < INSURANCE_PANEL_BLOCKS.Length; b++)
        {
            var rows = new List<InsurancePanelRow>();
            foreach (var spec in INSURANCE_PANEL_BLOCKS[b])
            {
                if (spec.Kind == InsurancePanelRowKind.Net)
                {
                    if (rows.Count == 0) continue;
                    List<string> unreadable;
                    var net = ComputeInsuranceNetValue(fields, out unreadable);
                    var row = net != null
                        ? MakePanelRow(spec.Label, false, new InsurancePanelPart { Text = net.Value.ToString("#,0"), Mono = true })
                        : MakePanelRow(spec.Label, false, new InsurancePanelPart { Text = "—", Mono = false });
                    row.ToolTip = net != null
                        ? INSURANCE_NET_FORMULA
                        : string.Join("・", unreadable.ToArray()) + "が数値として読めないため計算できません\n" + INSURANCE_NET_FORMULA;
                    rows.Add(row);
                    continue;
                }
                var built = BuildInsurancePanelRow(spec, fields);
                if (built != null) rows.Add(built);
            }
            if (rows.Count > 0) blocks.Add(new InsurancePanelBlock { Title = INSURANCE_PANEL_BLOCK_TITLES[b], Rows = rows });
        }
        return blocks;
    }

    // ---------- 生命保険の台帳: 行の一意キー ----------

    // 契約あり行の一意キー（宛名番号 × 保険会社名 × 証券番号 × 回答基準日）。
    // 詳細パネルのカードの照合に使う。回答基準日は正準形（FormatLedgerDate）で組み、
    // 台帳の日付形式が変わっても同一行と照合できる（預金の行キーと同じ考え方）
    public static string GetInsuranceLedgerRowKey(string[] fields, Dictionary<int, EraEntry> eraMap)
    {
        return fields[INSURANCE_COL_ADDRESS_NUM] + "|" + fields[INSURANCE_COL_COMPANY] + "|" +
               fields[INSURANCE_COL_POLICY_NUM] + "|" + FormatLedgerDate(fields[INSURANCE_COL_BASE_DATE], eraMap);
    }
}