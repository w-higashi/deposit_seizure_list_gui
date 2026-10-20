// ==============================================================================
// paddock_logic.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 画面を持たない処理のうち、台帳の列の並びと値の解釈（CSV・日付と和暦・年度・
// 行キー・シート名・残高と反対債権・検索の正規化・差押リストの対象の照合・
// 入力欄の補正・検出の解析・担当の表記ゆれ統合）。明細を見るの処理は
// paddock_detail_logic.cs にある。
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

// ==============================================================
// ビジネスロジック
// ==============================================================

public static partial class BusinessLogic
{
    // 預金の台帳の列の並び（JUDGE が書く28列・judge.ps1 の LEDGER_COLUMNS と同一順序）。
    // 台帳は見出しの列名で読み、この並びへ組み直して持つ（並びどおりなら組み直さない）
    public static readonly string[] DEPOSIT_LEDGER_COLUMNS =
    {
        "記録日時", "登録日", "回答基準日", "宛名番号", "名義人氏名", "カナ氏名", "届出住所",
        "金融機関名", "支店名", "支店番号", "口座種別", "口座番号", "残高", "最終取引日",
        "証書貸付", "手形貸付", "割引手形", "当座貸越", "取引有無", "年金", "給与", "保険",
        "定期入金", "定期出金", "登録推奨", "照会結果ファイル名", "シート名", "担当職員名"
    };

    // 見出しに要る列（欠ければその年度は読めない。担当職員名は列の追加前の台帳にないため含めない）
    public static readonly string[] DEPOSIT_REQUIRED_COLUMNS =
        { "登録日", "回答基準日", "宛名番号", "名義人氏名", "金融機関名", "支店名", "支店番号", "口座種別", "口座番号", "残高", "取引有無", "照会結果ファイル名", "シート名" };

    // この並びの中の位置（配列との対応は検査で確かめる。
    // 現行は末尾に担当職員名を加えた28列。旧27列の行も列数チェックで受け付ける）
    public const int DEPOSIT_COL_RECORDED_AT   = 0;   // 記録日時
    public const int DEPOSIT_COL_REG_DATE      = 1;   // 登録日
    public const int DEPOSIT_COL_BASE_DATE     = 2;   // 回答基準日
    public const int DEPOSIT_COL_ADDRESS_NUM   = 3;   // 宛名番号
    public const int DEPOSIT_COL_NAME          = 4;   // 名義人氏名
    public const int DEPOSIT_COL_KANA          = 5;   // カナ氏名
    public const int DEPOSIT_COL_ADDRESS       = 6;   // 届出住所
    public const int DEPOSIT_COL_INSTITUTION   = 7;   // 金融機関名
    public const int DEPOSIT_COL_BRANCH_NAME   = 8;   // 支店名
    public const int DEPOSIT_COL_BRANCH_NUM    = 9;   // 支店番号
    public const int DEPOSIT_COL_ACCOUNT_TYPE  = 10;  // 口座種別
    public const int DEPOSIT_COL_ACCOUNT_NUM   = 11;  // 口座番号
    public const int DEPOSIT_COL_BALANCE       = 12;  // 残高
    public const int DEPOSIT_COL_LAST_TX       = 13;  // 最終取引日
    public const int DEPOSIT_COL_LOAN_DEED     = 14;  // 証書貸付
    public const int DEPOSIT_COL_LOAN_BILL     = 15;  // 手形貸付
    public const int DEPOSIT_COL_LOAN_DISCOUNT = 16;  // 割引手形
    public const int DEPOSIT_COL_OVERDRAFT     = 17;  // 当座貸越
    public const int DEPOSIT_COL_HAS_ACCOUNT   = 18;  // 取引有無
    public const int DEPOSIT_COL_PENSION       = 19;  // 年金
    public const int DEPOSIT_COL_SALARY        = 20;  // 給与
    public const int DEPOSIT_COL_INSURANCE     = 21;  // 保険
    public const int DEPOSIT_COL_REG_DEPOSIT   = 22;  // 定期入金
    public const int DEPOSIT_COL_REG_WITHDRAW  = 23;  // 定期出金
    public const int DEPOSIT_COL_RECOMMEND     = 24;  // 登録推奨
    public const int DEPOSIT_COL_SOURCE_FILE   = 25;  // 照会結果ファイル名
    public const int DEPOSIT_COL_SHEET_NAME    = 26;  // シート名
    public const int DEPOSIT_COL_STAFF         = 27;  // 担当職員名（28列目。旧27列の行には存在しない）

    public const int DEPOSIT_LEDGER_COLUMN_COUNT = 27;    // 台帳行の最低列数（27列=担当職員名なしの旧形式。28列の現行形式も受ける）

    // 最近見た人の履歴に残す列（一覧の行を組む BuildDepositRecord と検出チップが読む列・行キーの列・登録日）。
    // ほかの列（記録日時・届出住所・照会結果ファイル名・シート名）は履歴に書かない
    public static readonly int[] RECENT_KEPT_DEPOSIT_COLUMNS =
    {
        DEPOSIT_COL_REG_DATE, DEPOSIT_COL_BASE_DATE, DEPOSIT_COL_ADDRESS_NUM, DEPOSIT_COL_NAME, DEPOSIT_COL_KANA,
        DEPOSIT_COL_INSTITUTION, DEPOSIT_COL_BRANCH_NAME, DEPOSIT_COL_BRANCH_NUM, DEPOSIT_COL_ACCOUNT_TYPE,
        DEPOSIT_COL_ACCOUNT_NUM, DEPOSIT_COL_BALANCE, DEPOSIT_COL_LAST_TX, DEPOSIT_COL_LOAN_DEED, DEPOSIT_COL_LOAN_BILL,
        DEPOSIT_COL_LOAN_DISCOUNT, DEPOSIT_COL_OVERDRAFT, DEPOSIT_COL_HAS_ACCOUNT, DEPOSIT_COL_PENSION, DEPOSIT_COL_SALARY,
        DEPOSIT_COL_INSURANCE, DEPOSIT_COL_REG_DEPOSIT, DEPOSIT_COL_REG_WITHDRAW, DEPOSIT_COL_RECOMMEND, DEPOSIT_COL_STAFF
    };

    // 台帳の行の配列の中で、同じ値の文字列をプールの 1 本に共有する（読み込みごとのプール・ワーカーの中で公開前に呼ぶ。
    // 空の要素は string.Empty のまま）。値は変えないので比較・表示・検索の意味は変わらない
    public static void InternFields(string[] fields, Dictionary<string, string> pool)
    {
        for (int i = 0; i < fields.Length; i++)
        {
            var value = fields[i];
            if (value.Length == 0) continue;
            string shared;
            if (pool.TryGetValue(value, out shared)) fields[i] = shared;
            else pool[value] = value;
        }
    }

    // 履歴に残す列だけを写した新しい配列を返す（残さない列は空。旧 27 列の行は列の範囲内だけ写す）。
    // 元の配列は台帳の行と共有されているため変えない
    public static string[] TrimRecentFields(bool insurance, string[] fields)
    {
        var kept = insurance ? RECENT_KEPT_INSURANCE_COLUMNS : RECENT_KEPT_DEPOSIT_COLUMNS;
        var result = new string[fields.Length];
        for (int i = 0; i < result.Length; i++) result[i] = "";
        foreach (var column in kept)
            if (column < fields.Length) result[column] = fields[column];
        return result;
    }

    // 日付の受理範囲（年）。区切りなしの数字を西暦と解釈するときの妥当性と、明細・台帳の日付の上限下限に共通
    public const int MIN_YEAR = 1900;
    public const int MAX_YEAR = 2099;
    public const int SHEET_NAME_MAX = 31;   // Excel のシート名の最大文字数（STALL の Get-SafeSheetName と同じ）

    // CSVテキスト全体をRFC 4180準拠でパースし、レコード（フィールド配列）のリストを返す
    // クォート内の改行・カンマ・エスケープ（""→"）に対応する
    // （台帳の書き手 ConvertTo-CsvField は改行もエスケープ対象とするため、
    //   行単位の分割では理論上破綻しうる。ファイル全体を1回走査する方式をとる）
    public static List<string[]> ParseCsvText(string text)
    {
        var records = new List<string[]>();
        var fields = new List<string>();
        var sb = new StringBuilder();
        bool inQuotes = false;
        bool fieldStarted = false;   // 現在のレコードに1文字でも内容があるか（空行スキップ用）

        int i = 0;
        while (i < text.Length)
        {
            char c = text[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"') { sb.Append('"'); i += 2; continue; }
                    inQuotes = false;
                    i++;
                    continue;
                }
                sb.Append(c);
                i++;
                continue;
            }

            if (c == '"') { inQuotes = true; fieldStarted = true; i++; continue; }
            if (c == ',')
            {
                fields.Add(sb.ToString());
                sb.Length = 0;
                fieldStarted = true;
                i++;
                continue;
            }
            if (c == '\r' || c == '\n')
            {
                // レコード終端（CRLF は2文字まとめて消費）
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i += 2; else i++;
                if (fieldStarted || sb.Length > 0)
                {
                    fields.Add(sb.ToString());
                    records.Add(fields.ToArray());
                    fields.Clear();
                    sb.Length = 0;
                    fieldStarted = false;
                }
                continue;
            }

            sb.Append(c);
            i++;
        }

        // 末尾に改行がないファイルの最終レコードを回収
        if (fieldStarted || sb.Length > 0)
        {
            fields.Add(sb.ToString());
            records.Add(fields.ToArray());
        }

        return records;
    }

    // 見出しの行から台帳の並び columns への対応表を作る（map[i] = 並びの i 列目が見出しの何列目か。
    // 見出しにない列は -1）。同じ列名が2度あれば先の列を使い、見出しにあって並びにない列は読まない
    //（JUDGE が末尾に列を足しても動く）。先頭が並びの先頭の列名でなければ見出しでない（null・missing は null）、
    // 必須の列 required が欠ければ読めない（null・missing に欠けた列名）
    public static int[] BuildLedgerColumnMap(string[] header, string[] columns, string[] required, out List<string> missing)
    {
        missing = null;
        if (header == null || header.Length == 0 || header[0] != columns[0]) return null;

        var positions = new Dictionary<string, int>();
        for (int i = 0; i < header.Length; i++)
            if (!positions.ContainsKey(header[i])) positions[header[i]] = i;

        missing = new List<string>();
        foreach (var name in required)
            if (!positions.ContainsKey(name)) missing.Add(name);
        if (missing.Count > 0) return null;

        var map = new int[columns.Length];
        for (int i = 0; i < map.Length; i++)
        {
            int position;
            map[i] = positions.TryGetValue(columns[i], out position) ? position : -1;
        }
        return map;
    }

    // 対応表が台帳の並びどおり（見出しが minColumns 列以上で、見出しの範囲では位置ごとに同じ列名・
    // 見出しより後ろの列はない）か。そのときは行を組み直さずに持つ（生命保険の台帳は1行122列で数万行になるため）。
    // 並びの後ろに知らない列が続いていても位置は同じ。列の追加前の短い見出し（預金の27列）も並びどおり
    public static bool IsIdentityColumnMap(int[] map, int headerLength, int minColumns)
    {
        if (headerLength < minColumns) return false;
        for (int i = 0; i < map.Length; i++)
            if (map[i] != (i < headerLength ? i : -1)) return false;
        return true;
    }

    // 1行を台帳の並びへ組み直す（対応のない列は空欄）。fields の列数は見出しと同じであること
    public static string[] ReorderLedgerFields(string[] fields, int[] map)
    {
        var result = new string[map.Length];
        for (int i = 0; i < map.Length; i++)
            result[i] = map[i] >= 0 ? fields[map[i]] : "";
        return result;
    }

    // 7桁和暦 → DateTime 変換（era_mapping.json 使用）
    public static DateTime? WarekiToDate(string wareki, Dictionary<int, EraEntry> eraMap)
    {
        if (string.IsNullOrWhiteSpace(wareki) || wareki.Length != 7) return null;
        int eraCode, year, month, day;
        if (!int.TryParse(wareki.Substring(0, 1), out eraCode)) return null;
        if (!int.TryParse(wareki.Substring(1, 2), out year)) return null;
        if (!int.TryParse(wareki.Substring(3, 2), out month)) return null;
        if (!int.TryParse(wareki.Substring(5, 2), out day)) return null;

        EraEntry era;
        if (!eraMap.TryGetValue(eraCode, out era)) return null;

        try { return new DateTime(era.StartYear + year - 1, month, day); }
        catch { return null; }
    }

    // 台帳の日付文字列 → DateTime 変換（形式非依存）。次の順に解釈を試みる。
    //  1) 数字7桁  → 7桁和暦（WarekiToDate）
    //  2) 数字8桁  → 西暦 yyyyMMdd（年が1900〜2099の場合のみ。NormalizeDateText の年ガードと同じ範囲）
    //  3) それ以外 → 西暦 yyyy/MM/dd・yyyy/M/d（ParseDisplayDate）
    // 7桁は和暦としてのみ解釈するため、西暦と取り違えることはない
    //（西暦は8桁または区切り付きに限る）。いずれも不成立・空欄は null
    public static DateTime? ParseLedgerDate(string text, Dictionary<int, EraEntry> eraMap)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        var trimmed = text.Trim();

        if (Regex.IsMatch(trimmed, "^[0-9]{7}$")) return WarekiToDate(trimmed, eraMap);

        if (Regex.IsMatch(trimmed, "^[0-9]{8}$"))
        {
            int year = int.Parse(trimmed.Substring(0, 4), CultureInfo.InvariantCulture);
            if (year < MIN_YEAR || year > MAX_YEAR) return null;
            DateTime value;
            if (DateTime.TryParseExact(trimmed, "yyyyMMdd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out value)) return value;
            return null;
        }

        return ParseDisplayDate(trimmed);
    }

    // 台帳の日付文字列 → 表示・比較用の正準形（yyyy/MM/dd）変換。
    // 正準形は文字列の Ordinal 比較がそのまま時系列順になる。
    // 変換不能時は生の値をそのまま返す（表示から情報を隠さない）
    public static string FormatLedgerDate(string text, Dictionary<int, EraEntry> eraMap)
    {
        if (string.IsNullOrEmpty(text)) return "";
        var date = ParseLedgerDate(text, eraMap);
        if (date == null) return text;
        return date.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
    }

    // 日付から年度（4月始まり）を取得する（1〜3月は前年の年度に属する）。
    // 年度の区切りは JUDGE の台帳・照会結果フォルダの年度と一致していなければならない
    public static int GetFiscalYear(DateTime date)
    {
        if (date.Month >= 4) return date.Year;
        return date.Year - 1;
    }

    // 年度の鍵（「2025年度」）。台帳ファイル名・照会結果フォルダの設定・年度セレクタ・履歴の見出しで同じ字句
    public static string FormatYearKey(int year)
    {
        return year + "年度";
    }

    // 年度の鍵を年度に戻す（「年度」を除いて整数に読めなければ null）
    public static int? ParseYearKey(string key)
    {
        int year;
        if (key != null && int.TryParse(key.Replace("年度", ""), out year)) return year;
        return null;
    }

    // 台帳の行の登録日から年度を導出する（台帳ファイルの年度帰属・照会結果フォルダの年度・
    // 最近見た人の年度と同じ規則。登録日が読めなければ null）
    public static int? GetRegistrationFiscalYear(string regDateText, Dictionary<int, EraEntry> eraMap)
    {
        var regDate = ParseLedgerDate(regDateText, eraMap);
        if (regDate == null) return null;
        return GetFiscalYear(regDate.Value);
    }

    // 台帳の行を登録日で並べる鍵（読めない登録日は最小値＝降順の末尾）
    public static DateTime GetRegistrationSortKey(string regDateText, Dictionary<int, EraEntry> eraMap)
    {
        return ParseLedgerDate(regDateText, eraMap) ?? DateTime.MinValue;
    }

    // 登録日から照会結果フォルダの年度の鍵（「2025年度」）を作る（導出できなければ null＝全年度を対象にする）
    public static string GetRegistrationYearKey(string regDateText, Dictionary<int, EraEntry> eraMap)
    {
        var year = GetRegistrationFiscalYear(regDateText, eraMap);
        return year == null ? null : FormatYearKey(year.Value);
    }

    // 登録日から照会結果フォルダを絞る年度の鍵を作る（導出できない・設定にない年度は null＝全年度を対象にする。
    // 設定そのものがなくても null）
    public static string GetSearchYearKey(string regDateText, Dictionary<int, EraEntry> eraMap, Dictionary<string, List<string>> folders)
    {
        var yearKey = GetRegistrationYearKey(regDateText, eraMap);
        return yearKey != null && folders != null && folders.ContainsKey(yearKey) ? yearKey : null;
    }

    // 口座あり行の一意キー（JUDGE が台帳へ追記するときの重複排除キーと同一の組:
    //   宛名番号 × 金融機関名 × 支店番号 × 口座種別 × 口座番号 × 回答基準日）
    // 行を年度をまたいで一意に識別する（最近見た人の重複繰り上げ・カード照合に使用）。
    // 回答基準日は正準形（FormatLedgerDate）で組む。同じ日付なら和暦7桁の行と
    // 西暦表記の行が同一キーになり、台帳の日付形式が変わっても同一行と照合できる
    public static string GetDepositLedgerRowKey(string[] fields, Dictionary<int, EraEntry> eraMap)
    {
        return fields[DEPOSIT_COL_ADDRESS_NUM] + "|" + fields[DEPOSIT_COL_INSTITUTION] + "|" +
               fields[DEPOSIT_COL_BRANCH_NUM] + "|" + fields[DEPOSIT_COL_ACCOUNT_TYPE] + "|" +
               fields[DEPOSIT_COL_ACCOUNT_NUM] + "|" + FormatLedgerDate(fields[DEPOSIT_COL_BASE_DATE], eraMap);
    }

    // Excel シート名の安全化（STALL の Get-SafeSheetName と同一ロジック）
    // 禁止文字（\ / ? * [ ] :）を "_" に置換し、31文字に切り詰める。
    // 重複時の連番付与は生成側の実行文脈に依存するため、照合側の
    // SheetNameMatchesExpected が連番形も一致として吸収する
    public static string GetSafeSheetName(string baseName)
    {
        var name = Regex.Replace(baseName, @"[\\/\?\*\[\]:]", "_");
        if (name.Length > SHEET_NAME_MAX) name = name.Substring(0, SHEET_NAME_MAX);
        return name;
    }

    // 期待シート名 {照会結果ファイル名(拡張子なし)}_{シート名} の安全化（明細の読み取りと照会ファイルの照合で同じ規則。
    // どちらかが空なら作れない＝null）
    public static string GetExpectedSheetName(string sourceFile, string sheetName)
    {
        if (sourceFile == "" || sheetName == "") return null;
        return GetSafeSheetName(System.IO.Path.GetFileNameWithoutExtension(sourceFile) + "_" + sheetName);
    }

    // シート名が期待シート名（またはその重複連番形）と一致するかを判定する
    // 生成側は重複時に基本名を「31 − 連番部の長さ」まで詰めてから "_N" を付ける
    private static readonly Regex SheetSuffixRegex = new Regex(@"^(.*)_(\d+)$", RegexOptions.Compiled);

    public static bool SheetNameMatchesExpected(string sheetName, string expectedSheet)
    {
        if (sheetName == expectedSheet) return true;

        var match = SheetSuffixRegex.Match(sheetName);
        if (!match.Success) return false;
        var maxBase = SHEET_NAME_MAX - (match.Groups[2].Value.Length + 1);
        if (maxBase <= 0) return false;
        var expectedBase = expectedSheet.Length > maxBase ? expectedSheet.Substring(0, maxBase) : expectedSheet;
        return match.Groups[1].Value == expectedBase;
    }

    // 残高がセンチネル値（-999999始まり・金融機関の回答コード）かを判定する
    // 台帳の残高はカンマ整形済み（例: -999,999,999,999,999）のため、
    // 生成側（judge.ps1）と同じくカンマを除去してから前方一致で判定する
    public static bool IsSentinelBalance(string balance)
    {
        if (string.IsNullOrEmpty(balance)) return false;
        return balance.Trim().Replace(",", "").StartsWith("-999999");
    }

    // 検索用の正規化: NFKC 正規化（半角カナ→全角カナ・全角英数→半角英数・濁点の合成）
    // → 大文字化 → 空白（半角/全角）の除去。半角/全角・空白の揺れを吸収して部分一致
    // させるための片寄せで、表示・ソートには使わない。台帳読み込み時に検索対象4列の
    // 正規化キーを事前計算し、検索文字列側も同じ関数で正規化してから突き合わせる
    //（キー入力のたびに全行を正規化しない）
    public static string NormalizeForSearch(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        string normalized;
        try { normalized = value.Normalize(NormalizationForm.FormKC); }
        catch (ArgumentException) { normalized = value; }   // 不正なサロゲートペア等は生値のまま扱う
        normalized = normalized.ToUpperInvariant();

        // 空白を除き、かなを「カタカナ・大書き」へ寄せる（ひらがな→カタカナ、小書き ュ→ユ 等）。
        // 照会結果の摘要・カナ氏名は小書きを使わない電文由来のため、入力側の小書きも同じ形へ畳む
        var sb = new StringBuilder(normalized.Length);
        foreach (var c in normalized)
        {
            if (c == ' ' || c == '　') continue;
            var ch = c;
            if (ch >= 'ぁ' && ch <= 'ゖ') ch = (char)(ch + 0x60);   // ひらがな→カタカナ（コード点差 0x60）
            switch (ch)
            {
                case 'ァ': ch = 'ア'; break;
                case 'ィ': ch = 'イ'; break;
                case 'ゥ': ch = 'ウ'; break;
                case 'ェ': ch = 'エ'; break;
                case 'ォ': ch = 'オ'; break;
                case 'ッ': ch = 'ツ'; break;
                case 'ャ': ch = 'ヤ'; break;
                case 'ュ': ch = 'ユ'; break;
                case 'ョ': ch = 'ヨ'; break;
                case 'ヮ': ch = 'ワ'; break;
                case 'ヵ': ch = 'カ'; break;
                case 'ヶ': ch = 'ケ'; break;
            }
            sb.Append(ch);
        }
        return sb.ToString();
    }

    // 差押リストの対象（カンマ区切り）を照合パターンへ解析する。
    // 空トークンは読み飛ばす（落ちない）。0件（キー未設定・空欄）は null = 制御なし
    public static List<SeizureTargetPattern> ParseSeizureTargetBanks(string raw)
    {
        if (string.IsNullOrEmpty(raw)) return null;
        var result = new List<SeizureTargetPattern>();
        foreach (var token in raw.Split(','))
        {
            var key = NormalizeForSearch(token);
            if (key.Length == 0) continue;
            var pattern = new SeizureTargetPattern();
            if (key.IndexOf('*') >= 0 || key.IndexOf('?') >= 0)
            {
                // Regex.Escape でリテラル化してから * ? だけをワイルドカードへ戻す
                //（金融機関名に正規表現の記号が含まれても文字どおりに照合される）
                pattern.Glob = new Regex(
                    "^" + Regex.Escape(key).Replace("\\*", ".*").Replace("\\?", ".") + "$");
            }
            else
            {
                pattern.ExactKey = key;
            }
            result.Add(pattern);
        }
        return result.Count > 0 ? result : null;
    }

    // 金融機関の正規化キーが対象パターンのいずれかに一致するか。
    // 一致した金融機関は全部対象（カテゴリ指定のため一意性は要求しない——
    // 担当の表記ゆれ統合（BuildStaffClusters）とは意味論が異なり、ここでは用いない）
    public static bool MatchesSeizureTarget(string institutionKey, List<SeizureTargetPattern> patterns)
    {
        foreach (var pattern in patterns)
        {
            if (pattern.Glob != null ? pattern.Glob.IsMatch(institutionKey)
                                     : pattern.ExactKey == institutionKey)
                return true;
        }
        return false;
    }

    // 検出カテゴリの固定順（チップの走査順・詳細検索の検出条件と共有）
    public static readonly string[] DETECTION_NAMES =
        { "反対債権", "年金", "給与", "保険", "定期入金", "定期出金" };

    // 台帳の金額の列を整数にする（JUDGE はカンマ付きで書く。全角の数字・負号・カンマは半角に寄せ、
    // カンマを除いてから読む）。読めなければ null（「電子回答不可」などの生値・空欄）
    public static long? ParseLedgerAmount(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        var sb = new StringBuilder(text.Length);
        foreach (var c in text.Trim())
        {
            var ch = c;
            if (ch >= '０' && ch <= '９') ch = (char)(ch - '０' + '0');
            else if (ch == '－') ch = '-';
            else if (ch == '，') ch = ',';
            if (ch == ',') continue;
            sb.Append(ch);
        }
        long value;
        if (long.TryParse(sb.ToString(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value)) return value;
        return null;
    }

    // 残高の数値化（詳細検索の金額条件用）。センチネル・空・非数は null（=条件の対象外）
    public static long? ParseBalanceValue(string balance)
    {
        if (string.IsNullOrWhiteSpace(balance) || IsSentinelBalance(balance)) return null;
        return ParseLedgerAmount(balance);
    }

    // 表示形式（yyyy/MM/dd）の日付文字列を日付化する（詳細検索の範囲条件用）。不能は null
    public static DateTime? ParseDisplayDate(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        DateTime value;
        if (DateTime.TryParseExact(text.Trim(), new[] { "yyyy/MM/dd", "yyyy/M/d" },
                CultureInfo.InvariantCulture, DateTimeStyles.None, out value)) return value;
        return null;
    }

    // 残高欄の3桁区切り整形（詳細検索の金額条件用）。半角数字・カンマ・先頭マイナス
    // だけで構成される文字列を、カンマの位置を無視して数字列に戻してから3桁区切りに
    // 整え直す（位置違いのカンマを吸収）。先頭の余分な0は落とす（"0" 単独は残す）。
    // 対象外の文字を含む入力は null（呼び出し側は書き換えない）
    public static string FormatAmountText(string text)
    {
        if (text == null || !Regex.IsMatch(text, @"^-?[0-9,]*$")) return null;
        bool negative = text.StartsWith("-");
        var digits = Regex.Replace(text, "[^0-9]", "").TrimStart('0');
        if (digits.Length == 0 && text.IndexOf('0') >= 0) digits = "0";
        if (digits.Length == 0) return negative ? "-" : "";
        var grouped = new StringBuilder();
        for (int i = 0; i < digits.Length; i++)
        {
            if (i > 0 && (digits.Length - i) % 3 == 0) grouped.Append(',');
            grouped.Append(digits[i]);
        }
        return (negative ? "-" : "") + grouped.ToString();
    }

    // 残高欄のフォーカス喪失時の補正。全角数字・記号を NFKC で半角へ片寄せしてから
    // 3桁区切りに整形する（片寄せは NormalizeForSearch と同じ手法）。
    // 片寄せしても数字として解釈できない入力は null（コーラル枠の提示は ReadAmountInput が行う）
    public static string NormalizeAmountText(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        string normalized;
        try { normalized = text.Trim().Normalize(NormalizationForm.FormKC); }
        catch (ArgumentException) { return null; }   // 不正なサロゲートペア等は補正しない
        return FormatAmountText(normalized);
    }

    // 日付欄のフォーカス喪失時の補正。全角数字・全角スラッシュを NFKC で半角へ
    // 片寄せしたうえで、次の順に解釈を試みて正規形 yyyy/MM/dd を返す。
    //  1) yyyy/m/d           → ゼロ埋めに揃える
    //  2) 数字8桁 yyyyMMdd   → 区切りを補う
    //  3) 数字7桁            → yyyy+M+dd / yyyy+MM+d の両解釈のうち一方だけ
    //                          日付として成立する場合のみ採用する
    //                          （例: 2026115 は 1/15 と 11/5 の両方が成立するため不採用）
    // 区切りなしの数字のみは西暦として、かつ年が妥当な範囲（1900〜2099）の場合だけ
    // 解釈する。7桁和暦（例: 5080715=令和8年7月15日）を年5080等の西暦に取り違える
    // 事故を防ぐためで、7桁和暦としての入力自体にも対応しない。
    // 曖昧・不成立・他の文字混入は null（元の入力を書き換えず、コーラル枠の提示は
    // ReadDateInput が行う）
    public static string NormalizeDateText(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        string normalized;
        try { normalized = text.Trim().Normalize(NormalizationForm.FormKC); }
        catch (ArgumentException) { return null; }

        var parsed = ParseDisplayDate(normalized);
        if (parsed != null) return parsed.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);

        if (Regex.IsMatch(normalized, "^[0-9]{7,8}$"))
        {
            int year = int.Parse(normalized.Substring(0, 4), CultureInfo.InvariantCulture);
            if (year < MIN_YEAR || year > MAX_YEAR) return null;
        }

        if (Regex.IsMatch(normalized, "^[0-9]{8}$"))
            return TryYmdText(normalized.Substring(0, 4), normalized.Substring(4, 2), normalized.Substring(6, 2));

        if (Regex.IsMatch(normalized, "^[0-9]{7}$"))
        {
            var single = TryYmdText(normalized.Substring(0, 4), normalized.Substring(4, 1), normalized.Substring(5, 2));
            var dual   = TryYmdText(normalized.Substring(0, 4), normalized.Substring(4, 2), normalized.Substring(6, 1));
            if (single != null && dual == null) return single;
            if (dual != null && single == null) return dual;
            return null;
        }
        return null;
    }

    // 年・月・日の文字列を日付として検証し、成立すれば正規形 yyyy/MM/dd を返す。不能は null
    private static string TryYmdText(string year, string month, string day)
    {
        var value = ParseDisplayDate(year + "/" + month + "/" + day);
        return value == null ? null : value.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
    }

    // 反対債権フィールドに「値がある」かを判定する
    // 空文字列と "0" は値なし。カンマ付き数値をパースして正なら値あり。
    // パース不能な非空文字列は安全側で値ありとして扱う（見落としの害の方が大きい）
    public static bool HasCounterClaimValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        var trimmed = value.Trim();
        if (trimmed == "0") return false;

        decimal amount;
        if (decimal.TryParse(trimmed.Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out amount))
            return amount > 0;
        return true;
    }

    // 検出カテゴリの生文字列から末尾の「（N回・バケット）」を解析する
    // 生成側（judge.ps1 の明細分析）のフォーマットは
    //   名称（N回） または 名称（N回・バケット）
    // で、複数件は「、」区切り。先頭件（回数最多）のバケットを返す。
    // 解析できない場合は null を返す（呼び出し側はカテゴリ名のみのチップに退行）
    private static readonly Regex DetectSuffixRegex = new Regex(@"（([^（）]*)）\s*$", RegexOptions.Compiled);

    // 複数件の区切り。名称自体に「、」が含まれうる（例: 「テキヨウナシカ－ド、現払」）ため、
    // 単純な Split ではなく「）」の直後に現れる「、」のみを区切りとして扱う
    private static readonly Regex DetectEntrySplitRegex = new Regex(@"(?<=）)、", RegexOptions.Compiled);

    // 検出カテゴリの生文字列を件単位に分割する
    public static string[] SplitDetectionEntries(string raw)
    {
        return DetectEntrySplitRegex.Split(raw);
    }

    public static string ParseDetectionBucket(string firstEntry)
    {
        if (string.IsNullOrEmpty(firstEntry)) return null;
        var match = DetectSuffixRegex.Match(firstEntry.Trim());
        if (!match.Success) return null;

        var parts = match.Groups[1].Value.Split('・');
        if (parts.Length < 2) return null;   // 「（N回）」のみ＝バケットなし（年金は常にこれ）

        var bucket = parts[1].Trim();
        return bucket.Length > 0 ? bucket : null;
    }

    // 検出1件の文字列から名称（末尾の「（N回・…）」を除いた部分）を取り出す
    public static string ParseDetectionName(string entry)
    {
        if (string.IsNullOrEmpty(entry)) return "";
        return DetectSuffixRegex.Replace(entry.Trim(), "").Trim();
    }

    // --- 担当の表記ゆれ統合 ---
    //
    // 照会結果の担当欄は外字が置換文字に化ける（例: 「髙田　太郎」→「◆田　太郎」）ほか、
    // 姓のみの記入が混在する。担当フィルターの列挙・照合に限り、機械的に安全な範囲で
    // 同一人物として統合する:
    //   段1: 正規化キー（NormalizeForSearch）が一致する生値は同一
    //   段2: 化け置換文字を任意の1文字とみなし、同長で化け位置以外が全一致なら統合
    //        （化け文字がより少ない側が親）
    //   段3: 短いキーが長いキーの先頭部分に（化け文字込みで）一致すれば統合
    // 統合先の候補クラスタが2つ以上あるキーは統合しない（曖昧なら独立のまま。
    // 候補は統合後のクラスタ単位で数える——「◆田」の候補「髙田太郎」「◆田太郎」が
    // 同一クラスタなら一意とみなす）。親は「同長で化け文字が少ない」または「より長い」
    // キーに限られるため、統合の連鎖は循環しない

    // キーに含まれる化け置換文字の個数
    private static int CountFuzzyChars(string key, HashSet<char> fuzzyChars)
    {
        int count = 0;
        foreach (var ch in key)
            if (fuzzyChars.Contains(ch)) count++;
        return count;
    }

    // 同長ワイルドカード一致（化け置換文字は任意の1文字とみなす・対称）
    private static bool FuzzyEquals(string a, string b, HashSet<char> fuzzyChars)
    {
        if (a.Length != b.Length) return false;
        for (int i = 0; i < a.Length; i++)
            if (a[i] != b[i] && !fuzzyChars.Contains(a[i]) && !fuzzyChars.Contains(b[i])) return false;
        return true;
    }

    // 短いキーが長いキーの先頭部分にワイルドカード一致するか
    private static bool FuzzyPrefixMatches(string shortKey, string longKey, HashSet<char> fuzzyChars)
    {
        return shortKey.Length < longKey.Length
            && FuzzyEquals(shortKey, longKey.Substring(0, shortKey.Length), fuzzyChars);
    }

    // 担当職員名の distinct 生値をクラスタへ統合する。返却は代表名の Ordinal 昇順
    public static List<StaffCluster> BuildStaffClusters(List<string> rawValues, HashSet<char> fuzzyChars)
    {
        // 段1: 正規化キーで集約（キー → 生値リスト・初出順）
        var keyRaws = new Dictionary<string, List<string>>();
        var keyOrder = new List<string>();
        foreach (var raw in rawValues)
        {
            var key = NormalizeForSearch(raw);
            if (key == "") continue;
            List<string> raws;
            if (!keyRaws.TryGetValue(key, out raws))
            {
                raws = new List<string>();
                keyRaws[key] = raws;
                keyOrder.Add(key);
            }
            if (!raws.Contains(raw)) raws.Add(raw);
        }

        // union-find（キー単位・経路圧縮のみ）
        var parent = new Dictionary<string, string>();
        foreach (var key in keyOrder) parent[key] = key;
        Func<string, string> find = delegate(string key)
        {
            while (parent[key] != key)
            {
                parent[key] = parent[parent[key]];
                key = parent[key];
            }
            return key;
        };

        // 段2・段3: 統合先の候補クラスタが一意のキーだけを統合する。統合で候補の
        // 数えかたが変わりうるため、変化がなくなるまで反復する（キー数は職員数程度）
        bool changed = true;
        while (changed)
        {
            changed = false;
            var snapshot = new Dictionary<string, string>();
            foreach (var key in keyOrder) snapshot[key] = find(key);
            foreach (var v in keyOrder)
            {
                int fuzzyCount = CountFuzzyChars(v, fuzzyChars);
                var candidates = new HashSet<string>();
                foreach (var u in keyOrder)
                {
                    if (u == v || snapshot[u] == snapshot[v]) continue;
                    if (FuzzyEquals(v, u, fuzzyChars) && CountFuzzyChars(u, fuzzyChars) < fuzzyCount)
                        candidates.Add(snapshot[u]);          // 段2
                    else if (FuzzyPrefixMatches(v, u, fuzzyChars))
                        candidates.Add(snapshot[u]);          // 段3
                }
                if (candidates.Count == 1)
                {
                    parent[find(v)] = find(candidates.First());
                    changed = true;
                }
            }
        }

        // クラスタへ集約し、代表名を選出する
        var grouped = new Dictionary<string, StaffCluster>();
        foreach (var key in keyOrder)
        {
            var root = find(key);
            StaffCluster cluster;
            if (!grouped.TryGetValue(root, out cluster))
            {
                cluster = new StaffCluster { MemberKeys = new HashSet<string>(), RawValues = new List<string>() };
                grouped[root] = cluster;
            }
            cluster.MemberKeys.Add(key);
            cluster.RawValues.AddRange(keyRaws[key]);
        }
        var result = new List<StaffCluster>();
        foreach (var cluster in grouped.Values)
        {
            // 代表名: 化け文字を含まない最長の生値（同点は Ordinal 順の先頭）。
            // 全メンバーが化け文字を含む場合は 化け最少 → 最長 → Ordinal 順
            var clean = cluster.RawValues.Where(r => CountFuzzyChars(r, fuzzyChars) == 0).ToList();
            var pool = clean.Count > 0 ? clean : cluster.RawValues;
            cluster.Representative = pool
                .OrderBy(r => CountFuzzyChars(r, fuzzyChars))
                .ThenByDescending(r => r.Length)
                .ThenBy(r => r, StringComparer.Ordinal)
                .First();
            result.Add(cluster);
        }
        result.Sort(delegate(StaffCluster a, StaffCluster b)
        {
            return string.CompareOrdinal(a.Representative, b.Representative);
        });
        return result;
    }
}