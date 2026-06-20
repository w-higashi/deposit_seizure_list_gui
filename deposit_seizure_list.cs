// ==============================================================================
// deposit_seizure_list.cs
// 預金差押予定一覧 作成ツール（WPF GUI版）
//
// 【使用方法】
// 1. build.bat を実行して deposit_seizure_list.exe を生成
// 2a. exe をダブルクリック → 初期画面で D&D またはファイル選択
// 2b. exe にファイルを D&D → 直接処理開始
// 2c. file_search.exe から呼び出し → 引数のファイルを処理後にクローズ
//
// 【処理概要】
// 電子預金照会結果 .xlsm から必要情報を抽出し、
// 担当者の判断（執行日・口座選択）を加えて
// 預金差押予定一覧.csv に追記する
//
// 【ビルド方法】
// build.bat を実行（.NET Framework 4.0 の csc.exe を使用）
//
// 【必要ファイル（同じフォルダに配置）】
// ＜必須＞
// - deposit_seizure_list.cs  （ソースコード）
// - deposit_seizure_list_config.json （設定ファイル）
// - document_number_counter.json （文書番号カウンター）
// - era_mapping.json （元号マッピング）
// - build.bat （ビルドスクリプト）
// ＜任意＞
// - transfer_fee_mapping.json, yucho_center_mapping.json,
//   account_type_mapping.json, institution_name_mapping.json,
//   branch_name_mapping.json
// ==============================================================================

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Xml;

// ==============================================================
// データモデル
// ==============================================================

// プロファイル設定（config JSON の profiles 配列の1要素）
public class ProfileConfig
{
    public string Name { get; set; }                     // プロファイル名
    public string CoverCell { get; set; }                // 表紙検出セル列（例: "A2"）
    public string CoverValue { get; set; }               // 表紙検出値
    public string[] StopValues { get; set; }             // 明細ページ先頭の目印文字列（OR条件）
    public string AddressNumberCell { get; set; }        // 宛名番号セル
    public string StaffCell { get; set; }                // 処分担当セル
    public string NameCell { get; set; }                 // 氏名セル
    public string ResidenceAddressCell { get; set; }     // 住民票住所セル
    public string DeliveryAddressCell { get; set; }      // 届出住所セル
    public string FinancialInstitutionCell { get; set; } // 金融機関名セル
    public string FilterCell { get; set; }               // シートフィルタセル
    public string[] FilterValues { get; set; }             // シートフィルタ値（部分一致、OR条件）
    public int AccountStartRow { get; set; }             // 口座テーブル開始行
    public string BranchNameCol { get; set; }            // 支店名列
    public string BranchNumberCol { get; set; }          // 支店番号列
    public string AccountTypeCol { get; set; }           // 口座種別列
    public string LastTransactionCol { get; set; }       // 最終取引日列
    public string AccountNumberCol { get; set; }         // 口座番号列
    public string BalanceCol { get; set; }               // 残高列
    public string DetailAccountNumberCell { get; set; }  // 明細ページ口座番号セル（任意）
    public string DetailAccountTypeCell { get; set; }    // 明細ページ口座種別セル（任意）
    public string OutputFolder { get; set; }             // CSV出力先フォルダ
    public string PrintFolder { get; set; }              // 印刷用ファイル保存先フォルダ
}

// アプリ全体の設定
public class AppConfig
{
    public List<ProfileConfig> Profiles { get; set; }
    public string DefaultFolder { get; set; }            // 初期画面のファイル選択初期フォルダ（任意）

    public AppConfig()
    {
        Profiles = new List<ProfileConfig>();
        DefaultFolder = null;
    }
}

// 元号マッピング1件
public class EraEntry
{
    public string Name { get; set; }                     // 元号名（例: "令和"）
    public int StartYear { get; set; }                   // 元号元年の西暦（例: 2019）
}

// 口座テーブルの1行
public class AccountItem
{
    public string BranchName { get; set; }               // 支店名
    public string BranchNumber { get; set; }             // 支店番号
    public string AccountType { get; set; }              // 口座種別
    public string LastTransaction { get; set; }          // 最終取引日
    public string AccountNum { get; set; }               // 口座番号
    public string AccountNumDisplay                       // 口座番号（★付き表示用）
    {
        get { return HasSeizureHistory ? AccountNum + " \u2605" : AccountNum; }
    }
    public string SeizureTooltip                          // 差押実績のツールチップ
    {
        get { return HasSeizureHistory ? "差押実績あり（執行日: " + SeizureExecDate + " / 文書番号: " + SeizureDocNumber + "）" : null; }
    }
    public double BalanceValue { get; set; }             // 残高（数値、ソート用）
    public string Balance { get; set; }                  // 残高（表示用、通貨形式）
    public int CoverIndex { get; set; }                  // 所属する表紙ブロックのインデックス
    public string DeliveryAddressRaw { get; set; }        // 所属する表紙ブロックの届出住所（生値）
    public bool HasSeizureHistory { get; set; }          // 差押実績の有無
    public string SeizureDocNumber { get; set; }         // 差押実績の文書番号（あれば）
    public string SeizureExecDate { get; set; }          // 差押実績の執行日（あれば、表示用）
}

// ファイルの処理状態
public enum FileProcessState
{
    Pending,       // 未処理
    Added,         // 一覧に追加済み
    Skipped,       // スキップ
    Error          // エラー
}

// 処理対象ファイル1件の情報
public class FileEntry
{
    public string FilePath { get; set; }
    public FileProcessState State { get; set; }
}

// ==============================================================
// JSON パーサー（手書き・外部ライブラリ不要）
// LGWAN 環境では NuGet パッケージが使えないため手動パース
// ==============================================================

public static class JsonHelper
{
    // JSON文字列から指定キーの文字列値を取得
    public static string GetString(string json, string key)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return null;
        var colonIdx = json.IndexOf(':', keyIdx + key.Length + 2);
        if (colonIdx < 0) return null;

        var rest = json.Substring(colonIdx + 1).TrimStart();
        if (rest.Length == 0 || rest[0] != '"') return null;

        var sb = new StringBuilder();
        bool escaped = false;
        for (int i = 1; i < rest.Length; i++)
        {
            if (escaped) { sb.Append(rest[i]); escaped = false; continue; }
            if (rest[i] == '\\') { escaped = true; continue; }
            if (rest[i] == '"') break;
            sb.Append(rest[i]);
        }
        return sb.ToString();
    }

    // JSON文字列から指定キーの整数値を取得
    public static int GetInt(string json, string key, int defaultValue = 0)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return defaultValue;
        var colonIdx = json.IndexOf(':', keyIdx + key.Length + 2);
        if (colonIdx < 0) return defaultValue;

        var rest = json.Substring(colonIdx + 1).TrimStart();
        var numStr = new StringBuilder();
        foreach (var c in rest)
        {
            if (char.IsDigit(c) || c == '-') numStr.Append(c);
            else if (numStr.Length > 0) break;
        }
        int result;
        return int.TryParse(numStr.ToString(), out result) ? result : defaultValue;
    }

    // JSON文字列から指定キーの文字列配列を取得
    public static string[] GetStringArray(string json, string key)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return new string[0];
        var arrStart = json.IndexOf('[', keyIdx);
        if (arrStart < 0) return new string[0];
        var arrEnd = FindMatchingBracket(json, arrStart, '[', ']');
        if (arrEnd < 0) return new string[0];

        var inner = json.Substring(arrStart + 1, arrEnd - arrStart - 1);
        return ExtractQuotedStrings(inner).ToArray();
    }

    // JSON オブジェクト配列を取得（各要素を文字列として返す）
    public static List<string> GetObjectArray(string json, string key)
    {
        var result = new List<string>();
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return result;
        var arrStart = json.IndexOf('[', keyIdx);
        if (arrStart < 0) return result;
        var arrEnd = FindMatchingBracket(json, arrStart, '[', ']');
        if (arrEnd < 0) return result;

        var inner = json.Substring(arrStart + 1, arrEnd - arrStart - 1);
        int pos = 0;
        while (pos < inner.Length)
        {
            var objStart = inner.IndexOf('{', pos);
            if (objStart < 0) break;
            var objEnd = FindMatchingBracket(inner, objStart, '{', '}');
            if (objEnd < 0) break;
            result.Add(inner.Substring(objStart, objEnd - objStart + 1));
            pos = objEnd + 1;
        }
        return result;
    }

    // JSON オブジェクトをキーバリューの辞書として取得（値は文字列のみ対応）
    public static Dictionary<string, string> GetStringDictionary(string json)
    {
        var dict = new Dictionary<string, string>();
        int pos = 0;
        while (pos < json.Length)
        {
            var keyStart = json.IndexOf('"', pos);
            if (keyStart < 0) break;
            var keyEnd = json.IndexOf('"', keyStart + 1);
            if (keyEnd < 0) break;
            var key = json.Substring(keyStart + 1, keyEnd - keyStart - 1);

            if (key.StartsWith("_")) { pos = keyEnd + 1; continue; }

            var colonIdx = json.IndexOf(':', keyEnd + 1);
            if (colonIdx < 0) break;

            var valStart = json.IndexOf('"', colonIdx + 1);
            if (valStart < 0) { pos = colonIdx + 1; continue; }
            var valEnd = json.IndexOf('"', valStart + 1);
            if (valEnd < 0) break;
            var val = json.Substring(valStart + 1, valEnd - valStart - 1);

            dict[key] = val.Replace("\\\\", "\\");
            pos = valEnd + 1;
        }
        return dict;
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

    // JSON内の引用符で囲まれた文字列を全て抽出
    public static List<string> ExtractQuotedStrings(string content)
    {
        var result = new List<string>();
        int pos = 0;
        while (pos < content.Length)
        {
            var qStart = content.IndexOf('"', pos);
            if (qStart < 0) break;
            var qEnd = content.IndexOf('"', qStart + 1);
            if (qEnd < 0) break;
            result.Add(content.Substring(qStart + 1, qEnd - qStart - 1));
            pos = qEnd + 1;
        }
        return result;
    }

    // JSON出力用のエスケープ
    public static string Escape(string s)
    {
        if (s == null) return "";
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"")
                .Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
    }
}

// ==============================================================
// 設定ファイル読込
// ==============================================================

public static class ConfigLoader
{
    // deposit_seizure_list_config.json を読み込む
    public static AppConfig LoadConfig(string configPath)
    {
        var config = new AppConfig();
        if (!File.Exists(configPath)) return config;

        var json = File.ReadAllText(configPath, Encoding.UTF8);
        config.DefaultFolder = JsonHelper.GetString(json, "defaultFolder");

        foreach (var profileJson in JsonHelper.GetObjectArray(json, "profiles"))
        {
            var p = new ProfileConfig
            {
                Name                     = JsonHelper.GetString(profileJson, "name") ?? "",
                CoverCell                = JsonHelper.GetString(profileJson, "coverCell"),
                CoverValue               = JsonHelper.GetString(profileJson, "coverValue"),
                StopValues               = JsonHelper.GetStringArray(profileJson, "stopValues"),
                AddressNumberCell        = JsonHelper.GetString(profileJson, "addressNumberCell"),
                StaffCell                = JsonHelper.GetString(profileJson, "staffCell"),
                NameCell                 = JsonHelper.GetString(profileJson, "nameCell"),
                ResidenceAddressCell     = JsonHelper.GetString(profileJson, "residenceAddressCell"),
                DeliveryAddressCell      = JsonHelper.GetString(profileJson, "deliveryAddressCell"),
                FinancialInstitutionCell = JsonHelper.GetString(profileJson, "financialInstitutionCell"),
                FilterCell               = JsonHelper.GetString(profileJson, "filterCell"),
                AccountStartRow          = JsonHelper.GetInt(profileJson, "accountStartRow", 26),
                BranchNameCol            = JsonHelper.GetString(profileJson, "branchNameCol"),
                BranchNumberCol          = JsonHelper.GetString(profileJson, "branchNumberCol"),
                AccountTypeCol           = JsonHelper.GetString(profileJson, "accountTypeCol"),
                LastTransactionCol       = JsonHelper.GetString(profileJson, "lastTransactionCol"),
                AccountNumberCol         = JsonHelper.GetString(profileJson, "accountNumberCol"),
                BalanceCol               = JsonHelper.GetString(profileJson, "balanceCol"),
                DetailAccountNumberCell  = JsonHelper.GetString(profileJson, "detailAccountNumberCell"),
                DetailAccountTypeCell    = JsonHelper.GetString(profileJson, "detailAccountTypeCell"),
                OutputFolder             = JsonHelper.GetString(profileJson, "outputFolder"),
                PrintFolder              = JsonHelper.GetString(profileJson, "printFolder")
            };
            if (p.OutputFolder != null) p.OutputFolder = p.OutputFolder.Replace("\\\\", "\\");
            if (p.PrintFolder != null)  p.PrintFolder  = p.PrintFolder.Replace("\\\\", "\\");

            // filterValue: 文字列または配列をサポート（配列の場合はOR条件）
            // GetString で文字列として取得を試み、失敗したら GetStringArray で配列として取得
            string singleFilterValue = JsonHelper.GetString(profileJson, "filterValue");
            if (singleFilterValue != null)
                p.FilterValues = new[] { singleFilterValue };
            else
                p.FilterValues = JsonHelper.GetStringArray(profileJson, "filterValue");

            config.Profiles.Add(p);
        }

        return config;
    }

    // era_mapping.json を読み込む
    public static Dictionary<int, EraEntry> LoadEraMapping(string path)
    {
        var map = new Dictionary<int, EraEntry>();
        if (!File.Exists(path)) return map;

        var json = File.ReadAllText(path, Encoding.UTF8);
        // キーが "3", "4", "5" 等の数値文字列であるオブジェクトを解析
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
            if (name != null && startYear > 0)
            {
                map[code] = new EraEntry { Name = name, StartYear = startYear };
            }
        }
        return map;
    }

    // 汎用マッピングファイルの読込（キー: 文字列, 値: 文字列）
    // transfer_fee, institution_name, branch_name 等に使用
    public static Dictionary<string, string> LoadSimpleMapping(string path)
    {
        if (!File.Exists(path)) return null;
        var json = File.ReadAllText(path, Encoding.UTF8);
        return JsonHelper.GetStringDictionary(json);
    }

    // transfer_fee_mapping.json 用（値がオブジェクト { feeText: "..." }）
    public static Dictionary<string, string> LoadFeeMapping(string path)
    {
        var map = new Dictionary<string, string>();
        if (!File.Exists(path)) return null;

        var json = File.ReadAllText(path, Encoding.UTF8);
        int pos = 0;
        while (pos < json.Length)
        {
            var keyStart = json.IndexOf('"', pos);
            if (keyStart < 0) break;
            var keyEnd = json.IndexOf('"', keyStart + 1);
            if (keyEnd < 0) break;
            var key = json.Substring(keyStart + 1, keyEnd - keyStart - 1);
            if (key.StartsWith("_")) { pos = keyEnd + 1; continue; }

            var objStart = json.IndexOf('{', keyEnd);
            if (objStart < 0) break;
            var objEnd = JsonHelper.FindMatchingBracket(json, objStart, '{', '}');
            if (objEnd < 0) break;

            var objJson = json.Substring(objStart, objEnd - objStart + 1);
            var feeText = JsonHelper.GetString(objJson, "feeText");
            if (feeText != null) map[key] = feeText;
            pos = objEnd + 1;
        }
        return map.Count > 0 ? map : null;
    }

    // yucho_center_mapping.json 用（ネストされた辞書）
    public static Dictionary<string, Dictionary<string, string>> LoadYuchoCenterMapping(string path)
    {
        if (!File.Exists(path)) return null;
        var json = File.ReadAllText(path, Encoding.UTF8);
        var result = new Dictionary<string, Dictionary<string, string>>();

        // 最上位キー（"総合口座", "振替口座"）ごとに内側の辞書を取得
        foreach (var topKey in new[] { "総合口座", "振替口座" })
        {
            var keyIdx = json.IndexOf("\"" + topKey + "\"");
            if (keyIdx < 0) continue;
            var objStart = json.IndexOf('{', keyIdx);
            if (objStart < 0) continue;
            var objEnd = JsonHelper.FindMatchingBracket(json, objStart, '{', '}');
            if (objEnd < 0) continue;
            var objJson = json.Substring(objStart, objEnd - objStart + 1);
            result[topKey] = JsonHelper.GetStringDictionary(objJson);
        }
        return result.Count > 0 ? result : null;
    }

    // account_type_mapping.json 用（global + 金融機関別のネスト辞書）
    public static Dictionary<string, Dictionary<string, string>> LoadAccountTypeMapping(string path)
    {
        if (!File.Exists(path)) return null;
        var json = File.ReadAllText(path, Encoding.UTF8);
        var result = new Dictionary<string, Dictionary<string, string>>();

        // 各キー（"global", 金融機関名...）の内側辞書を取得
        int pos = 1; // 最初の { をスキップ
        while (pos < json.Length)
        {
            var keyStart = json.IndexOf('"', pos);
            if (keyStart < 0) break;
            var keyEnd = json.IndexOf('"', keyStart + 1);
            if (keyEnd < 0) break;
            var key = json.Substring(keyStart + 1, keyEnd - keyStart - 1);
            if (key.StartsWith("_")) { pos = keyEnd + 1; continue; }

            var objStart = json.IndexOf('{', keyEnd);
            if (objStart < 0) break;
            var objEnd = JsonHelper.FindMatchingBracket(json, objStart, '{', '}');
            if (objEnd < 0) break;

            var objJson = json.Substring(objStart, objEnd - objStart + 1);
            result[key] = JsonHelper.GetStringDictionary(objJson);
            pos = objEnd + 1;
        }
        return result.Count > 0 ? result : null;
    }

    // document_number_counter.json から次の番号を読み取る
    public static int LoadNextDocNumber(string path)
    {
        if (!File.Exists(path)) return 1;
        var json = File.ReadAllText(path, Encoding.UTF8);
        return JsonHelper.GetInt(json, "nextNumber", 1);
    }

    // document_number_counter.json に次の番号を書き込む
    public static void SaveNextDocNumber(string path, int nextNumber)
    {
        var json = "{\n    \"nextNumber\":  " + nextNumber + "\n}";
        File.WriteAllText(path, json, Encoding.UTF8);
    }
}

// ==============================================================
// ヘルパー（ビジネスロジック）
// ==============================================================

public static class BusinessLogic
{
    // 列アルファベットを列インデックスに変換（例: "C" → 3, "AA" → 27）
    public static int ColToIndex(string col)
    {
        col = col.ToUpper().Trim();
        int result = 0;
        foreach (var ch in col.ToCharArray())
            result = result * 26 + ((int)ch - (int)'A' + 1);
        return result;
    }

    // セルアドレスに行オフセットを加算（例: "X11" + 40 → "X51"）
    public static string GetOffsetCell(string cell, int rowOffset)
    {
        if (rowOffset == 0) return cell;
        // セルアドレスを列部分と行番号に分割
        int i = 0;
        while (i < cell.Length && char.IsLetter(cell[i])) i++;
        if (i == 0 || i >= cell.Length) return cell;
        var colPart = cell.Substring(0, i);
        int rowPart;
        if (!int.TryParse(cell.Substring(i), out rowPart)) return cell;
        return colPart + (rowPart + rowOffset);
    }

    // 宛名番号を10桁0埋め
    public static string FormatAddressNumber(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";
        var clean = value.Trim().Replace(" ", "").Replace("\u3000", "");
        long num;
        if (long.TryParse(clean, out num) && clean.Length <= 10)
            return clean.PadLeft(10, '0');
        return clean;
    }

    // 住所の基本加工（郵便番号除去・空白除去）
    public static string FormatAddress(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        // ① 郵便番号（〒xxx-xxxx）と直後のスペース・全角スペースを除去
        var addr = System.Text.RegularExpressions.Regex.Replace(raw, @"〒\d{3}-\d{4}[\s\u3000]*", "");
        // ② 空白（半角・全角）を除去
        addr = addr.Replace(" ", "").Replace("\u3000", "");
        return addr;
    }

    // 残高を通貨形式に整形（例: 17241 → "17,241円"）
    public static string FormatBalance(object value)
    {
        if (value == null) return "-";
        var str = value.ToString().Trim();
        if (string.IsNullOrEmpty(str)) return "-";
        double num;
        if (double.TryParse(str, out num))
            return string.Format("{0:N0}円", num);
        return str;
    }

    // CSVフィールドのエスケープ（RFC 4180準拠）
    public static string CsvEscape(string field)
    {
        if (field == null) return "";
        if (field.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0)
            return "\"" + field.Replace("\"", "\"\"") + "\"";
        return field;
    }

    // 半角英数字・半角カタカナを全角に変換
    public static string ToFullWidth(string str)
    {
        if (string.IsNullOrEmpty(str)) return "";

        // 半角カナ→全角カナ変換テーブル
        var kanaMap = new Dictionary<string, string>
        {
            {"ｶﾞ","ガ"},{"ｷﾞ","ギ"},{"ｸﾞ","グ"},{"ｹﾞ","ゲ"},{"ｺﾞ","ゴ"},
            {"ｻﾞ","ザ"},{"ｼﾞ","ジ"},{"ｽﾞ","ズ"},{"ｾﾞ","ゼ"},{"ｿﾞ","ゾ"},
            {"ﾀﾞ","ダ"},{"ﾁﾞ","ヂ"},{"ﾂﾞ","ヅ"},{"ﾃﾞ","デ"},{"ﾄﾞ","ド"},
            {"ﾊﾞ","バ"},{"ﾋﾞ","ビ"},{"ﾌﾞ","ブ"},{"ﾍﾞ","ベ"},{"ﾎﾞ","ボ"},
            {"ｳﾞ","ヴ"},{"ﾊﾟ","パ"},{"ﾋﾟ","ピ"},{"ﾌﾟ","プ"},{"ﾍﾟ","ペ"},{"ﾎﾟ","ポ"},
            {"ｦ","ヲ"},{"ｧ","ァ"},{"ｨ","ィ"},{"ｩ","ゥ"},{"ｪ","ェ"},{"ｫ","ォ"},
            {"ｬ","ャ"},{"ｭ","ュ"},{"ｮ","ョ"},{"ｯ","ッ"},{"ｰ","ー"},
            {"ｱ","ア"},{"ｲ","イ"},{"ｳ","ウ"},{"ｴ","エ"},{"ｵ","オ"},
            {"ｶ","カ"},{"ｷ","キ"},{"ｸ","ク"},{"ｹ","ケ"},{"ｺ","コ"},
            {"ｻ","サ"},{"ｼ","シ"},{"ｽ","ス"},{"ｾ","セ"},{"ｿ","ソ"},
            {"ﾀ","タ"},{"ﾁ","チ"},{"ﾂ","ツ"},{"ﾃ","テ"},{"ﾄ","ト"},
            {"ﾅ","ナ"},{"ﾆ","ニ"},{"ﾇ","ヌ"},{"ﾈ","ネ"},{"ﾉ","ノ"},
            {"ﾊ","ハ"},{"ﾋ","ヒ"},{"ﾌ","フ"},{"ﾍ","ヘ"},{"ﾎ","ホ"},
            {"ﾏ","マ"},{"ﾐ","ミ"},{"ﾑ","ム"},{"ﾒ","メ"},{"ﾓ","モ"},
            {"ﾔ","ヤ"},{"ﾕ","ユ"},{"ﾖ","ヨ"},
            {"ﾗ","ラ"},{"ﾘ","リ"},{"ﾙ","ル"},{"ﾚ","レ"},{"ﾛ","ロ"},
            {"ﾜ","ワ"},{"ﾝ","ン"},{"ﾞ","゛"},{"ﾟ","゜"}
        };

        // まず濁点・半濁点の結合（2文字→1文字）を先に処理
        foreach (var pair in kanaMap.Where(p => p.Key.Length == 2))
            str = str.Replace(pair.Key, pair.Value);

        var sb = new StringBuilder();
        foreach (var c in str.ToCharArray())
        {
            int code = (int)c;
            if (code >= 0x41 && code <= 0x5A)       // 半角英大文字 → 全角
                sb.Append((char)(code + 0xFEE0));
            else if (code >= 0x61 && code <= 0x7A)   // 半角英小文字 → 全角
                sb.Append((char)(code + 0xFEE0));
            else if (code >= 0x30 && code <= 0x39)   // 半角数字 → 全角
                sb.Append((char)(code + 0xFEE0));
            else if (kanaMap.ContainsKey(c.ToString()))
                sb.Append(kanaMap[c.ToString()]);
            else
                sb.Append(c);
        }
        return sb.ToString();
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
        int adYear = era.StartYear + year - 1;

        try { return new DateTime(adYear, month, day); }
        catch { return null; }
    }

    // DateTime → 7桁和暦変換（era_mapping.json 使用）
    public static string DateToWareki(DateTime date, Dictionary<int, EraEntry> eraMap)
    {
        // 西暦年が最も近い元号を逆引き（降順で最初にマッチしたもの）
        foreach (var pair in eraMap.OrderByDescending(p => p.Value.StartYear))
        {
            if (date.Year >= pair.Value.StartYear)
            {
                int warekiYear = date.Year - pair.Value.StartYear + 1;
                return string.Format("{0}{1:D2}{2:D2}{3:D2}",
                    pair.Key, warekiYear, date.Month, date.Day);
            }
        }
        return "";
    }

    // 入力文字列を DateTime に変換（複数形式対応）
    // 7桁和暦 / 8桁西暦 / yyyy/MM/dd / yyyy/M/d
    public static DateTime? ParseFlexibleDate(string input, Dictionary<int, EraEntry> eraMap)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        input = input.Trim();

        // スラッシュ区切り（yyyy/MM/dd or yyyy/M/d）
        if (input.Contains("/"))
        {
            DateTime dt;
            if (DateTime.TryParse(input, out dt)) return dt;
            return null;
        }

        // 7桁和暦
        if (input.Length == 7 && input.All(char.IsDigit))
            return WarekiToDate(input, eraMap);

        // 8桁西暦（yyyyMMdd）
        if (input.Length == 8 && input.All(char.IsDigit))
        {
            DateTime dt;
            if (DateTime.TryParseExact(input, "yyyyMMdd",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out dt))
                return dt;
        }

        return null;
    }
}

// ==============================================================
// メインアプリケーション（スタブ構造）
// 次のチャットで各メソッドの実装を完成させる
// ==============================================================

public class DepositSeizureApp : Application
{
    // --- 設定 ---
    private AppConfig config;
    private ProfileConfig activeProfile;                   // 選択中のプロファイル
    private string exeDir;                                 // exe と同階層のフォルダパス
    private Dictionary<int, EraEntry> eraMapping;          // 元号マッピング
    private Dictionary<string, string> feeMapping;         // 振込手数料マッピング
    private Dictionary<string, Dictionary<string, string>> yuchoMapping;  // ゆうちょセンター
    private Dictionary<string, Dictionary<string, string>> accountTypeMapping; // 口座種別補完
    private Dictionary<string, string> institutionNameMapping;  // 金融機関名補正
    private Dictionary<string, string> branchNameMapping;  // CSV支店名置換

    // --- 状態 ---
    private List<FileEntry> fileEntries = new List<FileEntry>();  // 処理対象ファイルリスト
    private int currentFileIndex = -1;                     // 現在処理中のファイルインデックス
    private List<AccountItem> currentAccounts = new List<AccountItem>();  // 現在の口座テーブル
    private List<int> coverOffsets = new List<int>();       // 表紙ブロックの行オフセット
    private List<int> stopRows = new List<int>();           // stopValues 検出行
    private DateTime? processingDate = null;                // 執行日（内部保持）
    private bool isFromFileSearch = false;                  // file_search からの遷移か
    private int lastDisplayedCoverIndex = -1;               // 届出住所の自動更新で使用するカバーインデックス
    private Dictionary<string, string[]> seizureHistory = new Dictionary<string, string[]>(); // 差押実績（値: [文書番号, 執行日表示]）
    private bool suppressSheetChange = false;               // PopulateForm中のSelectionChanged抑止

    // --- キャッシュ済みブラシ（ShowResult・バリデーション表示用） ---
    // 毎回 new SolidColorBrush するとGC負荷が増えるため、
    // 静的フィールドで保持し Freeze() で描画スレッドの排他を不要にする
    private static readonly SolidColorBrush BrushBorderNormal    = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D0D0D0"));
    private static readonly SolidColorBrush BrushValidationError = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D32F2F"));
    private static readonly SolidColorBrush BrushSuccessIcon     = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#107C41"));
    private static readonly SolidColorBrush BrushAccent          = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00897B"));
    private static readonly SolidColorBrush BrushIconBgSuccess   = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F5ED"));
    private static readonly SolidColorBrush BrushIconBgError     = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FCEBEB"));
    private static readonly SolidColorBrush BrushIconBgSkip      = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E0F2F1"));
    private static readonly SolidColorBrush BrushDetailMuted     = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#999"));

    // --- 定数 ---
    private const int ACCOUNT_TABLE_MAX_ROWS = 200;        // 口座テーブル最大読取り行数
    private const int CSV_WRITE_MAX_RETRY = 5;             // CSV書き込みリトライ回数
    private const int CSV_WRITE_RETRY_INTERVAL_MS = 500;   // CSV書き込みリトライ間隔
    private const int MAX_PATH = 260;                      // Windowsのパス長上限
    private const string CSV_FILENAME = "預金差押予定一覧.csv";
    private const string CSV_HEADER = "登録日時,宛名番号,氏名,職員名,執行日,住民票住所,銀行届出住所,金融機関名,支店名,支店番号,口座種別,口座番号,差押文言1,差押文言2,差押文言3,文書番号,照会結果ファイル名,処理済フラグ1,処理済フラグ2,処理済フラグ3";


    // --- UI要素 ---
    private Window window;
    private Grid initialPanel, mainPanel, overlayPanel, loadingOverlay, resultOverlay;
    private ComboBox sheetCombo;
    private TextBlock fileLink, statusLeft, statusRight, guideText, deliveryError;
    private TextBlock resultIcon, resultTitle, resultDetail, resultSub;
    private Border resultIconBg;                 // 結果アイコンの背景円
    private TextBox txtAddressNum, txtName, txtInstitution, txtStaff;
    private TextBox txtResidenceAddr, txtDeliveryAddr, txtExecDate;
    private CheckBox chkDeliveryOutput;
    private ListView accountList;
    private Button btnAdd, btnSkip, btnLoadFile, resultButton;
    private Button btnCalendar;                // カレンダーPopup表示ボタン
    private Popup calendarPopup;               // カレンダーPopup
    private System.Windows.Controls.Calendar dateCalendar;  // カレンダーコントロール
    private RotateTransform spinnerRotation;  // スピナーの回転トランスフォーム
    private dynamic excel;               // Excel.Application（COM late binding）
    private string currentFilePath;      // 現在処理中のファイルパス
    private string selectedSheetName;    // 選択中のシート名
    private int lastUsedRow;             // シートの最終使用行
    private string lastDocNumber;        // ロールバック用文書番号

    // ==============================================================
    // エントリポイント
    // ==============================================================

    [STAThread]
    public static void Main(string[] args)
    {
        // ブラシを Freeze して描画パフォーマンスを向上
        BrushBorderNormal.Freeze();
        BrushValidationError.Freeze();
        BrushSuccessIcon.Freeze();
        BrushAccent.Freeze();
        BrushIconBgSuccess.Freeze();
        BrushIconBgError.Freeze();
        BrushIconBgSkip.Freeze();
        BrushDetailMuted.Freeze();

        var app = new DepositSeizureApp();
        app.StartupArgs = args;
        app.Run();
    }

    private string[] StartupArgs { get; set; }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        exeDir = AppDomain.CurrentDomain.BaseDirectory;

        // --- 設定ファイル読込 ---
        var configPath = System.IO.Path.Combine(exeDir, "deposit_seizure_list_config.json");
        if (!File.Exists(configPath))
        { MessageBox.Show("設定ファイルが見つかりません。\n\n" + configPath, "エラー", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); return; }
        config = ConfigLoader.LoadConfig(configPath);
        if (config.Profiles.Count == 0)
        { MessageBox.Show("プロファイルが設定されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); return; }

        // --- 元号マッピング読込 ---
        eraMapping = ConfigLoader.LoadEraMapping(System.IO.Path.Combine(exeDir, "era_mapping.json"));
        if (eraMapping.Count == 0)
        { MessageBox.Show("era_mapping.json が見つからないか空です。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); return; }

        // --- 文書番号カウンター確認 ---
        if (!File.Exists(System.IO.Path.Combine(exeDir, "document_number_counter.json")))
        { MessageBox.Show("document_number_counter.json が見つかりません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); return; }

        // --- 任意マッピングファイル読込 ---
        feeMapping = ConfigLoader.LoadFeeMapping(System.IO.Path.Combine(exeDir, "transfer_fee_mapping.json"));
        yuchoMapping = ConfigLoader.LoadYuchoCenterMapping(System.IO.Path.Combine(exeDir, "yucho_center_mapping.json"));
        accountTypeMapping = ConfigLoader.LoadAccountTypeMapping(System.IO.Path.Combine(exeDir, "account_type_mapping.json"));
        institutionNameMapping = ConfigLoader.LoadSimpleMapping(System.IO.Path.Combine(exeDir, "institution_name_mapping.json"));
        branchNameMapping = ConfigLoader.LoadSimpleMapping(System.IO.Path.Combine(exeDir, "branch_name_mapping.json"));

        // --- プロファイル選択 ---
        activeProfile = config.Profiles[0]; // 複数時は最初を自動選択

        // --- プロファイルの最低限のバリデーション ---
        var validationErrors = new List<string>();
        if (string.IsNullOrWhiteSpace(activeProfile.CoverCell))
            validationErrors.Add("coverCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.CoverValue))
            validationErrors.Add("coverValue が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.AddressNumberCell))
            validationErrors.Add("addressNumberCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.NameCell))
            validationErrors.Add("nameCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.StaffCell))
            validationErrors.Add("staffCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.ResidenceAddressCell))
            validationErrors.Add("residenceAddressCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.FinancialInstitutionCell))
            validationErrors.Add("financialInstitutionCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.BranchNameCol))
            validationErrors.Add("branchNameCol が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.AccountNumberCol))
            validationErrors.Add("accountNumberCol が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.BalanceCol))
            validationErrors.Add("balanceCol が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.OutputFolder))
            validationErrors.Add("outputFolder が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.PrintFolder))
            validationErrors.Add("printFolder が未設定です");
        if (validationErrors.Count > 0)
        {
            MessageBox.Show(
                "プロファイル「" + activeProfile.Name + "」の設定に問題があります。\n\n" +
                string.Join("\n", validationErrors.ToArray()),
                "設定エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        // --- 出力先フォルダの自動作成 ---
        try
        {
            if (!Directory.Exists(activeProfile.OutputFolder))
                Directory.CreateDirectory(activeProfile.OutputFolder);
            if (!Directory.Exists(activeProfile.PrintFolder))
                Directory.CreateDirectory(activeProfile.PrintFolder);
        }
        catch (Exception dirEx)
        {
            MessageBox.Show(
                "出力先フォルダの作成に失敗しました。\n\n" + dirEx.Message,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        // --- 起動モード判定 ---
        if (StartupArgs != null && StartupArgs.Length > 0)
        {
            isFromFileSearch = true;
            foreach (var arg in StartupArgs)
                if (File.Exists(arg)) fileEntries.Add(new FileEntry { FilePath = arg, State = FileProcessState.Pending });
        }

        // --- 差押実績ルックアップ構築 ---
        BuildSeizureHistory();

        // --- ウィンドウ構築 ---
        window = BuildWindow();
        FindControls();
        SetupEvents();
        InitializeUI();
        if (isFromFileSearch && fileEntries.Count > 0)
            window.ContentRendered += delegate { LoadFileAtIndex(0); };
        window.Closed += delegate { CleanupExcel(); };
        window.Show();
    }

    private void FindControls()
    {
        initialPanel = (Grid)window.FindName("InitialPanel");
        mainPanel = (Grid)window.FindName("MainPanel");
        overlayPanel = (Grid)window.FindName("OverlayPanel");
        loadingOverlay = (Grid)window.FindName("LoadingOverlay");
        resultOverlay = (Grid)window.FindName("ResultOverlay");
        sheetCombo = (ComboBox)window.FindName("SheetCombo");
        fileLink = (TextBlock)window.FindName("FileLink");
        txtAddressNum = (TextBox)window.FindName("TxtAddressNum");
        txtName = (TextBox)window.FindName("TxtName");
        txtInstitution = (TextBox)window.FindName("TxtInstitution");
        txtStaff = (TextBox)window.FindName("TxtStaff");
        txtResidenceAddr = (TextBox)window.FindName("TxtResidenceAddr");
        txtDeliveryAddr = (TextBox)window.FindName("TxtDeliveryAddr");
        txtExecDate = (TextBox)window.FindName("TxtExecDate");
        chkDeliveryOutput = (CheckBox)window.FindName("ChkDeliveryOutput");
        accountList = (ListView)window.FindName("AccountList");
        btnAdd = (Button)window.FindName("BtnAdd");
        btnSkip = (Button)window.FindName("BtnSkip");
        btnLoadFile = (Button)window.FindName("BtnLoadFile");
        statusLeft = (TextBlock)window.FindName("StatusLeft");
        statusRight = (TextBlock)window.FindName("StatusRight");
        guideText = (TextBlock)window.FindName("GuideText");
        deliveryError = (TextBlock)window.FindName("DeliveryError");
        resultIcon = (TextBlock)window.FindName("ResultIcon");
        resultIconBg = (Border)window.FindName("ResultIconBg");
        resultTitle = (TextBlock)window.FindName("ResultTitle");
        resultDetail = (TextBlock)window.FindName("ResultDetail");
        resultButton = (Button)window.FindName("ResultButton");
        resultSub = (TextBlock)window.FindName("ResultSub");
        btnCalendar = (Button)window.FindName("BtnCalendar");
        calendarPopup = (Popup)window.FindName("CalendarPopup");
        dateCalendar = (System.Windows.Controls.Calendar)window.FindName("DateCalendar");

        // スピナーの RotateTransform を取得（回転アニメーションの開始/停止を制御するため）
        var spinnerElement = (FrameworkElement)window.FindName("SpinnerPath");
        if (spinnerElement != null)
            spinnerRotation = spinnerElement.RenderTransform as RotateTransform;
    }

    private void InitializeUI()
    {
        if (activeProfile.OutputFolder != null) statusRight.Text = activeProfile.OutputFolder;
        ShowState("initial");
    }

    // 画面状態の切替
    private void ShowState(string state)
    {
        initialPanel.Visibility = state == "initial" ? Visibility.Visible : Visibility.Collapsed;
        mainPanel.Visibility = state == "main" ? Visibility.Visible : Visibility.Collapsed;
        overlayPanel.Visibility = (state == "loading" || state == "result") ? Visibility.Visible : Visibility.Collapsed;
        loadingOverlay.Visibility = state == "loading" ? Visibility.Visible : Visibility.Collapsed;
        resultOverlay.Visibility = state == "result" ? Visibility.Visible : Visibility.Collapsed;
        if (state == "initial") statusLeft.Text = "";
    }

    // ==============================================================
    // イベントハンドラ
    // ==============================================================

    private void SetupEvents()
    {
        // D&D
        window.AllowDrop = true;
        window.Drop += delegate(object s, DragEventArgs de)
        {
            if (!de.Data.GetDataPresent(DataFormats.FileDrop)) return;
            var files = ((string[])de.Data.GetData(DataFormats.FileDrop))
                .Where(f => f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)).ToArray();
            if (files.Length > 0) StartFileProcessing(files);
        };
        window.DragOver += delegate(object s, DragEventArgs de)
        { de.Effects = de.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None; de.Handled = true; };

        // 初期画面ファイル選択
        var btnSelect = (Button)window.FindName("BtnSelectFile");
        if (btnSelect != null) btnSelect.Click += delegate { DoOpenFileDialog(true); };

        // メインフォームボタン
        btnLoadFile.Click += delegate { DoOpenFileDialog(false); };
        fileLink.MouseDown += delegate
        { if (!string.IsNullOrEmpty(currentFilePath) && File.Exists(currentFilePath)) try { Process.Start(new ProcessStartInfo(currentFilePath) { UseShellExecute = true }); } catch {} };
        var btnReload = (Button)window.FindName("BtnReload");
        if (btnReload != null) btnReload.Click += delegate { if (!string.IsNullOrEmpty(currentFilePath)) LoadSingleFile(currentFilePath); };

        // シート切替（PopulateForm 内での初期選択時は抑止）
        sheetCombo.SelectionChanged += delegate { if (!suppressSheetChange && sheetCombo.SelectedItem != null) { selectedSheetName = sheetCombo.SelectedItem.ToString(); ReloadSheetData(); } };

        // 口座選択・執行日・必須フィールドでボタン制御
        // 口座選択: ボタン制御 + 表紙ブロック変更時に届出住所を自動更新
        accountList.SelectionChanged += delegate
        {
            UpdateAddButton();
            var selectedAccount = accountList.SelectedItem as AccountItem;
            if (selectedAccount != null && selectedAccount.CoverIndex != lastDisplayedCoverIndex)
            {
                // 異なる表紙ブロックの口座が選択された → 届出住所をその表紙の値に更新
                lastDisplayedCoverIndex = selectedAccount.CoverIndex;
                txtDeliveryAddr.Text = BusinessLogic.FormatAddress(selectedAccount.DeliveryAddressRaw ?? "");
            }
        };
        txtName.TextChanged += delegate { UpdateAddButton(); };
        txtStaff.TextChanged += delegate { UpdateAddButton(); };
        txtResidenceAddr.TextChanged += delegate { UpdateAddButton(); };
        txtExecDate.LostFocus += delegate
        {
            var input = txtExecDate.Text.Trim();
            if (string.IsNullOrEmpty(input)) { processingDate = null; UpdateAddButton(); return; }
            var dt = BusinessLogic.ParseFlexibleDate(input, eraMapping);
            if (dt.HasValue)
            {
                processingDate = dt.Value;
                txtExecDate.Text = dt.Value.ToString("yyyy/MM/dd");
                txtExecDate.BorderBrush = BrushBorderNormal;
            }
            else
            {
                processingDate = null;
                txtExecDate.BorderBrush = BrushValidationError;
            }
            UpdateAddButton();
        };

        // 届出住所バリデーション
        txtDeliveryAddr.TextChanged += delegate { ValidateDelivery(); };
        chkDeliveryOutput.Checked += delegate { ValidateDelivery(); };
        chkDeliveryOutput.Unchecked += delegate { ValidateDelivery(); };

        // カレンダーPopup
        btnCalendar.Click += delegate
        {
            calendarPopup.PlacementTarget = btnCalendar;
            if (!calendarPopup.IsOpen)
            {
                // 前回の表示状態（月/年選択）が残るのを防止し、常に日付選択ビューで開く
                dateCalendar.DisplayMode = CalendarMode.Month;
                dateCalendar.DisplayDate = processingDate ?? DateTime.Today;
                // 選択をクリアすることで同じ日付の再クリックでも SelectedDatesChanged が発火する
                dateCalendar.SelectedDates.Clear();
            }
            calendarPopup.IsOpen = !calendarPopup.IsOpen;
        };
        dateCalendar.SelectedDatesChanged += delegate
        {
            if (dateCalendar.SelectedDate.HasValue)
            {
                var selectedDate = dateCalendar.SelectedDate.Value;
                processingDate = selectedDate;
                txtExecDate.Text = selectedDate.ToString("yyyy/MM/dd");
                txtExecDate.BorderBrush = BrushBorderNormal;
                calendarPopup.IsOpen = false;
                UpdateAddButton();
            }
        };

        // アクション
        btnAdd.Click += delegate { ExecuteAdd(); };
        btnSkip.Click += delegate { ExecuteSkip(); };
        resultButton.Click += delegate { ProceedToNext(); };

        // ショートカット
        window.InputBindings.Add(new KeyBinding(new RelayCommand(p => { if (overlayPanel.Visibility == Visibility.Visible) ProceedToNext(); }), new KeyGesture(Key.Escape)));
        window.InputBindings.Add(new KeyBinding(new RelayCommand(p => DoOpenFileDialog(initialPanel.Visibility == Visibility.Visible)), new KeyGesture(Key.O, ModifierKeys.Control)));

        // 空白クリックでテキストボックスからフォーカスを外す
        window.MouseDown += delegate(object s, MouseButtonEventArgs me)
        {
            if (me.OriginalSource is System.Windows.Controls.Panel ||
                me.OriginalSource is Border ||
                me.OriginalSource is Window)
            {
                FocusManager.SetFocusedElement(window, window);
                Keyboard.ClearFocus();
            }
        };

        // 口座テーブルの空白エリアクリックでもフォーカスを外す
        // VisualTree を辿って ListViewItem が見つからなければ空白エリアと判定
        accountList.PreviewMouseDown += delegate(object s, MouseButtonEventArgs me)
        {
            var hit = me.OriginalSource as DependencyObject;
            while (hit != null && hit != accountList)
            {
                if (hit is ListViewItem) return;
                hit = VisualTreeHelper.GetParent(hit);
            }
            FocusManager.SetFocusedElement(window, window);
            Keyboard.ClearFocus();
        };

        // 口座テーブルの列幅自動調整
        accountList.SizeChanged += delegate { AdjustAccountColumns(); };
        accountList.Loaded += delegate { AdjustAccountColumns(); };

        // スピナーアニメーション: LoadingOverlay の表示/非表示に連動
        loadingOverlay.IsVisibleChanged += delegate(object s, DependencyPropertyChangedEventArgs dpce)
        {
            if ((bool)dpce.NewValue) StartSpinner();
            else StopSpinner();
        };
    }

    private void UpdateAddButton()
    {
        // 口座が0件の場合は専用メッセージ
        if (currentAccounts.Count == 0 && mainPanel.Visibility == Visibility.Visible)
        {
            guideText.Text = "表示可能な口座がありません";
            btnAdd.IsEnabled = false;
            return;
        }
        string msg = "";
        if (string.IsNullOrWhiteSpace(txtName.Text)) msg = "氏名を入力してください";
        else if (string.IsNullOrWhiteSpace(txtStaff.Text)) msg = "処分担当を入力してください";
        else if (string.IsNullOrWhiteSpace(txtResidenceAddr.Text)) msg = "住民票住所を入力してください";
        else if (accountList.SelectedItem == null) msg = "口座を選択してください";
        else if (!processingDate.HasValue) msg = "執行日を入力してください";
        btnAdd.IsEnabled = string.IsNullOrEmpty(msg);
        guideText.Text = msg;
    }

    // 口座テーブルの列幅自動調整
    // 支店名のみ可変幅（残りスペースを全て吸収）、他5列は固定幅
    private void AdjustAccountColumns()
    {
        var gv = accountList.View as GridView;
        if (gv == null || gv.Columns.Count < 6) return;

        double total = accountList.ActualWidth - SystemParameters.VerticalScrollBarWidth - 4;
        if (total <= 0) return;

        double col1 = 105;  // 支店番号
        double col2 = 110;  // 口座種別
        double col3 = 150;  // 最終取引日(満期日)
        double col4 = 135;  // 口座番号
        double col5 = 160;  // 残高
        double branchW = total - col1 - col2 - col3 - col4 - col5;
        if (branchW < 120) branchW = 120;

        gv.Columns[0].Width = branchW;
        gv.Columns[1].Width = col1;
        gv.Columns[2].Width = col2;
        gv.Columns[3].Width = col3;
        gv.Columns[4].Width = col4;
        gv.Columns[5].Width = col5;
    }

    // スピナー回転アニメーションを開始（RepeatBehavior.Forever で無限回転）
    private void StartSpinner()
    {
        if (spinnerRotation == null) return;
        var anim = new DoubleAnimation();
        anim.By = 360;
        anim.Duration = new Duration(TimeSpan.FromSeconds(1));
        anim.RepeatBehavior = RepeatBehavior.Forever;
        spinnerRotation.BeginAnimation(RotateTransform.AngleProperty, anim);
    }

    // スピナー回転アニメーションを停止
    private void StopSpinner()
    {
        if (spinnerRotation == null) return;
        spinnerRotation.BeginAnimation(RotateTransform.AngleProperty, null);
    }

    // オーバーレイのフェードイン（Opacity 0→1）
    private void FadeInOverlay()
    {
        overlayPanel.BeginAnimation(UIElement.OpacityProperty, null);
        overlayPanel.Opacity = 0;
        overlayPanel.Visibility = Visibility.Visible;
        var anim = new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150)));
        overlayPanel.BeginAnimation(UIElement.OpacityProperty, anim);
    }

    // オーバーレイのフェードアウト（Opacity 1→0 → Collapsed）
    private void FadeOutOverlay(Action onComplete = null)
    {
        var anim = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(150)));
        anim.Completed += delegate
        {
            overlayPanel.Visibility = Visibility.Collapsed;
            overlayPanel.BeginAnimation(UIElement.OpacityProperty, null);
            overlayPanel.Opacity = 1;
            if (onComplete != null) onComplete();
        };
        overlayPanel.BeginAnimation(UIElement.OpacityProperty, anim);
    }

    // ローディング→結果カードのクロスフェード切替
    // オーバーレイを一瞬フェードアウト → コンテンツ差替 → フェードイン
    private void CrossFadeToResult(string type, string title, string detail)
    {
        var fadeOut = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(100)));
        fadeOut.Completed += delegate
        {
            ShowResult(type, title, detail);
            var fadeIn = new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(100)));
            overlayPanel.BeginAnimation(UIElement.OpacityProperty, fadeIn);
        };
        overlayPanel.BeginAnimation(UIElement.OpacityProperty, fadeOut);
    }

    private void ValidateDelivery()
    {
        if (chkDeliveryOutput.IsChecked == true)
        {
            int len = ("（届出：" + txtDeliveryAddr.Text.Trim() + "）").Length;
            if (len > 50) { deliveryError.Text = "50文字を超えています（現在: " + len + "文字）"; deliveryError.Visibility = Visibility.Visible; return; }
        }
        deliveryError.Visibility = Visibility.Collapsed;
    }

    private void DoOpenFileDialog(bool isInitial)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "Excel (*.xlsm)|*.xlsm", Multiselect = true, Title = "照会結果ファイルを選択" };
        if (!isInitial && !string.IsNullOrEmpty(currentFilePath)) dlg.InitialDirectory = System.IO.Path.GetDirectoryName(currentFilePath);
        else if (config.DefaultFolder != null && Directory.Exists(config.DefaultFolder)) dlg.InitialDirectory = config.DefaultFolder;
        if (dlg.ShowDialog() == true) StartFileProcessing(dlg.FileNames);
    }

    private void StartFileProcessing(string[] paths)
    {
        fileEntries.Clear();
        foreach (var p in paths) fileEntries.Add(new FileEntry { FilePath = p, State = FileProcessState.Pending });
        LoadFileAtIndex(0);
    }

    private void LoadFileAtIndex(int index)
    {
        if (index >= fileEntries.Count) { if (isFromFileSearch) Shutdown(); else ShowState("initial"); return; }
        currentFileIndex = index; currentFilePath = fileEntries[index].FilePath;
        statusLeft.Text = fileEntries.Count > 1 ? (index + 1) + " / " + fileEntries.Count + " 件目" : "";

        // 新しいファイルに切り替わるとき、執行日と届出住所チェックをリセット
        txtExecDate.Text = ""; processingDate = null;
        chkDeliveryOutput.IsChecked = false;
        deliveryError.Visibility = Visibility.Collapsed;
        lastDisplayedCoverIndex = -1;

        LoadSingleFile(currentFilePath);
    }

    // ==============================================================
    // ファイル読込（Excel COM・非同期）
    // ==============================================================

    private void LoadSingleFile(string filePath)
    {
        ShowState("main");
        mainPanel.Visibility = Visibility.Visible;
        loadingOverlay.Visibility = Visibility.Visible;
        resultOverlay.Visibility = Visibility.Collapsed;
        FadeInOverlay();

        var worker = new BackgroundWorker();
        worker.DoWork += delegate(object s, DoWorkEventArgs args) { args.Result = ReadExcelFile(filePath); };
        worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs args)
        {
            if (args.Error != null) { CrossFadeToResult("error", "読込失敗", args.Error.Message); return; }
            var data = args.Result as Dictionary<string, object>;
            if (data != null && data.ContainsKey("error")) { CrossFadeToResult("error", "読込失敗", data["error"].ToString()); return; }
            PopulateForm(data);
            FadeOutOverlay();
        };
        worker.RunWorkerAsync();
    }

    // Excel COM でファイルからデータを読み取る
    // BackgroundWorker から呼ばれるためUIスレッドからは分離されている
    private Dictionary<string, object> ReadExcelFile(string filePath)
    {
        var result = new Dictionary<string, object>();

        // Excel COM の初回起動（アプリ生存期間中インスタンスを保持）
        if (excel == null)
        {
            var t = Type.GetTypeFromProgID("Excel.Application");
            if (t == null)
                return new Dictionary<string, object> { { "error", "Excelがインストールされていません" } };

            excel = Activator.CreateInstance(t);
            excel.Visible = false;
            excel.DisplayAlerts = false;

            // マクロ無効化（msoAutomationSecurityForceDisable = 3）
            try { excel.AutomationSecurity = 3; } catch { }

            // 自動計算を無効化（xlCalculationManual = -4135）
            // ファイルを開いた際の自動計算による遅延を防止
            try { excel.Calculation = -4135; } catch { }

            // 画面更新を無効化（非表示なので不要だが、念のため明示的に設定）
            try { excel.ScreenUpdating = false; } catch { }

            // イベント発火を無効化（マクロ付きファイルの Workbook_Open 等を抑止）
            try { excel.EnableEvents = false; } catch { }
        }
        dynamic wb = null;
        try
        {
            wb = excel.Workbooks.Open(filePath, 0, true); // 読取専用

            // シートフィルタ: filterCell/filterValue で対象シートを絞り込む
            var sheets = new List<string>();
            for (int i = 1; i <= (int)wb.Worksheets.Count; i++)
            {
                dynamic ws = wb.Worksheets[i];
                try
                {
                    string name = (string)ws.Name;
                    if (!string.IsNullOrEmpty(activeProfile.FilterCell) &&
                        activeProfile.FilterValues != null && activeProfile.FilterValues.Length > 0)
                    {
                        // filterCell の値に filterValues のいずれかが部分一致するシートを対象（OR条件）
                        try
                        {
                            string cellValue = Convert.ToString(ws.Range[activeProfile.FilterCell].Value2 ?? "");
                            bool matches = false;
                            foreach (var fv in activeProfile.FilterValues)
                            {
                                if (cellValue.IndexOf(fv, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    matches = true;
                                    break;
                                }
                            }
                            if (matches)
                                sheets.Add(name);
                        }
                        catch { /* セル読取り失敗時はそのシートをスキップ */ }
                    }
                    else
                    {
                        // フィルタ未設定時は Visible シートを全て対象
                        if ((int)ws.Visible == -1) sheets.Add(name);
                    }
                }
                catch { /* シート情報取得失敗は無視 */ }
            }

            result["sheets"] = sheets;
            result["filePath"] = filePath;
            result["fileName"] = System.IO.Path.GetFileName(filePath);

            // 一致シートが0件の場合
            if (sheets.Count == 0)
            {
                result["noSheet"] = true;
                try { wb.Close(false); } catch { }
                wb = null;
                return result;
            }

            result["selectedSheet"] = sheets[0];

            // 先頭シートのデータを読み取る
            try
            {
                result["sheetData"] = ReadSheetData(wb, sheets[0]);
            }
            catch (Exception rex)
            {
                result["error"] = "シートデータの読取りに失敗: " + rex.Message;
            }
        }
        catch (Exception ex) { result["error"] = ex.Message; }
        finally
        {
            if (wb != null)
            {
                try { wb.Close(false); } catch { }
                // 注: ReleaseComObject は dynamic late binding 環境では
                // RPC_E_SERVER_DIED_DNE を引き起こすため使用しない。
                // ワークブックの解放は CleanupExcel() の GC.Collect で行う。
            }
        }
        return result;
    }

    // 指定シートから基本情報と口座テーブルを読み取る
    // 表紙ブロック検出 → 基本情報取得 → 口座テーブル読取り → マッピング補完 の順で処理
    private Dictionary<string, object> ReadSheetData(dynamic wb, string sheetName)
    {
        var data = new Dictionary<string, object>();
        dynamic ws = wb.Worksheets[sheetName];

        try
        {
            // ── 表紙ブロック検出 + stopRows 一括収集 ──
            // coverCell 列を行1～最終行まで走査し、coverValue / stopValues を同時に検出
            // コンソール版と同一のアプローチ（全行一括スキャン）
            var offsets = new List<int>();
            var stopRowList = new List<int>();
            try
            {
                lastUsedRow = (int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count - 1;

                // coverCell（例: "A2"）を列部分と行番号に分解
                int colEndIdx = 0;
                while (colEndIdx < activeProfile.CoverCell.Length && char.IsLetter(activeProfile.CoverCell[colEndIdx]))
                    colEndIdx++;
                string coverCol = activeProfile.CoverCell.Substring(0, colEndIdx);
                int coverRow = int.Parse(activeProfile.CoverCell.Substring(colEndIdx));

                // coverCell列を全行スキャンして coverOffsets と stopRows を同時に収集
                for (int r = 1; r <= lastUsedRow; r++)
                {
                    try
                    {
                        string cellValue = Convert.ToString(ws.Range[coverCol + r.ToString()].Value2 ?? "");
                        string trimmed = cellValue.Trim();

                        if (trimmed == activeProfile.CoverValue)
                        {
                            // 表紙の目印を検出（オフセット = 検出行 - テンプレート上のcoverCell行）
                            offsets.Add(r - coverRow);
                        }
                        else if (activeProfile.StopValues != null && trimmed.Length > 0)
                        {
                            // stopValues に一致する行を検出（明細ページヘッダー等）
                            foreach (var sv in activeProfile.StopValues)
                            {
                                if (trimmed == sv)
                                {
                                    stopRowList.Add(r);
                                    break;
                                }
                            }
                        }
                    }
                    catch { /* セル読取り失敗は無視 */ }
                }
            }
            catch { /* UsedRange 取得失敗時はオフセット0で続行 */ }

            if (offsets.Count == 0) offsets.Add(0);
            data["coverOffsets"] = offsets;

            // ── 基本情報取得（1つ目の表紙から） ──
            int firstOffset = offsets[0];
            data["addressNum"] = GetCell(ws, activeProfile.AddressNumberCell, firstOffset);
            data["name"] = GetCell(ws, activeProfile.NameCell, firstOffset);
            data["staff"] = GetCell(ws, activeProfile.StaffCell, firstOffset);
            data["residenceAddr"] = GetCell(ws, activeProfile.ResidenceAddressCell, firstOffset);
            data["deliveryAddr"] = GetCell(ws, activeProfile.DeliveryAddressCell, firstOffset);

            // 金融機関名の取得とマッピング補正
            string institution = GetCell(ws, activeProfile.FinancialInstitutionCell, firstOffset).Trim();
            if (institutionNameMapping != null)
            {
                string normalizedInst = BusinessLogic.ToFullWidth(institution);
                foreach (var kv in institutionNameMapping)
                {
                    if (BusinessLogic.ToFullWidth(kv.Key) == normalizedInst)
                    {
                        institution = kv.Value;
                        break;
                    }
                }
            }
            data["institution"] = institution;

            // ── 口座テーブル読取り ──
            // 各表紙ブロックについて、accountStartRow から stopValues 検出行まで走査
            var accounts = new List<AccountItem>();
            bool isYucho = institution.Contains("ゆうちょ") || institution.Contains("郵貯");

            foreach (var coverOffset in offsets)
            {
                int startRow = activeProfile.AccountStartRow + coverOffset;
                int coverIndex = offsets.IndexOf(coverOffset);

                // 各表紙ブロックの届出住所を取得（表紙ごとに異なる可能性があるため）
                string coverDeliveryAddr = GetCell(ws, activeProfile.DeliveryAddressCell, coverOffset);

                // 次の表紙ブロックの口座テーブル開始行を算出（越境防止）
                // コンソール版と同様に、次の表紙・stopRow・200行上限の最小値で制限
                int nextBlockStartRow = (coverIndex + 1 < offsets.Count)
                    ? activeProfile.AccountStartRow + offsets[coverIndex + 1]
                    : startRow + ACCOUNT_TABLE_MAX_ROWS;

                for (int r = startRow; r < startRow + ACCOUNT_TABLE_MAX_ROWS; r++)
                {
                    // 次の表紙ブロックに到達したら停止（防御的チェック）
                    if (r >= nextBlockStartRow)
                        break;

                    // stopValues チェック（口座テーブルの終端検出）
                    // ※ stopRowList は表紙ブロック検出時に一括収集済み。
                    //    ここでは口座テーブルの読取り終了判定のみ行う。
                    if (stopRowList.Contains(r))
                        break;

                    // 口座番号の取得
                    string accountNum = "";
                    try { accountNum = Convert.ToString(ws.Range[activeProfile.AccountNumberCol + r].Value2 ?? ""); }
                    catch { }
                    if (string.IsNullOrWhiteSpace(accountNum)) continue;

                    // 支店番号の取得
                    string branchNumber = "";
                    try { branchNumber = Convert.ToString(ws.Range[activeProfile.BranchNumberCol + r].Value2 ?? ""); }
                    catch { }

                    // 口座番号・支店番号がともに非数値の行はスキップ
                    // （継続ページのヘッダー行「支店名｜支店番号｜口座種別｜...」を読み飛ばす）
                    bool accountNumIsNumeric = accountNum.Trim().Length > 0 &&
                        accountNum.Trim().All(char.IsDigit);
                    bool branchNumIsNumeric = branchNumber.Trim().Length > 0 &&
                        branchNumber.Trim().All(char.IsDigit);
                    if (!accountNumIsNumeric && !branchNumIsNumeric) continue;

                    // 支店名の取得（内部空白も全て除去。コンソール版と同一の正規化）
                    string branchName = "";
                    try { branchName = Convert.ToString(ws.Range[activeProfile.BranchNameCol + r].Value2 ?? ""); }
                    catch { }
                    branchName = branchName.Trim().Replace(" ", "").Replace("\u3000", "");

                    // ゆうちょ銀行の場合、支店名が空なら貯金事務センター名で補完
                    if (isYucho && string.IsNullOrWhiteSpace(branchName) && yuchoMapping != null)
                    {
                        branchName = GetYuchoCenter(branchNumber.Trim()) ?? "";
                    }

                    // 支店名が空の口座は表示しない（GUI版ではダイアログ手入力を廃止）
                    if (string.IsNullOrWhiteSpace(branchName)) continue;

                    // 支店名「支店」サフィックス付与（ゆうちょ以外のみ）
                    if (!isYucho && !string.IsNullOrEmpty(branchName) &&
                        !IsBranchExempt(branchName) && !branchName.EndsWith("支店"))
                    {
                        branchName += "支店";
                    }

                    // 口座種別の取得
                    string accountType = "";
                    try { accountType = Convert.ToString(ws.Range[activeProfile.AccountTypeCol + r].Value2 ?? ""); }
                    catch { }

                    // 最終取引日の取得（Excelシリアル値または文字列に対応）
                    string lastTransaction = "";
                    try
                    {
                        var lastTransValue = ws.Range[activeProfile.LastTransactionCol + r].Value2;
                        if (lastTransValue is double)
                            lastTransaction = DateTime.FromOADate((double)lastTransValue).ToString("yyyy/MM/dd");
                        else if (lastTransValue != null)
                            lastTransaction = lastTransValue.ToString();
                    }
                    catch { }

                    // 残高の取得
                    double balanceValue = 0;
                    try { balanceValue = Convert.ToDouble(ws.Range[activeProfile.BalanceCol + r].Value2 ?? 0); }
                    catch { }

                    // AccountItem を構築
                    var item = new AccountItem
                    {
                        BranchName = branchName,
                        BranchNumber = branchNumber.Trim(),
                        AccountType = accountType.Trim(),
                        LastTransaction = lastTransaction,
                        AccountNum = accountNum.Trim(),
                        BalanceValue = balanceValue,
                        Balance = BusinessLogic.FormatBalance(balanceValue),
                        CoverIndex = coverIndex,
                        DeliveryAddressRaw = coverDeliveryAddr
                    };

                    // 差押実績ルックアップ（宛名番号＋口座番号をキーにCSV既存行と照合）
                    string historyKey = BusinessLogic.FormatAddressNumber(
                        (data["addressNum"] ?? "").ToString()) + "_" + item.AccountNum;
                    if (seizureHistory.ContainsKey(historyKey))
                    {
                        item.HasSeizureHistory = true;
                        var hist = seizureHistory[historyKey];
                        item.SeizureDocNumber = hist[0];
                        item.SeizureExecDate = hist[1];
                    }

                    accounts.Add(item);
                }
            }

            data["accounts"] = accounts;
            data["stopRows"] = stopRowList;
        }
        finally
        {
            // 注: ws の ReleaseComObject は dynamic late binding 環境では
            // Excel COM 接続を切断するため使用しない。
            // GC.Collect による自然回収に任せる。
        }

        return data;
    }

    // 指定セルの値を文字列として取得（行オフセット対応）
    private string GetCell(dynamic ws, string addr, int offset)
    {
        if (string.IsNullOrEmpty(addr)) return "";
        try
        {
            string offsetAddr = BusinessLogic.GetOffsetCell(addr, offset);
            return Convert.ToString(ws.Range[offsetAddr].Value2 ?? "");
        }
        catch { return ""; }
    }

    // ゆうちょ銀行の支店番号（記号）から貯金事務センター名を返す
    // 1桁目: "1"=総合口座, "0"=振替口座、2〜3桁目: 地域コード
    // 該当なしの場合は null を返す（GUI版では支店名空として表示対象外になる）
    private string GetYuchoCenter(string branchNumber)
    {
        if (yuchoMapping == null || branchNumber.Length != 5)
            return null;

        // 1桁目で口座種別を判定
        string accountCategory;
        if (branchNumber[0] == '1')
            accountCategory = "総合口座";
        else if (branchNumber[0] == '0')
            accountCategory = "振替口座";
        else
            return null;

        // 2〜3桁目の地域コードでセンター名を引く
        Dictionary<string, string> centerTable;
        if (!yuchoMapping.TryGetValue(accountCategory, out centerTable))
            return null;

        string center;
        return centerTable.TryGetValue(branchNumber.Substring(1, 2), out center) ? center : null;
    }

    // 「支店」サフィックスを付与しない支店名かどうかを判定
    // 末尾が「営業部」「出張所」「公務部」「本店」で終わる支店名は「支店」を付けない
    private bool IsBranchExempt(string branchName)
    {
        return branchName.EndsWith("営業部") ||
               branchName.EndsWith("出張所") ||
               branchName.EndsWith("公務部") ||
               branchName.EndsWith("本店");
    }

    // フォームにデータ反映
    private void PopulateForm(Dictionary<string, object> data)
    {
        if (data == null) return;
        fileLink.Text = (data["fileName"] ?? "").ToString();
        var sheets = data["sheets"] as List<string>;
        sheetCombo.Items.Clear();
        if (sheets != null) foreach (var s in sheets) sheetCombo.Items.Add(s);
        if (data.ContainsKey("noSheet") && (bool)data["noSheet"])
        {
            // 対象シートなし → メインフォーム上にメッセージ、スキップで次へ
            txtAddressNum.Text = ""; txtName.Text = ""; txtInstitution.Text = "";
            txtStaff.Text = ""; txtResidenceAddr.Text = ""; txtDeliveryAddr.Text = "";
            currentAccounts.Clear(); accountList.ItemsSource = null;
            guideText.Text = "対象シートがありません（フィルタ条件: " +
                (activeProfile.FilterValues != null ? string.Join(", ", activeProfile.FilterValues) : "") + "）";
            btnAdd.IsEnabled = false; return;
        }
        // 初期選択時は SelectionChanged による ReloadSheetData を抑止
        // （ReadExcelFile で既にシートデータを読み取り済みのため二重読込を回避）
        suppressSheetChange = true;
        if (sheetCombo.Items.Count > 0) sheetCombo.SelectedIndex = 0;
        suppressSheetChange = false;
        if (sheetCombo.SelectedItem != null) selectedSheetName = sheetCombo.SelectedItem.ToString();
        if (data.ContainsKey("sheetData")) ApplySheet(data["sheetData"] as Dictionary<string, object>);
    }

    private void ApplySheet(Dictionary<string, object> d)
    {
        if (d == null) return;
        txtAddressNum.Text = BusinessLogic.FormatAddressNumber((d["addressNum"]??"").ToString());
        txtName.Text = (d["name"]??"").ToString().Trim();
        txtInstitution.Text = (d["institution"]??"").ToString().Trim();
        txtStaff.Text = (d["staff"]??"").ToString().Trim();
        txtResidenceAddr.Text = BusinessLogic.FormatAddress((d["residenceAddr"]??"").ToString());
        txtDeliveryAddr.Text = BusinessLogic.FormatAddress((d["deliveryAddr"]??"").ToString());
        // 届出住所は最初の表紙ブロックの値を初期表示。
        // 口座選択時に異なる表紙ブロックが選ばれたら自動更新される。
        lastDisplayedCoverIndex = 0;
        coverOffsets = d["coverOffsets"] as List<int> ?? new List<int>();
        stopRows = d["stopRows"] as List<int> ?? new List<int>();
        currentAccounts = d["accounts"] as List<AccountItem> ?? new List<AccountItem>();
        accountList.ItemsSource = null; accountList.ItemsSource = currentAccounts;
        if (currentAccounts.Count == 0)
            guideText.Text = "表示可能な口座がありません";
        UpdateAddButton();
    }

    // シート切替時にデータを再読み込みする
    // 現在のファイルを再度開いて、選択中のシートからデータを取得し直す
    private void ReloadSheetData()
    {
        if (excel == null || string.IsNullOrEmpty(currentFilePath)) return;

        // オーバーレイ表示（処理中スピナー）
        loadingOverlay.Visibility = Visibility.Visible;
        resultOverlay.Visibility = Visibility.Collapsed;
        FadeInOverlay();

        var worker = new BackgroundWorker();
        worker.DoWork += delegate(object s, DoWorkEventArgs args)
        {
            dynamic wb = null;
            try
            {
                wb = excel.Workbooks.Open(currentFilePath, 0, true);  // 読取専用
                args.Result = ReadSheetData(wb, selectedSheetName);
            }
            finally
            {
                if (wb != null)
                    try { wb.Close(false); } catch { }
            }
        };
        worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs args)
        {
            if (args.Error == null)
                ApplySheet(args.Result as Dictionary<string, object>);
            FadeOutOverlay();
        };
        worker.RunWorkerAsync();
    }

    // ==============================================================
    // 処理実行（一覧に追加・スキップ）
    // ==============================================================

    // 「一覧に追加」ボタン押下時の処理
    // バリデーション → UI入力値をDictionaryに収集 → BackgroundWorkerで非同期処理
    private void ExecuteAdd()
    {
        // 口座選択と執行日の最終チェック
        var selectedAccount = accountList.SelectedItem as AccountItem;
        if (selectedAccount == null || !processingDate.HasValue) return;

        // 届出住所50文字チェック（チェックONの場合のみ）
        if (chkDeliveryOutput.IsChecked == true)
        {
            string deliveryFull = "（届出：" + txtDeliveryAddr.Text.Trim() + "）";
            if (deliveryFull.Length > 50)
            {
                deliveryError.Visibility = Visibility.Visible;
                return;
            }
        }

        // オーバーレイ表示（処理中スピナー）
        loadingOverlay.Visibility = Visibility.Visible;
        resultOverlay.Visibility = Visibility.Collapsed;
        FadeInOverlay();

        // UI入力値をDictionaryに収集（BackgroundWorkerに渡すため）
        string deliveryAddr = (chkDeliveryOutput.IsChecked == true)
            ? "（届出：" + txtDeliveryAddr.Text.Trim() + "）"
            : "";
        var addData = new Dictionary<string, string>
        {
            { "addressNum",    txtAddressNum.Text.Trim() },
            { "name",          txtName.Text.Trim() },
            { "staff",         txtStaff.Text.Trim() },
            { "institution",   txtInstitution.Text.Trim() },
            { "residenceAddr", txtResidenceAddr.Text.Trim() },
            { "deliveryAddr",  deliveryAddr },
            { "execDate",      BusinessLogic.DateToWareki(processingDate.Value, eraMapping) },
            { "branchName",    selectedAccount.BranchName },
            { "branchNumber",  selectedAccount.BranchNumber },
            { "accountType",   selectedAccount.AccountType },
            { "accountNum",    selectedAccount.AccountNum },
            { "filePath",      currentFilePath },
            { "fileName",      System.IO.Path.GetFileName(currentFilePath) }
        };
        int coverIndex = selectedAccount.CoverIndex;

        // BackgroundWorker で非同期実行（Excel COM操作を含むため）
        var worker = new BackgroundWorker();
        worker.DoWork += delegate(object s, DoWorkEventArgs args)
        {
            args.Result = ProcessAdd(addData, coverIndex);
        };
        worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs args)
        {
            if (args.Error != null)
            {
                CrossFadeToResult("error", "処理失敗", args.Error.Message);
                return;
            }

            var result = args.Result as Dictionary<string, string>;
            if (result["status"] == "ok")
            {
                fileEntries[currentFileIndex].State = FileProcessState.Added;
                CrossFadeToResult("success", "一覧に追加しました",
                    "文書番号: " + result["docNumber"]);
            }
            else
            {
                CrossFadeToResult("error", "処理失敗", result["message"]);
            }
        };
        worker.RunWorkerAsync();
    }

    // 一覧への追加処理を実行する（BackgroundWorker から呼ばれる）
    // 処理フロー:
    //   1. 文書番号の排他ロック採番
    //   2. 差押文言3行の生成
    //   3. CSVブランチ名マッピング適用
    //   4. CSV行の構築（20列、RFC 4180準拠エスケープ）
    //   5. CSV追記（排他ロック付き FileStream、リトライ最大5回）
    //   6. 印刷用ファイル保存
    private Dictionary<string, string> ProcessAdd(Dictionary<string, string> addData, int coverIdx)
    {
        var result = new Dictionary<string, string>();

        // 1. 文書番号の排他ロック採番
        string docNumber;
        if (!AllocateDocNumber(out docNumber))
        {
            result["status"] = "error";
            result["message"] = "文書番号の取得に失敗";
            return result;
        }

        // 2. 差押文言3行の生成
        var seizureText = GenerateSeizureText(
            addData["institution"],
            addData["branchName"],
            addData["branchNumber"],
            addData["accountType"],
            addData["accountNum"]);

        // 3. CSVブランチ名マッピング適用
        // 銀行ごとにCSV出力の支店名列を別の文言に置き換える
        // （差押文言・確認表示には反映せず、CSVの支店名列のみに使用）
        string csvBranchName = addData["branchName"];
        if (branchNameMapping != null)
        {
            string normalizedInst = BusinessLogic.ToFullWidth(addData["institution"]);
            foreach (var kv in branchNameMapping)
            {
                if (BusinessLogic.ToFullWidth(kv.Key) == normalizedInst)
                {
                    csvBranchName = kv.Value;
                    break;
                }
            }
        }

        // 4. CSV行の構築（20列）
        string printFileName = docNumber + ".xlsm";
        var csvFields = new[]
        {
            DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),  // 登録日時
            addData["addressNum"],                           // 宛名番号
            addData["name"],                                 // 氏名
            addData["staff"],                                // 職員名（処分担当）
            addData["execDate"],                             // 執行日（7桁和暦）
            addData["residenceAddr"],                        // 住民票住所
            addData["deliveryAddr"],                          // 銀行届出住所（空 or「（届出：...）」）
            addData["institution"],                           // 金融機関名
            csvBranchName,                                    // 支店名（マッピング適用後）
            addData["branchNumber"],                          // 支店番号
            addData["accountType"],                           // 口座種別
            addData["accountNum"],                            // 口座番号
            seizureText["Line1"],                             // 差押文言1
            seizureText["Line2"],                             // 差押文言2（振込手数料）
            seizureText["Line3"],                             // 差押文言3（残高・差押額テンプレ）
            docNumber,                                        // 文書番号
            printFileName,                                    // 照会結果ファイル名
            "",                                               // 処理済フラグ1（空）
            "",                                               // 処理済フラグ2（空）
            ""                                                // 処理済フラグ3（空）
        };
        string csvLine = string.Join(",", csvFields.Select(f => BusinessLogic.CsvEscape(f)));

        // 5. CSV追記
        string csvPath = System.IO.Path.Combine(activeProfile.OutputFolder, CSV_FILENAME);
        if (!WriteCsvLine(csvPath, csvLine))
        {
            RollbackDocNumber();
            result["status"] = "error";
            result["message"] = "CSV書き込み失敗";
            return result;
        }

        // 6. 印刷用ファイル保存
        string printFilePath = System.IO.Path.Combine(activeProfile.PrintFolder, printFileName);
        SavePrintFile(addData["filePath"], printFilePath, addData["accountNum"], addData["accountType"], coverIdx);

        result["status"] = "ok";
        result["docNumber"] = docNumber;
        result["printFile"] = printFileName;
        return result;
    }

    // 現在のファイルをスキップし、スキップオーバーレイを表示する
    private void ExecuteSkip()
    {
        if (currentFileIndex >= 0 && currentFileIndex < fileEntries.Count)
            fileEntries[currentFileIndex].State = FileProcessState.Skipped;

        ShowResult("skip", "スキップしました", System.IO.Path.GetFileName(currentFilePath));
        FadeInOverlay();
    }

    // ==============================================================
    // 処理結果オーバーレイ
    // ==============================================================

    private void ShowResult(string type, string title, string detail)
    {
        overlayPanel.Visibility = Visibility.Visible; loadingOverlay.Visibility = Visibility.Collapsed; resultOverlay.Visibility = Visibility.Visible;
        resultTitle.Text = title;
        resultDetail.Text = detail;
        resultDetail.Visibility = string.IsNullOrEmpty(detail) ? Visibility.Collapsed : Visibility.Visible;
        bool last = currentFileIndex >= fileEntries.Count - 1;
        if (type == "success")
        {
            resultIcon.Text = "\u2713";
            resultIcon.Foreground = BrushSuccessIcon;
            resultIconBg.Background = BrushIconBgSuccess;
            resultDetail.Foreground = BrushAccent;
        }
        else if (type == "skip")
        {
            resultIcon.Text = "\u2192";
            resultIcon.Foreground = BrushAccent;
            resultIconBg.Background = BrushIconBgSkip;
            resultDetail.Foreground = BrushDetailMuted;
        }
        else
        {
            resultIcon.Text = "\u2717";
            resultIcon.Foreground = BrushValidationError;
            resultIconBg.Background = BrushIconBgError;
            resultDetail.Foreground = BrushValidationError;
        }
        resultButton.Content = last ? "完了" : "次のファイルへ \u2192";
        resultSub.Text = last ? "" : ((currentFileIndex + 2) + " / " + fileEntries.Count + " 件目へ進みます");
        resultSub.Visibility = last ? Visibility.Collapsed : Visibility.Visible;
    }

    // 次のファイルへ進む（オーバーレイをフェードアウトしてから次のインデックスを読み込む）
    private void ProceedToNext()
    {
        FadeOutOverlay(delegate { LoadFileAtIndex(currentFileIndex + 1); });
    }

    // ==============================================================
    // 差押文言生成
    // ==============================================================

    // 口座情報とマッピング情報から差押文言3行を生成する
    //
    // Line1: 滞納者・金融機関・支店・口座種別・口座番号を埋め込んだ差押条項
    //        ゆうちょ銀行→「記号番号：支店番号－口座番号」形式
    //        その他     →「口座番号：口座番号」形式
    // Line2: 振込手数料の文言（マッピングにあればその値、なければ空）
    // Line3: 「債権差押通知書到達日現在の残高 〜 円 差押額 〜 円」の固定文言
    private Dictionary<string, string> GenerateSeizureText(
        string institutionName, string branchName, string branchNumber,
        string accountType, string accountNumber)
    {
        // 全角正規化した金融機関名（マッピング照合用・差押文言表示用）
        string normalizedInst = BusinessLogic.ToFullWidth((institutionName ?? "").Trim());

        // 口座種別の差押文言用表示値を補完
        // 金融機関固有ルール → global の順で検索、どちらにもなければ元の値
        string displayAccountType = accountType ?? "";
        if (accountTypeMapping != null)
        {
            string matched = FindAccountTypeMatch(normalizedInst, displayAccountType);
            if (matched != null)
                displayAccountType = matched;
        }

        // 振込手数料の文言を取得
        string feeText = "";
        if (feeMapping != null)
        {
            foreach (var kv in feeMapping)
            {
                if (BusinessLogic.ToFullWidth(kv.Key) == normalizedInst)
                {
                    feeText = kv.Value ?? "";
                    break;
                }
            }
        }

        // 手数料ありの場合、Line1に「と手数料の合計額」を挿入
        string feeClause = !string.IsNullOrEmpty(feeText) ? "と手数料の合計額" : "";

        // ゆうちょ銀行かどうかの判定
        bool isYucho = (normalizedInst == BusinessLogic.ToFullWidth("ゆうちょ銀行"));
        string fullWidthAccountNum = BusinessLogic.ToFullWidth(accountNumber.Trim());

        // Line1 の組み立て
        string line1;
        if (isYucho)
        {
            // ゆうちょ：記号番号形式
            string fullWidthBranchNum = BusinessLogic.ToFullWidth(branchNumber.Trim());
            line1 = "\u3000上記滞納者が、債務者であるゆうちょ銀行（" + branchName + "扱）に対して有する"
                + displayAccountType + "（記号番号：" + fullWidthBranchNum + "－" + fullWidthAccountNum
                + "）の払戻請求権及びこれに対する債権差押通知書到達日までの約定利息の支払請求権。ただし、滞納金額"
                + feeClause + "に充るまでとし、残高が３，０００円未満の場合は差押しない。";
        }
        else
        {
            // ゆうちょ以外：口座番号形式
            line1 = "\u3000上記滞納者が、債務者である" + normalizedInst + "（" + branchName.Trim() + "扱）に対して有する"
                + displayAccountType + "（口座番号：" + fullWidthAccountNum
                + "）の払戻請求権及びこれに対する債権差押通知書到達日までの約定利息の支払請求権。ただし、滞納金額"
                + feeClause + "に充るまでとし、残高が３，０００円未満の場合は差押しない。";
        }

        return new Dictionary<string, string>
        {
            { "Line1", line1 },
            { "Line2", feeText },
            { "Line3", "債権差押通知書到達日現在の残高\u3000\u3000\u3000\u3000\u3000\u3000円\u3000差押額\u3000\u3000\u3000\u3000\u3000\u3000円" }
        };
    }

    // 口座種別マッピングから一致するルールを検索する
    // 検索順序: ① 金融機関固有ルール → ② global ルール → ③ 該当なし(null)
    private string FindAccountTypeMatch(string normalizedInst, string accountType)
    {
        // ① 金融機関固有ルールを先に探す
        foreach (var kv in accountTypeMapping)
        {
            if (kv.Key == "global") continue;
            if (BusinessLogic.ToFullWidth(kv.Key) == normalizedInst)
            {
                string matched = MatchAccountTypeRules(kv.Value, accountType);
                if (matched != null) return matched;
                break;  // 金融機関が見つかったがルール不一致の場合は global へ
            }
        }

        // ② global ルールで引く
        Dictionary<string, string> globalRules;
        if (accountTypeMapping.TryGetValue("global", out globalRules))
        {
            string matched = MatchAccountTypeRules(globalRules, accountType);
            if (matched != null) return matched;
        }

        return null;
    }

    // ルール辞書から accountType にマッチする値を探す
    // キー末尾が「*」なら前方一致、それ以外は完全一致
    private string MatchAccountTypeRules(Dictionary<string, string> rules, string target)
    {
        foreach (var kv in rules)
        {
            bool isWildcard = kv.Key.EndsWith("*");
            string pattern = isWildcard ? kv.Key.TrimEnd('*') : kv.Key;

            if (isWildcard ? target.StartsWith(pattern) : target == kv.Key)
                return kv.Value;
        }
        return null;
    }

    // ==============================================================
    // CSV 操作
    // ==============================================================

    // 差押実績ルックアップテーブルを構築する
    // 既存CSVから「宛名番号_口座番号 → 文書番号」のハッシュテーブルを作成
    // 口座テーブルの各行に差押実績マーク（★）を表示するために使用
    private void BuildSeizureHistory()
    {
        seizureHistory.Clear();
        if (activeProfile.OutputFolder == null) return;

        string csvPath = System.IO.Path.Combine(activeProfile.OutputFolder, CSV_FILENAME);
        if (!File.Exists(csvPath)) return;

        try
        {
            var lines = File.ReadAllLines(csvPath, Encoding.UTF8);
            for (int i = 1; i < lines.Length; i++)  // ヘッダー行をスキップ
            {
                var fields = ParseCsvLine(lines[i]);
                // f[1]=宛名番号, f[4]=執行日(7桁和暦), f[11]=口座番号, f[15]=文書番号
                if (fields.Length >= 16)
                {
                    string key = fields[1].Trim() + "_" + fields[11].Trim();
                    string docNum = fields[15].Trim();
                    // 7桁和暦 → "yyyy年M月d日" 表示形式に変換
                    string execDateDisplay = "";
                    var dt = BusinessLogic.WarekiToDate(fields[4].Trim(), eraMapping);
                    if (dt.HasValue)
                        execDateDisplay = dt.Value.Year + "年" + dt.Value.Month + "月" + dt.Value.Day + "日";
                    seizureHistory[key] = new string[] { docNum, execDateDisplay };  // 後勝ち＝直近
                }
            }
        }
        catch { /* CSVがロック中などの場合は実績なしとして続行 */ }
    }

    // CSV行をRFC 4180準拠でパースし、フィールドの配列として返す
    // ダブルクォートで囲まれたフィールド内のエスケープ（""→"）にも対応
    private string[] ParseCsvLine(string line)
    {
        var fields = new List<string>();
        int pos = 0;

        while (pos < line.Length)
        {
            if (line[pos] == '"')
            {
                // ダブルクォートで囲まれたフィールド
                pos++;  // 開始クォートをスキップ
                var sb = new StringBuilder();
                while (pos < line.Length)
                {
                    if (line[pos] == '"')
                    {
                        if (pos + 1 < line.Length && line[pos + 1] == '"')
                        {
                            // エスケープされたダブルクォート（"" → "）
                            sb.Append('"');
                            pos += 2;
                        }
                        else
                        {
                            // フィールド終了の閉じクォート
                            pos++;
                            break;
                        }
                    }
                    else
                    {
                        sb.Append(line[pos]);
                        pos++;
                    }
                }
                fields.Add(sb.ToString());
                // 閉じクォートの後のカンマをスキップ
                if (pos < line.Length && line[pos] == ',')
                    pos++;
            }
            else
            {
                // 通常フィールド（クォートなし）
                int nextComma = line.IndexOf(',', pos);
                if (nextComma < 0)
                {
                    fields.Add(line.Substring(pos));
                    break;
                }
                fields.Add(line.Substring(pos, nextComma - pos));
                pos = nextComma + 1;
            }
        }

        return fields.ToArray();
    }

    // CSV行を追記する（排他ロック付き、リトライ最大5回）
    // ファイルが存在しない場合はヘッダー行を先に書き込む
    // エンコーディング: BOM付きUTF-8（WinActor読み込み対応）
    private bool WriteCsvLine(string csvPath, string csvLine)
    {
        var dir = System.IO.Path.GetDirectoryName(csvPath);
        if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        bool isNewFile = !File.Exists(csvPath);

        for (int retry = 1; retry <= CSV_WRITE_MAX_RETRY; retry++)
        {
            try
            {
                using (var fileStream = new FileStream(
                    csvPath, FileMode.Append, FileAccess.Write, FileShare.None))
                using (var writer = new StreamWriter(fileStream, new UTF8Encoding(true)))
                {
                    if (isNewFile)
                        writer.WriteLine(CSV_HEADER);
                    writer.WriteLine(csvLine);
                    writer.Flush();
                }
                return true;
            }
            catch
            {
                // リトライ（最後の試行では待たない）
                if (retry < CSV_WRITE_MAX_RETRY)
                    System.Threading.Thread.Sleep(CSV_WRITE_RETRY_INTERVAL_MS);
            }
        }
        return false;
    }

    // ==============================================================
    // 文書番号管理
    // ==============================================================

    // 文書番号を排他ロックで採番する
    // ファイルを排他ロックしたまま「読む→採番→書き戻す」をアトミックに実行し、
    // 複数人の同時使用による重複採番を防止する
    private bool AllocateDocNumber(out string docNum)
    {
        docNum = "";
        string counterPath = System.IO.Path.Combine(exeDir, "document_number_counter.json");

        for (int retry = 1; retry <= CSV_WRITE_MAX_RETRY; retry++)
        {
            try
            {
                using (var fileStream = new FileStream(
                    counterPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // 読み込み（BOMがあれば除去）
                    var bytes = new byte[fileStream.Length];
                    fileStream.Read(bytes, 0, bytes.Length);
                    var json = Encoding.UTF8.GetString(bytes).TrimStart('\uFEFF');

                    // 現在の番号を取得
                    int nextNumber = JsonHelper.GetInt(json, "nextNumber", 1);
                    docNum = nextNumber.ToString();
                    lastDocNumber = docNum;

                    // +1 して書き戻す
                    fileStream.Seek(0, SeekOrigin.Begin);
                    fileStream.SetLength(0);
                    var newBytes = Encoding.UTF8.GetBytes(
                        "{\n    \"nextNumber\":  " + (nextNumber + 1) + "\n}");
                    fileStream.Write(newBytes, 0, newBytes.Length);
                }
                return true;
            }
            catch
            {
                if (retry < CSV_WRITE_MAX_RETRY)
                    System.Threading.Thread.Sleep(CSV_WRITE_RETRY_INTERVAL_MS);
            }
        }
        return false;
    }

    // 文書番号カウンターをロールバックする（CSV書き込み失敗時に採番を元に戻す）
    private void RollbackDocNumber()
    {
        if (string.IsNullOrEmpty(lastDocNumber)) return;
        string counterPath = System.IO.Path.Combine(exeDir, "document_number_counter.json");

        for (int retry = 1; retry <= CSV_WRITE_MAX_RETRY; retry++)
        {
            try
            {
                using (var fileStream = new FileStream(
                    counterPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    fileStream.Seek(0, SeekOrigin.Begin);
                    fileStream.SetLength(0);
                    var bytes = Encoding.UTF8.GetBytes(
                        "{\n    \"nextNumber\":  " + lastDocNumber + "\n}");
                    fileStream.Write(bytes, 0, bytes.Length);
                }
                return;
            }
            catch
            {
                if (retry < CSV_WRITE_MAX_RETRY)
                    System.Threading.Thread.Sleep(CSV_WRITE_RETRY_INTERVAL_MS);
            }
        }
    }

    // ==============================================================
    // 印刷用ファイル保存
    // ==============================================================

    // 印刷用ファイルを保存する
    // 処理フロー:
    //   1. 元ファイルを読み書き可能で開き直し、SaveAs で複製
    //   2. 選択シート以外を削除（VeryHidden シートは残す）
    //   3. 複数表紙時: 選択表紙ブロック以外の行を削除
    //   4. 明細ページ削除（detailAccountNumberCell 設定時）: 選択口座以外の明細を削除
    //   5. 全ての削除は行番号の大きい順に実行（行ずれ防止）
    //   6. 表示位置を A1・スクロール先頭にリセット
    private void SavePrintFile(string sourcePath, string destPath,
        string selectedAccountNum, string selectedAccountType, int selectedCoverIndex)
    {
        // MAX_PATH チェック
        if (destPath.Length > MAX_PATH) return;

        var destDir = System.IO.Path.GetDirectoryName(destPath);
        if (!Directory.Exists(destDir))
            Directory.CreateDirectory(destDir);

        dynamic printWorkbook = null;
        try
        {
            // 1. 元ファイルを読み書き可能で開き、SaveAs で複製
            excel.Visible = false;
            printWorkbook = excel.Workbooks.Open(sourcePath, 0, false);  // 読み書き可能

            // SaveAs 前にブック保護を解除
            try { if ((bool)printWorkbook.ProtectStructure) printWorkbook.Unprotect(); }
            catch { /* 保護解除失敗は無視 */ }

            // SaveAs（xlOpenXMLWorkbookMacroEnabled = 52）
            excel.DisplayAlerts = false;
            printWorkbook.SaveAs(destPath, 52);
            excel.DisplayAlerts = true;

            // SaveAs 後にも再度ブック保護を解除
            try { if ((bool)printWorkbook.ProtectStructure) printWorkbook.Unprotect(); }
            catch { /* 保護解除失敗は無視 */ }

            // 2. 選択シート以外を削除（VeryHidden シートは残す）
            excel.DisplayAlerts = false;
            for (int sheetIdx = (int)printWorkbook.Worksheets.Count; sheetIdx >= 1; sheetIdx--)
            {
                dynamic ws = printWorkbook.Worksheets[sheetIdx];
                try
                {
                    // VeryHidden（Visible=2）以外で、選択シート名と異なるシートを削除
                    // xlSheetVeryHidden = 2（非表示かつマクロからのみ操作可能なシート）
                    if ((int)ws.Visible != 2 && (string)ws.Name != selectedSheetName)
                        ws.Delete();
                }
                catch { /* シート削除失敗は無視 */ }
            }
            excel.DisplayAlerts = true;

            // 3 & 4. 削除対象範囲を集約してから後方からまとめて削除
            dynamic printSheet = printWorkbook.Worksheets[selectedSheetName];
            try
            {
                var deleteRanges = new List<int[]>();

                // シートの実際の最終行を再取得（SaveAs後に変わる可能性）
                int printLastRow;
                try
                {
                    printLastRow = (int)printSheet.UsedRange.Row +
                                   (int)printSheet.UsedRange.Rows.Count - 1;
                }
                catch { printLastRow = lastUsedRow; }

                // 選択された表紙ブロックの保持範囲を算出
                int keepStart, keepEnd;
                if (coverOffsets.Count > 1)
                {
                    keepStart = coverOffsets[selectedCoverIndex] + 1;
                    keepEnd = (selectedCoverIndex + 1 < coverOffsets.Count)
                        ? coverOffsets[selectedCoverIndex + 1]
                        : printLastRow;

                    // 後方ブロック範囲を削除対象に追加
                    if (keepEnd < printLastRow)
                        deleteRanges.Add(new[] { keepEnd + 1, printLastRow });
                    // 前方ブロック範囲を削除対象に追加
                    if (keepStart > 1)
                        deleteRanges.Add(new[] { 1, keepStart - 1 });
                }
                else
                {
                    // 表紙が1つしかない場合はシート全体を保持範囲とする
                    keepStart = 1;
                    keepEnd = printLastRow;
                }

                // 4. 明細ページ削除（detailAccountNumberCell が設定されている場合のみ）
                // detailAccountNumberCell / detailAccountTypeCell は
                // 「stopValues セルを A1 起点とした相対座標」として解釈
                if (!string.IsNullOrEmpty(activeProfile.DetailAccountNumberCell) && stopRows.Count >= 2)
                {
                    var detailMatch = System.Text.RegularExpressions.Regex.Match(
                        activeProfile.DetailAccountNumberCell, @"^([A-Za-z]{1,3})(\d+)$");
                    if (detailMatch.Success)
                    {
                        string detailCol = detailMatch.Groups[1].Value;
                        int detailRowOffset = int.Parse(detailMatch.Groups[2].Value) - 1;

                        // 口座種別セルの設定（任意）
                        string typeCol = null;
                        int typeRowOffset = 0;
                        if (!string.IsNullOrEmpty(activeProfile.DetailAccountTypeCell))
                        {
                            var typeMatch = System.Text.RegularExpressions.Regex.Match(
                                activeProfile.DetailAccountTypeCell, @"^([A-Za-z]{1,3})(\d+)$");
                            if (typeMatch.Success)
                            {
                                typeCol = typeMatch.Groups[1].Value;
                                typeRowOffset = int.Parse(typeMatch.Groups[2].Value) - 1;
                            }
                        }

                        // 保持範囲内の stopRows のみを対象
                        var relevantStopRows = stopRows
                            .Where(x => x >= keepStart && x <= keepEnd).ToList();

                        if (relevantStopRows.Count >= 2)
                        {
                            // ページサイズを算出（2つのstopRows間の行数）
                            int pageRows = relevantStopRows[1] - relevantStopRows[0];

                            // stopValues セルがページ内の何行目にあるかを剰余から動的に算出
                            int modResult = relevantStopRows[0] % pageRows;
                            int stopValueRowInPage = (modResult == 0) ? pageRows : modResult;
                            int rowsBeforeStopValue = stopValueRowInPage - 1;

                            for (int i = 0; i < relevantStopRows.Count; i++)
                            {
                                int pageStart = relevantStopRows[i] - rowsBeforeStopValue;
                                int pageEnd = pageStart + pageRows - 1;

                                // 口座番号セルの読み取り
                                string accountNumValue = "";
                                try
                                {
                                    accountNumValue = Convert.ToString(
                                        printSheet.Range[detailCol + (relevantStopRows[i] + detailRowOffset)].Value2 ?? "");
                                }
                                catch { /* セル読み取り失敗時はデフォルト値を使用 */ }
                                bool accountNumMatches = (accountNumValue.Trim() == selectedAccountNum);

                                // 口座種別セルの読み取り（設定されている場合のみ）
                                bool accountTypeMatches = true;
                                if (typeCol != null)
                                {
                                    string typeValue = "";
                                    try
                                    {
                                        typeValue = Convert.ToString(
                                            printSheet.Range[typeCol + (relevantStopRows[i] + typeRowOffset)].Value2 ?? "");
                                    }
                                    catch { /* セル読み取り失敗時はデフォルト値を使用 */ }
                                    accountTypeMatches = (typeValue.Trim() == selectedAccountType);
                                }

                                // 口座番号と口座種別の両方が一致した場合のみ残す
                                if (!(accountNumMatches && accountTypeMatches))
                                    deleteRanges.Add(new[] { pageStart, pageEnd });
                            }
                        }
                    }
                }

                // 行が大きい順にソートして削除（行ずれ防止）
                deleteRanges.Sort((a, b) => b[0].CompareTo(a[0]));
                excel.DisplayAlerts = false;
                foreach (var range in deleteRanges)
                {
                    try
                    {
                        printSheet.Rows[range[0] + ":" + range[1]].Delete();
                    }
                    catch { /* 行削除失敗は無視 */ }
                }
                excel.DisplayAlerts = true;

                // 表示位置を A1・スクロール先頭にリセット
                try
                {
                    printSheet.Activate();
                    printWorkbook.Application.ActiveWindow.ScrollRow = 1;
                    printWorkbook.Application.ActiveWindow.ScrollColumn = 1;
                    printSheet.Range["A1"].Select();
                }
                catch { /* 表示位置リセット失敗は無視 */ }
            }
            finally
            {
                // printSheet の ReleaseComObject は使用しない（dynamic環境のCOM安定性のため）
            }

            printWorkbook.Save();
        }
        catch { /* 保存失敗は警告のみ（CSVには既に追記済み） */ }
        finally
        {
            if (printWorkbook != null)
                try { printWorkbook.Close(false); } catch { }
        }
    }

    // ==============================================================
    // Excel COM クリーンアップ
    // ==============================================================

    // COMオブジェクトを安全に1回解放する（参照カウントを1つ減らす）
    // C# の dynamic late binding では通常1回の ReleaseComObject で十分。
    // PowerShell版のようなループ解放は、PSランタイムが暗黙的に参照カウントを
    // 増やす問題への対処であり、C#では不要。
    private void ReleaseComSafe(object obj)
    {
        if (obj == null) return;
        try { Marshal.ReleaseComObject(obj); }
        catch { /* COM分離済みの場合は無視 */ }
    }

    // アプリ終了時にExcelプロセスを完全に解放する
    // excel インスタンスはアプリ生存期間中保持し、ファイルごとの起動/終了は行わない設計
    private void CleanupExcel()
    {
        if (excel == null) return;
        try
        {
            // 開いたままのワークブックがあれば全て閉じる
            try
            {
                while ((int)excel.Workbooks.Count > 0)
                {
                    dynamic wb = excel.Workbooks[1];
                    try { wb.Close(false); } catch { }
                }
            }
            catch { /* Workbooks アクセス失敗は無視 */ }

            // 元の設定を復元してから終了
            try { excel.ScreenUpdating = true; } catch { }
            try { excel.DisplayAlerts = true; } catch { }

            excel.Quit();
            ReleaseComSafe(excel);
        }
        catch { /* Quit 失敗は無視 */ }
        finally
        {
            excel = null;
            // GC を2回呼び出して COM の Release を確実に完了させる
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    // ==============================================================
    // XAML 定義
    // ==============================================================

    private Window BuildWindow()
    {
        string xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    Title='預金差押予定一覧 作成ツール' Width='1000' Height='700' MinWidth='900' MinHeight='520'
    WindowStartupLocation='CenterScreen' Background='#F9F9F9' FontFamily='Meiryo UI'
    UseLayoutRounding='True' SnapsToDevicePixels='True'>
<Window.Resources>
    <!-- TextBox 角丸化（全TextBoxに自動適用） -->
    <Style TargetType='TextBox'>
        <Setter Property='Foreground' Value='#333'/>
        <Setter Property='Template'><Setter.Value>
            <ControlTemplate TargetType='TextBox'>
                <Border Background='{TemplateBinding Background}'
                        BorderBrush='{TemplateBinding BorderBrush}'
                        BorderThickness='{TemplateBinding BorderThickness}'
                        CornerRadius='4' Padding='{TemplateBinding Padding}'
                        SnapsToDevicePixels='True'>
                    <ScrollViewer x:Name='PART_ContentHost' Focusable='False'/></Border>
            </ControlTemplate>
        </Setter.Value></Setter>
    </Style>
    <!-- ComboBox フラットデザイン（TextBoxと統一した白背景＋角丸） -->
    <Style TargetType='ComboBox'>
        <Setter Property='Foreground' Value='#333'/>
        <Setter Property='Background' Value='White'/>
        <Setter Property='BorderBrush' Value='#D0D0D0'/>
        <Setter Property='BorderThickness' Value='1'/>
        <Setter Property='Padding' Value='6,5'/>
        <Setter Property='Cursor' Value='Hand'/>
        <Setter Property='Template'><Setter.Value>
            <ControlTemplate TargetType='ComboBox'>
                <Grid x:Name='comboRoot'>
                    <!-- 透明ToggleButton: ドロップダウンの開閉を制御 -->
                    <ToggleButton BorderThickness='0' Background='Transparent' Focusable='False' ClickMode='Press'
                        IsChecked='{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}'>
                        <ToggleButton.Template><ControlTemplate TargetType='ToggleButton'>
                            <Border Background='Transparent'/></ControlTemplate></ToggleButton.Template>
                    </ToggleButton>
                    <!-- 表示用Border: 白背景＋角丸＋ドロップダウン矢印 -->
                    <Border x:Name='bd' Background='{TemplateBinding Background}'
                            BorderBrush='{TemplateBinding BorderBrush}'
                            BorderThickness='{TemplateBinding BorderThickness}'
                            CornerRadius='4' IsHitTestVisible='False'>
                        <Grid Margin='{TemplateBinding Padding}'>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width='*'/><ColumnDefinition Width='20'/></Grid.ColumnDefinitions>
                            <ContentPresenter Content='{TemplateBinding SelectionBoxItem}'
                                ContentTemplate='{TemplateBinding SelectionBoxItemTemplate}'
                                HorizontalAlignment='Left' VerticalAlignment='Center'/>
                            <Path Grid.Column='1' Data='M0,0 L4,4 8,0' Stroke='#888'
                                  StrokeThickness='1.5' VerticalAlignment='Center' HorizontalAlignment='Center'/>
                        </Grid>
                    </Border>
                    <!-- ドロップダウンPopup -->
                    <Popup x:Name='PART_Popup' AllowsTransparency='True' Placement='Bottom'
                           IsOpen='{TemplateBinding IsDropDownOpen}'>
                        <Border Background='White' BorderBrush='#D0D0D0' BorderThickness='1'
                                CornerRadius='4' Margin='0,2,0,0' Padding='0,4'
                                MinWidth='{Binding ActualWidth, ElementName=comboRoot}'>
                            <Border.Effect><DropShadowEffect BlurRadius='8' ShadowDepth='2' Opacity='0.12'/></Border.Effect>
                            <ScrollViewer MaxHeight='200'>
                                <StackPanel IsItemsHost='True'/></ScrollViewer>
                        </Border>
                    </Popup>
                </Grid>
                <ControlTemplate.Triggers>
                    <Trigger Property='IsMouseOver' Value='True'>
                        <Setter TargetName='bd' Property='BorderBrush' Value='#00897B'/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate>
        </Setter.Value></Setter>
    </Style>
    <!-- アクセントボタン（青背景） -->
    <Style x:Key='AB' TargetType='Button'>
        <Setter Property='Background' Value='#00897B'/><Setter Property='Foreground' Value='White'/>
        <Setter Property='FontSize' Value='12'/><Setter Property='Padding' Value='16,8'/>
        <Setter Property='Cursor' Value='Hand'/><Setter Property='BorderThickness' Value='0'/>
        <Setter Property='Template'><Setter.Value><ControlTemplate TargetType='Button'>
            <Border x:Name='bd' Background='{TemplateBinding Background}' CornerRadius='4' Padding='{TemplateBinding Padding}'>
                <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='bd' Property='Background' Value='#00796B'/></Trigger>
                <Trigger Property='IsEnabled' Value='False'><Setter TargetName='bd' Property='Background' Value='#CCC'/><Setter Property='Foreground' Value='#999'/></Trigger>
            </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style>
    <!-- グレーボタン -->
    <Style x:Key='GB' TargetType='Button'>
        <Setter Property='Background' Value='White'/><Setter Property='Foreground' Value='#555'/>
        <Setter Property='FontSize' Value='12'/><Setter Property='Padding' Value='16,8'/>
        <Setter Property='Cursor' Value='Hand'/><Setter Property='BorderBrush' Value='#D0D0D0'/><Setter Property='BorderThickness' Value='1'/>
        <Setter Property='Template'><Setter.Value><ControlTemplate TargetType='Button'>
            <Border x:Name='bd' Background='{TemplateBinding Background}' BorderBrush='{TemplateBinding BorderBrush}'
                BorderThickness='{TemplateBinding BorderThickness}' CornerRadius='4' Padding='{TemplateBinding Padding}'>
                <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='bd' Property='Background' Value='#EEF6F5'/></Trigger>
            </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style>
    <!-- CheckBox 青スタイル（チェック時に青背景＋白チェックマーク） -->
    <Style TargetType='CheckBox'>
        <Setter Property='Foreground' Value='#333'/>
        <Setter Property='Cursor' Value='Hand'/>
        <Setter Property='Template'><Setter.Value>
            <ControlTemplate TargetType='CheckBox'>
                <StackPanel Orientation='Horizontal'>
                    <Border x:Name='cbBox' Width='16' Height='16' CornerRadius='3'
                            Background='White' BorderBrush='#C8C8C8' BorderThickness='1'
                            VerticalAlignment='Center' Margin='0,0,6,0'>
                        <Path x:Name='cbCheck' Data='M2.5,7 L5.5,10 L11.5,3.5' Stroke='White'
                              StrokeThickness='2' Visibility='Collapsed'/></Border>
                    <ContentPresenter VerticalAlignment='Center'/></StackPanel>
                <ControlTemplate.Triggers>
                    <Trigger Property='IsChecked' Value='True'>
                        <Setter TargetName='cbBox' Property='Background' Value='#00897B'/>
                        <Setter TargetName='cbBox' Property='BorderBrush' Value='#00897B'/>
                        <Setter TargetName='cbCheck' Property='Visibility' Value='Visible'/></Trigger>
                    <Trigger Property='IsMouseOver' Value='True'>
                        <Setter TargetName='cbBox' Property='BorderBrush' Value='#00897B'/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate>
        </Setter.Value></Setter>
    </Style>
    <!-- GridViewColumnHeader: フラットデザイン（中央揃え） -->
    <Style TargetType='GridViewColumnHeader'>
        <Setter Property='Background' Value='#F5F7FA'/>
        <Setter Property='Foreground' Value='#777'/>
        <Setter Property='FontSize' Value='11'/>
        <Setter Property='Padding' Value='8,8'/>
        <Setter Property='HorizontalContentAlignment' Value='Center'/>
        <Setter Property='Template'><Setter.Value>
            <ControlTemplate TargetType='GridViewColumnHeader'>
                <Border Background='{TemplateBinding Background}'
                        BorderBrush='#E8E8E8' BorderThickness='0,0,0,1'
                        Padding='{TemplateBinding Padding}'>
                    <ContentPresenter HorizontalAlignment='{TemplateBinding HorizontalContentAlignment}'
                                      VerticalAlignment='Center'/></Border>
            </ControlTemplate>
        </Setter.Value></Setter>
    </Style>
</Window.Resources>
<DockPanel>
    <Border DockPanel.Dock='Top' Background='#00897B' Padding='18,10'>
        <TextBlock Text='預金差押予定一覧 作成ツール' FontSize='13' FontWeight='Medium' Foreground='White'/></Border>
    <Border DockPanel.Dock='Bottom' Background='#F0F0F0' BorderBrush='#E0E0E0' BorderThickness='0,1,0,0' Padding='18,4'>
        <DockPanel>
                <StackPanel DockPanel.Dock='Right' Orientation='Horizontal'>
                    <TextBlock Text='出力先: ' FontSize='11' Foreground='#666'/>
                    <TextBlock x:Name='StatusRight' FontSize='11' Foreground='#666'/></StackPanel>
                <TextBlock x:Name='StatusLeft' FontSize='11' Foreground='#666'/></DockPanel></Border>
    <Grid>
        <!-- 初期画面 -->
        <Grid x:Name='InitialPanel'>
            <Border Background='White' BorderBrush='#B2CECB' BorderThickness='2' CornerRadius='8'
                    Margin='80,60' VerticalAlignment='Center' HorizontalAlignment='Center' Padding='60,40'>
                <StackPanel HorizontalAlignment='Center'>
                    <TextBlock Text='&#x1F4C2;' FontSize='36' HorizontalAlignment='Center' Margin='0,0,0,12'/>
                    <TextBlock Text='ここにファイルをドラッグ＆ドロップ' FontSize='14' Foreground='#666' HorizontalAlignment='Center' Margin='0,0,0,16'/>
                    <Button x:Name='BtnSelectFile' Style='{StaticResource AB}' HorizontalAlignment='Center'>
                        <TextBlock Text='ファイルを選択' FontSize='13'/></Button>
                </StackPanel></Border></Grid>
        <!-- メインフォーム -->
        <Grid x:Name='MainPanel' Visibility='Collapsed' Margin='18,14,18,12'>
            <Grid.RowDefinitions>
                <RowDefinition Height='Auto'/><RowDefinition Height='Auto'/>
                <RowDefinition Height='*'/><RowDefinition Height='Auto'/></Grid.RowDefinitions>
            <!-- シート選択 -->
            <DockPanel Grid.Row='0' Margin='0,0,0,10'>
                <TextBlock Text='シート:' VerticalAlignment='Center' Foreground='#555' FontSize='11' Margin='0,0,6,0'/>
                <ComboBox x:Name='SheetCombo' MinWidth='180' FontSize='12'/>
                <StackPanel DockPanel.Dock='Right' Orientation='Horizontal' HorizontalAlignment='Right'>
                    <TextBlock Text='&#x1F4C4; ' FontSize='11' Foreground='#00897B' VerticalAlignment='Center'/>
                    <TextBlock x:Name='FileLink' FontSize='11' Foreground='#00897B'
                               Cursor='Hand' TextDecorations='Underline' VerticalAlignment='Center'/>
                    <Button x:Name='BtnReload' Style='{StaticResource GB}' Padding='8,4' Margin='8,0,0,0' FontSize='11'>
                        <TextBlock Text='&#x1F504; 再読み込み'/></Button>
                </StackPanel>
            </DockPanel>
            <!-- 基本情報 -->
            <Border Grid.Row='1' Background='White' BorderBrush='#E0E0E0' BorderThickness='1' CornerRadius='6' Padding='16,14' Margin='0,0,0,10'>
                <StackPanel><TextBlock Text='&#x1F464; 基本情報' FontSize='13' Foreground='#00897B' FontWeight='Medium' Margin='0,0,0,10'/>
                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width='*'/><ColumnDefinition Width='16'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>
                    <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='6'/>
                        <RowDefinition Height='Auto'/><RowDefinition Height='6'/><RowDefinition Height='Auto'/></Grid.RowDefinitions>
                    <!-- 1行目: 宛名番号+氏名 | 金融機関+処分担当 -->
                    <Grid Grid.Row='0' Grid.Column='0'><Grid.ColumnDefinitions><ColumnDefinition Width='120'/><ColumnDefinition Width='16'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>
                        <StackPanel><TextBlock Text='宛名番号' FontSize='11' Foreground='#777' Margin='0,0,0,3'/>
                            <TextBox x:Name='TxtAddressNum' IsReadOnly='True' Background='#F3F3F3' BorderBrush='#E8E8E8' FontFamily='Consolas' FontSize='12' Padding='5,4'/></StackPanel>
                        <StackPanel Grid.Column='2'><TextBlock FontSize='11' Foreground='#777' Margin='0,0,0,3'>氏名 &#x270E;</TextBlock>
                            <TextBox x:Name='TxtName' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/></StackPanel></Grid>
                    <Grid Grid.Row='0' Grid.Column='2'><Grid.ColumnDefinitions><ColumnDefinition Width='*'/><ColumnDefinition Width='16'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>
                        <StackPanel><TextBlock Text='金融機関' FontSize='11' Foreground='#777' Margin='0,0,0,3'/>
                            <TextBox x:Name='TxtInstitution' IsReadOnly='True' Background='#F3F3F3' BorderBrush='#E8E8E8' FontSize='12' Padding='5,4'/></StackPanel>
                        <StackPanel Grid.Column='2'><TextBlock FontSize='11' Foreground='#777' Margin='0,0,0,3'>処分担当 &#x270E;</TextBlock>
                            <TextBox x:Name='TxtStaff' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/></StackPanel></Grid>
                    <!-- 2行目: 住民票住所 | 届出住所+バリデーション -->
                    <StackPanel Grid.Row='2' Grid.Column='0'><TextBlock FontSize='11' Foreground='#777' Margin='0,0,0,3'>住民票住所 &#x270E;</TextBlock>
                        <TextBox x:Name='TxtResidenceAddr' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/></StackPanel>
                    <StackPanel Grid.Row='2' Grid.Column='2'><TextBlock FontSize='11' Foreground='#777' Margin='0,0,0,3'>届出住所 &#x270E;</TextBlock>
                        <TextBox x:Name='TxtDeliveryAddr' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/>
                        <TextBlock x:Name='DeliveryError' Foreground='#D32F2F' FontSize='10' Visibility='Collapsed' Margin='0,1,0,0'/></StackPanel>
                    <!-- 3行目: 執行日+カレンダー | チェックボックス -->
                    <StackPanel Grid.Row='4' Grid.Column='0'><TextBlock Text='執行日' FontSize='11' Foreground='#777' Margin='0,0,0,3'/>
                        <StackPanel Orientation='Horizontal'>
                            <TextBox x:Name='TxtExecDate' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0' FontFamily='Consolas' Width='120'/>
                            <Button x:Name='BtnCalendar' Style='{StaticResource GB}' Padding='6,4' Margin='4,0,0,0'>
                                <TextBlock Text='&#x1F4C5;' FontSize='13'/></Button>
                            <Popup x:Name='CalendarPopup' StaysOpen='False' Placement='Bottom' AllowsTransparency='True'>
                                <Border Background='White' BorderBrush='#D0D0D0' BorderThickness='1'
                                        CornerRadius='6' Padding='8' Margin='0,4,0,0'>
                                    <Border.Effect><DropShadowEffect BlurRadius='12' ShadowDepth='3' Opacity='0.15'/></Border.Effect>
                                    <Calendar x:Name='DateCalendar' SelectionMode='SingleDate'/></Border>
                            </Popup>
                        </StackPanel></StackPanel>
                    <StackPanel Grid.Row='4' Grid.Column='2' VerticalAlignment='Top' Margin='0,18,0,0'>
                        <CheckBox x:Name='ChkDeliveryOutput' Content='届出住所を差押通知書に出力する' FontSize='12'/></StackPanel>
                </Grid></StackPanel></Border>
            <!-- 口座選択 -->
            <Border Grid.Row='2' Background='White' BorderBrush='#E0E0E0' BorderThickness='1' CornerRadius='6' Padding='16,14' Margin='0,0,0,10'>
                <DockPanel><TextBlock DockPanel.Dock='Top' Text='&#x2261; 口座選択' FontSize='13' Foreground='#00897B' FontWeight='Medium' Margin='0,0,0,8'/>
                    <Border BorderBrush='#E0E0E0' BorderThickness='1' CornerRadius='4' ClipToBounds='True'>
                    <ListView x:Name='AccountList' BorderThickness='0' Background='White' FontSize='12' SelectionMode='Single'>
                        <ListView.ItemContainerStyle>
                            <Style TargetType='ListViewItem'>
                                <Setter Property='ToolTip' Value='{Binding SeizureTooltip}'/>
                                <Setter Property='Cursor' Value='Hand'/>
                                <Setter Property='Foreground' Value='#333'/>
                                <Setter Property='HorizontalContentAlignment' Value='Stretch'/>
                                <Setter Property='Padding' Value='0'/>
                                <Setter Property='Margin' Value='0'/>
                                <Setter Property='Template'><Setter.Value>
                                    <ControlTemplate TargetType='ListViewItem'>
                                        <Grid>
                                            <Border x:Name='rowBd' Background='White'
                                                    BorderBrush='#F0F0F0' BorderThickness='0,0,0,1'>
                                                <GridViewRowPresenter VerticalAlignment='Center'
                                                    Margin='0,7,0,7'/>
                                            </Border>
                                            <Border x:Name='accent' HorizontalAlignment='Left'
                                                    Width='3' Background='Transparent'/>
                                        </Grid>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property='IsMouseOver' Value='True'>
                                                <Setter TargetName='rowBd' Property='Background' Value='#F1F9F8'/></Trigger>
                                            <Trigger Property='IsSelected' Value='True'>
                                                <Setter TargetName='accent' Property='Background' Value='#00897B'/>
                                                <Setter TargetName='rowBd' Property='Background' Value='#E0F2F1'/></Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Setter.Value></Setter>
                            </Style>
                        </ListView.ItemContainerStyle>
                        <ListView.View><GridView>
                            <GridViewColumn Header='支店名' Width='250'>
                                <GridViewColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text='{Binding BranchName}' HorizontalAlignment='Center'/></DataTemplate></GridViewColumn.CellTemplate></GridViewColumn>
                            <GridViewColumn Header='支店番号' Width='95'>
                                <GridViewColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text='{Binding BranchNumber}' HorizontalAlignment='Center'/></DataTemplate></GridViewColumn.CellTemplate></GridViewColumn>
                            <GridViewColumn Header='口座種別' Width='110'>
                                <GridViewColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text='{Binding AccountType}' HorizontalAlignment='Center'/></DataTemplate></GridViewColumn.CellTemplate></GridViewColumn>
                            <GridViewColumn Header='最終取引日(満期日)' Width='150'>
                                <GridViewColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text='{Binding LastTransaction}' HorizontalAlignment='Center'/></DataTemplate></GridViewColumn.CellTemplate></GridViewColumn>
                            <GridViewColumn Header='口座番号' Width='120'>
                                <GridViewColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text='{Binding AccountNumDisplay}' HorizontalAlignment='Center'/></DataTemplate></GridViewColumn.CellTemplate></GridViewColumn>
                            <GridViewColumn Header='残高' Width='145'>
                                <GridViewColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text='{Binding Balance}' HorizontalAlignment='Center'/></DataTemplate></GridViewColumn.CellTemplate></GridViewColumn>
                        </GridView></ListView.View></ListView></Border></DockPanel></Border>
            <!-- アクションボタン -->
            <DockPanel Grid.Row='3'>
                <Button x:Name='BtnLoadFile' DockPanel.Dock='Left' Style='{StaticResource GB}'><TextBlock Text='ファイルを読み込む'/></Button>
                <StackPanel DockPanel.Dock='Right' Orientation='Horizontal' HorizontalAlignment='Right'>
                    <TextBlock x:Name='GuideText' VerticalAlignment='Center' FontSize='11' Foreground='#D32F2F' Margin='0,0,12,0'/>
                    <Button x:Name='BtnSkip' Style='{StaticResource GB}' Margin='0,0,8,0'><TextBlock Text='&#x2192; スキップ'/></Button>
                    <Button x:Name='BtnAdd' Style='{StaticResource AB}' IsEnabled='False'><TextBlock Text='&#xFF0B; 一覧に追加'/></Button>
                </StackPanel></DockPanel>
        </Grid>
        <!-- オーバーレイ -->
        <Grid x:Name='OverlayPanel' Visibility='Collapsed' Background='#CCFFFFFF'>
            <Grid x:Name='LoadingOverlay' Visibility='Collapsed' HorizontalAlignment='Center' VerticalAlignment='Center'>
                <Path x:Name='SpinnerPath' Data='M 20,2 A 18,18 0 1 1 2,20'
                      Stroke='#00897B' StrokeThickness='3'
                      StrokeStartLineCap='Round' StrokeEndLineCap='Round'
                      Width='40' Height='40' Stretch='None'
                      RenderTransformOrigin='0.5,0.5'>
                    <Path.RenderTransform><RotateTransform/></Path.RenderTransform>
                </Path></Grid>
            <Grid x:Name='ResultOverlay' Visibility='Collapsed' HorizontalAlignment='Center' VerticalAlignment='Center'>
                <Border Background='White' CornerRadius='10' Padding='48,36' MinWidth='350'>
                    <Border.Effect><DropShadowEffect BlurRadius='16' ShadowDepth='4' Opacity='0.12'/></Border.Effect>
                    <StackPanel HorizontalAlignment='Center'>
                        <Border x:Name='ResultIconBg' Width='64' Height='64' CornerRadius='32' Background='#E8F5ED' HorizontalAlignment='Center' Margin='0,0,0,20'>
                            <TextBlock x:Name='ResultIcon' Text='&#x2713;' FontSize='32' HorizontalAlignment='Center' VerticalAlignment='Center' Foreground='#107C41'/></Border>
                        <TextBlock x:Name='ResultTitle' Text='' FontSize='17' FontWeight='Medium' HorizontalAlignment='Center' Margin='0,0,0,8'/>
                        <TextBlock x:Name='ResultDetail' FontSize='12' Foreground='#999' HorizontalAlignment='Center' Margin='0,0,0,0'/>
                        <Button x:Name='ResultButton' Style='{StaticResource AB}' HorizontalAlignment='Center' Padding='24,10' Margin='0,24,0,0'>
                            <TextBlock Text='次のファイルへ' FontSize='13'/></Button>
                        <TextBlock x:Name='ResultSub' FontSize='11' Foreground='#999' HorizontalAlignment='Center' Margin='0,12,0,0'/>
                    </StackPanel></Border></Grid>
        </Grid>
    </Grid>
</DockPanel></Window>";
        using (var reader = XmlReader.Create(new StringReader(xaml))) { return (Window)XamlReader.Load(reader); }
    }
}

// ==============================================================
// RelayCommand
// ==============================================================

public class RelayCommand : ICommand
{
    private Action<object> execute;
    public RelayCommand(Action<object> execute) { this.execute = execute; }
    public event EventHandler CanExecuteChanged { add {} remove {} }
    public bool CanExecute(object p) { return true; }
    public void Execute(object p) { execute(p); }
}