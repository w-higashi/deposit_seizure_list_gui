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
using System.Windows.Shapes;
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
    public string FilterValue { get; set; }              // シートフィルタ値（部分一致）
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
        get { return HasSeizureHistory ? "差押実績あり（文書番号: " + SeizureDocNumber + "）" : null; }
    }
    public double BalanceValue { get; set; }             // 残高（数値、ソート用）
    public string Balance { get; set; }                  // 残高（表示用、通貨形式）
    public int CoverIndex { get; set; }                  // 所属する表紙ブロックのインデックス
    public bool HasSeizureHistory { get; set; }          // 差押実績の有無
    public string SeizureDocNumber { get; set; }         // 差押実績の文書番号（あれば）
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
    public string DocNumber { get; set; }                // 採番された文書番号（追加済みの場合）
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
                FilterValue              = JsonHelper.GetString(profileJson, "filterValue"),
                AccountStartRow          = JsonHelper.GetInt(profileJson, "accountStartRow"),
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
    private Dictionary<string, string> seizureHistory = new Dictionary<string, string>(); // 差押実績

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
    private TextBlock resultIcon, resultTitle, resultDocNum, resultDetail, resultSub;
    private TextBox txtAddressNum, txtName, txtInstitution, txtStaff;
    private TextBox txtResidenceAddr, txtDeliveryAddr, txtExecDate;
    private CheckBox chkDeliveryOutput;
    private ListView accountList;
    private Button btnAdd, btnSkip, btnLoadFile, resultButton;
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
        resultTitle = (TextBlock)window.FindName("ResultTitle");
        resultDocNum = (TextBlock)window.FindName("ResultDocNum");
        resultDetail = (TextBlock)window.FindName("ResultDetail");
        resultButton = (Button)window.FindName("ResultButton");
        resultSub = (TextBlock)window.FindName("ResultSub");
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

        // シート切替
        sheetCombo.SelectionChanged += delegate { if (sheetCombo.SelectedItem != null) { selectedSheetName = sheetCombo.SelectedItem.ToString(); ReloadSheetData(); } };

        // 口座選択・執行日・必須フィールドでボタン制御
        accountList.SelectionChanged += delegate { UpdateAddButton(); };
        txtName.TextChanged += delegate { UpdateAddButton(); };
        txtStaff.TextChanged += delegate { UpdateAddButton(); };
        txtResidenceAddr.TextChanged += delegate { UpdateAddButton(); };
        txtExecDate.LostFocus += delegate
        {
            var input = txtExecDate.Text.Trim();
            if (string.IsNullOrEmpty(input)) { processingDate = null; UpdateAddButton(); return; }
            var dt = BusinessLogic.ParseFlexibleDate(input, eraMapping);
            if (dt.HasValue) { processingDate = dt.Value; txtExecDate.Text = dt.Value.ToString("yyyy/MM/dd"); txtExecDate.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D0D0D0")); }
            else { processingDate = null; txtExecDate.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D32F2F")); }
            UpdateAddButton();
        };

        // 届出住所バリデーション
        txtDeliveryAddr.TextChanged += delegate { ValidateDelivery(); };
        chkDeliveryOutput.Checked += delegate { ValidateDelivery(); };
        chkDeliveryOutput.Unchecked += delegate { ValidateDelivery(); };

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

        // 口座テーブルの列幅自動調整
        accountList.SizeChanged += delegate { AdjustAccountColumns(); };
        accountList.Loaded += delegate { AdjustAccountColumns(); };
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

    // 口座テーブルの列幅自動調整（余白をゼロにする）
    private void AdjustAccountColumns()
    {
        var gv = accountList.View as GridView;
        if (gv == null || gv.Columns.Count < 6) return;
        double total = accountList.ActualWidth - SystemParameters.VerticalScrollBarWidth - 10;
        if (total <= 0) return;
        // 支店番号・口座種別・最終取引日・口座番号は固定比率、支店名と残高が可変
        double fixed4 = 80 + 90 + 120 + 100; // 支店番号+口座種別+最終取引日(満期日)+口座番号
        double remaining = total - fixed4;
        if (remaining < 200) remaining = 200;
        double branchW = remaining * 0.6;   // 支店名に60%
        double balanceW = remaining * 0.4;  // 残高に40%
        gv.Columns[0].Width = branchW;
        gv.Columns[1].Width = 80;
        gv.Columns[2].Width = 90;
        gv.Columns[3].Width = 120;
        gv.Columns[4].Width = 100;
        gv.Columns[5].Width = balanceW;
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
        if (fileEntries.Count > 1) statusLeft.Text = (index + 1) + " / " + fileEntries.Count + " 件目";

        // 新しいファイルに切り替わるとき、執行日と届出住所チェックをリセット
        txtExecDate.Text = ""; processingDate = null;
        chkDeliveryOutput.IsChecked = false;
        deliveryError.Visibility = Visibility.Collapsed;

        LoadSingleFile(currentFilePath);
    }

    // ==============================================================
    // ファイル読込（Excel COM・非同期）
    // ==============================================================

    private void LoadSingleFile(string filePath)
    {
        ShowState("main");
        mainPanel.Visibility = Visibility.Visible;
        overlayPanel.Visibility = Visibility.Visible;
        loadingOverlay.Visibility = Visibility.Visible;
        resultOverlay.Visibility = Visibility.Collapsed;

        var worker = new BackgroundWorker();
        worker.DoWork += delegate(object s, DoWorkEventArgs args) { args.Result = ReadExcelFile(filePath); };
        worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs args)
        {
            overlayPanel.Visibility = Visibility.Collapsed;
            if (args.Error != null) { ShowResult("error", "読込失敗", "", args.Error.Message); return; }
            var data = args.Result as Dictionary<string, object>;
            if (data != null && data.ContainsKey("error")) { ShowResult("error", "読込失敗", "", data["error"].ToString()); return; }
            PopulateForm(data);
        };
        worker.RunWorkerAsync();
    }

    // Excel COM でファイルからデータを読み取る
    private Dictionary<string, object> ReadExcelFile(string filePath)
    {
        var result = new Dictionary<string, object>();
        if (excel == null)
        {
            var t = Type.GetTypeFromProgID("Excel.Application");
            if (t == null) return new Dictionary<string, object> { { "error", "Excelがインストールされていません" } };
            excel = Activator.CreateInstance(t);
            excel.Visible = false; excel.DisplayAlerts = false;
            try { excel.AutomationSecurity = 3; } catch { } // msoAutomationSecurityForceDisable
        }
        dynamic wb = null;
        try
        {
            wb = excel.Workbooks.Open(filePath, 0, true); // 読取専用
            // シートフィルタ
            var sheets = new List<string>();
            for (int i = 1; i <= (int)wb.Worksheets.Count; i++)
            {
                dynamic ws = wb.Worksheets[i];
                string name = (string)ws.Name;
                if (!string.IsNullOrEmpty(activeProfile.FilterCell) && !string.IsNullOrEmpty(activeProfile.FilterValue))
                {
                    try { string v = Convert.ToString(ws.Range[activeProfile.FilterCell].Value2 ?? "");
                          if (v.IndexOf(activeProfile.FilterValue, StringComparison.OrdinalIgnoreCase) >= 0) sheets.Add(name); } catch { }
                }
                else { if ((int)ws.Visible == -1) sheets.Add(name); }
            }
            result["sheets"] = sheets; result["filePath"] = filePath; result["fileName"] = System.IO.Path.GetFileName(filePath);
            if (sheets.Count == 0) { result["noSheet"] = true; try { wb.Close(false); } catch {} wb = null; return result; }
            result["selectedSheet"] = sheets[0];
            try { result["sheetData"] = ReadSheetData(wb, sheets[0]); }
            catch (Exception rex) { result["error"] = "シートデータの読取りに失敗: " + rex.Message; }
        }
        catch (Exception ex) { result["error"] = ex.Message; }
        finally { if (wb != null) try { wb.Close(false); } catch { } }
        return result;
    }

    // 指定シートから基本情報と口座テーブルを読み取る
    private Dictionary<string, object> ReadSheetData(dynamic wb, string sheetName)
    {
        var data = new Dictionary<string, object>();
        dynamic ws = wb.Worksheets[sheetName];
        // 表紙ブロック検出
        var offsets = new List<int>();
        try
        {
            lastUsedRow = (int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count - 1;
            string coverCol = activeProfile.CoverCell.Substring(0, activeProfile.CoverCell.Length - activeProfile.CoverCell.TrimStart(new char[]{'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'}).Length);
            int ci = 0; while (ci < activeProfile.CoverCell.Length && char.IsLetter(activeProfile.CoverCell[ci])) ci++;
            coverCol = activeProfile.CoverCell.Substring(0, ci);
            int coverRow = int.Parse(activeProfile.CoverCell.Substring(ci));
            for (int r = coverRow; r <= lastUsedRow; r++)
            {
                try { string v = Convert.ToString(ws.Range[coverCol + r.ToString()].Value2 ?? "");
                      if (v.Trim() == activeProfile.CoverValue) offsets.Add(r - coverRow); } catch { }
            }
        } catch { }
        if (offsets.Count == 0) offsets.Add(0);
        data["coverOffsets"] = offsets;

        int offset = offsets[0];
        data["addressNum"] = GetCell(ws, activeProfile.AddressNumberCell, offset);
        data["name"] = GetCell(ws, activeProfile.NameCell, offset);
        data["staff"] = GetCell(ws, activeProfile.StaffCell, offset);
        data["residenceAddr"] = GetCell(ws, activeProfile.ResidenceAddressCell, offset);
        data["deliveryAddr"] = GetCell(ws, activeProfile.DeliveryAddressCell, offset);
        string inst = GetCell(ws, activeProfile.FinancialInstitutionCell, offset).Trim();
        // 金融機関名マッピング補正
        if (institutionNameMapping != null)
        {
            var ni = BusinessLogic.ToFullWidth(inst);
            foreach (var kv in institutionNameMapping) { if (BusinessLogic.ToFullWidth(kv.Key) == ni) { inst = kv.Value; break; } }
        }
        data["institution"] = inst;

        // 口座テーブル読取り
        var accounts = new List<AccountItem>();
        var sRows = new List<int>();
        foreach (var co in offsets)
        {
            int startRow = activeProfile.AccountStartRow + co;
            int cidx = offsets.IndexOf(co);
            for (int r = startRow; r < startRow + ACCOUNT_TABLE_MAX_ROWS; r++)
            {
                // stopValuesチェック
                bool stopped = false;
                foreach (var sv in activeProfile.StopValues)
                { try { string v = Convert.ToString(ws.Range["A" + r].Value2 ?? ""); if (v.Trim() == sv) { sRows.Add(r); stopped = true; } } catch { } }
                if (stopped) break;

                string accNum = ""; try { accNum = Convert.ToString(ws.Range[activeProfile.AccountNumberCol + r].Value2 ?? ""); } catch { }
                if (string.IsNullOrWhiteSpace(accNum)) continue;
                string bn = ""; try { bn = Convert.ToString(ws.Range[activeProfile.BranchNameCol + r].Value2 ?? ""); } catch { }
                bn = bn.Trim();
                // ゆうちょ補完
                string bnum = ""; try { bnum = Convert.ToString(ws.Range[activeProfile.BranchNumberCol + r].Value2 ?? ""); } catch { }
                bool isYucho = inst.Contains("ゆうちょ") || inst.Contains("郵貯");
                if (isYucho && string.IsNullOrWhiteSpace(bn) && yuchoMapping != null)
                { bn = GetYuchoCenter(bnum.Trim()) ?? ""; }
                if (string.IsNullOrWhiteSpace(bn)) continue; // 支店名空は除外
                // 支店サフィックス（ゆうちょ以外のみ）
                if (!isYucho && !string.IsNullOrEmpty(bn) && !IsBranchExempt(bn) && !bn.EndsWith("支店")) bn += "支店";

                string at = ""; try { at = Convert.ToString(ws.Range[activeProfile.AccountTypeCol + r].Value2 ?? ""); } catch { }
                string lt = ""; try { var lv = ws.Range[activeProfile.LastTransactionCol + r].Value2;
                    if (lv is double) lt = DateTime.FromOADate((double)lv).ToString("yyyy/MM/dd"); else if (lv != null) lt = lv.ToString(); } catch { }
                double bv = 0; try { bv = Convert.ToDouble(ws.Range[activeProfile.BalanceCol + r].Value2 ?? 0); } catch { }

                var item = new AccountItem { BranchName=bn, BranchNumber=bnum.Trim(), AccountType=at.Trim(),
                    LastTransaction=lt, AccountNum=accNum.Trim(), BalanceValue=bv, Balance=BusinessLogic.FormatBalance(bv), CoverIndex=cidx };
                string hk = BusinessLogic.FormatAddressNumber((data["addressNum"]??"").ToString()) + "_" + item.AccountNum;
                if (seizureHistory.ContainsKey(hk)) { item.HasSeizureHistory = true; item.SeizureDocNumber = seizureHistory[hk]; }
                accounts.Add(item);
            }
        }
        data["accounts"] = accounts; data["stopRows"] = sRows;
        return data;
    }

    private string GetCell(dynamic ws, string addr, int offset)
    { if (string.IsNullOrEmpty(addr)) return ""; try { return Convert.ToString(ws.Range[BusinessLogic.GetOffsetCell(addr, offset)].Value2 ?? ""); } catch { return ""; } }

    private string GetYuchoCenter(string bn)
    {
        if (yuchoMapping == null || bn.Length != 5) return null;
        string k = (bn[0] == '1') ? "総合口座" : (bn[0] == '0') ? "振替口座" : null;
        if (k == null) return null; Dictionary<string, string> t; if (!yuchoMapping.TryGetValue(k, out t)) return null;
        string c; return t.TryGetValue(bn.Substring(1, 2), out c) ? c : null;
    }

    private bool IsBranchExempt(string n) { return n.EndsWith("営業部") || n.EndsWith("出張所") || n.EndsWith("公務部") || n.EndsWith("本店"); }

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
            guideText.Text = "対象シートがありません（フィルタ条件: " + (activeProfile.FilterValue ?? "") + "）";
            btnAdd.IsEnabled = false; return;
        }
        if (sheetCombo.Items.Count > 0) sheetCombo.SelectedIndex = 0;
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
        coverOffsets = d["coverOffsets"] as List<int> ?? new List<int>();
        stopRows = d["stopRows"] as List<int> ?? new List<int>();
        currentAccounts = d["accounts"] as List<AccountItem> ?? new List<AccountItem>();
        accountList.ItemsSource = null; accountList.ItemsSource = currentAccounts;
        if (currentAccounts.Count == 0)
            guideText.Text = "表示可能な口座がありません";
        UpdateAddButton();
    }

    private void ReloadSheetData()
    {
        if (excel == null || string.IsNullOrEmpty(currentFilePath)) return;
        overlayPanel.Visibility = Visibility.Visible; loadingOverlay.Visibility = Visibility.Visible; resultOverlay.Visibility = Visibility.Collapsed;
        var worker = new BackgroundWorker();
        worker.DoWork += delegate(object s, DoWorkEventArgs args)
        { dynamic wb = null; try { wb = excel.Workbooks.Open(currentFilePath, 0, true); args.Result = ReadSheetData(wb, selectedSheetName); } finally { if (wb != null) try { wb.Close(false); } catch {} } };
        worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs args)
        { overlayPanel.Visibility = Visibility.Collapsed; if (args.Error == null) ApplySheet(args.Result as Dictionary<string, object>); };
        worker.RunWorkerAsync();
    }

    // ==============================================================
    // 処理実行（一覧に追加・スキップ）
    // ==============================================================

    private void ExecuteAdd()
    {
        var sel = accountList.SelectedItem as AccountItem;
        if (sel == null || !processingDate.HasValue) return;
        if (chkDeliveryOutput.IsChecked == true && ("（届出：" + txtDeliveryAddr.Text.Trim() + "）").Length > 50)
        { deliveryError.Visibility = Visibility.Visible; return; }

        overlayPanel.Visibility = Visibility.Visible; loadingOverlay.Visibility = Visibility.Visible; resultOverlay.Visibility = Visibility.Collapsed;
        var d = new Dictionary<string, string> {
            {"addressNum",txtAddressNum.Text.Trim()},{"name",txtName.Text.Trim()},{"staff",txtStaff.Text.Trim()},
            {"institution",txtInstitution.Text.Trim()},{"residenceAddr",txtResidenceAddr.Text.Trim()},
            {"deliveryAddr",chkDeliveryOutput.IsChecked==true?"（届出："+txtDeliveryAddr.Text.Trim()+"）":""},
            {"execDate",BusinessLogic.DateToWareki(processingDate.Value,eraMapping)},
            {"branchName",sel.BranchName},{"branchNumber",sel.BranchNumber},
            {"accountType",sel.AccountType},{"accountNum",sel.AccountNum},
            {"filePath",currentFilePath},{"fileName",System.IO.Path.GetFileName(currentFilePath)} };
        int ci = sel.CoverIndex;
        var worker = new BackgroundWorker();
        worker.DoWork += delegate(object s, DoWorkEventArgs args) { args.Result = ProcessAdd(d, ci); };
        worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs args)
        {
            overlayPanel.Visibility = Visibility.Collapsed;
            if (args.Error != null) { ShowResult("error","処理失敗","",args.Error.Message); return; }
            var r = args.Result as Dictionary<string, string>;
            if (r["status"]=="ok") { fileEntries[currentFileIndex].State=FileProcessState.Added; ShowResult("success","一覧に追加しました",r["docNumber"],"照会結果を保存しました: "+r["printFile"]); }
            else ShowResult("error","処理失敗","",r["message"]);
        };
        worker.RunWorkerAsync();
    }

    private Dictionary<string, string> ProcessAdd(Dictionary<string, string> d, int coverIdx)
    {
        var r = new Dictionary<string, string>();
        string docNum; if (!AllocateDocNumber(out docNum)) { r["status"]="error"; r["message"]="文書番号の取得に失敗"; return r; }
        var st = GenerateSeizureText(d["institution"],d["branchName"],d["branchNumber"],d["accountType"],d["accountNum"]);
        string csvBranch = d["branchName"];
        if (branchNameMapping != null) { var ni = BusinessLogic.ToFullWidth(d["institution"]); foreach (var kv in branchNameMapping) if (BusinessLogic.ToFullWidth(kv.Key)==ni) { csvBranch=kv.Value; break; } }
        string pf = docNum + ".xlsm";
        var fields = new[] { DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), d["addressNum"],d["name"],d["staff"],d["execDate"],d["residenceAddr"],d["deliveryAddr"],
            d["institution"],csvBranch,d["branchNumber"],d["accountType"],d["accountNum"], st["Line1"],st["Line2"],st["Line3"], docNum,pf,"","","" };
        string csvLine = string.Join(",", fields.Select(f => BusinessLogic.CsvEscape(f)));
        string csvPath = System.IO.Path.Combine(activeProfile.OutputFolder, CSV_FILENAME);
        if (!WriteCsvLine(csvPath, csvLine)) { RollbackDocNumber(); r["status"]="error"; r["message"]="CSV書き込み失敗"; return r; }
        string pp = System.IO.Path.Combine(activeProfile.PrintFolder, pf);
        SavePrintFile(d["filePath"], pp, d["accountNum"], d["accountType"], coverIdx);
        r["status"]="ok"; r["docNumber"]=docNum; r["printFile"]=pf; return r;
    }

    private void ExecuteSkip()
    {
        if (currentFileIndex >= 0 && currentFileIndex < fileEntries.Count) fileEntries[currentFileIndex].State = FileProcessState.Skipped;
        ShowResult("skip", "スキップしました", "", System.IO.Path.GetFileName(currentFilePath));
    }

    // ==============================================================
    // 処理結果オーバーレイ
    // ==============================================================

    private void ShowResult(string type, string title, string docNum, string detail)
    {
        overlayPanel.Visibility = Visibility.Visible; loadingOverlay.Visibility = Visibility.Collapsed; resultOverlay.Visibility = Visibility.Visible;
        resultTitle.Text = title;
        resultDocNum.Text = !string.IsNullOrEmpty(docNum) ? "文書番号: " + docNum : "";
        resultDocNum.Visibility = string.IsNullOrEmpty(docNum) ? Visibility.Collapsed : Visibility.Visible;
        resultDetail.Text = detail;
        bool last = currentFileIndex >= fileEntries.Count - 1;
        if (type == "success") { resultIcon.Text = "\u2713"; resultIcon.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#107C41")); }
        else if (type == "skip") { resultIcon.Text = "\u2192"; resultIcon.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#005FB8")); }
        else { resultIcon.Text = "\u2717"; resultIcon.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D32F2F")); }
        resultButton.Content = last ? "完了" : "次のファイルへ \u2192";
        resultSub.Text = last ? "" : ((currentFileIndex + 2) + " / " + fileEntries.Count + " 件目へ進みます");
    }

    private void ProceedToNext() { overlayPanel.Visibility = Visibility.Collapsed; LoadFileAtIndex(currentFileIndex + 1); }

    // ==============================================================
    // 差押文言生成
    // ==============================================================

    private Dictionary<string, string> GenerateSeizureText(string inst, string branch, string branchNum, string accType, string accNum)
    {
        var ni = BusinessLogic.ToFullWidth((inst ?? "").Trim());
        string dat = accType ?? "";
        if (accountTypeMapping != null) { var m = FindTypeMatch(ni, dat); if (m != null) dat = m; }
        string fee = "";
        if (feeMapping != null) foreach (var kv in feeMapping) { if (BusinessLogic.ToFullWidth(kv.Key) == ni) { fee = kv.Value ?? ""; break; } }
        string fc = !string.IsNullOrEmpty(fee) ? "と手数料の合計額" : "";
        bool yu = ni == BusinessLogic.ToFullWidth("ゆうちょ銀行");
        string fa = BusinessLogic.ToFullWidth(accNum.Trim());
        string l1;
        if (yu) { string fb = BusinessLogic.ToFullWidth(branchNum.Trim());
            l1 = "\u3000上記滞納者が、債務者であるゆうちょ銀行（"+branch+"扱）に対して有する"+dat+"（記号番号："+fb+"－"+fa+"）の払戻請求権及びこれに対する債権差押通知書到達日までの約定利息の支払請求権。ただし、滞納金額"+fc+"に充るまでとし、残高が３，０００円未満の場合は差押しない。"; }
        else { l1 = "\u3000上記滞納者が、債務者である"+ni+"（"+branch.Trim()+"扱）に対して有する"+dat+"（口座番号："+fa+"）の払戻請求権及びこれに対する債権差押通知書到達日までの約定利息の支払請求権。ただし、滞納金額"+fc+"に充るまでとし、残高が３，０００円未満の場合は差押しない。"; }
        return new Dictionary<string, string> { {"Line1",l1}, {"Line2",fee},
            {"Line3","債権差押通知書到達日現在の残高\u3000\u3000\u3000\u3000\u3000\u3000円\u3000差押額\u3000\u3000\u3000\u3000\u3000\u3000円"} };
    }

    private string FindTypeMatch(string ni, string at)
    {
        foreach (var kv in accountTypeMapping) { if (kv.Key=="global") continue;
            if (BusinessLogic.ToFullWidth(kv.Key)==ni) { var m=MatchRules(kv.Value,at); if (m!=null) return m; break; } }
        Dictionary<string,string> g; if (accountTypeMapping.TryGetValue("global",out g)) { var m=MatchRules(g,at); if (m!=null) return m; }
        return null;
    }

    private string MatchRules(Dictionary<string,string> rules, string target)
    { foreach (var kv in rules) { bool w=kv.Key.EndsWith("*"); string p=w?kv.Key.TrimEnd('*'):kv.Key;
        if (w?target.StartsWith(p):target==kv.Key) return kv.Value; } return null; }

    // ==============================================================
    // CSV 操作
    // ==============================================================

    private void BuildSeizureHistory()
    {
        seizureHistory.Clear();
        if (activeProfile.OutputFolder == null) return;
        string p = System.IO.Path.Combine(activeProfile.OutputFolder, CSV_FILENAME);
        if (!File.Exists(p)) return;
        try { var lines = File.ReadAllLines(p, Encoding.UTF8);
            for (int i = 1; i < lines.Length; i++) { var f = ParseCsv(lines[i]);
                if (f.Length >= 16) seizureHistory[f[1].Trim()+"_"+f[11].Trim()] = f[15].Trim(); } } catch { }
    }

    private string[] ParseCsv(string line)
    {
        var fs = new List<string>(); int p = 0;
        while (p < line.Length)
        {
            if (line[p] == '"') { p++; var sb = new StringBuilder();
                while (p < line.Length) { if (line[p]=='"') { if (p+1<line.Length&&line[p+1]=='"') { sb.Append('"'); p+=2; } else { p++; break; } } else { sb.Append(line[p]); p++; } }
                fs.Add(sb.ToString()); if (p < line.Length && line[p]==',') p++; }
            else { int n = line.IndexOf(',', p); if (n<0) { fs.Add(line.Substring(p)); break; } fs.Add(line.Substring(p, n-p)); p = n+1; }
        }
        return fs.ToArray();
    }

    private bool WriteCsvLine(string path, string line)
    {
        var dir = System.IO.Path.GetDirectoryName(path);
        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        bool isNew = !File.Exists(path);
        for (int r = 1; r <= CSV_WRITE_MAX_RETRY; r++)
        { try { using (var fs = new FileStream(path,FileMode.Append,FileAccess.Write,FileShare.None))
            using (var sw = new StreamWriter(fs,new UTF8Encoding(true)))
            { if (isNew) sw.WriteLine(CSV_HEADER); sw.WriteLine(line); sw.Flush(); } return true;
          } catch { if (r < CSV_WRITE_MAX_RETRY) System.Threading.Thread.Sleep(CSV_WRITE_RETRY_INTERVAL_MS); } }
        return false;
    }

    // ==============================================================
    // 文書番号管理
    // ==============================================================

    private bool AllocateDocNumber(out string docNum)
    {
        docNum = ""; string cp = System.IO.Path.Combine(exeDir, "document_number_counter.json");
        for (int r = 1; r <= CSV_WRITE_MAX_RETRY; r++)
        { try { using (var fs = new FileStream(cp,FileMode.Open,FileAccess.ReadWrite,FileShare.None))
            { var b = new byte[fs.Length]; fs.Read(b,0,b.Length); var j = Encoding.UTF8.GetString(b).TrimStart('\uFEFF');
              int n = JsonHelper.GetInt(j,"nextNumber",1); docNum = n.ToString(); lastDocNumber = docNum;
              fs.Seek(0,SeekOrigin.Begin); fs.SetLength(0);
              var nb = Encoding.UTF8.GetBytes("{\n    \"nextNumber\":  "+(n+1)+"\n}"); fs.Write(nb,0,nb.Length); } return true;
          } catch { if (r < CSV_WRITE_MAX_RETRY) System.Threading.Thread.Sleep(CSV_WRITE_RETRY_INTERVAL_MS); } }
        return false;
    }

    private void RollbackDocNumber()
    {
        if (string.IsNullOrEmpty(lastDocNumber)) return;
        string cp = System.IO.Path.Combine(exeDir, "document_number_counter.json");
        for (int r = 1; r <= CSV_WRITE_MAX_RETRY; r++)
        { try { using (var fs = new FileStream(cp,FileMode.Open,FileAccess.ReadWrite,FileShare.None))
            { fs.Seek(0,SeekOrigin.Begin); fs.SetLength(0);
              var b = Encoding.UTF8.GetBytes("{\n    \"nextNumber\":  "+lastDocNumber+"\n}"); fs.Write(b,0,b.Length); } return;
          } catch { if (r < CSV_WRITE_MAX_RETRY) System.Threading.Thread.Sleep(CSV_WRITE_RETRY_INTERVAL_MS); } }
    }

    // ==============================================================
    // 印刷用ファイル保存
    // ==============================================================

    private void SavePrintFile(string src, string dst, string selAccNum, string selAccType, int selCoverIdx)
    {
        if (dst.Length > MAX_PATH) return;
        var dir = System.IO.Path.GetDirectoryName(dst); if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        dynamic pwb = null;
        try
        {
            excel.Visible = false; pwb = excel.Workbooks.Open(src, 0, false);
            try { if ((bool)pwb.ProtectStructure) pwb.Unprotect(); } catch {}
            excel.DisplayAlerts = false; pwb.SaveAs(dst, 52); excel.DisplayAlerts = true;
            try { if ((bool)pwb.ProtectStructure) pwb.Unprotect(); } catch {}
            // 選択シート以外を削除
            excel.DisplayAlerts = false;
            for (int j = (int)pwb.Worksheets.Count; j >= 1; j--)
            { dynamic ws = pwb.Worksheets[j]; if ((int)ws.Visible != -1 && (string)ws.Name != selectedSheetName) try { ws.Delete(); } catch {} }
            excel.DisplayAlerts = true;
            // 削除対象範囲を集約
            dynamic pws = pwb.Worksheets[selectedSheetName];
            var delRanges = new List<int[]>();
            int plr; try { plr = (int)pws.UsedRange.Row + (int)pws.UsedRange.Rows.Count - 1; } catch { plr = lastUsedRow; }
            int ks, ke;
            if (coverOffsets.Count > 1) {
                ks = coverOffsets[selCoverIdx] + 1;
                ke = (selCoverIdx+1 < coverOffsets.Count) ? coverOffsets[selCoverIdx+1] : plr;
                if (ke < plr) delRanges.Add(new[]{ke+1, plr});
                if (ks > 1) delRanges.Add(new[]{1, ks-1});
            } else { ks = 1; ke = plr; }
            // 明細ページ削除
            if (!string.IsNullOrEmpty(activeProfile.DetailAccountNumberCell) && stopRows.Count >= 2)
            {
                var m = System.Text.RegularExpressions.Regex.Match(activeProfile.DetailAccountNumberCell, @"^([A-Za-z]{1,3})(\d+)$");
                if (m.Success) {
                    string dc = m.Groups[1].Value; int dro = int.Parse(m.Groups[2].Value) - 1;
                    string tc = null; int tro = 0;
                    if (!string.IsNullOrEmpty(activeProfile.DetailAccountTypeCell)) {
                        var tm = System.Text.RegularExpressions.Regex.Match(activeProfile.DetailAccountTypeCell, @"^([A-Za-z]{1,3})(\d+)$");
                        if (tm.Success) { tc = tm.Groups[1].Value; tro = int.Parse(tm.Groups[2].Value) - 1; } }
                    var rs = stopRows.Where(x => x >= ks && x <= ke).ToList();
                    if (rs.Count >= 2) {
                        int pr = rs[1] - rs[0]; int md = rs[0] % pr; int sip = (md==0)?pr:md; int rb = sip - 1;
                        for (int i = 0; i < rs.Count; i++) {
                            int ps = rs[i] - rb; int pe = ps + pr - 1;
                            string av = ""; try { av = Convert.ToString(pws.Range[dc+(rs[i]+dro)].Value2 ?? ""); } catch {}
                            bool am = av.Trim() == selAccNum;
                            bool tm2 = true; if (tc != null) { string tv = ""; try { tv = Convert.ToString(pws.Range[tc+(rs[i]+tro)].Value2 ?? ""); } catch {} tm2 = tv.Trim() == selAccType; }
                            if (!(am && tm2)) delRanges.Add(new[]{ps, pe}); } } } }
            // 後方から削除
            delRanges.Sort((a,b) => b[0].CompareTo(a[0]));
            excel.DisplayAlerts = false;
            foreach (var rng in delRanges) try { pws.Rows[rng[0]+":"+rng[1]].Delete(); } catch {}
            excel.DisplayAlerts = true;
            // 表示位置リセット
            try { pws.Activate(); pwb.Application.ActiveWindow.ScrollRow=1; pwb.Application.ActiveWindow.ScrollColumn=1; pws.Range["A1"].Select(); } catch {}
            pwb.Save();
        }
        catch { /* 保存失敗は警告のみ */ }
        finally { if (pwb != null) try { pwb.Close(false); } catch {} }
    }

    // ==============================================================
    // Excel COM クリーンアップ
    // ==============================================================

    private void CleanupExcel()
    {
        if (excel == null) return;
        try { excel.Quit(); Marshal.ReleaseComObject(excel); } catch {}
        finally { excel = null; GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); }
    }

    // ==============================================================
    // XAML 定義
    // ==============================================================

    private Window BuildWindow()
    {
        string xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    Title='預金差押予定一覧 作成ツール' Width='1000' Height='620' MinWidth='900' MinHeight='520'
    WindowStartupLocation='CenterScreen' Background='#F9F9F9' FontFamily='Meiryo UI'
    UseLayoutRounding='True' SnapsToDevicePixels='True'>
<Window.Resources>
    <Style x:Key='AB' TargetType='Button'>
        <Setter Property='Background' Value='#005FB8'/><Setter Property='Foreground' Value='White'/>
        <Setter Property='FontSize' Value='12'/><Setter Property='Padding' Value='16,8'/>
        <Setter Property='Cursor' Value='Hand'/><Setter Property='BorderThickness' Value='0'/>
        <Setter Property='Template'><Setter.Value><ControlTemplate TargetType='Button'>
            <Border x:Name='bd' Background='{TemplateBinding Background}' CornerRadius='4' Padding='{TemplateBinding Padding}'>
                <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='bd' Property='Background' Value='#004FA0'/></Trigger>
                <Trigger Property='IsEnabled' Value='False'><Setter TargetName='bd' Property='Background' Value='#CCC'/><Setter Property='Foreground' Value='#999'/></Trigger>
            </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style>
    <Style x:Key='GB' TargetType='Button'>
        <Setter Property='Background' Value='White'/><Setter Property='Foreground' Value='#555'/>
        <Setter Property='FontSize' Value='12'/><Setter Property='Padding' Value='16,8'/>
        <Setter Property='Cursor' Value='Hand'/><Setter Property='BorderBrush' Value='#D0D0D0'/><Setter Property='BorderThickness' Value='1'/>
        <Setter Property='Template'><Setter.Value><ControlTemplate TargetType='Button'>
            <Border x:Name='bd' Background='{TemplateBinding Background}' BorderBrush='{TemplateBinding BorderBrush}'
                BorderThickness='{TemplateBinding BorderThickness}' CornerRadius='4' Padding='{TemplateBinding Padding}'>
                <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='bd' Property='Background' Value='#F0F4F8'/></Trigger>
            </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style>
</Window.Resources>
<DockPanel>
    <Border DockPanel.Dock='Top' Background='#005FB8' Padding='18,10'>
        <TextBlock Text='預金差押予定一覧 作成ツール' FontSize='13' FontWeight='Medium' Foreground='White'/></Border>
    <Border DockPanel.Dock='Bottom' Background='#F0F0F0' BorderBrush='#E0E0E0' BorderThickness='0,1,0,0' Padding='18,4'>
        <DockPanel><TextBlock x:Name='StatusRight' DockPanel.Dock='Right' FontSize='11' Foreground='#666'/>
            <TextBlock x:Name='StatusLeft' FontSize='11' Foreground='#666'/></DockPanel></Border>
    <Grid>
        <!-- 初期画面 -->
        <Grid x:Name='InitialPanel'>
            <Border Background='White' BorderBrush='#C0C8D0' BorderThickness='2' CornerRadius='8'
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
                    <TextBlock x:Name='FileLink' FontSize='11' Foreground='#005FB8'
                               Cursor='Hand' TextDecorations='Underline' VerticalAlignment='Center'/>
                    <Button x:Name='BtnReload' Style='{StaticResource GB}' Padding='8,4' Margin='8,0,0,0' FontSize='11'>
                        <TextBlock Text='再読み込み'/></Button>
                </StackPanel>
            </DockPanel>
            <!-- 基本情報 -->
            <Border Grid.Row='1' Background='White' BorderBrush='#E0E0E0' BorderThickness='1' CornerRadius='6' Padding='16,14' Margin='0,0,0,10'>
                <StackPanel><TextBlock Text='基本情報' FontSize='11' Foreground='#005FB8' FontWeight='Medium' Margin='0,0,0,10'/>
                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width='*'/><ColumnDefinition Width='16'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>
                    <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='6'/><RowDefinition Height='Auto'/></Grid.RowDefinitions>
                    <Grid Grid.Row='0' Grid.Column='0'><Grid.ColumnDefinitions><ColumnDefinition Width='120'/><ColumnDefinition Width='8'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>
                        <StackPanel><TextBlock Text='宛名番号' FontSize='10' Foreground='#888' Margin='0,0,0,3'/>
                            <TextBox x:Name='TxtAddressNum' IsReadOnly='True' Background='#F3F3F3' BorderBrush='#E8E8E8' FontFamily='Consolas' FontSize='12' Padding='5,4'/></StackPanel>
                        <StackPanel Grid.Column='2'><TextBlock FontSize='10' Foreground='#888' Margin='0,0,0,3'>氏名 &#x270E;</TextBlock>
                            <TextBox x:Name='TxtName' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/></StackPanel></Grid>
                    <Grid Grid.Row='0' Grid.Column='2'><Grid.ColumnDefinitions><ColumnDefinition Width='*'/><ColumnDefinition Width='8'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>
                        <StackPanel><TextBlock Text='金融機関' FontSize='10' Foreground='#888' Margin='0,0,0,3'/>
                            <TextBox x:Name='TxtInstitution' IsReadOnly='True' Background='#F3F3F3' BorderBrush='#E8E8E8' FontSize='12' Padding='5,4'/></StackPanel>
                        <StackPanel Grid.Column='2'><TextBlock FontSize='10' Foreground='#888' Margin='0,0,0,3'>処分担当 &#x270E;</TextBlock>
                            <TextBox x:Name='TxtStaff' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/></StackPanel></Grid>
                    <StackPanel Grid.Row='2' Grid.Column='0'><TextBlock FontSize='10' Foreground='#888' Margin='0,0,0,3'>住民票住所 &#x270E;</TextBlock>
                        <TextBox x:Name='TxtResidenceAddr' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/></StackPanel>
                    <StackPanel Grid.Row='2' Grid.Column='2'><TextBlock FontSize='10' Foreground='#888' Margin='0,0,0,3'>届出住所 &#x270E;</TextBlock>
                        <TextBox x:Name='TxtDeliveryAddr' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/>
                        <CheckBox x:Name='ChkDeliveryOutput' Content='差押通知書に出力する' FontSize='11' Margin='0,3,0,0'/>
                        <TextBlock x:Name='DeliveryError' Foreground='#D32F2F' FontSize='10' Visibility='Collapsed' Margin='0,1,0,0'/></StackPanel>
                </Grid>
                <StackPanel Margin='0,6,0,0'><TextBlock Text='執行日' FontSize='10' Foreground='#888' Margin='0,0,0,3'/>
                    <TextBox x:Name='TxtExecDate' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0' Width='150' HorizontalAlignment='Left'/></StackPanel>
                </StackPanel></Border>
            <!-- 口座選択 -->
            <Border Grid.Row='2' Background='White' BorderBrush='#E0E0E0' BorderThickness='1' CornerRadius='6' Padding='16,14' Margin='0,0,0,10'>
                <DockPanel><TextBlock DockPanel.Dock='Top' Text='口座選択' FontSize='11' Foreground='#005FB8' FontWeight='Medium' Margin='0,0,0,10'/>
                    <ListView x:Name='AccountList' BorderThickness='0' Background='Transparent' FontSize='12' SelectionMode='Single'>
                        <ListView.ItemContainerStyle>
                            <Style TargetType='ListViewItem'>
                                <Setter Property='ToolTip' Value='{Binding SeizureTooltip}'/>
                                <Setter Property='Cursor' Value='Hand'/>
                            </Style>
                        </ListView.ItemContainerStyle>
                        <ListView.View><GridView>
                            <GridViewColumn Header='支店名' Width='250' DisplayMemberBinding='{Binding BranchName}'/>
                            <GridViewColumn Header='支店番号' Width='80' DisplayMemberBinding='{Binding BranchNumber}'/>
                            <GridViewColumn Header='口座種別' Width='90' DisplayMemberBinding='{Binding AccountType}'/>
                            <GridViewColumn Header='最終取引日(満期日)' Width='120' DisplayMemberBinding='{Binding LastTransaction}'/>
                            <GridViewColumn Header='口座番号' Width='100' DisplayMemberBinding='{Binding AccountNumDisplay}'/>
                            <GridViewColumn Header='残高' Width='140' DisplayMemberBinding='{Binding Balance}'/>
                        </GridView></ListView.View></ListView></DockPanel></Border>
            <!-- アクションボタン -->
            <DockPanel Grid.Row='3'>
                <Button x:Name='BtnLoadFile' DockPanel.Dock='Left' Style='{StaticResource GB}'><TextBlock Text='ファイルを読み込む'/></Button>
                <StackPanel DockPanel.Dock='Right' Orientation='Horizontal' HorizontalAlignment='Right'>
                    <TextBlock x:Name='GuideText' VerticalAlignment='Center' FontSize='11' Foreground='#D32F2F' Margin='0,0,12,0'/>
                    <Button x:Name='BtnSkip' Style='{StaticResource GB}' Margin='0,0,8,0'><TextBlock Text='スキップ'/></Button>
                    <Button x:Name='BtnAdd' Style='{StaticResource AB}' IsEnabled='False'><TextBlock Text='一覧に追加'/></Button>
                </StackPanel></DockPanel>
        </Grid>
        <!-- オーバーレイ -->
        <Grid x:Name='OverlayPanel' Visibility='Collapsed' Background='#CCFFFFFF'>
            <Grid x:Name='LoadingOverlay' Visibility='Collapsed' HorizontalAlignment='Center' VerticalAlignment='Center'>
                <Border Width='60' Height='60' CornerRadius='30' Background='#005FB8'>
                    <TextBlock Text='...' FontSize='20' Foreground='White' HorizontalAlignment='Center' VerticalAlignment='Center'/></Border></Grid>
            <Grid x:Name='ResultOverlay' Visibility='Collapsed' HorizontalAlignment='Center' VerticalAlignment='Center'>
                <Border Background='White' CornerRadius='10' Padding='44,32' MinWidth='350'>
                    <Border.Effect><DropShadowEffect BlurRadius='16' ShadowDepth='4' Opacity='0.12'/></Border.Effect>
                    <StackPanel HorizontalAlignment='Center'>
                        <Border Width='52' Height='52' CornerRadius='26' Background='#E8F5ED' HorizontalAlignment='Center' Margin='0,0,0,16'>
                            <TextBlock x:Name='ResultIcon' Text='&#x2713;' FontSize='28' HorizontalAlignment='Center' VerticalAlignment='Center' Foreground='#107C41'/></Border>
                        <TextBlock x:Name='ResultTitle' Text='' FontSize='16' FontWeight='Medium' HorizontalAlignment='Center' Margin='0,0,0,6'/>
                        <TextBlock x:Name='ResultDocNum' FontSize='18' FontWeight='Medium' HorizontalAlignment='Center' Margin='0,0,0,8'/>
                        <TextBlock x:Name='ResultDetail' FontSize='11' Foreground='#999' HorizontalAlignment='Center' Margin='0,0,0,20'/>
                        <Button x:Name='ResultButton' Style='{StaticResource AB}' HorizontalAlignment='Center' Padding='24,10'>
                            <TextBlock Text='次のファイルへ' FontSize='13'/></Button>
                        <TextBlock x:Name='ResultSub' FontSize='11' Foreground='#999' HorizontalAlignment='Center' Margin='0,10,0,0'/>
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