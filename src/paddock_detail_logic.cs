// ==============================================================================
// paddock_detail_logic.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 画面を持たない処理のうち、明細を見る（読み取り位置の設定値の解釈・照会ファイルの
// Open XML の読み取り・明細ページの特定と取引の抽出・口座種別の補正・推定残高・カテゴリ付け・集約）。
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
    // ---------- 明細を見る: 設定値の解釈 ----------

    // 列文字（A〜XFD の範囲で1〜3文字）を1始まりの列番号にする。JUDGE の Convert-ColToIndex と同式
    public static bool TryParseColumnLetters(string letters, out int index)
    {
        index = 0;
        if (letters == null) return false;
        var trimmed = letters.Trim();
        if (trimmed.Length < 1 || trimmed.Length > 3) return false;
        foreach (var c in trimmed)
        {
            if (c < 'A' || c > 'Z')
            {
                if (c < 'a' || c > 'z') { index = 0; return false; }
                index = index * 26 + (c - 'a' + 1);
            }
            else index = index * 26 + (c - 'A' + 1);
        }
        return index <= 16384;
    }

    // セル参照（例: "M3"）を列番号と行オフセット（1行目=0）にする。JUDGE の detail*Cell と同じ解釈
    public static bool TryParseCellRef(string cell, out int colIndex, out int rowOffset)
    {
        colIndex = 0;
        rowOffset = 0;
        if (cell == null) return false;
        var trimmed = cell.Trim();
        int split = 0;
        while (split < trimmed.Length && char.IsLetter(trimmed[split])) split++;
        if (split == 0 || split == trimmed.Length) return false;
        int row;
        if (!int.TryParse(trimmed.Substring(split), NumberStyles.None, CultureInfo.InvariantCulture, out row) || row < 1)
            return false;
        if (!TryParseColumnLetters(trimmed.Substring(0, split), out colIndex)) return false;
        rowOffset = row - 1;
        return true;
    }

    // ---------- 明細を見る: 照会ファイル（OPC）の読み取り ----------
    // Excel を起動せず System.IO.Packaging でシート XML を直接読む。読む部品は
    // workbook.xml（シート名→関係ID）・workbook.xml.rels（関係ID→パート）・sharedStrings.xml・
    // 期待シートの XML の4つで、DOM を作るのは小さな前2者のみ。styles.xml は読まない

    // Open XML の名前空間（シート照合の FileHasExpectedSheet と共用）
    public const string NS_SPREADSHEET    = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string NS_OFFICE_RELS   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string NS_PACKAGE_RELS  = "http://schemas.openxmlformats.org/package/2006/relationships";

    // 照会ファイルの骨格（シート名・共有文字列）を読む。以後のシート本体の読み取りに使う
    public static InquiryWorkbook OpenInquiryWorkbook(string path)
    {
        var result = new InquiryWorkbook { Path = path };
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var package = Package.Open(fs, FileMode.Open, FileAccess.Read))
        {
            // workbook.xml: シートの並びと関係ID・1904年始まり
            var wbUri = new Uri("/xl/workbook.xml", UriKind.Relative);
            if (!package.PartExists(wbUri)) throw new InvalidDataException("xl/workbook.xml がありません");
            var wbDoc = new XmlDocument();
            using (var stream = package.GetPart(wbUri).GetStream(FileMode.Open, FileAccess.Read)) wbDoc.Load(stream);
            var nsmgr = new XmlNamespaceManager(wbDoc.NameTable);
            nsmgr.AddNamespace("s", NS_SPREADSHEET);
            nsmgr.AddNamespace("r", NS_OFFICE_RELS);
            var pr = wbDoc.SelectSingleNode("//s:workbookPr", nsmgr);
            result.Date1904 = pr != null && pr.Attributes["date1904"] != null &&
                (pr.Attributes["date1904"].Value == "1" || pr.Attributes["date1904"].Value == "true");

            // workbook.xml.rels: 関係ID→パート（Target は xl/ からの相対か絶対パス）
            var relMap = new Dictionary<string, Uri>();
            var relUri = new Uri("/xl/_rels/workbook.xml.rels", UriKind.Relative);
            if (package.PartExists(relUri))
            {
                var relDoc = new XmlDocument();
                using (var stream = package.GetPart(relUri).GetStream(FileMode.Open, FileAccess.Read)) relDoc.Load(stream);
                var relNs = new XmlNamespaceManager(relDoc.NameTable);
                relNs.AddNamespace("p", NS_PACKAGE_RELS);
                foreach (XmlNode rel in relDoc.SelectNodes("//p:Relationship", relNs))
                {
                    var id = rel.Attributes["Id"];
                    var target = rel.Attributes["Target"];
                    if (id == null || target == null) continue;
                    var targetPath = target.Value.StartsWith("/") ? target.Value : "/xl/" + target.Value;
                    relMap[id.Value] = new Uri(targetPath, UriKind.Relative);
                }
            }
            foreach (XmlNode sheet in wbDoc.SelectNodes("//s:sheets/s:sheet", nsmgr))
            {
                var name = sheet.Attributes["name"];
                var rid = sheet.Attributes["id", NS_OFFICE_RELS];
                Uri partUri;
                if (name != null && rid != null && relMap.TryGetValue(rid.Value, out partUri) && package.PartExists(partUri))
                    result.Sheets.Add(new KeyValuePair<string, Uri>(name.Value, partUri));
            }

            var ssUri = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
            if (package.PartExists(ssUri))
                using (var stream = package.GetPart(ssUri).GetStream(FileMode.Open, FileAccess.Read))
                    result.SharedStrings = ReadSharedStrings(stream);
        }
        return result;
    }

    // sharedStrings.xml を先頭から流し読みする。1項目（si）の本文は直下の t と書式付き
    // 連（r/t）の連結で、ふりがな（rPh）の t は本文に含めない
    private static string[] ReadSharedStrings(Stream stream)
    {
        var result = new List<string>();
        var sb = new StringBuilder();
        bool inItem = false, inPhonetic = false;
        using (var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false, IgnoreComments = true }))
        {
            // 要素内容を読み取ったあとはリーダーが次のノードに進んでいるため、その回は Read() を挟まない
            bool advance = true;
            while (!advance || reader.Read())
            {
                advance = true;
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "si": inItem = true; sb.Length = 0; break;
                        case "rPh": inPhonetic = true; break;
                        case "t":
                            if (inItem && !inPhonetic)
                            {
                                sb.Append(reader.ReadElementContentAsString());
                                advance = false;
                            }
                            break;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.LocalName == "rPh") inPhonetic = false;
                    else if (reader.LocalName == "si") { result.Add(sb.ToString()); inItem = false; }
                }
            }
        }
        return result.ToArray();
    }

    // シート1枚のセルを流し読みする。wantedCols（1始まりの列番号）に含む列だけを保持し、
    // 数値セルは生値のまま・文字列セルは共有文字列/数式結果/インライン文字列を本文にする。
    // r 属性のないセルは同じ行の直前のセルの次の列とみなす
    public static SheetCells ReadSheetCells(InquiryWorkbook workbook, Uri partUri, HashSet<int> wantedCols)
    {
        var result = new SheetCells();
        using (var fs = new FileStream(workbook.Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var package = Package.Open(fs, FileMode.Open, FileAccess.Read))
        using (var stream = package.GetPart(partUri).GetStream(FileMode.Open, FileAccess.Read))
        using (var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false, IgnoreComments = true }))
        {
            int currentRow = 0, lastCol = 0;
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName == "row")
                {
                    int r;
                    var rAttr = reader.GetAttribute("r");
                    currentRow = rAttr != null && int.TryParse(rAttr, NumberStyles.None, CultureInfo.InvariantCulture, out r) ? r : currentRow + 1;
                    lastCol = 0;
                    continue;
                }
                if (reader.LocalName != "c") continue;

                int col, rowOffset, row = currentRow;
                var refAttr = reader.GetAttribute("r");
                if (refAttr != null && TryParseCellRef(refAttr, out col, out rowOffset)) row = rowOffset + 1;
                else col = lastCol + 1;
                lastCol = col;
                if (row > result.MaxRow) result.MaxRow = row;
                if (!wantedCols.Contains(col) || reader.IsEmptyElement) continue;

                var type = reader.GetAttribute("t");
                string text = null;
                using (var sub = reader.ReadSubtree())
                {
                    while (sub.Read())
                    {
                        if (sub.NodeType != XmlNodeType.Element) continue;
                        if (sub.LocalName == "v") { text = sub.ReadElementContentAsString(); break; }
                        if (sub.LocalName == "is")
                        {
                            // インライン文字列は t の連結（ふりがなは除く）
                            var sb = new StringBuilder();
                            bool inPhonetic = false, advance = true;
                            while (!advance || sub.Read())
                            {
                                advance = true;
                                if (sub.NodeType == XmlNodeType.Element)
                                {
                                    if (sub.LocalName == "rPh") inPhonetic = true;
                                    else if (sub.LocalName == "t" && !inPhonetic)
                                    {
                                        sb.Append(sub.ReadElementContentAsString());
                                        advance = false;
                                    }
                                }
                                else if (sub.NodeType == XmlNodeType.EndElement && sub.LocalName == "rPh") inPhonetic = false;
                            }
                            text = sb.ToString();
                            break;
                        }
                    }
                }
                if (text == null) continue;

                var cell = new SheetCell();
                if (type == "s")
                {
                    int index;
                    cell.Text = int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out index) &&
                        index >= 0 && index < workbook.SharedStrings.Length ? workbook.SharedStrings[index] : "";
                }
                else if (type == null || type == "n")
                {
                    cell.Text = text;
                    cell.IsNumeric = true;
                }
                else cell.Text = text;   // str（数式結果）・inlineStr・b・e はそのまま

                Dictionary<int, SheetCell> cols;
                if (!result.Rows.TryGetValue(row, out cols))
                {
                    cols = new Dictionary<int, SheetCell>();
                    result.Rows[row] = cols;
                }
                cols[col] = cell;
            }
        }
        return result;
    }

    // ---------- 明細を見る: 明細ページの特定と取引の抽出 ----------
    // 位置の規則は JUDGE（judge.ps1 の Invoke-DetailAnalysis）と同一: 目印列（A）の完全一致で
    // ページ先頭を集め、ページ長は隣接する目印行の差（単一ページはシート末尾まで）、口座は
    // 目印セル相対の口座番号・口座種別の両方一致、明細は固定の開始行から読み始めて
    // 摘要・出金・入金がすべて空の行で終わる（摘要だけ空の行は読み飛ばす）。
    // 口座種別だけは JUDGE と違い、台帳の表記（JUDGE が補正した後）とも照合する（MatchAccountPages）

    private const int DETAIL_MARKER_COL = 1;         // 目印文字列の列（A）

    // 読み取り対象の列（目印・口座番号・口座種別・明細5列）
    public static HashSet<int> GetDetailWantedColumns(DetailConfig detail)
    {
        return new HashSet<int>
        {
            DETAIL_MARKER_COL, detail.AccountNumberColIndex, detail.AccountTypeColIndex,
            detail.DateColIndex, detail.SummaryColIndex, detail.WithdrawalColIndex, detail.DepositColIndex, detail.BalanceColIndex
        };
    }

    // 目印文字列に一致する行（明細ページの先頭）を昇順で返す
    public static List<int> FindStopRows(SheetCells cells, DetailConfig detail)
    {
        var result = new List<int>();
        for (int row = 1; row <= cells.MaxRow; row++)
        {
            var text = cells.GetText(row, DETAIL_MARKER_COL).Trim();
            if (text != "" && detail.StopValues.Contains(text)) result.Add(row);
        }
        return result;
    }

    // 口座番号・口座種別の両方が一致する明細ページの先頭行を返す。口座種別は、明細ページの表記が
    // 台帳の表記と同じか、明細ページの表記を口座種別の補正表で補正すると台帳の表記になれば一致とする
    //（台帳には JUDGE が補正した後の表記が書かれている。補正に使う金融機関名は台帳の行の値。
    //  補正表がなければ補正しない）
    public static List<int> MatchAccountPages(SheetCells cells, List<int> stopRows, DetailConfig detail, string accountNumber, string accountType,
        string institutionName, AccountTypeMapping typeMapping)
    {
        var result = new List<int>();
        var wantedNumber = RemoveWhitespace(accountNumber);
        var wantedType = (accountType ?? "").Trim();
        foreach (var sr in stopRows)
        {
            var number = RemoveWhitespace(cells.GetText(sr + detail.AccountNumberRowOffset, detail.AccountNumberColIndex));
            if (number != wantedNumber) continue;
            var type = cells.GetText(sr + detail.AccountTypeRowOffset, detail.AccountTypeColIndex).Trim();
            if (type == wantedType || (typeMapping != null && ResolveAccountType(type, institutionName, typeMapping).Trim() == wantedType))
                result.Add(sr);
        }
        return result;
    }

    // 一致した明細ページの取引を読む（ページの掲載順のまま）
    public static List<DetailTransaction> ReadTransactions(SheetCells cells, List<int> stopRows, List<int> matchedStopRows,
        DetailConfig detail, bool date1904, Dictionary<int, EraEntry> eraMap)
    {
        var result = new List<DetailTransaction>();
        int pageRows = stopRows.Count >= 2 ? stopRows[1] - stopRows[0] : 0;
        foreach (var sr in matchedStopRows)
        {
            int startRow = sr + detail.DataStartRow - 1;
            int endRow = pageRows > 0 ? sr + pageRows - 1 : cells.MaxRow;
            for (int row = startRow; row <= endRow; row++)
            {
                var summary = cells.GetText(row, detail.SummaryColIndex);
                var withdrawal = cells.Get(row, detail.WithdrawalColIndex);
                var deposit = cells.Get(row, detail.DepositColIndex);
                if (string.IsNullOrWhiteSpace(summary) && IsEmptyCell(withdrawal) && IsEmptyCell(deposit)) break;
                if (string.IsNullOrWhiteSpace(summary)) continue;

                // 回答コード（-999999 始まり・一覧の「対象外」と同じ判定）のセルは金額として扱わない。
                // 出金・入金のどちらにも金額がなく回答コードを含む行は、取引でなく回答コードの行
                //（期間中に入出金がない旨の行など）として数えない
                var balance = cells.Get(row, detail.BalanceColIndex);
                bool withdrawalCode = IsSentinelCell(withdrawal), depositCode = IsSentinelCell(deposit), balanceCode = IsSentinelCell(balance);
                var withdrawalAmount = withdrawalCode ? null : ParseDetailAmount(withdrawal);
                var depositAmount = depositCode ? null : ParseDetailAmount(deposit);
                if ((withdrawalCode || depositCode || balanceCode) && withdrawalAmount == null && depositAmount == null) continue;

                var tx = new DetailTransaction
                {
                    Date = ParseDetailDate(cells.Get(row, detail.DateColIndex), date1904, eraMap),
                    Summary = summary.Trim(),
                    Withdrawal = withdrawalAmount,
                    Deposit = depositAmount,
                    BalanceMasked = balance != null && balance.Text.StartsWith("*"),
                    SheetRow = row
                };
                if (!tx.BalanceMasked && !balanceCode) tx.Balance = ParseDetailAmount(balance);
                result.Add(tx);
            }
        }
        return result;
    }

    private static bool IsEmptyCell(SheetCell cell)
    {
        return cell == null || cell.Text == "";
    }

    private static bool IsSentinelCell(SheetCell cell)
    {
        return cell != null && IsSentinelBalance(cell.Text);
    }

    private static string RemoveWhitespace(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        var sb = new StringBuilder(value.Length);
        foreach (var c in value)
            if (!char.IsWhiteSpace(c)) sb.Append(c);
        return sb.ToString();
    }

    // 取引日の解釈: 数値セルは OADate（1904年始まりのブックは 1462 日ずれる）・文字列は台帳の日付と同じ解釈。
    // 年は 1900〜2099 のみ受け取る（スイートの日付の受理範囲。範囲外の値で期間の計算が日付の上限を越えないように）
    public static DateTime? ParseDetailDate(SheetCell cell, bool date1904, Dictionary<int, EraEntry> eraMap)
    {
        if (cell == null || cell.Text == "") return null;
        DateTime? date;
        if (cell.IsNumeric)
        {
            double serial;
            if (!double.TryParse(cell.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out serial)) return null;
            if (date1904) serial += 1462;
            if (serial < 1 || serial >= 2958466) return null;   // Excel の日付範囲外
            try { date = DateTime.FromOADate(serial); }
            catch (ArgumentException) { return null; }
        }
        else date = ParseLedgerDate(cell.Text, eraMap);
        if (date == null || date.Value.Year < MIN_YEAR || date.Value.Year > MAX_YEAR) return null;
        return date;
    }

    // 金額の解釈: 桁区切り（半角・全角）と空白を除き、全角数字・全角マイナスを半角にして整数へ。
    // 数値セルは指数表記も受理する。解釈できなければ null
    public static long? ParseDetailAmount(SheetCell cell)
    {
        if (cell == null || cell.Text == "") return null;
        var sb = new StringBuilder(cell.Text.Length);
        foreach (var c in cell.Text)
        {
            if (c == ',' || c == '，' || char.IsWhiteSpace(c)) continue;
            if (c >= '０' && c <= '９') sb.Append((char)('0' + (c - '０')));
            else if (c == '－' || c == '−' || c == 'ー') sb.Append('-');
            else if (c == '．') sb.Append('.');
            else sb.Append(c);
        }
        decimal value;
        if (!decimal.TryParse(sb.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return null;
        if (value > long.MaxValue || value < long.MinValue) return null;
        return (long)Math.Round(value, MidpointRounding.AwayFromZero);
    }

    // 読み取ったシートのセルから口座1件の明細を取り出す。取引の並びは原本のまま（明細の掲載順が
    // 金融機関の記帳順で、残高の連なりもこの順。年末処理などで日付が前後する行があっても並べ替えない）
    public static DetailAccountData ExtractDetailAccount(SheetCells cells, string accountNumber, string accountType,
        string institutionName, AccountTypeMapping typeMapping, DetailConfig detail, bool date1904, Dictionary<int, EraEntry> eraMap)
    {
        var result = new DetailAccountData { Status = DetailStatus.NoDetailPage };
        var stopRows = FindStopRows(cells, detail);
        var matched = MatchAccountPages(cells, stopRows, detail, accountNumber, accountType, institutionName, typeMapping);
        if (matched.Count == 0) return result;

        result.Transactions = ReadTransactions(cells, stopRows, matched, detail, date1904, eraMap);
        foreach (var tx in result.Transactions)
            if (tx.BalanceMasked) result.MaskedCount++;
        result.Status = result.Transactions.Count == 0 ? DetailStatus.NoTransactions : DetailStatus.Loaded;
        return result;
    }

    // 期待シートを特定して口座1件の明細を読み取る（シートのセルを保持しない一括版・検証用起動が使う）
    public static DetailAccountData ReadDetailAccount(InquiryWorkbook workbook, string expectedSheet, string accountNumber, string accountType,
        string institutionName, AccountTypeMapping typeMapping, DetailConfig detail, Dictionary<int, EraEntry> eraMap)
    {
        var sheet = workbook.FindSheet(expectedSheet);
        if (sheet == null) return new DetailAccountData { SourcePath = workbook.Path, Status = DetailStatus.NoDetailPage };

        var cells = ReadSheetCells(workbook, sheet.Value.Value, GetDetailWantedColumns(detail));
        var result = ExtractDetailAccount(cells, accountNumber, accountType, institutionName, typeMapping, detail, workbook.Date1904, eraMap);
        result.SourcePath = workbook.Path;
        result.SheetName = sheet.Value.Key;
        return result;
    }

    // ---------- 明細を見る: 口座種別の補正 ----------
    // 台帳の口座種別は、JUDGE（judge.ps1 の Resolve-AccountType）が口座種別の補正表
    //（account_type_mapping.json）で補正した表記。明細ページの口座を台帳の行と照合するため、
    // 同じ補正をここで再現する。規則は JUDGE と同一: 金融機関の表を先に引き、金融機関の表が
    // あれば当たらなくても global へは行かない。表のキーは並びの順に見て最初に当たったものを
    // 採る（末尾が * のキーは前方一致・ほかは完全一致。大文字小文字は区別しない）。金融機関名は
    // Convert-ToFullWidth と同じ全角化をしてから比べる（口座種別は全角化しない）

    // 文字列でもオブジェクトでもない JSON の値（数値・真偽値・null・配列）
    private static readonly object JSON_NON_STRING_VALUE = new object();
    private static readonly Regex JsonNumberRegex = new Regex(@"\G-?(0|[1-9][0-9]*)(\.[0-9]+)?([eE][+-]?[0-9]+)?", RegexOptions.Compiled);

    // 補正表（JSON）を解析する。JSON として読めない・トップがオブジェクトでない・同じオブジェクトに
    // 大文字小文字を区別せずに同じキーが2つある・空のキーがある場合は null（補正しない）
    public static AccountTypeMapping ParseAccountTypeMapping(string json)
    {
        if (json == null) return null;
        int pos = 0;
        object root;
        if (!TryReadJsonValue(json, ref pos, out root, 0)) return null;
        SkipJsonSpace(json, ref pos);
        var top = root as List<KeyValuePair<string, object>>;
        if (pos != json.Length || top == null) return null;

        var result = new AccountTypeMapping();
        foreach (var member in top)
        {
            var rules = ParseAccountTypeRules(member.Value);
            if (TextEqualsIgnoreCase(member.Key, "global")) result.Global = rules;
            else result.Institutions.Add(new AccountTypeRuleTable { InstitutionKey = ToFullWidth(member.Key), Rules = rules });
        }
        return result;
    }

    // 金融機関（または global）の表の中身を読む。オブジェクトでなければ中身のない表として扱う
    private static List<AccountTypeRule> ParseAccountTypeRules(object value)
    {
        var rules = new List<AccountTypeRule>();
        var members = value as List<KeyValuePair<string, object>>;
        if (members == null) return rules;
        foreach (var member in members)
        {
            var rule = new AccountTypeRule { Key = member.Key, Value = member.Value as string };
            if (member.Key.EndsWith("*"))
            {
                // 末尾の * をすべて除いた残りで始まるか（残りの中の * ? はワイルドカード・ほかの記号は文字どおり）
                rule.Prefix = new Regex("^" + Regex.Escape(member.Key.TrimEnd('*')).Replace("\\*", ".*").Replace("\\?", "."),
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline);
            }
            rules.Add(rule);
        }
        return rules;
    }

    // 口座種別を補正表で補正する（当たらなければ入力のまま）
    public static string ResolveAccountType(string accountType, string institutionName, AccountTypeMapping mapping)
    {
        accountType = accountType ?? "";
        if (mapping == null) return accountType;
        var institution = ToFullWidth((institutionName ?? "").Trim());
        foreach (var table in mapping.Institutions)
        {
            if (TextEqualsIgnoreCase(table.InstitutionKey, institution))
                return ApplyAccountTypeRules(table.Rules, accountType);   // 金融機関の表があれば global へは行かない
        }
        return mapping.Global != null ? ApplyAccountTypeRules(mapping.Global, accountType) : accountType;
    }

    // 表のキーを並びの順に見て、最初に当たったキーの値を返す（値が文字列でなければ補正しない）
    private static string ApplyAccountTypeRules(List<AccountTypeRule> rules, string accountType)
    {
        foreach (var rule in rules)
        {
            bool hit = rule.Prefix != null ? rule.Prefix.IsMatch(accountType) : TextEqualsIgnoreCase(accountType, rule.Key);
            if (hit) return rule.Value ?? accountType;
        }
        return accountType;
    }

    // 大文字小文字を区別しない比較（JUDGE が使う Windows PowerShell の -eq と同じ比べ方）
    private static bool TextEqualsIgnoreCase(string a, string b)
    {
        return CultureInfo.InvariantCulture.CompareInfo.Compare(a, b, CompareOptions.IgnoreCase) == 0;
    }

    // JUDGE の Convert-ToFullWidth と同じ全角化: 半角の英字・数字を全角に、半角カナを全角カナに
    //（濁点・半濁点の付く2文字は1文字にまとめ、残った単独の ﾞ ﾟ は ゛ ゜ にする）。ほかの文字は変えない
    private const string HALF_WIDTH_KANA  = "ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ";
    private const string FULL_WIDTH_KANA  = "ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜";
    private const string KANA_VOICED      = "ウカキクケコサシスセソタチツテトハヒフヘホ";   // ﾞ が付くと濁音になる（ｳﾞ→ヴ）
    private const string KANA_SEMI_VOICED = "ハヒフヘホ";                   // ﾟ が付くと半濁音になる

    public static string ToFullWidth(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        var sb = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++)
        {
            var c = value[i];
            int kana = HALF_WIDTH_KANA.IndexOf(c);
            if (kana >= 0)
            {
                var full = FULL_WIDTH_KANA[kana];
                if (i + 1 < value.Length && value[i + 1] == 'ﾞ' && KANA_VOICED.IndexOf(full) >= 0)
                {
                    sb.Append(full == 'ウ' ? 'ヴ' : (char)(full + 1));
                    i++;
                }
                else if (i + 1 < value.Length && value[i + 1] == 'ﾟ' && KANA_SEMI_VOICED.IndexOf(full) >= 0)
                {
                    sb.Append((char)(full + 2));
                    i++;
                }
                else sb.Append(full);
            }
            else if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9')) sb.Append((char)(c + 0xFEE0));
            else sb.Append(c);
        }
        return sb.ToString();
    }

    // JSON の値を1つ読む（標準の文法どおり）。オブジェクトはキーと値の組の並び
    //（List<KeyValuePair<string, object>>・並びを保つ）、文字列は string、それ以外は JSON_NON_STRING_VALUE
    private static bool TryReadJsonValue(string json, ref int pos, out object value, int depth)
    {
        value = null;
        if (depth > 32) return false;   // 入れ子が深すぎるファイルは読まない（補正表は2段）
        SkipJsonSpace(json, ref pos);
        if (pos >= json.Length) return false;
        var c = json[pos];
        if (c == '{')
        {
            pos++;
            var members = new List<KeyValuePair<string, object>>();
            SkipJsonSpace(json, ref pos);
            if (pos < json.Length && json[pos] == '}') { pos++; value = members; return true; }
            while (true)
            {
                SkipJsonSpace(json, ref pos);
                string key;
                if (!TryReadJsonString(json, ref pos, out key) || key == "") return false;   // 空のキーは読めない扱い
                foreach (var member in members)
                    if (string.Equals(member.Key, key, StringComparison.OrdinalIgnoreCase)) return false;   // 大文字小文字だけが違うキーも重複として読まない
                SkipJsonSpace(json, ref pos);
                if (pos >= json.Length || json[pos] != ':') return false;
                pos++;
                object item;
                if (!TryReadJsonValue(json, ref pos, out item, depth + 1)) return false;
                members.Add(new KeyValuePair<string, object>(key, item));
                SkipJsonSpace(json, ref pos);
                if (pos < json.Length && json[pos] == ',') { pos++; continue; }
                if (pos < json.Length && json[pos] == '}') { pos++; value = members; return true; }
                return false;
            }
        }
        if (c == '[')
        {
            pos++;
            SkipJsonSpace(json, ref pos);
            if (pos < json.Length && json[pos] == ']') { pos++; value = JSON_NON_STRING_VALUE; return true; }
            while (true)
            {
                object item;
                if (!TryReadJsonValue(json, ref pos, out item, depth + 1)) return false;
                SkipJsonSpace(json, ref pos);
                if (pos < json.Length && json[pos] == ',') { pos++; continue; }
                if (pos < json.Length && json[pos] == ']') { pos++; value = JSON_NON_STRING_VALUE; return true; }
                return false;
            }
        }
        if (c == '"')
        {
            string text;
            if (!TryReadJsonString(json, ref pos, out text)) return false;
            value = text;
            return true;
        }
        foreach (var literal in new[] { "true", "false", "null" })
        {
            if (string.CompareOrdinal(json, pos, literal, 0, literal.Length) == 0)
            {
                pos += literal.Length;
                value = JSON_NON_STRING_VALUE;
                return true;
            }
        }
        var number = JsonNumberRegex.Match(json, pos);
        if (!number.Success || number.Length == 0) return false;
        pos += number.Length;
        value = JSON_NON_STRING_VALUE;
        return true;
    }

    // JSON の文字列を1つ読む（エスケープを解く。制御文字・不正なエスケープは読めない）
    private static bool TryReadJsonString(string json, ref int pos, out string text)
    {
        text = null;
        if (pos >= json.Length || json[pos] != '"') return false;
        var sb = new StringBuilder();
        for (int i = pos + 1; i < json.Length; i++)
        {
            var c = json[i];
            if (c == '"') { pos = i + 1; text = sb.ToString(); return true; }
            if (c < ' ') return false;
            if (c != '\\') { sb.Append(c); continue; }
            if (++i >= json.Length) return false;
            switch (json[i])
            {
                case '"':  sb.Append('"'); break;
                case '\\': sb.Append('\\'); break;
                case '/':  sb.Append('/'); break;
                case 'b':  sb.Append('\b'); break;
                case 'f':  sb.Append('\f'); break;
                case 'n':  sb.Append('\n'); break;
                case 'r':  sb.Append('\r'); break;
                case 't':  sb.Append('\t'); break;
                case 'u':
                    int code;
                    if (i + 4 >= json.Length ||
                        !int.TryParse(json.Substring(i + 1, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out code)) return false;
                    sb.Append((char)code);
                    i += 4;
                    break;
                default: return false;
            }
        }
        return false;
    }

    private static void SkipJsonSpace(string json, ref int pos)
    {
        while (pos < json.Length && (json[pos] == ' ' || json[pos] == '\t' || json[pos] == '\n' || json[pos] == '\r')) pos++;
    }

    // ---------- 明細を見る: 推定残高・カテゴリ付け・集約 ----------
    // いずれも原本の値と台帳の検出から算術のみで導く（評価はしない）

    private static readonly string[] DETAIL_CATEGORY_NAMES = { "年金", "給与", "保険", "定期入金", "定期出金" };
    private static readonly int[] DETAIL_CATEGORY_COLUMNS = { DEPOSIT_COL_PENSION, DEPOSIT_COL_SALARY, DEPOSIT_COL_INSURANCE, DEPOSIT_COL_REG_DEPOSIT, DEPOSIT_COL_REG_WITHDRAW };

    // 台帳の行から検出（カテゴリ・名称）を台帳の記載順に取り出す。突合用の正規化名が空になる件は除く
    public static List<DetailDetection> BuildDetailDetections(string[] fields)
    {
        var result = new List<DetailDetection>();
        for (int i = 0; i < DETAIL_CATEGORY_NAMES.Length; i++)
        {
            var raw = fields[DETAIL_CATEGORY_COLUMNS[i]];
            if (string.IsNullOrWhiteSpace(raw)) continue;
            foreach (var entry in SplitDetectionEntries(raw.Trim()))
            {
                var name = ParseDetectionName(entry);
                var normalized = NormalizeForSearch(name);
                if (normalized == "") continue;
                result.Add(new DetailDetection
                {
                    Category = DETAIL_CATEGORY_NAMES[i],
                    Name = name,
                    Bucket = ParseDetectionBucket(entry),
                    NormalizedName = normalized,
                    IsDeposit = i != 2 && i != 4   // 保険・定期出金は出金側
                });
            }
        }
        return result;
    }

    // 各取引に検出を当てる: 摘要と検出名称の両方を検索用正規化にかけ、名称を含む最初の検出
    //（台帳の記載順）を採用する。入出金の向きも一致条件（JUDGE の判定と同じ向き）
    public static void CategorizeTransactions(List<DetailTransaction> transactions, List<DetailDetection> detections)
    {
        foreach (var tx in transactions)
        {
            tx.Category = null;
            tx.DetectionName = null;
            if (detections.Count == 0) continue;
            bool isDeposit = tx.Deposit > 0, isWithdrawal = tx.Withdrawal > 0;
            if (!isDeposit && !isWithdrawal) continue;
            var summary = NormalizeForSearch(tx.Summary);
            foreach (var detection in detections)
            {
                if (!(detection.IsDeposit ? isDeposit : isWithdrawal)) continue;
                if (summary.Contains(detection.NormalizedName))
                {
                    tx.Category = detection.Category;
                    tx.DetectionName = detection.Name;
                    break;
                }
            }
        }
    }

    // 残高の無い行を入出金の積み上げで埋める。既知の残高を起点に前へ積み上げ、次の既知の残高で
    // 突き合わせて差額を残し、既知の値で積み直す。先頭に既知がない区間は最初の既知から逆算する。
    // 既知が1つもなければ推定しない。戻り値は差額のある点の数
    public static int ReconstructBalances(List<DetailTransaction> transactions)
    {
        int gapCount = 0;
        long? running = null;
        foreach (var tx in transactions)
        {
            tx.EstimatedBalance = null;
            tx.Gap = null;
            long? carried = running == null ? null : running - (tx.Withdrawal ?? 0) + (tx.Deposit ?? 0);
            if (tx.Balance != null)
            {
                if (carried != null && carried != tx.Balance)
                {
                    tx.Gap = tx.Balance - carried;
                    gapCount++;
                }
                running = tx.Balance;
            }
            else
            {
                tx.EstimatedBalance = carried;
                running = carried;
            }
        }

        // 先頭区間の逆算: 最初の既知の残高から、直前の行の残高 = 次の行の残高 + 次の行の出金 − 次の行の入金
        int firstKnown = transactions.FindIndex(t => t.Balance != null);
        if (firstKnown > 0)
        {
            long after = transactions[firstKnown].Balance.Value;
            for (int i = firstKnown - 1; i >= 0; i--)
            {
                var next = transactions[i + 1];
                after = after + (next.Withdrawal ?? 0) - (next.Deposit ?? 0);
                transactions[i].EstimatedBalance = after;
            }
        }
        return gapCount;
    }

    // 検出ごとの回数・1回あたり・直近を集める（当たった行が1件もない検出は含めない）。
    // 並びは合計の大きい順（同額は台帳の記載順）
    public static List<DetectionAggregate> AggregateDetections(List<DetailTransaction> transactions, List<DetailDetection> detections)
    {
        var result = new List<DetectionAggregate>();
        foreach (var detection in detections)
        {
            DetectionAggregate aggregate = null;
            foreach (var tx in transactions)
            {
                if (tx.Category != detection.Category || tx.DetectionName != detection.Name) continue;
                if (aggregate == null)
                {
                    aggregate = new DetectionAggregate { Category = detection.Category, Name = detection.Name };
                    result.Add(aggregate);
                }
                aggregate.Count++;
                aggregate.Total += detection.IsDeposit ? (tx.Deposit ?? 0) : -(tx.Withdrawal ?? 0);
                if (tx.Date != null && (aggregate.LastDate == null || tx.Date > aggregate.LastDate)) aggregate.LastDate = tx.Date;
            }
        }
        foreach (var aggregate in result)
            aggregate.PerAmount = (long)Math.Round(aggregate.Total / (decimal)aggregate.Count, MidpointRounding.AwayFromZero);
        return result.OrderByDescending(a => Math.Abs(a.Total)).ToList();   // OrderBy は安定ソート
    }

    // 取引日の範囲: 最初の取引日と最後の取引日（月初・月末には丸めない。横軸はこの範囲を
    // 描画域いっぱいに使う）。取引日のない行は無視し、全行になければ null
    public static void DeriveDetailPeriod(List<DetailTransaction> transactions, out DateTime? periodFrom, out DateTime? periodTo)
    {
        periodFrom = null;
        periodTo = null;
        foreach (var tx in transactions)
        {
            if (tx.Date == null) continue;
            if (periodFrom == null || tx.Date < periodFrom) periodFrom = tx.Date;
            if (periodTo == null || tx.Date > periodTo) periodTo = tx.Date;
        }
    }

    // 明細ウィンドウの絞り込み: 向き（検出のみ／入金／出金／すべて）と摘要の部分一致（検索用正規化）
    public enum DetailFilterMode { Hits, Deposits, Withdrawals, All }

    public static List<DetailTransaction> FilterDetailTransactions(List<DetailTransaction> transactions, DetailFilterMode mode, string query)
    {
        var normalizedQuery = NormalizeForSearch(query);
        var result = new List<DetailTransaction>();
        foreach (var tx in transactions)
        {
            switch (mode)
            {
                case DetailFilterMode.Hits: if (tx.Category == null) continue; break;
                case DetailFilterMode.Deposits: if (!(tx.Deposit > 0)) continue; break;
                case DetailFilterMode.Withdrawals: if (!(tx.Withdrawal > 0)) continue; break;
            }
            if (normalizedQuery != "" && !NormalizeForSearch(tx.Summary).Contains(normalizedQuery)) continue;
            result.Add(tx);
        }
        return result;
    }

    // 読み取った明細に検出を当て、推定残高・集約・期間を確定する
    public static void AnalyzeDetailAccount(DetailAccountData data, List<DetailDetection> detections)
    {
        CategorizeTransactions(data.Transactions, detections);
        data.GapCount = ReconstructBalances(data.Transactions);
        data.Aggregates = AggregateDetections(data.Transactions, detections);
        DateTime? periodFrom, periodTo;
        DeriveDetailPeriod(data.Transactions, out periodFrom, out periodTo);
        data.PeriodFrom = periodFrom;
        data.PeriodTo = periodTo;
    }
}