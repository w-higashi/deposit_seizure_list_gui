// ==============================================================================
// paddock_recent_view.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 最近見た人ビュー（ナビ「最近見た人」）と履歴の読み書き。
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

public partial class PaddockApp
{
    // ==============================================================
    // 最近見た人ビュー（ナビ「最近見た人」）
    // ==============================================================
    //
    // 照会ファイルを開いた・差押リストに追加した行の履歴をユーザー単位で永続化し、
    // 年度（行の登録日基準。台帳ファイルの年度帰属と同じ規則）で区切って表示する。
    // 履歴のクリックで当該人物の詳細パネル（年度横断ビュー）を開く。
    // 履歴に持つ列は一覧の行・行キー・年度の導出に要る分だけ（BusinessLogic.RECENT_KEPT_*_COLUMNS）。
    // 届出住所・生年月日などは履歴に書かない

    // 履歴ファイルのパス（history/{Windowsユーザー名}/recent_people.json）
    private string GetRecentPeoplePath()
    {
        return System.IO.Path.Combine(userHistDir, "recent_people.json");
    }

    // 起動時に履歴ファイルを読み込む。
    // ファイル全体が読めない・構造が読めない（entries がない・括弧が閉じない＝書いている途中で
    // 切れたなど）場合はセッション内メモリのみの動作に退行し、以降の保存も試行しない
    //（壊れたファイルへの上書きで状況を悪化させない）。空のファイルは「ない」のと同じ
    private void LoadRecentPeople()
    {
        var path = GetRecentPeoplePath();
        if (!File.Exists(path)) return;

        try
        {
            var json = File.ReadAllText(path, Encoding.UTF8);
            if (json.Trim().Length == 0) return;
            var entriesIdx = json.IndexOf("\"entries\"");
            if (entriesIdx < 0) throw new InvalidDataException("entries の構造が読めません");
            var arrStart = json.IndexOf('[', entriesIdx);
            if (arrStart < 0) throw new InvalidDataException("entries の構造が読めません");
            var arrEnd = JsonHelper.FindMatchingBracket(json, arrStart, '[', ']');
            if (arrEnd < 0) throw new InvalidDataException("entries の構造が読めません");

            // fields の値は台帳由来の任意文字列（波括弧を含みうる）のため、エントリの
            // 区切りは文字列リテラルを解する FindMatchingBracket で取る
            //（差押登録ログの IndexOf('}') 方式は値が単純なログ専用）
            int pos = arrStart + 1;
            int skipped = 0;
            while (pos < arrEnd)
            {
                var objStart = json.IndexOf('{', pos);
                if (objStart < 0 || objStart >= arrEnd) break;
                var objEnd = JsonHelper.FindMatchingBracket(json, objStart, '{', '}');
                if (objEnd < 0) break;

                var entry = ParseRecentEntry(json.Substring(objStart, objEnd - objStart + 1));
                if (entry != null) recentPeople.Add(entry);   // 形式不正はそのエントリのみ読み飛ばす（件数はフッターとログに）
                else skipped++;
                pos = objEnd + 1;
            }
            recentSkippedCount = skipped;
            if (skipped > 0) WriteErrorLog("履歴の読み込み", "読み飛ばし " + skipped + " 件");
        }
        catch (Exception ex)
        {
            recentPeople.Clear();
            recentSkippedCount = 0;   // 構造が読めないときは件数に意味がない
            persistRecent = false;
            WriteErrorLog("履歴の読み込み", ex);
        }
    }

    // 履歴エントリ1件を解析する（列の値が読めない・宛名番号が空の場合は null）。
    // 種別は列の値より前の部分にある kind で決める（預金のエントリは kind を持たない。列の値は
    // 台帳由来の任意文字列のため、kind・year・viewedAt は値より前に書き、先頭の部分から取る）
    private RecentPersonEntry ParseRecentEntry(string objJson)
    {
        var fieldsIdx = objJson.IndexOf("\"fields\"");
        var columnsIdx = objJson.IndexOf("\"columns\"");
        var headEnd = objJson.Length;
        if (fieldsIdx >= 0 && fieldsIdx < headEnd) headEnd = fieldsIdx;
        if (columnsIdx >= 0 && columnsIdx < headEnd) headEnd = columnsIdx;
        bool insurance = JsonHelper.GetString(objJson.Substring(0, headEnd), "kind") == KIND_INSURANCE;

        var fields = insurance ? ParseRecentColumns(objJson, columnsIdx) : ParseRecentFields(objJson, fieldsIdx);
        if (fields == null) return null;

        var entry = new RecentPersonEntry
        {
            IsInsurance = insurance,
            Fields = BusinessLogic.TrimRecentFields(insurance, fields),   // 旧い形（全列）のエントリもここで最小の列になり、次の保存で新しい形へ
            Year = JsonHelper.GetInt(objJson, "year", -1),
            ViewedAtRaw = JsonHelper.GetString(objJson, "viewedAt") ?? ""
        };
        DateTime viewedAt;
        entry.ViewedAt = DateTime.TryParse(entry.ViewedAtRaw, CultureInfo.InvariantCulture, DateTimeStyles.None, out viewedAt)
            ? viewedAt : DateTime.MinValue;
        return entry;
    }

    // 預金のエントリの fields（台帳の位置の配列）を読む（最低列数に満たない・宛名番号が空の場合は null。
    // 27列時代に記録したエントリと現行28列のエントリの両方を受ける）
    private static string[] ParseRecentFields(string objJson, int fieldsIdx)
    {
        if (fieldsIdx < 0) return null;
        var arrStart = objJson.IndexOf('[', fieldsIdx);
        if (arrStart < 0) return null;
        var arrEnd = JsonHelper.FindMatchingBracket(objJson, arrStart, '[', ']');
        if (arrEnd < 0) return null;

        var fields = new List<string>();
        var arrJson = objJson.Substring(arrStart, arrEnd - arrStart + 1);
        foreach (Match m in Regex.Matches(arrJson, "\"((?:[^\"\\\\]|\\\\.)*)\""))
            fields.Add(JsonHelper.UnescapeJsonString(m.Groups[1].Value));
        if (fields.Count < BusinessLogic.DEPOSIT_LEDGER_COLUMN_COUNT) return null;
        if (fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM] == "") return null;
        return fields.ToArray();
    }

    // 生命保険のエントリの columns（列名→値。空の列は書かれていない）を台帳の並びの配列に組み直して読む
    //（知らない列名は捨て、ない列は空。宛名番号が空の場合は null）
    private static string[] ParseRecentColumns(string objJson, int columnsIdx)
    {
        if (columnsIdx < 0) return null;
        var objStart = objJson.IndexOf('{', columnsIdx);
        if (objStart < 0) return null;
        var objEnd = JsonHelper.FindMatchingBracket(objJson, objStart, '{', '}');
        if (objEnd < 0) return null;

        // 文字列の並びを列名・値の組として読む（列名も値も fields と同じエスケープ）
        var values = new Dictionary<string, string>();
        var pairJson = objJson.Substring(objStart, objEnd - objStart + 1);
        string key = null;
        foreach (Match m in Regex.Matches(pairJson, "\"((?:[^\"\\\\]|\\\\.)*)\""))
        {
            var s = JsonHelper.UnescapeJsonString(m.Groups[1].Value);
            if (key == null) key = s;
            else { values[key] = s; key = null; }
        }

        var fields = new string[BusinessLogic.INSURANCE_LEDGER_COLUMN_COUNT];
        for (int i = 0; i < fields.Length; i++)
        {
            string value;
            fields[i] = values.TryGetValue(BusinessLogic.INSURANCE_LEDGER_COLUMNS[i], out value) ? value : "";
        }
        if (fields[BusinessLogic.INSURANCE_COL_ADDRESS_NUM] == "") return null;
        return fields;
    }

    // 履歴を保存する（記録の都度呼び出す —
    // ファイルは小さく、異常終了しても直前までの履歴が残る）
    private void SaveRecentPeople()
    {
        if (!persistRecent) return;

        try
        {
            var historyDir = System.IO.Path.GetDirectoryName(GetRecentPeoplePath());
            if (!Directory.Exists(historyDir))
                Directory.CreateDirectory(historyDir);

            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine("  \"entries\": [");
            for (int i = 0; i < recentPeople.Count; i++)
            {
                var entry = recentPeople[i];
                var viewedAt = entry.ViewedAt == DateTime.MinValue
                    ? entry.ViewedAtRaw : entry.ViewedAt.ToString("o");
                sb.Append("    { \"viewedAt\": \"" + EscapeJson(viewedAt) + "\"");
                sb.Append(", \"year\": " + entry.Year);
                if (entry.IsInsurance)
                {
                    // 生命保険の行は列名→値（空の列は書かない）。kind は columns より前に置く（読む側は先頭の部分から取る）
                    sb.Append(", \"kind\": \"" + KIND_INSURANCE + "\", \"columns\": {");
                    bool first = true;
                    for (int j = 0; j < entry.Fields.Length; j++)
                    {
                        if (entry.Fields[j] == "") continue;
                        if (!first) sb.Append(", ");
                        first = false;
                        sb.Append("\"" + EscapeJson(BusinessLogic.INSURANCE_LEDGER_COLUMNS[j]) + "\": \"" + EscapeJson(entry.Fields[j]) + "\"");
                    }
                    sb.AppendLine(i < recentPeople.Count - 1 ? "} }," : "} }");
                }
                else
                {
                    // 預金の行は台帳の位置の配列（種別は書かない）
                    sb.Append(", \"fields\": [");
                    for (int j = 0; j < entry.Fields.Length; j++)
                    {
                        if (j > 0) sb.Append(", ");
                        sb.Append("\"" + EscapeJson(entry.Fields[j]) + "\"");
                    }
                    sb.AppendLine(i < recentPeople.Count - 1 ? "] }," : "] }");
                }
            }
            sb.AppendLine("  ]");
            sb.AppendLine("}");

            WriteUserFileAtomic(GetRecentPeoplePath(), sb.ToString());
        }
        catch
        {
            persistRecent = false;   // 書き込み失敗時は以降の保存をスキップ
        }
    }

    // JSON 文字列値のエスケープ
    private static string EscapeJson(string s)
    {
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"")
                .Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
    }

    // 照会ファイルを開いた・差押リストに追加した預金の行を履歴に記録する
    //（終端アクションの成立時のみ呼ばれる。口座なし行はファイル解決のガードで
    //  弾かれるため、履歴には口座あり行しか入らない）
    private void RecordRecentDepositPerson(string[] fields)
    {
        if (fields == null || fields.Length < BusinessLogic.DEPOSIT_LEDGER_COLUMN_COUNT) return;

        // 年度は行の登録日から導出する（台帳ファイルの年度帰属・ドリルダウンの
        // 検索先と同じ規則。導出できない行は「年度不明」= -1）
        int year = BusinessLogic.GetRegistrationFiscalYear(fields[BusinessLogic.DEPOSIT_COL_REG_DATE], eraMapping) ?? -1;

        RecordRecentEntry(false, fields, year, BusinessLogic.GetDepositLedgerRowKey(fields, eraMapping));
    }

    // 照会ファイルを開いた・差押リストに追加した生命保険の行を履歴に記録する
    //（終端アクションの成立時のみ呼ばれる。契約なし行と照会シートのない電子回答不可行は
    //  ファイル解決のガードで弾かれるため、履歴には契約あり行と短いシートを開いた電子回答不可行が入る）
    private void RecordRecentInsurancePerson(string[] fields)
    {
        if (fields == null || fields.Length < BusinessLogic.INSURANCE_LEDGER_COLUMN_COUNT) return;

        // 年度は行の登録日から導出する（預金と同じ規則）
        int year = BusinessLogic.GetRegistrationFiscalYear(fields[BusinessLogic.INSURANCE_COL_REG_DATE], eraMapping) ?? -1;

        RecordRecentEntry(true, fields, year, BusinessLogic.GetInsuranceLedgerRowKey(fields, eraMapping));
    }

    // 履歴に1行を記録する（預金・生命保険で共通の後段）。
    // 同一の行（同じ種別で同じ一意キー）の再記録は閲覧日時の更新のみとする
    //（表示はセクション内で閲覧日時の降順のため、先頭に浮上する）
    private void RecordRecentEntry(bool insurance, string[] fields, int year, string key)
    {
        RecentPersonEntry entry = null;
        foreach (var e in recentPeople)
        {
            if (e.IsInsurance == insurance && GetRecentEntryKey(e) == key) { entry = e; break; }
        }
        if (entry == null)
        {
            entry = new RecentPersonEntry();
            recentPeople.Add(entry);
        }
        entry.IsInsurance = insurance;
        entry.Fields = BusinessLogic.TrimRecentFields(insurance, fields);   // 履歴に残す列だけの写し（台帳の行の配列は共有されているため、その場で空にしない）
        entry.Year = year;
        entry.ViewedAt = DateTime.Now;
        entry.ViewedAtRaw = null;

        // 年度ごとの保持上限（config の recentMax）を超えた分は閲覧日時の古いものから捨てる
        var sameYear = recentPeople.Where(x => x.Year == year)
                                   .OrderBy(x => x.ViewedAt).ToList();
        for (int i = 0; i < sameYear.Count - config.RecentMax; i++)
            recentPeople.Remove(sameYear[i]);

        SaveRecentPeople();
    }

    // 履歴エントリの一意キー（種別ごとの行キー。種別が違う行どうしは比べない）
    private string GetRecentEntryKey(RecentPersonEntry e)
    {
        return e.IsInsurance ? BusinessLogic.GetInsuranceLedgerRowKey(e.Fields, eraMapping)
                             : BusinessLogic.GetDepositLedgerRowKey(e.Fields, eraMapping);
    }

    // 最近見た人ビューの中身（年度セクションの並び）を組み立てる
    //（差押登録ビューの BuildSeizureLogView と同一構造）
    private void BuildRecentView()
    {
        recentSectionsPanel.Children.Clear();
        recentSelectedBorder = null;

        if (recentPeople.Count == 0)
        {
            recentCard.Visibility = Visibility.Collapsed;
            recentEmptyText.Text = "履歴はまだありません。\n照会ファイルを開く・差押リストに追加すると、ここに表示されます。";
            recentEmptyText.Visibility = Visibility.Visible;
            SetStatusText(recentFooterText, RecentCountText());
            return;
        }
        recentCard.Visibility = Visibility.Visible;
        recentEmptyText.Visibility = Visibility.Collapsed;

        // 年度（行の登録日基準）でグルーピングする。導出できなかったエントリは
        // 「年度不明」として末尾に置く（隠さない）
        var groups = new Dictionary<string, List<RecentPersonEntry>>();
        foreach (var entry in recentPeople)
        {
            var key = entry.Year < 0 ? "年度不明" : BusinessLogic.FormatYearKey(entry.Year);
            if (!groups.ContainsKey(key)) groups[key] = new List<RecentPersonEntry>();
            groups[key].Add(entry);
        }

        // 年度の降順（「年度不明」は末尾）。セクション内は閲覧日時の降順
        var orderedKeys = groups.Keys
            .OrderBy(k => k == "年度不明" ? 1 : 0)
            .ThenByDescending(k => k, StringComparer.Ordinal)
            .ToList();
        var currentYearKey = GetCurrentYearKey();

        foreach (var key in orderedKeys)
        {
            var entries = groups[key].OrderByDescending(e => e.ViewedAt).ToList();
            recentSectionsPanel.Children.Add(
                BuildYearSection(key, entries.Count, key == currentYearKey, delegate(StackPanel body)
                {
                    body.Children.Add(BuildRecentHeaderRow());
                    foreach (var entry in entries)
                        body.Children.Add(BuildRecentEntryRow(entry));
                }));
        }

        SetStatusText(recentFooterText, RecentCountText() + "　·　" + GetRecentPeoplePath());
    }

    // フッターの件数「N 件の履歴」。起動時に読み飛ばした履歴があれば「（読めない M 件）」を添える（読み飛ばしを隠さない）
    private string RecentCountText()
    {
        return recentPeople.Count.ToString("N0") + " 件の履歴" +
               (recentSkippedCount > 0 ? "（読めない " + recentSkippedCount.ToString("N0") + " 件）" : "");
    }

    // 最近見た人ビューの列構成（ヘッダー行と履歴行で共有する）
    // 台帳一覧の11列から カナ氏名・最終取引日 を除き、先頭に閲覧日時と種別を加えた11列。
    // 生命保険の行は意味の近い列に入れる（金融機関・保険会社／口座種別・契約種類／口座番号・証券番号／
    // 残高・解約返戻金／検出・注記）。列幅は台帳一覧と同じ固定値+末尾の余白列（Star 列が画面の余りを
    // 吸収して間延びしないように）。残高と検出の間には24pxのスペーサー列を挟む
    //（台帳一覧では間の最終取引日列が緩衝になっている隣接を補う）
    private static Grid CreateRecentRowGrid()
    {
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(92) });                       // 0: 閲覧日時
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(56) });                       // 1: 種別
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(104) });                      // 2: 回答基準日
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(104) });                      // 3: 宛名番号
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });                      // 4: 氏名
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });                      // 5: 金融機関・保険会社
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });                      // 6: 支店名
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });                      // 7: 口座種別・契約種類
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });                      // 8: 口座番号・証券番号
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(104) });                      // 9: 残高・解約返戻金
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(24) });                       // 10: スペーサー
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(330) });                      // 11: 検出・注記
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });     // 余白
        return grid;
    }

    // 列ヘッダー行（一覧テーブルのヘッダーと同じ配色・サイズ）
    private UIElement BuildRecentHeaderRow()
    {
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = BrushLine,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(14, 7, 14, 7)
        };
        var grid = CreateRecentRowGrid();
        var labels = new[] { "閲覧日時", "種別", "回答基準日", "宛名番号", "氏名", "金融機関・保険会社", "支店名", "口座種別・契約種類", "口座番号・証券番号", "残高・解約返戻金", "検出・注記" };
        for (int i = 0; i < labels.Length; i++)
        {
            var label = new TextBlock { Text = labels[i], FontSize = 11, Foreground = BrushTextMuted };
            if (labels[i] == "残高・解約返戻金")
            {
                // 残高の見出しのみ中央揃え（一覧のヘッダーと同じ。データは右揃えのまま。
                // 右余白 8px はデータセル側と同じに保ち、データ域の中央に一致させる）
                label.HorizontalAlignment = HorizontalAlignment.Center;
                label.Margin = new Thickness(0, 0, 8, 0);
            }
            Grid.SetColumn(label, labels[i] == "検出・注記" ? i + 1 : i);   // スペーサー列(10)を飛ばす
            grid.Children.Add(label);
        }
        border.Child = grid;
        return border;
    }

    // 履歴1件の行。預金の行は台帳一覧と同じ表示規約（元号変換・「対象外」置換・検出チップ+
    // 暗色ツールチップ）を BuildDepositRecord の再利用で、生命保険の行は生命保険の一覧と同じ規約
    //（注記チップ・薄文字色）を BuildInsuranceRecord の再利用で再現する。クリックで詳細パネルを開く
    private UIElement BuildRecentEntryRow(RecentPersonEntry entry)
    {
        // 外側=下罫線、内側=左3pxの選択バー+ホバー/選択背景（一覧の行と同じ見え方）。
        // 行高は固定32px（台帳一覧と統一。チップ有無で行高が揺れない）
        var outer = new Border
        {
            Background = Brushes.White,
            BorderBrush = BrushLineWeak,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Height = 32
        };
        var inner = new Border
        {
            Background = Brushes.Transparent,
            BorderBrush = Brushes.Transparent,
            BorderThickness = new Thickness(3, 0, 0, 0),
            Padding = new Thickness(11, 0, 14, 0),
            Cursor = Cursors.Hand,
            ClipToBounds = true
        };
        var grid = CreateRecentRowGrid();
        grid.VerticalAlignment = VerticalAlignment.Center;

        // 閲覧日時: 年はセクション見出しにあるため MM/dd HH:mm。読めない日時は生値のまま
        var viewedAt = entry.ViewedAt == DateTime.MinValue
            ? entry.ViewedAtRaw : entry.ViewedAt.ToString("MM/dd HH:mm", CultureInfo.InvariantCulture);
        AddRowCell(grid, 0, viewedAt, true, false);
        AddKindChipCell(grid, 1, entry.IsInsurance);
        if (entry.IsInsurance) AddInsuranceRecentCells(grid, entry.Fields);
        else AddDepositRecentCells(grid, entry.Fields);

        inner.Child = grid;
        outer.Child = inner;

        // ホバー表示（選択中の行はそのまま）
        inner.MouseEnter += delegate
        {
            if (!ReferenceEquals(inner, recentSelectedBorder)) inner.Background = BrushRowHover;
        };
        inner.MouseLeave += delegate
        {
            if (!ReferenceEquals(inner, recentSelectedBorder)) inner.Background = Brushes.Transparent;
        };

        // クリックで詳細パネルを開く（このビューに留まったまま右側にスライドイン）。
        // キー（Enter／Space）も同じだが、開いている人の行では開いている方のパネルのカードへ移る（一覧の Space と同じ）
        inner.MouseLeftButtonUp += delegate { OpenRecentEntry(entry, inner); };
        MakeKeyActivatable(inner, delegate
        {
            if (isDepositPanelOpen && ReferenceEquals(recentSelectedBorder, inner)) FocusDepositPanelCard();
            else if (insuranceTab.isInsurancePanelOpen && ReferenceEquals(recentSelectedBorder, inner)) insuranceTab.FocusInsurancePanelCard();
            else OpenRecentEntry(entry, inner);
        }, keyFocusInsetStyle);

        return outer;
    }

    // 預金の行のセル（回答基準日〜検出。列番号は CreateRecentRowGrid の並び）
    private void AddDepositRecentCells(Grid grid, string[] fields)
    {
        var record = BuildDepositRecord(fields);
        AddRowCell(grid, 2, record.DisplayBaseDate, true, false);
        AddRowCell(grid, 3, record.AddressNumber, true, false);
        AddNameCellWithNoteMark(grid, 4, record.Name, record.HasNotes, record.NotesToolTip, BrushTextMain);
        AddRowCell(grid, 5, record.InstitutionName, false, true);
        AddRowCell(grid, 6, record.BranchName, false, true);
        AddRowCell(grid, 7, record.AccountType, false, true);
        AddRowCell(grid, 8, record.AccountNumber, true, false);

        // 残高（右寄せ・センチネル値は「対象外」+生値ツールチップ。一覧と同じ）
        var balance = new TextBlock
        {
            Text = record.DisplayBalance,
            FontSize = 12,
            FontFamily = FontMono,
            Foreground = BrushTextMain,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 0, 8, 0),
            VerticalAlignment = VerticalAlignment.Center,
            ToolTip = record.BalanceToolTip
        };
        Grid.SetColumn(balance, 9);
        grid.Children.Add(balance);

        // 検出チップ（第1層）+ 暗色ツールチップ（第2層）。一覧の検出列と同じ2層方式
        var chips = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            VerticalAlignment = VerticalAlignment.Center,
            Background = Brushes.Transparent   // 余白へのホバーでもツールチップが出るように
        };
        foreach (var chip in record.Chips)
        {
            var chipBorder = new Border
            {
                Background = chip.Bg,
                CornerRadius = new CornerRadius(9),
                Padding = new Thickness(8, 1, 8, 2),
                Margin = new Thickness(0, 0, 4, 0)
            };
            chipBorder.Child = new TextBlock { Text = chip.Text, FontSize = 11, Foreground = chip.Fg };
            chips.Children.Add(chipBorder);
        }
        if (record.HasDetection)
        {
            var tip = new ToolTip { Style = (Style)window.FindResource("DetectTip") };
            tip.LayoutTransform = displayScaleTransform;   // 表示倍率に追従
            tip.Content = BuildDetectTipContent(record.TooltipLines);
            chips.ToolTip = tip;
            ToolTipService.SetInitialShowDelay(chips, 350);
            ToolTipService.SetShowDuration(chips, 60000);
        }
        Grid.SetColumn(chips, 11);   // 10 はスペーサー列
        grid.Children.Add(chips);
    }

    // 生命保険の行のセル（回答基準日〜注記。値の作り方は生命保険の一覧の行と同じ）。
    // 支店名の列は生命保険にないため空ける。値のない項目はセルを置かない（空文字列を ToolTip に
    // 付けると空のツールチップが出るため。電子回答不可の行は契約種類・証券番号・解約返戻金が空）。
    // 電子回答不可の行は一覧と同じく薄文字色
    private void AddInsuranceRecentCells(Grid grid, string[] fields)
    {
        var record = insuranceTab.BuildInsuranceRecord(fields);
        var text = record.IsMutedRow ? BrushTextMutedRow : BrushTextMain;
        AddRowCell(grid, 2, record.DisplayBaseDate, true, false, text);
        AddRowCell(grid, 3, record.AddressNumber, true, false, text);
        AddNameCellWithNoteMark(grid, 4, record.Name ?? "", record.HasNotes, record.NotesToolTip, text);   // 電子回答不可の行は氏名が空（印は出す）
        if (record.CompanyName != null) AddRowCell(grid, 5, record.CompanyName, false, true, text);
        if (record.ContractType != null) AddRowCell(grid, 7, record.ContractType, false, true, text);
        if (record.PolicyNumber != null) AddRowCell(grid, 8, record.PolicyNumber, true, false, text);

        // 解約返戻金（右寄せ・数値として読めない生値は薄文字色のまま。「対象外」の置き換えはしない）
        if (record.DisplaySurrenderValue != "")
        {
            var surrender = new TextBlock
            {
                Text = record.DisplaySurrenderValue,
                FontSize = 12,
                FontFamily = FontMono,
                Foreground = record.IsSurrenderValueText ? BrushTextMutedRow : text,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(surrender, 9);
            grid.Children.Add(surrender);
        }

        // 注記チップ（第1層・一覧と同じ最大3個+「+N」）+ 暗色ツールチップ（第2層）。
        // 電子回答不可の行は注記の代わりにピル（一覧と同じ）
        var chips = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            VerticalAlignment = VerticalAlignment.Center,
            Background = Brushes.Transparent   // 余白へのホバーでもツールチップが出るように
        };
        if (record.MutedPillText != null)
        {
            var pill = new Border
            {
                Background = BrushChipNeutralBg,
                CornerRadius = new CornerRadius(9),
                Padding = new Thickness(8, 1, 8, 2)
            };
            pill.Child = new TextBlock { Text = record.MutedPillText, FontSize = 11, Foreground = BrushTextMuted };
            chips.Children.Add(pill);
        }
        else
        {
            foreach (var chip in record.Chips)
            {
                var chipBorder = new Border
                {
                    Background = chip.Bg,
                    CornerRadius = new CornerRadius(9),
                    Padding = new Thickness(8, 1, 8, 2),
                    Margin = new Thickness(0, 0, 4, 0)
                };
                chipBorder.Child = new TextBlock { Text = chip.Text, FontSize = 11, Foreground = chip.Fg };
                chips.Children.Add(chipBorder);
            }
            if (record.HasNote)
            {
                var tip = new ToolTip { Style = (Style)window.FindResource("DetectTip") };
                tip.LayoutTransform = displayScaleTransform;   // 表示倍率に追従
                tip.Content = BuildDetectTipContent(record.TooltipLines);
                chips.ToolTip = tip;
                ToolTipService.SetInitialShowDelay(chips, 350);
                ToolTipService.SetShowDuration(chips, 60000);
            }
        }
        Grid.SetColumn(chips, 11);   // 10 はスペーサー列
        grid.Children.Add(chips);
    }

    // 検出セルの暗色ツールチップの中身（カテゴリチップ+原文の行形式）を組み立てる
    //（一覧の検出列の DataTemplate と同じ見え方をコードから構築する）
    private UIElement BuildDetectTipContent(List<TooltipLineInfo> lines)
    {
        var stack = new StackPanel { MaxWidth = 460 };
        foreach (var line in lines)
        {
            var grid = new Grid { Margin = new Thickness(0, 3, 0, 3) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var chip = new Border
            {
                Background = line.ChipBg,
                CornerRadius = new CornerRadius(9),
                Padding = new Thickness(9, 1, 9, 2),
                VerticalAlignment = VerticalAlignment.Top
            };
            chip.Child = new TextBlock { Text = line.ChipText, FontSize = 11, Foreground = line.ChipFg };
            grid.Children.Add(chip);

            var content = new TextBlock
            {
                Text = line.Content,
                FontSize = 12,
                Foreground = BrushTipText,
                FontFamily = line.ContentFont,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(9, 1, 0, 0)
            };
            Grid.SetColumn(content, 1);
            grid.Children.Add(content);

            stack.Children.Add(grid);
        }
        return stack;
    }

    // 履歴行の選択状態を切り替える（選択中の行=詳細パネルに表示中の行）
    private void SelectRecentRow(Border row)
    {
        if (recentSelectedBorder != null)
        {
            recentSelectedBorder.Background = Brushes.Transparent;
            recentSelectedBorder.BorderBrush = Brushes.Transparent;
        }
        recentSelectedBorder = row;
        if (row != null)
        {
            row.Background = BrushSelectedCardBg;
            row.BorderBrush = BrushAccent;
        }
    }

    // 履歴行のクリック: 当該人物の、行の種別の詳細パネルを開く。中身は一覧から開いた場合と
    // 同一（全年度横断ビュー）で、記録した行のカードが選択状態になる
    //（カードの照合は ApplyDepositCardSelection／ApplyInsuranceCardSelection の行キーフォールバックによる）。
    // 同時に開くのは一方だけ（もう一方のパネルが開いていれば先に閉じる）
    private void OpenRecentEntry(RecentPersonEntry entry, Border row)
    {
        SelectRecentRow(row);
        if (entry.IsInsurance)
        {
            insuranceTab.OpenInsurancePanelFromRecent(entry.Fields);
            return;
        }
        if (insuranceTab.isInsurancePanelOpen) insuranceTab.CloseInsuranceDetailPanel(false);
        isDepositPanelOpen = true;
        SwitchDepositPanelPerson(entry.Fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM], entry.Fields);
        AnimateDetailPanel(depositDetailPanel, DETAIL_PANEL_WIDTH);
    }
}