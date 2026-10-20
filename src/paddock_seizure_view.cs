// ==============================================================================
// paddock_seizure_view.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 差押登録ビュー（ナビ「差押登録」）と差押登録ログの読み込み。年度の開閉セクション・
// 角丸のクリップ・行のセルなど、ほかのビューも使う部品を含む。
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
    // 差押登録ビュー（ナビ「差押登録」）
    // ==============================================================
    //
    // 差押登録ログを年度（4月始まり・登録日時基準）で区切って全件表示する。
    // 当年度のセクションのみ開いた状態で表示し、ヘッダークリックで開閉できる。

    // 差押登録ログを読み込む（専用ビューで全件を表示するため、件数の上限は設けない）
    private void LoadSeizureLog()
    {
        seizureLog.Clear();
        try
        {
            var logPath = GetSeizureLogPath();
            if (!File.Exists(logPath)) return;

            var json = File.ReadAllText(logPath, Encoding.UTF8);
            var entriesIdx = json.IndexOf("\"entries\"");
            if (entriesIdx < 0) return;
            var arrStart = json.IndexOf('[', entriesIdx);
            if (arrStart < 0) return;
            var arrEnd = JsonHelper.FindMatchingBracket(json, arrStart, '[', ']');
            if (arrEnd < 0) return;
            var arrContent = json.Substring(arrStart + 1, arrEnd - arrStart - 1);

            int pos = 0;
            while (pos < arrContent.Length)
            {
                var objStart = arrContent.IndexOf('{', pos);
                if (objStart < 0) break;
                var objEnd = arrContent.IndexOf('}', objStart);
                if (objEnd < 0) break;
                var objContent = arrContent.Substring(objStart + 1, objEnd - objStart - 1);

                var entry = new SeizureLogEntry
                {
                    IsInsurance     = JsonHelper.GetString(objContent, "kind") == KIND_INSURANCE,   // kind のない行は預金
                    AddressNumber   = JsonHelper.GetString(objContent, "addressNumber") ?? "",
                    Name            = JsonHelper.GetString(objContent, "name") ?? "",
                    InstitutionName = JsonHelper.GetString(objContent, "institutionName") ?? "",
                    ExecutionDate   = JsonHelper.GetString(objContent, "executionDate") ?? "",
                    DocumentNumber  = JsonHelper.GetString(objContent, "documentNumber") ?? "",
                    RegisteredAtRaw = JsonHelper.GetString(objContent, "registeredAt") ?? ""
                };
                DateTime regDt;
                entry.RegisteredAt = DateTime.TryParse(entry.RegisteredAtRaw, CultureInfo.InvariantCulture, DateTimeStyles.None, out regDt)
                    ? regDt : DateTime.MinValue;

                seizureLog.Add(entry);
                pos = objEnd + 1;
            }
        }
        catch (Exception ex)
        {
            // 読み込み失敗時は空のまま（0件の案内表示に退行する。落ちない・隠さない）
            WriteErrorLog("差押登録ログの読み込み", ex);
        }
    }

    // ログを読み直して差押登録ビューを再構築する（子プロセス終了時・ナビ切替時）
    // 非表示時は何もしない: Exited は台帳ビューで作業中に飛んでくることが多く、
    // 見えないビューの構築で UI スレッドを使わない。ShowSeizureView が必ず読み直すため整合は保たれる
    private void ReloadSeizureLog()
    {
        if (seizureView.Visibility != Visibility.Visible) return;
        LoadSeizureLog();
        BuildSeizureLogView();
    }

    // 差押登録ビューの中身（年度セクションの並び）を組み立てる
    private void BuildSeizureLogView()
    {
        seizureSectionsPanel.Children.Clear();

        if (seizureLog.Count == 0)
        {
            seizureCard.Visibility = Visibility.Collapsed;
            seizureEmptyText.Text = "登録履歴はまだありません。\nPADDOCK から差押予定一覧に登録すると、ここに表示されます。";
            seizureEmptyText.Visibility = Visibility.Visible;
            SetStatusText(seizureFooterText, "0 件の登録");
            return;
        }
        seizureCard.Visibility = Visibility.Visible;
        seizureEmptyText.Visibility = Visibility.Collapsed;

        // 年度（4月始まり・registeredAt 基準）でグルーピングする。
        // 日時が読めないエントリは「年度不明」として末尾に置く（隠さない）
        var groups = new Dictionary<string, List<SeizureLogEntry>>();
        foreach (var entry in seizureLog)
        {
            var key = entry.RegisteredAt == DateTime.MinValue
                ? "年度不明"
                : BusinessLogic.FormatYearKey(BusinessLogic.GetFiscalYear(entry.RegisteredAt));
            if (!groups.ContainsKey(key)) groups[key] = new List<SeizureLogEntry>();
            groups[key].Add(entry);
        }

        // 年度の降順（「年度不明」は末尾）。セクション内は登録日時の降順
        //（ログファイル自体も先頭追記の降順だが、並びに依存せずここで揃える。
        //  OrderByDescending は安定ソートのため、日時が読めない行は元の並びを保つ）
        var orderedKeys = groups.Keys
            .OrderBy(k => k == "年度不明" ? 1 : 0)
            .ThenByDescending(k => k, StringComparer.Ordinal)
            .ToList();
        var currentYearKey = GetCurrentYearKey();

        foreach (var key in orderedKeys)
        {
            var entries = groups[key].OrderByDescending(e => e.RegisteredAt).ToList();
            seizureSectionsPanel.Children.Add(
                BuildYearSection(key, entries.Count, key == currentYearKey, delegate(StackPanel body)
                {
                    body.Children.Add(BuildSeizureHeaderRow());
                    foreach (var entry in entries)
                        body.Children.Add(BuildSeizureEntryRow(entry));
                }));
        }

        SetStatusText(seizureFooterText, seizureLog.Count.ToString("N0") + " 件の登録　·　" + GetSeizureLogPath());
    }

    // 年度セクション1つ（開閉ヘッダー+本文）を組み立てる（差押登録・最近見た人ビューで共有）
    // expanded=true（当年度）のみ開いた状態で生成し、本文の行は buildRows が組み立てる
    private UIElement BuildYearSection(string yearLabel, int count, bool expanded, Action<StackPanel> buildRows)
    {
        var section = new StackPanel();

        // --- 開閉ヘッダー（シェブロン+年度+件数チップ）---
        var header = new Border
        {
            Background = BrushSurfaceAlt,
            BorderBrush = BrushLine,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(14, 8, 14, 8),
            Cursor = Cursors.Hand
        };
        var headerStack = new StackPanel { Orientation = Orientation.Horizontal };

        var chevron = new TextBlock
        {
            Text = expanded ? "▾" : "▸",
            FontSize = 10,
            Foreground = BrushTextMuted,
            Width = 14,
            VerticalAlignment = VerticalAlignment.Center
        };
        headerStack.Children.Add(chevron);

        headerStack.Children.Add(new TextBlock
        {
            Text = yearLabel,
            FontSize = 12,
            Foreground = BrushTextMain,
            VerticalAlignment = VerticalAlignment.Center
        });

        // 件数チップ（当年度=バイオレット / 他年度=ニュートラル。件数という事実の提示）
        var countChip = new Border
        {
            Background = expanded ? BrushChipVioletBg : BrushChipNeutralBg,
            CornerRadius = new CornerRadius(9),
            Padding = new Thickness(8, 1, 8, 2),
            Margin = new Thickness(8, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        countChip.Child = new TextBlock
        {
            Text = count.ToString("N0") + "件",
            FontSize = 10.5,
            Foreground = expanded ? BrushChipVioletFg : BrushChipNeutralFg
        };
        headerStack.Children.Add(countChip);
        header.Child = headerStack;
        section.Children.Add(header);

        // --- セクション本体（列ヘッダー+エントリ行）---
        // 行の生成は初回展開時まで遅延する（遅延生成）。ログ・履歴は年々累積するため、
        // 閉じたままの過去年度の行を毎回作らない（詳細パネルの人物都度読込と同じ思想）
        var body = new StackPanel { Visibility = expanded ? Visibility.Visible : Visibility.Collapsed };
        bool bodyBuilt = false;
        Action buildBody = delegate
        {
            if (bodyBuilt) return;
            bodyBuilt = true;
            buildRows(body);
        };
        if (expanded) buildBody();
        section.Children.Add(body);

        // ヘッダーのクリック（またはフォーカスして Enter／Space）で開閉を切り替える
        Action toggle = delegate
        {
            bool opening = body.Visibility != Visibility.Visible;
            if (opening) buildBody();
            body.Visibility = opening ? Visibility.Visible : Visibility.Collapsed;
            chevron.Text = opening ? "▾" : "▸";
        };
        header.MouseLeftButtonUp += delegate { toggle(); };
        MakeKeyActivatable(header, toggle, keyFocusInsetStyle);

        return section;
    }

    // カードの中身を角丸の形に沿って切り抜く（差押登録・最近見た人・設定の白カードで共有）。
    // ClipToBounds は矩形でしか切れないため、カードの先頭・末尾に来たセクションヘッダーの
    // 背景が四隅にはみ出して角張って見える。サイズ追従の角丸 Clip でカードの形どおりに切る
    private static void ApplyRoundedClip(Border card, double radius)
    {
        card.SizeChanged += delegate
        {
            var clip = new RectangleGeometry(
                new Rect(0, 0, card.ActualWidth, card.ActualHeight), radius, radius);
            clip.Freeze();
            card.Clip = clip;
        };
    }

    // 差押登録ビューの列構成（ヘッダー行とエントリ行で共有する）
    private static Grid CreateSeizureRowGrid()
    {
        var grid = new Grid();
        // 列幅は固定+末尾の余白列（最近見た人と同じ方針。氏名・金融機関・保険会社は
        // 列数が少ないぶん台帳一覧の150pxより広い200pxをとる）
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(118) });                      // 登録日時
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(56) });                       // 種別
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(104) });                      // 宛名番号
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) });                      // 氏名
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) });                      // 金融機関・保険会社
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(96) });                       // 執行日
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(130) });                      // 文書番号
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });     // 余白
        return grid;
    }

    // 列ヘッダー行（一覧テーブルのヘッダーと同じ配色・サイズ）
    private UIElement BuildSeizureHeaderRow()
    {
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = BrushLine,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(14, 7, 14, 7)
        };
        var grid = CreateSeizureRowGrid();
        var labels = new[] { "登録日時", "種別", "宛名番号", "氏名", "金融機関・保険会社", "執行日", "文書番号" };
        for (int i = 0; i < labels.Length; i++)
        {
            var label = new TextBlock { Text = labels[i], FontSize = 11, Foreground = BrushTextMuted };
            Grid.SetColumn(label, i);
            grid.Children.Add(label);
        }
        border.Child = grid;
        return border;
    }

    // エントリ1件の行。数値・日付・番号は Consolas（スイート規約）
    private UIElement BuildSeizureEntryRow(SeizureLogEntry entry)
    {
        // 行高は固定32px（台帳一覧・最近見た人の行と統一）
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = BrushLineWeak,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(14, 0, 14, 0),
            Height = 32
        };
        var grid = CreateSeizureRowGrid();
        grid.VerticalAlignment = VerticalAlignment.Center;

        // 登録日時: 年はセクション見出しにあるため MM/dd HH:mm。読めない日時は生値のまま
        var registered = entry.RegisteredAt == DateTime.MinValue
            ? entry.RegisteredAtRaw
            : entry.RegisteredAt.ToString("MM/dd HH:mm", CultureInfo.InvariantCulture);
        AddRowCell(grid, 0, registered, true, false);
        AddKindChipCell(grid, 1, entry.IsInsurance);
        AddRowCell(grid, 2, entry.AddressNumber, true, false);
        AddRowCell(grid, 3, entry.Name, false, true);
        AddRowCell(grid, 4, entry.InstitutionName, false, true);
        AddRowCell(grid, 5, FormatExecutionDate(entry.ExecutionDate), true, false);
        AddRowCell(grid, 6, entry.DocumentNumber, true, true);

        border.Child = grid;
        return border;
    }

    // 種別のチップ（預金=ティール／生保=アンバー。ZEKKEN の最近開いたファイルの種別チップと同じ色・
    // 形は年度セクションの件数チップと同じ）を列に置く（差押登録・最近見た人ビューで共有）
    private static void AddKindChipCell(Grid grid, int column, bool insurance)
    {
        var chip = new Border
        {
            Background = insurance ? BrushChipAmberBg : BrushChipTealBg,
            CornerRadius = new CornerRadius(9),
            Padding = new Thickness(8, 1, 8, 2),
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Center
        };
        chip.Child = new TextBlock
        {
            Text = insurance ? "生保" : "預金",
            FontSize = 10.5,
            Foreground = insurance ? BrushChipAmberFg : BrushChipTealFg
        };
        Grid.SetColumn(chip, column);
        grid.Children.Add(chip);
    }

    // 最近見た人の氏名セル（一覧の氏名列と同じく、メモのある人物は氏名の直後に印 ✎）。
    // 氏名の幅は列幅から右余白と印の分（定数）を引いて、長い氏名でも印が切れないようにする。
    // 文字色は行に合わせる（生命保険の電子回答不可の行＝薄文字色）
    private static void AddNameCellWithNoteMark(Grid grid, int column, string name, bool hasNotes, string notesToolTip, Brush foreground)
    {
        var cell = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
        cell.Children.Add(new TextBlock
        {
            Text = name,
            FontSize = 12,
            Foreground = foreground,
            MaxWidth = NAME_COLUMN_WIDTH - NAME_CELL_PADDING_R - (hasNotes ? NOTE_MARK_WIDTH : 0),
            TextTrimming = TextTrimming.CharacterEllipsis,
            ToolTip = name
        });
        if (hasNotes)
        {
            cell.Children.Add(new TextBlock
            {
                Text = "✎",
                FontSize = 10,
                Foreground = BrushTextMuted,
                Margin = new Thickness(5, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                ToolTip = notesToolTip
            });
        }
        Grid.SetColumn(cell, column);
        grid.Children.Add(cell);
    }

    // エントリ行のセル1個を追加する（差押登録・最近見た人ビューで共有）
    private void AddRowCell(Grid grid, int column, string text, bool mono, bool trim)
    {
        AddRowCell(grid, column, text, mono, trim, BrushTextMain);
    }

    // 同・文字色を指定する形（最近見た人の生命保険の電子回答不可の行＝薄文字色）
    private void AddRowCell(Grid grid, int column, string text, bool mono, bool trim, Brush foreground)
    {
        var cell = new TextBlock
        {
            Text = text,
            FontSize = 12,
            Foreground = foreground,
            Margin = new Thickness(0, 0, 8, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        if (mono) cell.FontFamily = FontMono;
        if (trim)
        {
            cell.TextTrimming = TextTrimming.CharacterEllipsis;
            cell.ToolTip = text;
        }
        Grid.SetColumn(cell, column);
        grid.Children.Add(cell);
    }

    // 執行日（yyyy-MM-dd で記録される）を表示形式 yyyy/MM/dd に整える
    // 読めない場合は生値をそのまま返す（落ちない・隠さない）
    private static string FormatExecutionDate(string raw)
    {
        DateTime date;
        if (DateTime.TryParseExact(raw, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            return date.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
        return raw;
    }
}