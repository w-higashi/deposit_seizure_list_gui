// ==============================================================================
// paddock_settings.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 設定ビュー（ナビ「設定」）と設定ファイルの保存。
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
    // 設定ビュー（ナビ「設定」）
    // ==============================================================
    //
    // paddock_config.json の内容と保管ファイルの状態を提示し、設定値の編集・保存を行う。
    // 保存は保存バーの「保存」の明示操作のみで、ファイル全体をツールが書き直す
    //（_README はコード内の定型文で毎回再生成する — BuildConfigJson 参照）。
    // config は係全体で共有される1ファイルのため、書き込みは保存操作に限定する。ビューはファイルの今の値から組み、
    // 保存の前に本文を照合する（ほかの職員の保存を後勝ちで消さない）。
    // フォルダ・ファイルの存在チェックは事実の提示に徹し、保存を妨げない（判定しない）。

    // ビューの中身（6枚のカード）を設定ファイルの今の値（読めなければメモリ上の config）から組み立て直す。
    // 組み立て時点の入力値がスナップショットとなり、変更検出と「変更を破棄」の戻し先になる
    private void BuildSettingsView()
    {
        // ファイルの今の値を示す（ほかの職員の保存を映す）。読めない・解釈できないときはメモリ上の値で組み、
        // 保存前の照合（SaveSettings）は行わない
        settingsShownConfig = LoadConfig(System.IO.Path.Combine(exeDir, CONFIG_FILE), out settingsConfigText);
        if (settingsConfigText == null) settingsShownConfig = config;

        suppressSettingsDirty = true;
        settingsTypeMappingChecked = null;   // 組み立て直すときは補正表を読み直す
        settingsPanel.Children.Clear();
        depositSettingsRows.FolderRows.Clear();
        insuranceSettingsRows.FolderRows.Clear();

        BuildSettingsKindCard(depositSettingsRows, settingsShownConfig.DepositLedgerFolder, settingsShownConfig.DepositSearchFolders,
            settingsShownConfig.DepositSeizureListScript, settingsShownConfig.DepositSeizureTargetBanks);
        BuildSettingsKindCard(insuranceSettingsRows, settingsShownConfig.InsuranceLedgerFolder, settingsShownConfig.InsuranceSearchFolders,
            settingsShownConfig.InsuranceSeizureListScript, settingsShownConfig.InsuranceSeizureTargetCompanies);
        BuildSettingsCommonCard();
        BuildSettingsUserCard();
        BuildSettingsHistoryCard();
        BuildSettingsInfoCard();

        // スナップショットの確定（現在の入力値 = 保存済みの値）
        TakeSettingsSnapshot(depositSettingsRows);
        TakeSettingsSnapshot(insuranceSettingsRows);
        settingsSnapRecentMax = settingsRecentMaxBox.Text;
        settingsSnapFuzzy = settingsFuzzyBox.Text;
        settingsSnapTypeMapping = settingsTypeMappingBox.Text;

        suppressSettingsDirty = false;
        settingsDirty = false;
        SetSettingsSaveBarVisible(false);

        UpdateSettingsChips();

        settingsFooterNormal = "設定ファイル: " + System.IO.Path.Combine(exeDir, CONFIG_FILE);
        settingsFooterLeft.Text = settingsFooterNormal;
        settingsFooterRight.Text = "PADDOCK " + APP_VERSION;
    }

    // 種別ごとの4行の入力値を保存済みの値として控える
    private static void TakeSettingsSnapshot(SettingsKindRows rows)
    {
        rows.SnapLedger = rows.LedgerBox.Text;
        rows.SnapScript = rows.ScriptBox.Text;
        rows.SnapTarget = rows.TargetBox.Text;
        rows.SnapFolders = CollectFolderPairs(rows);
    }

    // 1・2枚目: 預金・生命保険のカード（台帳フォルダ／照会結果フォルダ／差押リスト連携／差押リストの対象）。
    // 行の部品は入れ物 rows に控え、現在値は config の種別の値を渡す
    private void BuildSettingsKindCard(SettingsKindRows rows, string ledgerFolderPath, Dictionary<string, List<string>> folderMap,
        string seizureListScript, string seizureTargets)
    {
        StackPanel body;
        settingsPanel.Children.Add(MakeSettingsCard(rows.Kind, CONFIG_FILE, out body));

        // --- 台帳フォルダ（空欄 = 既定: exe と同一フォルダから検出）---
        var ledgerValue = new StackPanel();
        var ledgerLine = new Grid();
        ledgerLine.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        ledgerLine.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var ledgerStack = new Grid();   // 入力欄とプレースホルダの重ね合わせ（検索ボックスと同じ方式）
        ledgerStack.Children.Add(MakeSettingsInput(ledgerFolderPath ?? "", out rows.LedgerBox, 0));
        rows.LedgerPlaceholder = new TextBlock
        {
            Text = "（未設定 — paddock.exe と同一フォルダから検出）",
            FontSize = 11.5,
            Foreground = BrushTextMuted,
            Margin = new Thickness(10, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            IsHitTestVisible = false
        };
        rows.LedgerPlaceholder.Visibility =
            rows.LedgerBox.Text.Length > 0 ? Visibility.Collapsed : Visibility.Visible;
        rows.LedgerBox.TextChanged += delegate
        {
            rows.LedgerPlaceholder.Visibility =
                rows.LedgerBox.Text.Length > 0 ? Visibility.Collapsed : Visibility.Visible;
        };
        ledgerStack.Children.Add(rows.LedgerPlaceholder);
        ledgerLine.Children.Add(ledgerStack);

        var ledgerBrowse = MakeMiniButton("参照…");
        ledgerBrowse.Margin = new Thickness(8, 0, 0, 0);
        ledgerBrowse.Click += delegate { BrowseForFolder(rows.LedgerBox); };
        Grid.SetColumn(ledgerBrowse, 1);
        ledgerLine.Children.Add(ledgerBrowse);
        ledgerValue.Children.Add(ledgerLine);

        rows.LedgerNote = MakeSettingsNote("");
        ledgerValue.Children.Add(rows.LedgerNote);
        AddSettingsRow(body, "台帳フォルダ", rows.LedgerKey, ledgerValue, false);

        // --- 照会結果フォルダ（年度→フォルダのマッピング。行の追加・削除で編集する）---
        var foldersValue = new StackPanel();
        rows.FolderRowsPanel = new StackPanel();
        if (folderMap != null)
        {
            foreach (var pair in folderMap.OrderBy(p => p.Key, StringComparer.Ordinal))
                foreach (var folder in pair.Value)
                    AddFolderRow(rows, pair.Key, folder);
        }
        foldersValue.Children.Add(rows.FolderRowsPanel);
        foldersValue.Children.Add(MakeAddFolderLink(rows));
        foldersValue.Children.Add(MakeSettingsNote(
            "年度キーは台帳の年度表記（例: 2027年度）に合わせてください。" +
            "フォルダの有無は入力のたびに確認して表示します（保存は妨げません）。" +
            "年度・パスのどちらかが空欄の行は保存されません。"));
        AddSettingsRow(body, "照会結果フォルダ", rows.FoldersKey, foldersValue, false);

        // --- 差押リスト連携（連携先の exe のパス）---
        var scriptValue = new StackPanel();
        var scriptLine = new Grid();
        scriptLine.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        scriptLine.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        scriptLine.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        scriptLine.Children.Add(MakeSettingsInput(seizureListScript ?? "", out rows.ScriptBox, 0));

        rows.ScriptChip = MakeChip();
        rows.ScriptChip.Margin = new Thickness(8, 0, 0, 0);
        Grid.SetColumn(rows.ScriptChip, 1);
        scriptLine.Children.Add(rows.ScriptChip);

        var scriptBrowse = MakeMiniButton("参照…");
        scriptBrowse.Margin = new Thickness(8, 0, 0, 0);
        scriptBrowse.Click += delegate
        {
            BrowseForFile(rows.ScriptBox, "実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*", rows.ToolName + "を選択");
        };
        Grid.SetColumn(scriptBrowse, 2);
        scriptLine.Children.Add(scriptBrowse);
        scriptValue.Children.Add(scriptLine);

        scriptValue.Children.Add(MakeSettingsNote(
            "この項目の変更は PADDOCK の再起動後に反映されます（連携ボタン・メニューの表示が切り替わるため）。"));
        AddSettingsRow(body, "差押リスト連携", rows.ScriptKey, scriptValue, false);

        // --- 差押リストの対象金融機関・対象保険会社 ---
        var targetValue = new StackPanel();
        var targetLine = new Grid();
        targetLine.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        targetLine.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        targetLine.Children.Add(MakeSettingsInput(seizureTargets ?? "", out rows.TargetBox, 0));
        rows.TargetChip = MakeChip();
        rows.TargetChip.Margin = new Thickness(8, 0, 0, 0);
        Grid.SetColumn(rows.TargetChip, 1);
        targetLine.Children.Add(rows.TargetChip);
        targetValue.Children.Add(targetLine);

        // 一致の開示（淡アイリス面）: 読み込み済みの台帳と照合した事実の提示のみで保存は妨げない
        rows.TargetMatchText = new TextBlock
        {
            FontSize = 11,
            Foreground = BrushTextSub,
            TextWrapping = TextWrapping.Wrap
        };
        rows.TargetMissText = new TextBlock
        {
            FontSize = 11,
            Foreground = BrushTextMuted,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 3, 0, 0)
        };
        var targetStack = new StackPanel();
        targetStack.Children.Add(rows.TargetMatchText);
        targetStack.Children.Add(rows.TargetMissText);
        rows.TargetDisclosure = new Border
        {
            Background = BrushIrisPanelBg,
            BorderBrush = BrushSelectedLine,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(10, 6, 10, 7),
            Margin = new Thickness(0, 6, 0, 0),
            Child = targetStack
        };
        targetValue.Children.Add(rows.TargetDisclosure);

        targetValue.Children.Add(MakeSettingsNote(
            "カンマ区切りで指定します。対象を限定する場合に指定し、* は任意の文字列・" +
            "? は任意の1文字（例: *信用金庫。全角の＊？も可）。一致しない" + rows.TargetNoun + "の行では" +
            "「差押リストに追加」が無効になります。未設定・空欄の場合は制御しません。" +
            "全ユーザー共通の設定です。"));
        AddSettingsRow(body, "差押リストの対象" + rows.TargetNoun, rows.TargetKey, targetValue, true);
    }

    // 3枚目: 共通のカード（最近見た人の保持件数／担当の化け置換文字／口座種別の補正表／明細の読み取り位置。
    // 後ろの2つは預金だけの機能で、明細の読み取り位置は現在値の開示のみ）
    private void BuildSettingsCommonCard()
    {
        StackPanel body;
        settingsPanel.Children.Add(MakeSettingsCard("共通", CONFIG_FILE, out body));

        // --- 最近見た人の保持件数 ---
        var recentValue = new StackPanel();
        var recentLine = new StackPanel { Orientation = Orientation.Horizontal };
        recentLine.Children.Add(MakeSettingsInput(settingsShownConfig.RecentMax.ToString(), out settingsRecentMaxBox, 76));
        settingsRecentMaxBox.TextAlignment = TextAlignment.Right;
        recentLine.Children.Add(new TextBlock
        {
            Text = "件（年度セクションごと）",
            FontSize = 12,
            Foreground = BrushTextSub,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 0, 0, 0)
        });
        recentValue.Children.Add(recentLine);
        recentValue.Children.Add(MakeSettingsNote(
            "超えた分は閲覧日時の古いものから削除されます。全ユーザー共通の設定です。"));
        AddSettingsRow(body, "最近見た人の保持件数", "recentMax", recentValue, false);

        // --- 担当の化け置換文字 ---
        var fuzzyValue = new StackPanel();
        var fuzzyLine = new StackPanel { Orientation = Orientation.Horizontal };
        fuzzyLine.Children.Add(MakeSettingsInput(
            FormatStaffFuzzyChars(settingsShownConfig.StaffFuzzyChars), out settingsFuzzyBox, 160));
        settingsFuzzyChip = MakeChip();
        settingsFuzzyChip.Margin = new Thickness(8, 0, 0, 0);
        fuzzyLine.Children.Add(settingsFuzzyChip);
        fuzzyValue.Children.Add(fuzzyLine);
        fuzzyValue.Children.Add(MakeSettingsNote(
            "カンマ区切りで指定します。照会結果で外字が置き換わる文字（例: ◆,●）を" +
            "指定すると、詳細検索「担当」の列挙でこれらを任意の1文字とみなして" +
            "表記ゆれを統合します。空欄にすると化け文字による統合を行いません。" +
            "全ユーザー共通の設定です。"));
        AddSettingsRow(body, "担当の化け置換文字", "staffFuzzyChars", fuzzyValue, false);

        // --- 口座種別の補正表（JUDGE の account_type_mapping.json のパス）---
        var typeMappingValue = new StackPanel();
        var typeMappingLine = new Grid();
        typeMappingLine.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        typeMappingLine.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        typeMappingLine.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        typeMappingLine.Children.Add(MakeSettingsInput(settingsShownConfig.AccountTypeMappingFile ?? "", out settingsTypeMappingBox, 0));

        settingsTypeMappingChip = MakeChip();
        settingsTypeMappingChip.Margin = new Thickness(8, 0, 0, 0);
        Grid.SetColumn(settingsTypeMappingChip, 1);
        typeMappingLine.Children.Add(settingsTypeMappingChip);

        var typeMappingBrowse = MakeMiniButton("参照…");
        typeMappingBrowse.Margin = new Thickness(8, 0, 0, 0);
        typeMappingBrowse.Click += delegate
        {
            BrowseForFile(settingsTypeMappingBox, "JSON ファイル (*.json)|*.json|すべてのファイル (*.*)|*.*", "口座種別の補正表を選択");
        };
        Grid.SetColumn(typeMappingBrowse, 2);
        typeMappingLine.Children.Add(typeMappingBrowse);
        typeMappingValue.Children.Add(typeMappingLine);

        typeMappingValue.Children.Add(MakeSettingsNote(
            "預金のみ。JUDGE の口座種別の補正表（judge.ps1 と同じフォルダの account_type_mapping.json）を指定します。" +
            "台帳の口座種別は JUDGE がこの表で補正した表記のため、指定すると明細を見るときに照会ファイルの" +
            "明細ページを補正前・補正後のどちらの表記でも見つけます。未設定・読めないときは明細ページの表記のまま照合します。" +
            "全ユーザー共通の設定です。"));
        AddSettingsRow(body, "口座種別の補正表", "accountTypeMappingFile", typeMappingValue, false);

        // --- 明細の読み取り位置: ファイルの今の値の開示のみ（編集はテキストエディタ）---
        var detail = settingsShownConfig.Detail;
        var detailValue = new StackPanel();
        detailValue.Children.Add(MakeSettingsValueLine("目印", string.Join("／", detail.StopValues)));
        detailValue.Children.Add(MakeSettingsValueLine(
            "口座番号", detail.AccountNumberCell, "口座種別", detail.AccountTypeCell, "開始行", detail.DataStartRow.ToString()));
        detailValue.Children.Add(MakeSettingsValueLine(
            "取引日", detail.DateCol, "摘要", detail.SummaryCol, "出金", detail.WithdrawalCol,
            "入金", detail.DepositCol, "残高", detail.BalanceCol));
        detailValue.Children.Add(MakeSettingsNote(
            "預金のみ・表示のみ。詳細パネルの口座カードから照会シートの取引明細を読む位置です（明細ページ先頭の目印の行を1行目とした相対指定）。" +
            "JUDGE の設定と同じ位置を指すようにしてください。この項目はここでは変更できません。" +
            "変更はテキストエディタで行い、再起動後に反映されます。"));
        AddSettingsRow(body, "明細の読み取り位置", "detail", detailValue, true);
    }

    // 「項目名 値　項目名 値 …」の1行（項目名は副テキスト色・値は等幅）を作る
    private static TextBlock MakeSettingsValueLine(params string[] labelsAndValues)
    {
        var line = new TextBlock { FontSize = 12, Margin = new Thickness(0, 3, 0, 0), TextWrapping = TextWrapping.Wrap };
        for (int i = 0; i + 1 < labelsAndValues.Length; i += 2)
        {
            if (i > 0) line.Inlines.Add(new Run("　"));
            line.Inlines.Add(new Run(labelsAndValues[i] + " ") { Foreground = BrushTextSub });
            line.Inlines.Add(new Run(labelsAndValues[i + 1]) { FontFamily = FontMono, Foreground = BrushTextMain });
        }
        return line;
    }

    // 4枚目: このユーザーの設定カード（user_prefs.json の編集）。
    // 表示倍率は押した瞬間に反映し、そのまま自動保存する — 倍率は即時に見えないと
    // 調整できず、保存先がユーザー単位ファイルのため係共有 config のような書き込み
    // 競合もない。保存バー・変更の要約には関与しない
    private TextBlock scaleStepperValue;            // ステッパーの現在値表示
    private Border scaleStepperChip;                // 既定/縮小/拡大チップ
    private Button scaleStepDownBtn;                // −（最小の80%で無効）
    private Button scaleStepUpBtn;                  // ＋（最大の120%で無効）

    private void BuildSettingsUserCard()
    {
        StackPanel body;
        settingsPanel.Children.Add(MakeSettingsCard("このユーザーの設定",
            System.IO.Path.Combine("history", Environment.UserName, "user_prefs.json") +
            " — この Windows ユーザーにのみ適用", out body));

        var value = new StackPanel();
        var line = new StackPanel { Orientation = Orientation.Horizontal };

        // ステッパー（−/現在値/＋）: 1本の枠の中を3分割する
        var stepper = new Border
        {
            Background = Brushes.White,
            BorderBrush = BrushLine,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(7),
            VerticalAlignment = VerticalAlignment.Center
        };
        ApplyRoundedClip(stepper, 7);   // ボタンのホバー面を角丸に沿って切り抜く
        var stepGrid = new Grid();
        stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        scaleStepDownBtn = MakeScaleStepButton("−");
        scaleStepDownBtn.Click += delegate { StepDisplayScale(-1); };
        stepGrid.Children.Add(scaleStepDownBtn);

        var valueBorder = new Border
        {
            BorderBrush = BrushLine,
            BorderThickness = new Thickness(1, 0, 1, 0),
            Width = 64
        };
        scaleStepperValue = new TextBlock
        {
            FontFamily = FontMono,
            FontSize = 12.5,
            Foreground = BrushTextMain,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        valueBorder.Child = scaleStepperValue;
        Grid.SetColumn(valueBorder, 1);
        stepGrid.Children.Add(valueBorder);

        scaleStepUpBtn = MakeScaleStepButton("＋");
        scaleStepUpBtn.Click += delegate { StepDisplayScale(+1); };
        Grid.SetColumn(scaleStepUpBtn, 2);
        stepGrid.Children.Add(scaleStepUpBtn);
        stepper.Child = stepGrid;
        line.Children.Add(stepper);

        scaleStepperChip = MakeChip();
        scaleStepperChip.Margin = new Thickness(8, 0, 0, 0);
        line.Children.Add(scaleStepperChip);
        value.Children.Add(line);

        value.Children.Add(MakeSettingsNote(
            "80〜120% の9段階（5%刻み）。文字・余白・一覧・パネルがまとめて拡大縮小されます。" +
            "変更は押した瞬間に画面へ反映され、そのまま自動で保存されます（次回起動時もこの倍率で開きます）。"));

        AddSettingsRow(body, "画面の表示倍率", "displayScale", value, true);
        UpdateScaleStepperUi();
    }

    private Button MakeScaleStepButton(string label)
    {
        return new Button { Content = label, Style = (Style)window.FindResource("ScaleStepBtn") };
    }

    // 表示倍率を1段動かして即時反映・即時保存する
    private void StepDisplayScale(int direction)
    {
        var index = Array.IndexOf(DisplayScaleSteps, displayScale) + direction;
        if (index < 0 || index >= DisplayScaleSteps.Length) return;
        displayScale = DisplayScaleSteps[index];
        ApplyDisplayScale();
        WriteUserPrefs();
        UpdateScaleStepperUi();
    }

    // ステッパーの現在値・チップ・端点の無効化を現在の倍率に合わせる
    private void UpdateScaleStepperUi()
    {
        if (scaleStepperValue == null) return;
        scaleStepperValue.Text = displayScale + "%";
        if (displayScale == 100)
            SetChip(scaleStepperChip, "既定", BrushChipNeutralBg, BrushChipNeutralFg);
        else
            SetChip(scaleStepperChip, displayScale < 100 ? "縮小" : "拡大", BrushChipTealBg, BrushChipTealFg);
        scaleStepDownBtn.IsEnabled = displayScale != DisplayScaleSteps[0];
        scaleStepUpBtn.IsEnabled = displayScale != DisplayScaleSteps[DisplayScaleSteps.Length - 1];
    }

    // 5枚目: 履歴ファイルカード（history/ の状態の提示。読み取り専用）
    private void BuildSettingsHistoryCard()
    {
        StackPanel body;
        settingsPanel.Children.Add(MakeSettingsCard("履歴ファイル", "history／このユーザーの分のみ", out body));

        AddSettingsRow(body, "最近見た人", null,
            MakePathWithChip(GetRecentPeoplePath(), recentPeople.Count.ToString("N0") + " 件"), false);

        // 件数の提示用にログを読み直す（ログは小さい。差押登録ビューは表示のたびに再読込するため影響しない）
        LoadSeizureLog();
        AddSettingsRow(body, "差押登録ログ", null,
            MakePathWithChip(GetSeizureLogPath(), seizureLog.Count.ToString("N0") + " 件"), false);

        // 件数の提示用にメモのキャッシュを確保する（初回のみ読み込み・以降はキャッシュ）
        EnsureNotesLoaded();
        AddSettingsRow(body, "人物メモ", null,
            MakePathWithChip(GetPersonNotesPath(),
                myPersonNotes.Values.Sum(l => l.Count).ToString("N0") + " 件"), false);

        AddSettingsRow(body, "お知らせの既読情報", null,
            MakePathWithChip(GetNoticeStatePath(),
                lastSeenVersion == "" ? "未確認" : "ver " + lastSeenVersion + " まで確認済み"), false);

        // エラーログ（記録の件数=見出し行「[日時] PADDOCK …」の数。読めなければ「なし」）
        int errorCount = 0;
        try
        {
            if (File.Exists(GetErrorLogPath()))
                foreach (var line in File.ReadAllLines(GetErrorLogPath(), Encoding.UTF8))
                    if (line.StartsWith("[") && line.Contains("] PADDOCK ")) errorCount++;
        }
        catch { }   // 読めなければ件数は「なし」
        AddSettingsRow(body, "エラーログ", null,
            MakePathWithChip(GetErrorLogPath(), errorCount == 0 ? "なし" : errorCount.ToString("N0") + " 件"), true);
    }

    // 6枚目: ツール情報カード（バージョン+お知らせの再表示）
    private void BuildSettingsInfoCard()
    {
        StackPanel body;
        settingsPanel.Children.Add(MakeSettingsCard("ツール情報", null, out body));

        var verLine = new Grid();
        verLine.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        verLine.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        verLine.Children.Add(new TextBlock
        {
            Text = "PADDOCK " + APP_VERSION,
            FontFamily = FontMono,
            FontSize = 12,
            Foreground = BrushTextMain,
            VerticalAlignment = VerticalAlignment.Center
        });

        var replayButton = MakeMiniButton("お知らせを再表示");
        bool noticeExists = File.Exists(System.IO.Path.Combine(exeDir, NOTICE_FILE));
        replayButton.IsEnabled = noticeExists;
        if (!noticeExists)
        {
            replayButton.ToolTip = "update_notice.json が見つかりません";
            ToolTipService.SetShowOnDisabled(replayButton, true);
        }
        replayButton.Click += delegate { ReplayUpdateNotice(); };
        Grid.SetColumn(replayButton, 1);
        verLine.Children.Add(replayButton);

        AddSettingsRow(body, "バージョン", null, verLine, true);
    }

    // 照会結果フォルダの編集行1本（年度キー+パス+存在チップ+参照+削除）を種別の入れ物へ追加する
    private void AddFolderRow(SettingsKindRows rows, string year, string path)
    {
        var grid = new Grid { Margin = new Thickness(0, 0, 0, 6) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(84) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var row = new SettingsFolderRow();

        TextBox yearBox;
        var yearInput = MakeSettingsInput(year, out yearBox, 0);
        yearInput.Margin = new Thickness(0, 0, 8, 0);
        grid.Children.Add(yearInput);

        TextBox pathBox;
        var pathInput = MakeSettingsInput(path, out pathBox, 0);
        Grid.SetColumn(pathInput, 1);
        grid.Children.Add(pathInput);

        var chip = MakeChip();
        chip.Margin = new Thickness(8, 0, 0, 0);
        Grid.SetColumn(chip, 2);
        grid.Children.Add(chip);

        var browse = MakeMiniButton("参照…");
        browse.Margin = new Thickness(8, 0, 0, 0);
        browse.Click += delegate { BrowseForFolder(pathBox); };
        Grid.SetColumn(browse, 3);
        grid.Children.Add(browse);

        var del = MakeRowDeleteButton();
        del.Margin = new Thickness(4, 0, 0, 0);
        // キーで消したときは、繰り上がった行（なければ1つ上の行）の年度の欄へフォーカスを移す
        Action remove = delegate
        {
            bool refocus = del.IsKeyboardFocused;
            int index = rows.FolderRows.IndexOf(row);
            rows.FolderRows.Remove(row);
            rows.FolderRowsPanel.Children.Remove(grid);
            UpdateSettingsDirty();
            if (refocus && rows.FolderRows.Count > 0)
                FocusLater(delegate { rows.FolderRows[Math.Min(index, rows.FolderRows.Count - 1)].YearBox.Focus(); });
        };
        del.MouseLeftButtonUp += delegate { remove(); };
        MakeKeyActivatable(del, remove, keyFocusPillStyle);
        Grid.SetColumn(del, 4);
        grid.Children.Add(del);

        row.YearBox = yearBox;
        row.PathBox = pathBox;
        row.Chip = chip;
        rows.FolderRows.Add(row);
        rows.FolderRowsPanel.Children.Add(grid);
    }

    // 「＋ 年度フォルダを追加」リンク（クリック・Enter／Space で空の行を追加し年度キーへフォーカス）
    private UIElement MakeAddFolderLink(SettingsKindRows rows)
    {
        var panel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Cursor = Cursors.Hand,
            HorizontalAlignment = HorizontalAlignment.Left,
            Margin = new Thickness(0, 2, 0, 0)
        };
        panel.Children.Add(new Border
        {
            Width = 18,
            Height = 18,
            BorderBrush = BrushAccent,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(5),
            Child = new TextBlock
            {
                Text = "＋",
                FontSize = 10,
                Foreground = BrushAccent,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            }
        });
        panel.Children.Add(new TextBlock
        {
            Text = "年度フォルダを追加",
            FontSize = 11.5,
            Foreground = BrushAccent,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(6, 0, 0, 0)
        });
        Action add = delegate
        {
            AddFolderRow(rows, "", "");
            UpdateSettingsDirty();
            rows.FolderRows[rows.FolderRows.Count - 1].YearBox.Focus();
        };
        panel.MouseLeftButtonUp += delegate { add(); };
        MakeKeyActivatable(panel, add);
        return panel;
    }

    // 入力の変化: 変更検出と存在チェックの再実行。存在チェックは250msデバウンス
    //（F: はネットワークドライブのため、1文字ごとの Directory.Exists を避ける。
    //  詳細パネルの年度横断読み込みと同じ規約）
    private void OnSettingsInputChanged()
    {
        if (suppressSettingsDirty) return;
        UpdateSettingsDirty();

        if (settingsChipTimer == null)
        {
            settingsChipTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
            settingsChipTimer.Tick += delegate
            {
                settingsChipTimer.Stop();
                UpdateSettingsChips();
            };
        }
        settingsChipTimer.Stop();
        settingsChipTimer.Start();
    }

    // 現在の入力値をスナップショットと突き合わせ、保存バーの表示と変更の要約を更新する
    private void UpdateSettingsDirty()
    {
        var changed = new List<string>();
        CollectSettingsChanges(depositSettingsRows, changed);
        CollectSettingsChanges(insuranceSettingsRows, changed);
        if (settingsRecentMaxBox.Text != settingsSnapRecentMax) changed.Add("最近見た人の保持件数");
        if (settingsFuzzyBox.Text != settingsSnapFuzzy) changed.Add("担当の化け置換文字");
        if (settingsTypeMappingBox.Text != settingsSnapTypeMapping) changed.Add("口座種別の補正表");

        settingsDirty = changed.Count > 0;
        SetSettingsSaveBarVisible(settingsDirty);
        if (settingsDirty)
            settingsChangeSummary.Text = "変更 " + changed.Count + "件: " + string.Join("　·　", changed.ToArray());
    }

    // 種別ごとの4行の変更を要約へ集める（行の名前に「（預金）」「（生命保険）」を添える）
    private static void CollectSettingsChanges(SettingsKindRows rows, List<string> changed)
    {
        var suffix = "（" + rows.Kind + "）";
        if (rows.LedgerBox.Text != rows.SnapLedger) changed.Add("台帳フォルダ" + suffix);
        if (!FolderPairsEqual(CollectFolderPairs(rows), rows.SnapFolders)) changed.Add("照会結果フォルダ" + suffix);
        if (rows.ScriptBox.Text != rows.SnapScript) changed.Add("差押リスト連携" + suffix);
        if (rows.TargetBox.Text != rows.SnapTarget) changed.Add("差押リストの対象" + rows.TargetNoun);
    }

    // 保存バーの表示切替（下から生えて下へ消える開閉アニメーション・150ms）
    private void SetSettingsSaveBarVisible(bool visible)
    {
        if (visible == settingsSaveBarShown) return;
        settingsSaveBarShown = visible;
        if (visible) AnimateExpand(settingsSaveBar);
        else AnimateCollapse(settingsSaveBar);
    }

    // 種別の照会結果フォルダの編集行を {年度, パス} の組の並びとして取り出す
    private static List<string[]> CollectFolderPairs(SettingsKindRows rows)
    {
        var pairs = new List<string[]>();
        foreach (var row in rows.FolderRows)
            pairs.Add(new[] { row.YearBox.Text, row.PathBox.Text });
        return pairs;
    }

    private static bool FolderPairsEqual(List<string[]> a, List<string[]> b)
    {
        if (a.Count != b.Count) return false;
        for (int i = 0; i < a.Count; i++)
            if (a[i][0] != b[i][0] || a[i][1] != b[i][1]) return false;
        return true;
    }

    // 存在チェックのチップと台帳フォルダの注記を更新する（事実の提示。見つからなくても保存は妨げない）
    private void UpdateSettingsChips()
    {
        UpdateSettingsKindChips(depositSettingsRows);
        UpdateSettingsKindChips(insuranceSettingsRows);
        UpdateFuzzyChip();
        UpdateTypeMappingChip();
    }

    // 種別ごとの4行のチップと注記
    private void UpdateSettingsKindChips(SettingsKindRows rows)
    {
        foreach (var row in rows.FolderRows)
        {
            var path = row.PathBox.Text.Trim();
            if (path == "") SetChip(row.Chip, null, null, null);   // 空欄はチップなし
            else if (Directory.Exists(path)) SetChip(row.Chip, "フォルダあり", BrushChipTealBg, BrushChipTealFg);
            else SetChip(row.Chip, "見つかりません", BrushChipCoralBg, BrushChipCoralFg);
        }
        UpdateScriptChip(rows);
        UpdateSeizureTargetNote(rows);
        UpdateLedgerNote(rows);
    }

    // 差押リスト連携のチップ。保存済み・再起動待ちの間は「再起動後に反映」を提示する
    //（入力を再度変えたら通常の存在チェックに戻る）
    private static void UpdateScriptChip(SettingsKindRows rows)
    {
        var path = rows.ScriptBox.Text.Trim();
        if (rows.ScriptRestartPending && rows.ScriptBox.Text == rows.SnapScript)
            SetChip(rows.ScriptChip, "再起動後に反映", BrushChipAmberBg, BrushChipAmberFg);
        else if (path == "")
            SetChip(rows.ScriptChip, "未設定（連携なし）", BrushChipNeutralBg, BrushChipNeutralFg);
        else if (File.Exists(path))
            SetChip(rows.ScriptChip, "ファイルあり", BrushChipTealBg, BrushChipTealFg);
        else
            SetChip(rows.ScriptChip, "見つかりません", BrushChipCoralBg, BrushChipCoralFg);
    }

    // 差押リストの対象のチップと一致の開示を現在の入力値で更新する。
    // 読み込み済みの台帳の金融機関・保険会社（distinct）と照合した結果の提示で、年度を
    // 切り替えたときも読み込み完了時に組み直される。設定項目ごとに一致の有無を
    // 数え、一致0件の項目だけを「効いていない設定」として1行で示す。
    // 生命保険の台帳を読み込む前（生命保険タブを一度も表示していない）は、設定のために台帳は読まず、
    // 語があるときだけ読み込むと一致を表示する旨の1行を出す
    private void UpdateSeizureTargetNote(SettingsKindRows rows)
    {
        var tokens = new List<string>();
        foreach (var token in rows.TargetBox.Text.Split(','))
        {
            var trimmed = token.Trim();
            if (trimmed.Length > 0) tokens.Add(trimmed);
        }
        if (tokens.Count == 0)
        {
            SetChip(rows.TargetChip, "未設定 — 制御なし", BrushChipNeutralBg, BrushChipNeutralFg);
            rows.TargetDisclosure.Visibility = Visibility.Collapsed;
            return;
        }

        // 読み込み済みの行の distinct の名前（照合は事前計算済みキー・表示は生値の初出）
        var targetNames = new Dictionary<string, string>();
        var keyOrder = new List<string>();
        if (rows == depositSettingsRows)
        {
            foreach (var record in depositAllRecords)
                AddSeizureTargetName(targetNames, keyOrder, record.SearchInstitution, record.InstitutionName);
        }
        else if (insuranceTab.IsLoaded)
        {
            foreach (var record in insuranceTab.insuranceAllRecords)
                AddSeizureTargetName(targetNames, keyOrder, record.SearchCompany, record.CompanyName);
        }
        else
        {
            SetChip(rows.TargetChip, null, null, null);
            rows.TargetMatchText.Text = "生命保険の台帳を読み込むと一致を表示します（生命保険タブを開くと読み込みます）";
            rows.TargetMissText.Visibility = Visibility.Collapsed;
            rows.TargetDisclosure.Visibility = Visibility.Visible;
            return;
        }

        var matchedKeys = new HashSet<string>();
        var missTokens = new List<string>();
        foreach (var token in tokens)
        {
            var patterns = BusinessLogic.ParseSeizureTargetBanks(token);
            bool any = false;
            if (patterns != null)
            {
                foreach (var key in keyOrder)
                {
                    if (!BusinessLogic.MatchesSeizureTarget(key, patterns)) continue;
                    matchedKeys.Add(key);
                    any = true;
                }
            }
            if (!any) missTokens.Add(token);
        }
        var matchedNames = new List<string>();
        foreach (var key in keyOrder)
            if (matchedKeys.Contains(key)) matchedNames.Add(targetNames[key]);

        if (matchedNames.Count == 0)
            SetChip(rows.TargetChip, "一致なし", BrushChipNeutralBg, BrushChipNeutralFg);
        else
            SetChip(rows.TargetChip, matchedNames.Count + " 種類が一致", BrushChipTealBg, BrushChipTealFg);

        rows.TargetMatchText.Text = matchedNames.Count == 0
            ? "一致: この年度の台帳に一致する" + rows.TargetNoun + "はありません"
            : "一致: " + string.Join("・", matchedNames.ToArray());
        if (missTokens.Count == 0)
        {
            rows.TargetMissText.Visibility = Visibility.Collapsed;
        }
        else
        {
            rows.TargetMissText.Visibility = Visibility.Visible;
            rows.TargetMissText.Text = "効いていない設定: " + string.Join("・", missTokens.ToArray());
        }
        rows.TargetDisclosure.Visibility = Visibility.Visible;
    }

    // 一致の開示に使う名前を集める（キーの初出だけ。空のキーは対象外）
    private static void AddSeizureTargetName(Dictionary<string, string> names, List<string> keyOrder, string key, string name)
    {
        if (string.IsNullOrEmpty(key) || names.ContainsKey(key)) return;
        names[key] = name ?? "";
        keyOrder.Add(key);
    }

    // 化け置換文字のチップ（採用される文字数の提示。空欄=統合なしも事実として示す）
    private void UpdateFuzzyChip()
    {
        var chars = ParseStaffFuzzyChars(settingsFuzzyBox.Text);
        if (chars.Count == 0)
            SetChip(settingsFuzzyChip, "化け文字の統合なし", BrushChipNeutralBg, BrushChipNeutralFg);
        else
            SetChip(settingsFuzzyChip, chars.Count + " 文字", BrushChipTealBg, BrushChipTealFg);
    }

    // 口座種別の補正表のチップ（読めるかどうかと金融機関の表の数の提示。読めなくても保存は妨げない）。
    // ファイルは入力が前回の確かめから変わったときだけ読む（ほかの行を打つたびに共有フォルダの
    // ファイルを読み直さない）
    private void UpdateTypeMappingChip()
    {
        var path = settingsTypeMappingBox.Text.Trim();
        if (path == settingsTypeMappingChecked) return;
        settingsTypeMappingChecked = path;
        if (path == "")
        {
            SetChip(settingsTypeMappingChip, "未設定（補正なし）", BrushChipNeutralBg, BrushChipNeutralFg);
            return;
        }
        if (!File.Exists(path))
        {
            SetChip(settingsTypeMappingChip, "見つかりません", BrushChipCoralBg, BrushChipCoralFg);
            return;
        }
        var mapping = LoadAccountTypeMapping(path);
        if (mapping == null)
            SetChip(settingsTypeMappingChip, "読み込めません", BrushChipCoralBg, BrushChipCoralFg);
        else
            SetChip(settingsTypeMappingChip, "金融機関 " + mapping.Institutions.Count + (mapping.Global != null ? "・global" : ""),
                BrushChipTealBg, BrushChipTealFg);
    }

    // 台帳フォルダの注記（現在の入力値での検出先と台帳ファイル数）を更新する
    private void UpdateLedgerNote(SettingsKindRows rows)
    {
        var folder = rows.LedgerBox.Text.Trim();
        var resolved = folder == "" ? exeDir : folder;
        string detected;
        try
        {
            detected = Directory.Exists(resolved)
                ? "台帳 " + Directory.GetFiles(resolved, rows.LedgerPattern).Length + " ファイル"
                : "フォルダが見つかりません";
        }
        catch
        {
            detected = "フォルダを確認できません";
        }
        rows.LedgerNote.Text = "検出先: " + resolved + "　·　" + detected;
    }

    // 保存: 入力値の整形 → config JSON の全文再生成 → アトミック書き込み → 反映
    private void SaveSettings()
    {
        // recentMax は LoadConfig と同じ規則で正規化する（非数・1未満は既定の100へ）
        int recentMax;
        if (!int.TryParse(settingsRecentMaxBox.Text.Trim(), out recentMax) || recentMax < 1)
            recentMax = 100;

        // 照会結果フォルダ: 年度・パスのどちらかが空欄の行は保存対象外（注記で案内済み）
        var depositFolders = CollectSavedFolderPairs(depositSettingsRows);
        var insuranceFolders = CollectSavedFolderPairs(insuranceSettingsRows);

        // 保存する値を新しい config に組む（対象は生値のまま保持し、照合時に正規化する）
        var next = new AppConfig
        {
            DepositLedgerFolder = depositSettingsRows.LedgerBox.Text.Trim(),
            InsuranceLedgerFolder = insuranceSettingsRows.LedgerBox.Text.Trim(),
            DepositSearchFolders = FolderPairsToMap(depositFolders),
            InsuranceSearchFolders = FolderPairsToMap(insuranceFolders),
            DepositSeizureListScript = depositSettingsRows.ScriptBox.Text.Trim(),
            InsuranceSeizureListScript = insuranceSettingsRows.ScriptBox.Text.Trim(),
            DepositSeizureTargetBanks = depositSettingsRows.TargetBox.Text.Trim(),
            InsuranceSeizureTargetCompanies = insuranceSettingsRows.TargetBox.Text.Trim(),
            RecentMax = recentMax,
            StaffFuzzyChars = ParseStaffFuzzyChars(settingsFuzzyBox.Text),   // LoadConfig と同じ規則で正規化
            AccountTypeMappingFile = settingsTypeMappingBox.Text.Trim(),
            Detail = settingsShownConfig.Detail   // ビューが示しているファイルの値（手で直した detail を保存で巻き戻さない）
        };
        var configPath = System.IO.Path.Combine(exeDir, CONFIG_FILE);

        // ほかの職員が先に保存していれば止める（ビューを組んだ時点の本文と今の本文の照合。
        // 組んだときに読めなかった・いま読めないときは照合せずに保存する）
        if (settingsConfigText != null)
        {
            string current;
            try { current = File.ReadAllText(configPath, Encoding.UTF8); }
            catch { current = null; }   // 読めないときは照合しない（書けなければ下の「保存できませんでした」）
            if (current != null && current != settingsConfigText)
            {
                MessageBox.Show(window,
                    "ほかの人が設定を保存したため、保存しませんでした。\n" + configPath + "\n\n" +
                    "「変更を破棄」で最新の設定を読み直してから、もう一度変更してください。",
                    "設定の保存", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;   // 編集状態はそのまま残す（控えて書き直せる）
            }
        }

        if (!WriteConfigFile(next, depositFolders, insuranceFolders))
        {
            MessageBox.Show(window,
                "設定を保存できませんでした。\n" +
                configPath + "\n\n" +
                "ファイルが他のプログラムで開かれていないか確認してください。",
                "設定の保存", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;   // 編集状態はそのまま残す（やり直せる）
        }

        // --- 変わった項目を控えてから、メモリ上の config を差し替える ---
        //（比べる相手はこの PC で有効なメモリ上の値。ほかの職員が変えた項目も、この保存で反映すべき差として拾う）
        bool depositLedgerChanged = (config.DepositLedgerFolder ?? "").Trim() != next.DepositLedgerFolder;
        bool insuranceLedgerChanged = (config.InsuranceLedgerFolder ?? "").Trim() != next.InsuranceLedgerFolder;
        bool depositScriptChanged = (config.DepositSeizureListScript ?? "").Trim() != next.DepositSeizureListScript;
        bool insuranceScriptChanged = (config.InsuranceSeizureListScript ?? "").Trim() != next.InsuranceSeizureListScript;
        bool fuzzyChanged = FormatStaffFuzzyChars(config.StaffFuzzyChars) != FormatStaffFuzzyChars(next.StaffFuzzyChars);
        bool depositTargetChanged = (config.DepositSeizureTargetBanks ?? "").Trim() != next.DepositSeizureTargetBanks;
        bool insuranceTargetChanged = (config.InsuranceSeizureTargetCompanies ?? "").Trim() != next.InsuranceSeizureTargetCompanies;
        bool typeMappingChanged = (config.AccountTypeMappingFile ?? "").Trim() != next.AccountTypeMappingFile;
        config = next;

        // --- 即時反映 ---
        // 照会結果フォルダはファイル解決の都度 config を参照するため差し替えのみで反映される
        UpdateRecentSubtitle();
        if (depositScriptChanged) depositSettingsRows.ScriptRestartPending = true;     // 連携UIは起動時に構築されるため次回起動で反映
        if (insuranceScriptChanged) insuranceSettingsRows.ScriptRestartPending = true;
        if (fuzzyChanged)
        {
            // 担当クラスタとトグル群を組み直す（選択中の条件は同一クラスタが残れば維持——
            // 年度切替と同じ規則。台帳の再読込は不要で、列挙のやり直しだけで反映される）。
            // 生命保険の分は読み込み済みのときだけ（未読み込みなら読み込むときに新しい設定で組まれる）
            RebuildDepositStaffToggles();
            OnDepositAdvancedFilterChanged();
            if (insuranceTab.IsLoaded)
            {
                insuranceTab.RebuildInsuranceStaffToggles();
                insuranceTab.OnInsuranceAdvancedFilterChanged();
            }
        }
        if (depositTargetChanged)
        {
            // 対象金融機関の照合パターンを組み直す（右クリックメニューは開くたびに
            // 判定するため差し替えのみ・詳細パネルを開いたままの保存はボタンの追従が必要）
            RebuildDepositSeizureTargets();
            if (isDepositPanelOpen) UpdateDepositSeizureButton();
        }
        if (insuranceTargetChanged)
        {
            insuranceTab.RebuildInsuranceSeizureTargets();
            if (insuranceTab.isInsurancePanelOpen) insuranceTab.UpdateInsuranceSeizureButton();
        }
        if (typeMappingChanged)
        {
            // 明細の口座種別の照合に使う表を読み直し、古い表で読んだ明細の結果とウィンドウを捨てる
            detailTypeMapping = LoadAccountTypeMapping(config.AccountTypeMappingFile);
            ResetDetailResults();
        }
        if (depositLedgerChanged)
        {
            // 台帳の検出をやり直す。人物キャッシュは別フォルダの内容と混ざらないよう捨てる
            depositLedgerFolder = config.DepositLedgerFolder == "" ? exeDir : config.DepositLedgerFolder;
            depositOtherYearCache.Clear();
            ResetDetailResults();
            if (isDepositPanelOpen) CloseDepositDetailPanel();
            depositPanelAddressNumberShown = null;   // 開き直したら組み直す（捨てる前に組んだパネルを使い回さない）
            DetectDepositLedgerFilesThenPopulate(null);
        }
        if (insuranceLedgerChanged)
            insuranceTab.OnInsuranceLedgerFolderChanged(config.InsuranceLedgerFolder == "" ? exeDir : config.InsuranceLedgerFolder);

        // 保存済みの状態でビューを組み直す（スナップショット更新・保存バー消灯・チップ再判定）
        BuildSettingsView();
        ShowSettingsFooterTransient(
            "保存しました（" + DateTime.Now.ToString("yyyy/MM/dd HH:mm", CultureInfo.InvariantCulture) + "）　·　" + CONFIG_FILE);
    }

    // 保存する照会結果フォルダの {年度, パス} の組（年度・パスのどちらかが空欄の行は除く）
    private static List<string[]> CollectSavedFolderPairs(SettingsKindRows rows)
    {
        var folders = new List<string[]>();
        foreach (var row in rows.FolderRows)
        {
            var year = row.YearBox.Text.Trim();
            var path = row.PathBox.Text.Trim();
            if (year == "" || path == "") continue;
            folders.Add(new[] { year, path });
        }
        return folders;
    }

    // {年度, パス} の組を config の形（年度 → フォルダの並び）にする。0件は null（LoadConfig と同じ扱い）
    private static Dictionary<string, List<string>> FolderPairsToMap(List<string[]> folders)
    {
        var map = new Dictionary<string, List<string>>();
        foreach (var pair in folders)
        {
            if (!map.ContainsKey(pair[0])) map[pair[0]] = new List<string>();
            map[pair[0]].Add(pair[1]);
        }
        return map.Count > 0 ? map : null;
    }

    // paddock_config.json を全文再生成して書き込む。一時ファイルへ書いてから
    // 置き換えるアトミック方式+リトライ5回（台帳追記と同じ規約。共有フォルダ上の
    // 設定ファイルを中途半端な状態にしない）
    private bool WriteConfigFile(AppConfig next, List<string[]> depositFolders, List<string[]> insuranceFolders)
    {
        var json = BuildConfigJson(next, depositFolders, insuranceFolders);
        var path = System.IO.Path.Combine(exeDir, CONFIG_FILE);
        var temp = path + ".tmp";

        for (int attempt = 0; attempt < 5; attempt++)
        {
            try
            {
                File.WriteAllText(temp, json, Encoding.UTF8);
                if (!File.Exists(path))
                    File.Move(temp, path);
                else
                {
                    // File.Replace が使えないボリューム（一部のネットワーク共有）では上書きコピーに退行する
                    //（コピーも失敗する場合＝ロック等は外側の catch でリトライに回る）
                    try { File.Replace(temp, path, null); }
                    catch
                    {
                        File.Copy(temp, path, true);
                        File.Delete(temp);
                    }
                }
                return true;
            }
            catch
            {
                System.Threading.Thread.Sleep(100);
            }
        }
        try { if (File.Exists(temp)) File.Delete(temp); }
        catch { /* 一時ファイルの掃除失敗は無視する */ }
        return false;
    }

    // 設定ファイルの本文を組み立てる。_README はこの定型文で毎回上書きされる
    //（解説の変更はここを直す — 手書きの config と二重管理にしない）。
    // 照会結果フォルダは行の並びを保って書くため、種別ごとの {年度, パス} の組を別に受け取る
    private static string BuildConfigJson(AppConfig c, List<string[]> depositFolders, List<string[]> insuranceFolders)
    {
        var readme = new[]
        {
            "【設定項目の解説】",
            "この設定はナビ「設定」から編集・保存できる（保存するとファイル全体が書き直され、この解説も定型文に戻る）。",
            "テキストエディタでの直接編集も引き続き可能。",
            "預金の項目は deposit、生命保険の項目は insurance で始まる（台帳フォルダ・照会結果フォルダ・",
            "差押リスト連携・差押リストの対象の4つを対で持つ）。",
            "",
            "■ depositLedgerFolder（任意）",
            "   預金の台帳CSV（財産調査台帳_YYYY年度.csv）の検出フォルダの絶対パス。",
            "   未設定または空欄の場合は paddock.exe と同一フォルダから検出する。",
            "   JUDGE（judge.ps1）の預金のプロファイルの ledgerFolder 設定と同じフォルダを指すこと。",
            "",
            "■ depositSearchFolders（ドリルダウン機能で使用・任意）",
            "   年度→預金の照会結果フォルダの明示マッピング。毎年1エントリ追記する運用。",
            "   値は文字列（単一フォルダ）または配列（複数フォルダ）で指定可能。",
            "   未設定の場合、預金の照会ファイルを開く操作と差押リスト連携は無効化される",
            "   （一覧・検索・詳細パネルは動作する）。",
            "",
            "■ depositSeizureListScript（差押リスト連携で使用・任意）",
            "   預金差押予定一覧 作成ツール（entry_deposit.exe）の絶対パス。",
            "   未設定または該当ファイルが存在しない場合、預金の連携系UI（右クリックメニューの",
            "   「差押リストに追加」・詳細パネルのボタン）は非表示になる。",
            "   ナビ「差押登録」は預金・生命保険のどちらかの連携先があれば表示される。",
            "   差押登録ログ（history/{ユーザー名}/seizure_log.json）は PADDOCK 起点の履歴で、",
            "   ZEKKEN のログとは独立に保管される。",
            "",
            "■ depositSeizureTargetBanks（差押リスト連携で使用・任意）",
            "   カンマ区切りで指定する。対象を限定する場合に指定し、ワイルドカード可",
            "   （* = 任意の文字列・? = 任意の1文字。例: *信用金庫。全角の＊？も可）。",
            "   ワイルドカードを含まない項目は完全一致（全半角・空白のゆれは吸収）。",
            "   一致しない金融機関の行では「差押リストに追加」が無効になる。",
            "   未設定または空欄の場合は制御しない（全行で追加可）。",
            "",
            "■ insuranceLedgerFolder（任意）",
            "   生命保険の台帳CSV（生命保険調査台帳_YYYY年度.csv）の検出フォルダの絶対パス。",
            "   未設定または空欄の場合は paddock.exe と同一フォルダから検出する。",
            "   JUDGE（judge.ps1）の生命保険のプロファイルの ledgerFolder 設定と同じフォルダを指すこと。",
            "",
            "■ insuranceSearchFolders（ドリルダウン機能で使用・任意）",
            "   年度→生命保険の照会結果フォルダの明示マッピング（書き方は depositSearchFolders と同じ）。",
            "   未設定の場合、生命保険の照会ファイルを開く操作と差押リスト連携は無効化される",
            "   （一覧・検索・詳細パネルは動作する）。",
            "   預金の照会結果フォルダの配下には置かない（候補の収集はフォルダ配下を再帰的に集めるため）。",
            "",
            "■ insuranceSeizureListScript（差押リスト連携で使用・任意）",
            "   生命保険差押予定一覧 作成ツール（entry_insurance.exe）の絶対パス。",
            "   未設定または該当ファイルが存在しない場合、生命保険の連携系UI（右クリックメニューの",
            "   「差押リストに追加」・詳細パネルのボタン）は非表示になる。",
            "",
            "■ insuranceSeizureTargetCompanies（差押リスト連携で使用・任意）",
            "   差押リストの対象保険会社。書き方と扱いは depositSeizureTargetBanks と同じ（保険会社名で照合する）。",
            "",
            "■ recentMax（任意・既定値: 100）",
            "   ナビ「最近見た人」の履歴の保持件数の上限。年度セクションごとに適用され、",
            "   超えた分は閲覧日時の古いものから削除される。",
            "   履歴（history/{ユーザー名}/recent_people.json）はユーザー単位で保管される。",
            "",
            "■ staffFuzzyChars（担当の絞り込みで使用・任意・既定値: ◆）",
            "   カンマ区切りで指定する。照会結果の担当欄で外字が置き換わる文字",
            "   （化け置換文字）を指定すると、詳細検索「担当」の列挙でこれらを",
            "   任意の1文字とみなして表記ゆれを統合する（預金・生命保険とも）。",
            "   空欄にすると化け文字による統合を行わない（空白の揺れの吸収と",
            "   姓のみの前方一致による統合は引き続き行われる）。",
            "",
            "■ accountTypeMappingFile（明細を見るで使用・任意・預金のみ）",
            "   JUDGE（judge.ps1）の口座種別の補正表（account_type_mapping.json）の絶対パス。",
            "   台帳の口座種別は JUDGE がこの表で補正した表記のため、指定すると明細を見るときに",
            "   照会ファイルの明細ページを補正前・補正後のどちらの表記でも見つけられる。",
            "   未設定・空欄・読めない場合は明細ページの表記のまま照合する。PADDOCK はこの表を書き換えない。",
            "",
            "■ detail（明細を見るで使用・任意・預金のみ・通常は変更不要）",
            "   詳細パネルの口座カードから照会シートの取引明細を読む位置。明細ページ先頭の",
            "   目印セル（A列）を1行目とした相対指定で、JUDGE（judge_config.json）の",
            "   detail* 設定と同じ位置を指すこと。stopValues=目印文字列の配列、",
            "   accountNumberCell/accountTypeCell=口座番号・口座種別のセル、",
            "   dataStartRow=明細の開始行、dateCol〜balanceCol=取引日・摘要・出金・入金・残高の列。",
            "   キー単位で省略でき、省略・不正値は既定（JUDGE の標準設定と同値）に戻る。",
            "   この項目はナビ「設定」では表示のみ。変更はテキストエディタで行い、再起動後に反映される。",
            "",
            "※ アップデート告知（update_notice.json）の既読情報はこのファイルではなく",
            "   history/{ユーザー名}/notice_state.json にユーザー単位で保管される。"
        };

        var sb = new StringBuilder();
        sb.AppendLine("{");
        sb.AppendLine("  \"_README\": [");
        for (int i = 0; i < readme.Length; i++)
            sb.AppendLine("    \"" + EscapeJson(readme[i]) + "\"" + (i < readme.Length - 1 ? "," : ""));
        sb.AppendLine("  ],");

        // 預金 → 生命保険 → 共通の並び（設定ビューのカードの並びと同じ）
        sb.AppendLine("  \"depositLedgerFolder\": \"" + EscapeJson(c.DepositLedgerFolder) + "\",");
        AppendSearchFoldersJson(sb, "depositSearchFolders", depositFolders);
        sb.AppendLine("  \"depositSeizureListScript\": \"" + EscapeJson(c.DepositSeizureListScript) + "\",");
        sb.AppendLine("  \"depositSeizureTargetBanks\": \"" + EscapeJson(c.DepositSeizureTargetBanks) + "\",");
        sb.AppendLine("  \"insuranceLedgerFolder\": \"" + EscapeJson(c.InsuranceLedgerFolder) + "\",");
        AppendSearchFoldersJson(sb, "insuranceSearchFolders", insuranceFolders);
        sb.AppendLine("  \"insuranceSeizureListScript\": \"" + EscapeJson(c.InsuranceSeizureListScript) + "\",");
        sb.AppendLine("  \"insuranceSeizureTargetCompanies\": \"" + EscapeJson(c.InsuranceSeizureTargetCompanies) + "\",");
        sb.AppendLine("  \"recentMax\": " + c.RecentMax + ",");
        sb.AppendLine("  \"staffFuzzyChars\": \"" + EscapeJson(FormatStaffFuzzyChars(c.StaffFuzzyChars)) + "\",");
        sb.AppendLine("  \"accountTypeMappingFile\": \"" + EscapeJson(c.AccountTypeMappingFile) + "\",");

        // detail（設定ビューでは編集しないため、読み込んだ値をそのまま書き戻す）
        var detail = c.Detail;
        sb.AppendLine("  \"detail\": {");
        sb.Append("    \"stopValues\": [");
        for (int i = 0; i < detail.StopValues.Count; i++)
        {
            if (i > 0) sb.Append(", ");
            sb.Append("\"" + EscapeJson(detail.StopValues[i]) + "\"");
        }
        sb.AppendLine("],");
        sb.AppendLine("    \"accountNumberCell\": \"" + EscapeJson(detail.AccountNumberCell) + "\",");
        sb.AppendLine("    \"accountTypeCell\": \"" + EscapeJson(detail.AccountTypeCell) + "\",");
        sb.AppendLine("    \"dataStartRow\": " + detail.DataStartRow + ",");
        sb.AppendLine("    \"dateCol\": \"" + EscapeJson(detail.DateCol) + "\",");
        sb.AppendLine("    \"summaryCol\": \"" + EscapeJson(detail.SummaryCol) + "\",");
        sb.AppendLine("    \"withdrawalCol\": \"" + EscapeJson(detail.WithdrawalCol) + "\",");
        sb.AppendLine("    \"depositCol\": \"" + EscapeJson(detail.DepositCol) + "\",");
        sb.AppendLine("    \"balanceCol\": \"" + EscapeJson(detail.BalanceCol) + "\"");
        sb.AppendLine("  }");
        sb.AppendLine("}");
        return sb.ToString();
    }

    // 照会結果フォルダ（年度 → フォルダ）を書き出す。0件のときはキー自体を書かない（未設定と同じ扱い）。
    // 行の並びを保ったまま年度ごとにまとめる（同一年度の複数行は配列になる）
    private static void AppendSearchFoldersJson(StringBuilder sb, string key, List<string[]> folders)
    {
        if (folders.Count == 0) return;
        var order = new List<string>();
        var map = new Dictionary<string, List<string>>();
        foreach (var pair in folders)
        {
            if (!map.ContainsKey(pair[0])) { map[pair[0]] = new List<string>(); order.Add(pair[0]); }
            map[pair[0]].Add(pair[1]);
        }
        sb.AppendLine("  \"" + key + "\": {");
        for (int i = 0; i < order.Count; i++)
        {
            var list = map[order[i]];
            sb.Append("     \"" + EscapeJson(order[i]) + "\": ");
            if (list.Count == 1)
                sb.Append("\"" + EscapeJson(list[0]) + "\"");
            else
            {
                sb.Append("[");
                for (int j = 0; j < list.Count; j++)
                {
                    if (j > 0) sb.Append(", ");
                    sb.Append("\"" + EscapeJson(list[j]) + "\"");
                }
                sb.Append("]");
            }
            sb.AppendLine(i < order.Count - 1 ? "," : "");
        }
        sb.AppendLine("  },");
    }

    // 設定フッター左側の一時表示（台帳ビューの ShowTransientStatus と同じ規約・5秒で戻す）
    private void ShowSettingsFooterTransient(string message)
    {
        if (settingsFooterTimer == null)
        {
            settingsFooterTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
            settingsFooterTimer.Tick += delegate
            {
                settingsFooterTimer.Stop();
                settingsFooterLeft.Text = settingsFooterNormal;
            };
        }
        settingsFooterLeft.Text = message;
        settingsFooterTimer.Stop();
        settingsFooterTimer.Start();
    }

    // --- 設定ビューの UI 部品 ---

    // カード1枚（セクションヘッダー+本文）を作る。差押登録ビューの白カードと同じ意匠
    private static Border MakeSettingsCard(string title, string note, out StackPanel body)
    {
        var card = new Border
        {
            Background = Brushes.White,
            BorderBrush = BrushLine,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(10),
            Margin = new Thickness(0, 0, 0, 12)
        };
        ApplyRoundedClip(card, 10);
        var stack = new StackPanel();

        var head = new Border
        {
            Background = BrushSurfaceAlt,
            BorderBrush = BrushLineWeak,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(16, 8, 16, 8)
        };
        var headGrid = new Grid();
        headGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        headGrid.Children.Add(new TextBlock
        {
            Text = title,
            FontSize = 12,
            Foreground = BrushTextMain
        });
        if (note != null)
        {
            var noteBlock = new TextBlock
            {
                Text = note,
                FontSize = 10.5,
                Foreground = BrushTextMuted,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(noteBlock, 1);
            headGrid.Children.Add(noteBlock);
        }
        head.Child = headGrid;
        stack.Children.Add(head);

        body = new StackPanel();
        stack.Children.Add(body);
        card.Child = stack;
        return card;
    }

    // カード内の1行（左: 項目名+configキー / 右: 値）を追加する
    private static void AddSettingsRow(StackPanel body, string label, string cfgKey, UIElement value, bool isLast)
    {
        var row = new Border
        {
            BorderBrush = BrushLineWeak,
            BorderThickness = new Thickness(0, 0, 0, isLast ? 0 : 1),
            Padding = new Thickness(16, 10, 16, 10)
        };
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(190) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var key = new StackPanel { Margin = new Thickness(0, 3, 12, 0) };
        key.Children.Add(new TextBlock { Text = label, FontSize = 12, Foreground = BrushTextSub });
        if (cfgKey != null)
        {
            key.Children.Add(new TextBlock
            {
                Text = cfgKey,
                FontSize = 10,
                FontFamily = FontMono,
                Foreground = BrushTextMuted,
                Margin = new Thickness(0, 2, 0, 0)
            });
        }
        grid.Children.Add(key);

        Grid.SetColumn(value, 1);
        grid.Children.Add(value);
        row.Child = grid;
        body.Children.Add(row);
    }

    // 値の下に置く注記（淡色・折り返しあり）を作る
    private static TextBlock MakeSettingsNote(string text)
    {
        return new TextBlock
        {
            Text = text,
            FontSize = 10.5,
            Foreground = BrushTextMuted,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 5, 0, 0)
        };
    }

    // 設定の入力欄（白地+罫線+角丸7。フォーカス中はアイリス枠）を作る
    private Border MakeSettingsInput(string text, out TextBox box, double fixedWidth)
    {
        var textBox = new TextBox
        {
            Text = text,
            BorderThickness = new Thickness(0),
            Background = Brushes.Transparent,
            FontFamily = FontMono,
            FontSize = 12,
            Foreground = BrushTextMain,
            VerticalContentAlignment = VerticalAlignment.Center
        };
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = BrushLine,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(7),
            Padding = new Thickness(9, 5, 9, 5),
            Child = textBox,
            VerticalAlignment = VerticalAlignment.Center
        };
        if (fixedWidth > 0) border.Width = fixedWidth;
        textBox.GotKeyboardFocus += delegate { border.BorderBrush = BrushAccent; };
        textBox.LostKeyboardFocus += delegate { border.BorderBrush = BrushLine; };
        textBox.TextChanged += delegate { OnSettingsInputChanged(); };
        box = textBox;
        return border;
    }

    // 小型ボタン（参照… 等。スタイルは XAML の SettingsMiniBtn）
    private Button MakeMiniButton(string text)
    {
        return new Button
        {
            Content = text,
            Style = (Style)window.FindResource("SettingsMiniBtn"),
            VerticalAlignment = VerticalAlignment.Center
        };
    }

    // 行削除の ✕（詳細パネルの閉じるボタンと同じ意匠。Border+イベントで組む）
    private static Border MakeRowDeleteButton()
    {
        var btn = new Border
        {
            Width = 22,
            Height = 22,
            CornerRadius = new CornerRadius(11),
            Background = Brushes.Transparent,
            Cursor = Cursors.Hand,
            VerticalAlignment = VerticalAlignment.Center,
            ToolTip = "この年度を削除",
            Child = new TextBlock
            {
                Text = "✕",
                FontSize = 10,
                Foreground = BrushTextSub,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            }
        };
        btn.MouseEnter += delegate { btn.Background = BrushChipNeutralBg; };
        btn.MouseLeave += delegate { btn.Background = Brushes.Transparent; };
        return btn;
    }

    // チップ（年度セクションの件数チップと同じ形状）を作る。文言と配色は SetChip で入れる
    //（SetChip が呼ばれるまでは非表示 — 空のチップを描かない）
    private static Border MakeChip()
    {
        var chip = new Border
        {
            CornerRadius = new CornerRadius(9),
            Padding = new Thickness(8, 1, 8, 2),
            VerticalAlignment = VerticalAlignment.Center,
            Visibility = Visibility.Collapsed
        };
        chip.Child = new TextBlock { FontSize = 10 };
        return chip;
    }

    // チップの文言と配色を設定する（text=null で非表示）
    private static void SetChip(Border chip, string text, Brush bg, Brush fg)
    {
        if (text == null)
        {
            chip.Visibility = Visibility.Collapsed;
            return;
        }
        chip.Visibility = Visibility.Visible;
        chip.Background = bg;
        var label = (TextBlock)chip.Child;
        label.Text = text;
        label.Foreground = fg;
    }

    // パス+ニュートラルチップの1行（履歴ファイルカードの読み取り専用の提示用）
    private static UIElement MakePathWithChip(string path, string chipText)
    {
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.Children.Add(new TextBlock
        {
            Text = path,
            FontFamily = FontMono,
            FontSize = 12,
            Foreground = BrushTextMain,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            ToolTip = path,
            Margin = new Thickness(0, 2, 0, 2)
        });
        var chip = MakeChip();
        SetChip(chip, chipText, BrushChipNeutralBg, BrushChipNeutralFg);
        chip.Margin = new Thickness(8, 0, 0, 0);
        Grid.SetColumn(chip, 1);
        grid.Children.Add(chip);
        return grid;
    }

    // フォルダ参照ダイアログ（System.Windows.Forms.dll が必要 — paddock_build.bat に設定済み）
    private static void BrowseForFolder(TextBox target)
    {
        var dlg = new System.Windows.Forms.FolderBrowserDialog();
        var current = target.Text.Trim();
        if (current != "" && Directory.Exists(current)) dlg.SelectedPath = current;
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            target.Text = dlg.SelectedPath;
    }

    // ファイル参照ダイアログ（差押リスト連携先の exe・口座種別の補正表）
    private static void BrowseForFile(TextBox target, string filter, string title)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = filter,
            Title = title
        };
        var current = target.Text.Trim();
        if (current != "" && File.Exists(current))
            dlg.InitialDirectory = System.IO.Path.GetDirectoryName(current);
        if (dlg.ShowDialog() == true) target.Text = dlg.FileName;
    }
}