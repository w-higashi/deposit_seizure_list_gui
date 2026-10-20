// ==============================================================================
// paddock_deposit_ledger.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 預金タブ: 年度の選択・台帳の読み込み・一覧の行と検出チップの組み立て・検出列の幅。
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
    // 台帳ファイルの検出・年度セレクタ
    // ==============================================================

    // 台帳ファイルの検出をワーカーで行い、終わってから年度セレクタを組んで読む（起動・F5・台帳フォルダの
    // 変更の入口。共有フォルダが遅い・切断中でもウィンドウは止まらない）。入口で世代を進めるため、走っていた
    // 古い検出・読み込み・走査の完了は合流しない。検出に失敗したら台帳 0 本として整えてから文言を差し替える
    private void DetectDepositLedgerFilesThenPopulate(string preferredItem)
    {
        var epoch = ++depositLedgerEpoch;
        var folder = depositLedgerFolder;
        depositLedgerFiles = new Dictionary<int, string>();   // 検出が終わるまで台帳 0 本（検出中に開いた人物の走査が旧フォルダのファイルを読まない）
        depositYearCombo.IsEnabled = false;
        depositEmptyStateText.Visibility = Visibility.Collapsed;
        ShowLoadingOverlay(depositLoadingOverlay, depositSpinnerRotation);

        depositDetectWorker = new BackgroundWorker();
        depositDetectWorker.DoWork += delegate(object s, DoWorkEventArgs e)
        {
            e.Result = DetectLedgerFiles(folder, DEPOSIT_LEDGER_PATTERN, DepositLedgerYearRegex);
        };
        depositDetectWorker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs e)
        {
            if (epoch != depositLedgerEpoch) return;   // そのあとに F5・台帳フォルダの変更があった古い検出は合流しない
            if (e.Error != null)
            {
                WriteErrorLog("台帳の検出（預金）", e.Error);
                PopulateDepositYearSelector(null);
                ShowEmptyState(depositEmptyStateText, "台帳フォルダを読めませんでした:\n" + e.Error.Message);
                SetStatusText(depositStatusLeft, "読み込みエラー");
                return;
            }
            depositLedgerFiles = (Dictionary<int, string>)e.Result;
            PopulateDepositYearSelector(preferredItem);
            if (isDepositPanelOpen)   // 検出の前に開いた人物の他年度を改めて走査する
            {
                depositCrossYearTimer.Stop();
                depositCrossYearTimer.Start();
            }
        };
        depositDetectWorker.RunWorkerAsync();
    }

    // 年度セレクタを構築して読み込む。既定は「すべての年度」
    //（年度の帰属は登録日基準のため年度替わり直後は当年度の台帳がほぼ空になる。
    //  既定を全年度にして、探すたびに前年度へ切り替えずに済むようにする）。
    // preferredItem に選択中だった項目を渡すと、残っていればそれを選び直す（再読み込み用）
    private void PopulateDepositYearSelector(string preferredItem)
    {
        suppressDepositYearChange = true;
        depositYearCombo.Items.Clear();

        var years = depositLedgerFiles.Keys.OrderByDescending(y => y).ToList();
        if (years.Count > 0) depositYearCombo.Items.Add(ALL_YEARS_LABEL);
        foreach (var year in years)
            depositYearCombo.Items.Add(BusinessLogic.FormatYearKey(year));

        if (years.Count == 0)
        {
            suppressDepositYearChange = false;
            depositYearCombo.IsEnabled = false;
            UpdateRailYearLabel();         // 台帳なしの表示（—）に更新する
            ResetDepositLedgerToEmpty();   // 再読み込みで台帳が消えた場合に前の行を残さない
            ShowEmptyState(depositEmptyStateText, "台帳CSVが見つかりません。\nフォルダ: " + depositLedgerFolder +
                           "\n\nJUDGE（judge.ps1）の預金のプロファイルの ledgerFolder 設定と\n" +
                           "paddock_config.json の depositLedgerFolder が同じフォルダを指しているか確認してください。");
            SetStatusText(depositStatusLeft, "台帳なし");
            HideLoadingOverlay(depositLoadingOverlay, depositSpinnerRotation);   // 検出のワーカーが出したオーバーレイを消す
            return;
        }

        var index = preferredItem != null ? depositYearCombo.Items.IndexOf(preferredItem) : -1;
        depositYearCombo.SelectedIndex = index >= 0 ? index : 0;
        suppressDepositYearChange = false;

        LoadDepositLedger(GetSelectedYear(depositYearCombo));
    }

    // 台帳を読み直す（F5）。ファイルの検出からやり直して年度セレクタを組み直し、選んで
    // いた項目は残っていれば維持する（消えていれば「すべての年度」）。検索欄・詳細検索の
    // 条件は入力欄が持っているため保持され、読み込み完了時の絞り込みで再適用される。
    // 読み込み中に押しても検出からやり直す（古い読み込みの完了は世代で捨てる）
    private void ReloadDepositLedger()
    {
        var current = depositYearCombo.SelectedItem != null ? depositYearCombo.SelectedItem.ToString() : null;
        depositOtherYearCache.Clear();   // 他年度の追記も拾えるよう人物単位のキャッシュは捨てる
        detailTypeMapping = LoadAccountTypeMapping(config.AccountTypeMappingFile);   // JUDGE の補正表の改訂も拾う
        ResetDetailResults();            // 台帳の検出が変わりうるため明細の結果とウィンドウも捨てる
        if (isDepositPanelOpen) CloseDepositDetailPanel();
        depositPanelAddressNumberShown = null;   // 開き直したら組み直す（捨てる前に組んだパネルを使い回さない）
        DetectDepositLedgerFilesThenPopulate(current);
    }

    // 一覧の状態を「台帳0行」へ戻す（読み込みエラー時・台帳ファイルが1本もないとき）。
    // 担当トグルを空にし、解除された担当条件をチップ・バッジへ反映する
    //（一覧の表示はこの後の空状態・エラー文言が担うため絞り込みは行わない）
    private void ResetDepositLedgerToEmpty()
    {
        depositAllRecords = new List<DepositLedgerRecord>();
        depositSkippedRowCount = 0;
        depositLoadFailedYears = new List<int>();
        UpdateDepositDetectColumnWidth();   // 台帳なし=検出列を下限幅へ戻す
        depositStaffRawValues = new List<string>();
        RebuildDepositStaffToggles();
        if (depositSettingsRows.TargetBox != null) UpdateSeizureTargetNote(depositSettingsRows);   // 台帳0行としての一致開示に更新
        ReadDepositAdvancedInputs();
        RebuildDepositFilterChips();
        UpdateDepositAdvToggleVisual();
        depositLedgerList.ItemsSource = null;
    }

    // 年度セレクタ（預金・生命保険のコンボ）の選択中の項目から年度を取得する（未選択・「すべての年度」は null）
    private static int? GetSelectedYear(ComboBox combo)
    {
        if (combo.SelectedItem == null) return null;
        var text = combo.SelectedItem.ToString();
        if (text == ALL_YEARS_LABEL) return null;
        return BusinessLogic.ParseYearKey(text);
    }

    // ==============================================================
    // 台帳の読み込み（BackgroundWorker）
    // ==============================================================

    // 読み込み結果（DoWork → RunWorkerCompleted の受け渡し用）
    private class LoadDepositResult
    {
        public List<DepositLedgerRecord> Records;
        public int SkippedRows;
        public string Label;                  // フッター右の表示（年度指定=ファイル名／すべての年度=範囲と本数）
        public string Error;
        public List<string> StaffRawValues;   // 担当職員名の distinct 生値（初出順・担当トグルの列挙用）
        public List<int> FailedYears;         // 読めなかった年度（すべての年度で一部のファイルが読めなかった場合）
    }

    // 台帳CSVを非同期で読み込む。year=null は「すべての年度」（台帳フォルダの全ファイルを結合）
    private void LoadDepositLedger(int? year)
    {
        // 読む対象を年度の降順に並べる（結合後は安定ソートのため、同一基準日内の並びは
        // 新しい年度→古い年度・各ファイル内は追記順になる）
        var targets = new List<KeyValuePair<int, string>>();
        if (year == null)
        {
            foreach (var y in depositLedgerFiles.Keys.OrderByDescending(k => k))
                targets.Add(new KeyValuePair<int, string>(y, depositLedgerFiles[y]));
        }
        else
        {
            string path;
            if (!depositLedgerFiles.TryGetValue(year.Value, out path)) return;
            targets.Add(new KeyValuePair<int, string>(year.Value, path));
        }
        if (targets.Count == 0) return;
        var epoch = depositLedgerEpoch;   // この読み込みを始めた世代（F5・台帳フォルダの変更で進む。古い世代の完了は合流しない）

        depositYearCombo.IsEnabled = false;
        depositEmptyStateText.Visibility = Visibility.Collapsed;
        ShowLoadingOverlay(depositLoadingOverlay, depositSpinnerRotation);

        depositLoadWorker = new BackgroundWorker();
        depositLoadWorker.DoWork += delegate(object s, DoWorkEventArgs e)
        {
            var result = new LoadDepositResult { FailedYears = new List<int>() };
            result.Label = year != null
                ? System.IO.Path.GetFileName(targets[0].Value)
                : "財産調査台帳（" + (targets.Count == 1
                    ? BusinessLogic.FormatYearKey(targets[0].Key)
                    : targets[targets.Count - 1].Key + "〜" + BusinessLogic.FormatYearKey(targets[0].Key)) +
                  "・" + targets.Count + "本）";

            var records = new List<DepositLedgerRecord>();
            int skipped = 0;
            var staffSeen = new HashSet<string>();
            var staffRaws = new List<string>();
            var pool = new Dictionary<string, string>();   // 同じ値の文字列を 1 本に共有する（この読み込みの間だけ）

            foreach (var target in targets)
            {
                try
                {
                    // ファイル全体を読み込んで1回走査でパースする（途中で失敗したファイルの行は載せない）
                    var text = ReadLedgerText(target.Value);
                    var rawRecords = BusinessLogic.ParseCsvText(text);
                    var fileRecords = new List<DepositLedgerRecord>(rawRecords.Count);
                    int fileSkipped = 0;

                    // 1行目の見出しから台帳の並びへの対応表を作る。見出しがない・必須の列が欠ける
                    // ファイルはその年度を読めない扱いにする。見出しが台帳の並びどおり（列の追加前の
                    // 27列を含む）なら組み直さず、行の列数の規則も従来どおり（27列未満は読み飛ばし・余剰の列は無視）
                    List<string> missing;
                    var header = rawRecords.Count > 0 ? rawRecords[0] : new string[0];
                    var map = BusinessLogic.BuildLedgerColumnMap(header, BusinessLogic.DEPOSIT_LEDGER_COLUMNS, BusinessLogic.DEPOSIT_REQUIRED_COLUMNS, out missing);
                    if (map == null)
                    {
                        throw new InvalidDataException(missing == null
                            ? "見出しがありません"
                            : "見出しに必須の列がありません: " + string.Join("・", missing.ToArray()));
                    }
                    bool identity = BusinessLogic.IsIdentityColumnMap(map, header.Length, BusinessLogic.DEPOSIT_LEDGER_COLUMN_COUNT);

                    for (int i = 1; i < rawRecords.Count; i++)
                    {
                        var fields = rawRecords[i];

                        // 見出しと同じ字句の行は読まない（先頭の列名で見分ける）
                        if (fields.Length > 0 && fields[0] == header[0]) continue;

                        // 列数の合わない行は読み飛ばす（件数のみ記録して落とさない。並びどおりなら27列未満だけ、
                        // 組み直すときは見出しと違う列数の行）
                        if (identity ? fields.Length < BusinessLogic.DEPOSIT_LEDGER_COLUMN_COUNT : fields.Length != header.Length) { fileSkipped++; continue; }
                        if (!identity) fields = BusinessLogic.ReorderLedgerFields(fields, map);

                        BusinessLogic.InternFields(fields, pool);
                        fileRecords.Add(BuildDepositRecord(fields));

                        // 担当トグルの列挙用に生値の distinct を集める（28列目・空欄は対象外）
                        if (fields.Length > BusinessLogic.DEPOSIT_COL_STAFF)
                        {
                            var staffRaw = fields[BusinessLogic.DEPOSIT_COL_STAFF].Trim();
                            if (staffRaw != "" && staffSeen.Add(staffRaw)) staffRaws.Add(staffRaw);
                        }
                    }
                    records.AddRange(fileRecords);
                    skipped += fileSkipped;
                }
                catch (Exception ex)
                {
                    // 年度単位で控えて続行する。読めた分だけ表示し、全て読めなければ読み込みエラー
                    result.FailedYears.Add(target.Key);
                    result.Error = ex.Message;
                }
            }

            if (result.FailedYears.Count < targets.Count)
            {
                result.Error = null;

                // 既定ソート: 回答基準日の降順（口座なし行は登録日で代替）
                // SortKey は正準形（yyyy/MM/dd）のため文字列比較で時系列順になる
                // ※ 安定ソート（OrderBy）を使い、同一基準日内はファイルの追記順
                //   （＝同一人物の口座が連続する並び）を保持する
                result.Records = records.OrderByDescending(r => r.SortKey, StringComparer.Ordinal).ToList();
                result.SkippedRows = skipped;
                result.StaffRawValues = staffRaws;
            }
            e.Result = result;
        };
        depositLoadWorker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs e)
        {
            if (epoch != depositLedgerEpoch) return;   // 新しい読み込みが始まっている（オーバーレイと年度セレクタはそちらが面倒を見る）
            depositYearCombo.IsEnabled = true;

            // DoWork の外で起きた例外（e.Result は null）は読み込みの失敗と同じ経路に流す
            if (e.Error != null) WriteErrorLog("台帳の読み込み（預金）", e.Error);
            var result = e.Error == null ? (LoadDepositResult)e.Result : new LoadDepositResult { Error = e.Error.Message };
            if (result.Error != null)
            {
                ResetDepositLedgerToEmpty();
                ShowEmptyState(depositEmptyStateText, "台帳の読み込みに失敗しました:\n" + result.Error);
                SetStatusText(depositStatusLeft, "読み込みエラー");
                depositStatusRight.Text = result.Label;
                HideLoadingOverlay(depositLoadingOverlay, depositSpinnerRotation);
                return;
            }

            depositAllRecords = result.Records;
            depositSkippedRowCount = result.SkippedRows;
            depositLoadFailedYears = result.FailedYears;
            depositStatusRight.Text = result.Label;

            // 検出列の幅を新しい台帳の最長チップ列に合わせる
            UpdateDepositDetectColumnWidth();

            // 担当トグルを新しい年度の台帳から組み直してから絞り込む（選択中の
            // 担当が新年度に存在しない場合の条件解除をチップ・バッジへ反映する）
            depositStaffRawValues = result.StaffRawValues;
            RebuildDepositStaffToggles();
            OnDepositAdvancedFilterChanged();
            if (depositSettingsRows.TargetBox != null) UpdateSeizureTargetNote(depositSettingsRows);   // 設定の一致開示を新年度で組み直す

            // 一覧への反映が終わってから消灯し、反映後の一覧をフェードイン
            //（反映が終わるまでオーバーレイが覆い続ける）
            HideLoadingOverlay(depositLoadingOverlay, depositSpinnerRotation);
            depositLedgerList.BeginAnimation(UIElement.OpacityProperty,
                new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(100))));
        };
        depositLoadWorker.RunWorkerAsync();
    }

    // ==============================================================
    // 行データの構築（生フィールド → 表示用レコード）
    // ==============================================================

    private DepositLedgerRecord BuildDepositRecord(string[] fields)
    {
        var record = new DepositLedgerRecord();
        record.RawFields = fields;

        var baseDate = fields[BusinessLogic.DEPOSIT_COL_BASE_DATE];
        record.IsNoAccount = fields[BusinessLogic.DEPOSIT_COL_HAS_ACCOUNT] == "0";

        // 回答基準日は正準形（yyyy/MM/dd）に変換して表示する。台帳の値は7桁和暦・
        // 西暦のどちらでもよい（形式非依存）。口座なし行は常に空欄
        record.DisplayBaseDate = BusinessLogic.FormatLedgerDate(baseDate, eraMapping);

        // ソートキー: 回答基準日が空（口座なし行）は登録日で代替。正準形の Ordinal
        // 比較が時系列順になるため、台帳内で日付形式が混在しても並びが崩れない
        //（変換不能な値は生値のままで、並び位置は保証しない）
        record.SortKey = BusinessLogic.FormatLedgerDate(
            baseDate != "" ? baseDate : fields[BusinessLogic.DEPOSIT_COL_REG_DATE], eraMapping);

        // 文字列セルは空のとき null で保持する（空文字列を ToolTip にバインドすると
        // 空のツールチップが表示されるため。null なら非表示になる）
        record.AddressNumber   = fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM];
        record.Name            = NullIfEmpty(fields[BusinessLogic.DEPOSIT_COL_NAME]);
        record.Kana            = NullIfEmpty(fields[BusinessLogic.DEPOSIT_COL_KANA]);
        record.InstitutionName = NullIfEmpty(fields[BusinessLogic.DEPOSIT_COL_INSTITUTION]);
        record.BranchName      = NullIfEmpty(fields[BusinessLogic.DEPOSIT_COL_BRANCH_NAME]);
        record.AccountType     = NullIfEmpty(fields[BusinessLogic.DEPOSIT_COL_ACCOUNT_TYPE]);
        record.AccountNumber   = fields[BusinessLogic.DEPOSIT_COL_ACCOUNT_NUM];
        record.DisplayLastTransaction = fields[BusinessLogic.DEPOSIT_COL_LAST_TX];

        // 検索用の正規化キー（半角/全角・空白の揺れを吸収した部分一致用）
        record.SearchAddressNumber = BusinessLogic.NormalizeForSearch(fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM]);
        record.SearchName          = BusinessLogic.NormalizeForSearch(fields[BusinessLogic.DEPOSIT_COL_NAME]);
        record.SearchKana          = BusinessLogic.NormalizeForSearch(fields[BusinessLogic.DEPOSIT_COL_KANA]);
        record.SearchInstitution   = BusinessLogic.NormalizeForSearch(fields[BusinessLogic.DEPOSIT_COL_INSTITUTION]);

        // 詳細検索用の事前計算。日付は正準形（yyyy/MM/dd）に変換済みの
        // 表示値を日付化する（パース不能・空欄は null=当該条件の設定時に対象外）
        record.Classification = fields[BusinessLogic.DEPOSIT_COL_RECOMMEND];
        record.BalanceValue   = BusinessLogic.ParseBalanceValue(fields[BusinessLogic.DEPOSIT_COL_BALANCE]);
        record.BaseDateValue  = BusinessLogic.ParseDisplayDate(record.DisplayBaseDate);
        record.LastTxValue    = BusinessLogic.ParseDisplayDate(record.DisplayLastTransaction);

        // 担当職員名（28列目）。旧27列の行・空欄は null（=担当条件の対象外）
        record.StaffKey = fields.Length > BusinessLogic.DEPOSIT_COL_STAFF
            ? NullIfEmpty(BusinessLogic.NormalizeForSearch(fields[BusinessLogic.DEPOSIT_COL_STAFF]))
            : null;

        // 残高: センチネル値（金融機関の回答コード）は中立ラベルに置換し、生値はツールチップで提示する
        var balance = fields[BusinessLogic.DEPOSIT_COL_BALANCE];
        if (BusinessLogic.IsSentinelBalance(balance))
        {
            record.DisplayBalance = "対象外";
            record.BalanceToolTip = "回答コード: " + balance;
        }
        else
        {
            record.DisplayBalance = balance;
            record.BalanceToolTip = null;
        }

        BuildDepositDetectionChips(record, fields);

        // 人物メモの有無（起動時に読み込んだ全ユーザー分の件数表）
        int noteCount;
        if (noteCountByAddress.TryGetValue(fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM], out noteCount) && noteCount > 0)
        {
            record.HasNotes = true;
            record.NotesToolTip = "メモ " + noteCount + "件";
        }
        return record;
    }

    // 検出チップ（第1層）とツールチップ行（第2層）を構築する
    // 走査順は固定: 反対債権 → 年金 → 給与 → 保険 → 定期入金 → 定期出金
    private void BuildDepositDetectionChips(DepositLedgerRecord record, string[] fields)
    {
        var chips = new List<ChipInfo>();
        var lines = new List<TooltipLineInfo>();
        var flags = new bool[6];   // 検出の有無（DETECTION_NAMES と同順。詳細検索の検出条件用）

        // --- 反対債権（4種のいずれかに値があれば1チップに集約）---
        var counterClaims = new[]
        {
            new { Name = "証書貸付", Value = fields[BusinessLogic.DEPOSIT_COL_LOAN_DEED] },
            new { Name = "手形貸付", Value = fields[BusinessLogic.DEPOSIT_COL_LOAN_BILL] },
            new { Name = "割引手形", Value = fields[BusinessLogic.DEPOSIT_COL_LOAN_DISCOUNT] },
            new { Name = "当座貸越", Value = fields[BusinessLogic.DEPOSIT_COL_OVERDRAFT] }
        };
        bool hasCounterClaim = false;
        foreach (var claim in counterClaims)
        {
            if (!BusinessLogic.HasCounterClaimValue(claim.Value)) continue;
            hasCounterClaim = true;
            lines.Add(new TooltipLineInfo
            {
                ChipText = claim.Name, ChipBg = BrushTipCoralBg, ChipFg = BrushTipCoralFg,
                Content = claim.Value.Trim(), ContentFont = FontMono
            });
        }
        if (hasCounterClaim)
            chips.Add(new ChipInfo { Text = "反対債権", Bg = BrushChipCoralBg, Fg = BrushChipCoralFg });
        flags[0] = hasCounterClaim;

        // --- 取引5カテゴリ（1カテゴリ=1チップ・生バケットを無加工で表示）---
        var categories = new[]
        {
            new { Name = "年金",     Value = fields[BusinessLogic.DEPOSIT_COL_PENSION],      Bg = BrushChipTealBg,    Fg = BrushChipTealFg,    TipBg = BrushTipTealBg,    TipFg = BrushTipTealFg },
            new { Name = "給与",     Value = fields[BusinessLogic.DEPOSIT_COL_SALARY],       Bg = BrushChipVioletBg,  Fg = BrushChipVioletFg,  TipBg = BrushTipVioletBg,  TipFg = BrushTipVioletFg },
            new { Name = "保険",     Value = fields[BusinessLogic.DEPOSIT_COL_INSURANCE],    Bg = BrushChipAmberBg,   Fg = BrushChipAmberFg,   TipBg = BrushTipAmberBg,   TipFg = BrushTipAmberFg },
            new { Name = "定期入金", Value = fields[BusinessLogic.DEPOSIT_COL_REG_DEPOSIT],  Bg = BrushChipNeutralBg, Fg = BrushChipNeutralFg, TipBg = BrushTipNeutralBg, TipFg = BrushTipNeutralFg },
            new { Name = "定期出金", Value = fields[BusinessLogic.DEPOSIT_COL_REG_WITHDRAW], Bg = BrushChipNeutralBg, Fg = BrushChipNeutralFg, TipBg = BrushTipNeutralBg, TipFg = BrushTipNeutralFg }
        };
        for (int i = 0; i < categories.Length; i++)
        {
            var category = categories[i];
            if (string.IsNullOrWhiteSpace(category.Value)) continue;
            var raw = category.Value.Trim();
            flags[i + 1] = true;   // DETECTION_NAMES は先頭が反対債権のため1つずれる

            // チップ文字列: カテゴリ名 + 先頭件のバケット（解析不能時はカテゴリ名のみに退行）
            // + 複数件なら「×N」を付記
            var entries = BusinessLogic.SplitDetectionEntries(raw);
            var bucket = BusinessLogic.ParseDetectionBucket(entries[0]);
            var text = category.Name;
            if (bucket != null) text += " " + bucket;
            if (entries.Length >= 2) text += " ×" + entries.Length;

            chips.Add(new ChipInfo { Text = text, Bg = category.Bg, Fg = category.Fg });
            lines.Add(new TooltipLineInfo
            {
                ChipText = category.Name, ChipBg = category.TipBg, ChipFg = category.TipFg,
                Content = raw, ContentFont = FontUi   // 原文は UI フォント（金額行のみ Consolas）
            });
        }

        // --- 第1層はチップ最大3個+超過件数「+N」---
        if (chips.Count > 3)
        {
            var overflow = chips.Count - 3;
            chips = chips.Take(3).ToList();
            chips.Add(new ChipInfo { Text = "+" + overflow, Bg = BrushChipNeutralBg, Fg = BrushChipNeutralFg });
        }

        record.Chips = chips;
        record.TooltipLines = lines;
        record.HasDetection = lines.Count > 0;
        record.DetectionFlags = flags;
    }

    // 検出列の幅を、読み込み済みの全行のチップ列が収まる幅に更新する。
    // GridView の列幅は固定値のため、列幅より広いチップ列はセルの右端で切り
    // 落とされ、横スクロールしても見えない（スクロール範囲は列幅の合計で決まる）。
    // 年度の読み込みごとに最長行を実測して合わせる。絞り込みでは変えない
    //（条件の付け外しで列幅が揺れないようにするため）
    private void UpdateDepositDetectColumnWidth()
    {
        if (depositDetectColumn == null) return;

        // 加算値は XAML の寸法を写した定数（チップの左右の Padding と右の Margin・セルの左右の Margin・描画誤差の余裕）
        double maxRow = 0;
        foreach (var record in depositAllRecords)
        {
            double row = 0;
            if (record.IsNoAccount)
                row = MeasureChipText("口座なし") + CHIP_PADDING_H;
            else if (record.Chips != null)
                foreach (var chip in record.Chips)
                    row += MeasureChipText(chip.Text) + CHIP_PADDING_H + CHIP_MARGIN_R;
            if (row > maxRow) maxRow = row;
        }

        var width = maxRow + CELL_MARGIN_H + CHIP_COLUMN_SLACK;
        depositDetectColumn.Width = width > 330 ? width : 330;   // 下限は 330（チップの短い年度でも列幅を狭めない）
    }
}