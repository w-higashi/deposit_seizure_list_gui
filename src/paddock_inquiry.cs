// ==============================================================================
// paddock_inquiry.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 照会ファイルの解決（候補の収集・期待シートの照合・候補ダイアログ・紙照会の PDF）と
// Excel で開く処理・差押リスト連携。
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
    // 照会ファイル解決（ドリルダウン）
    // ==============================================================
    //
    // 台帳の行から仕分け後の照会ファイルを解決する。
    //   1. 行の登録日から年度を導出し、config の depositSearchFolders から検索フォルダを決める
    //   2. フォルダ配下を再帰走査し、ファイル名に宛名番号を含む .xlsm を候補にする
    //   3. 期待シート名（原本ファイル名_元シート名の安全化）を workbook.xml の照合で確認し、
    //      1件に特定できたら直行、できなければ候補ダイアログで職員が選ぶ（判定しない）
    // 「照会ファイルを開く」と「差押リストに追加」がこの解決パイプラインを
    // 共有し、終端アクション（Excel で開く / 作成ツールへ引き渡す）のみが異なる。

    // 候補ダイアログの案内の文（年度限定で 0 件のときの 2 行目・全年度で 0 件のとき）
    private const string SEARCH_ALL_HINT = "設定済みのすべての年度から探すには「全年度から探す」を押してください。";
    private const string NOT_FOUND_ALL_YEARS = "該当するファイルが見つかりませんでした。\n照会結果の保存年限（約3年）を経過している可能性があります。";
    private const int COLLECT_GRACE_MS = 300;   // 候補の窓を出す前にワーカーの完了を待つ最長時間（ms。間に合えば窓を出さずに分岐する）

    // 照会ファイルの候補1件
    private class InquiryFileCandidate
    {
        public string FullPath;          // 絶対パス
        public string RelativePath;      // 検索フォルダからの相対フォルダ（取組フォルダが読める形）
        public DateTime Modified;        // 更新日時
        public bool HasExpectedSheet;    // 期待シート名との照合結果
    }

    // 解決済みの照会ファイル（パイプラインの出力）
    private class ResolvedInquiryFile
    {
        public string FullPath;          // 絶対パス
        public string ExpectedSheet;     // シートジャンプに使う期待シート名（照合一致なしで開く場合は null）
    }

    // 台帳の行から照会ファイルを解決して開く（一覧・パネル・全ての「開く」動線の入口）。
    // 電子行は Excel 連携（シートジャンプつき）・紙照会行は既定アプリでの単純オープン
    private void OpenDepositInquiryFileForRow(string[] fields)
    {
        var resolved = ResolveDepositInquiryFileForRow(fields,
            "照会ファイルの候補", "開く", "開くファイルを選んでください");
        if (resolved != null)
        {
            RecordRecentDepositPerson(fields);   // 実際に開く行のみ記録する（キャンセル・0件時は記録しない）
            if (IsPaperRow(fields)) OpenPaperPdf(resolved.FullPath);
            else OpenInquiryFileInExcel(resolved.FullPath, resolved.ExpectedSheet);
        }
    }

    // 紙照会行か（照会結果ファイル名が .pdf — RACECARD 由来の行）
    private static bool IsPaperRow(string[] fields)
    {
        return fields != null &&
            fields[BusinessLogic.DEPOSIT_COL_SOURCE_FILE].EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
    }

    // 紙照会の PDF を既定アプリで開く（Excel 連携は経由しない。
    // 失敗時の文言は OpenInquiryFileInExcel の最終退行と同じ規約）
    private void OpenPaperPdf(string path)
    {
        try
        {
            Process.Start(path);
        }
        catch (Exception ex)
        {
            MessageBox.Show(window, "照会ファイルを開けませんでした。\n\n" + path + "\n" + ex.Message,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    // 預金の台帳の行から照会ファイルを解決する（口座なし・照会結果フォルダ未設定・紙照会を
    // ここで振り分け、電子照会の行は ResolveInquiryFile へ渡す）
    // 候補ダイアログの見出し・プライマリボタン・選択案内の文言は終端アクションに合わせて
    // 呼び出し側が渡す。解決できなかった・職員がキャンセルした場合は null を返す
    private ResolvedInquiryFile ResolveDepositInquiryFileForRow(
        string[] fields, string dialogTitle, string primaryButtonText, string selectPrompt)
    {
        if (fields == null) return null;
        if (fields[BusinessLogic.DEPOSIT_COL_HAS_ACCOUNT] == "0")
        {
            ShowTransientStatus("口座なし回答に照会シートはありません");
            return null;
        }
        if (config.DepositSearchFolders == null)
        {
            ShowTransientStatus("config の depositSearchFolders が未設定のため、照会ファイルを開く機能は無効です");
            return null;
        }

        // 紙照会行（RACECARD 由来・照会結果ファイル名が .pdf）: シート照合でなく
        // ファイル名完全一致で PDF を解決する（xlsm 行の経路には一切影響しない）
        if (IsPaperRow(fields))
            return ResolvePaperInquiryFile(fields, dialogTitle, primaryButtonText);

        return ResolveInquiryFile(config.DepositSearchFolders,
            fields[BusinessLogic.DEPOSIT_COL_ADDRESS_NUM], fields[BusinessLogic.DEPOSIT_COL_SOURCE_FILE],
            fields[BusinessLogic.DEPOSIT_COL_SHEET_NAME], fields[BusinessLogic.DEPOSIT_COL_REG_DATE],
            FindRememberedInquiryFile(fields), dialogTitle, primaryButtonText, selectPrompt);
    }

    // 照会ファイルの解決の本体（照会ファイルを開く・差押リストに追加・預金と生命保険で共有）:
    // 登録日の年度フォルダで宛名番号を含むファイルを集め、期待シート名の照合で1件に特定できれば
    // それを返す。候補の収集は候補ダイアログの中でワーカーが行う（探している間もキャンセルできる）。
    // 1件に特定できればダイアログは自分で閉じてそれを返し、0件なら案内（年度限定なら全年度への拡大）を
    // 出し、特定できなければ候補を並べて職員が選ぶ。remembered（明細の読み取りで特定済みのファイル。
    // 預金だけ・なければ null）がまだ在れば集めずにそれを返す
    private ResolvedInquiryFile ResolveInquiryFile(Dictionary<string, List<string>> folders,
        string addressNumber, string sourceFile, string sheetName, string regDateText,
        InquiryFileCandidate remembered, string dialogTitle, string primaryButtonText, string selectPrompt)
    {
        var expectedSheet = BusinessLogic.GetExpectedSheetName(sourceFile, sheetName);

        // 行の登録日から年度フォルダを決める（導出できない・設定にない年度は全年度を対象にする）
        var yearKey = BusinessLogic.GetSearchYearKey(regDateText, eraMapping, folders);

        var chosen = ShowCandidateDialog(folders, null, yearKey, addressNumber, expectedSheet, null,
            remembered, dialogTitle, primaryButtonText, selectPrompt);
        if (chosen == null) return null;
        return new ResolvedInquiryFile
        {
            FullPath = chosen.FullPath,
            ExpectedSheet = chosen.HasExpectedSheet ? expectedSheet : null
        };
    }

    // 紙照会行の解決: 登録日から年度フォルダを決め、配下をファイル名
    // 完全一致で探す。年度の導出規則・0件時の全年度への拡大・最終0件の
    // 文言は電子行（xlsm）の解決と同一で、探している間と0件の案内は同じ候補ダイアログに出す
    //（完全一致のため候補は並べず、見つかればダイアログが閉じてそのパスを返す）
    private ResolvedInquiryFile ResolvePaperInquiryFile(string[] fields, string dialogTitle, string primaryButtonText)
    {
        var fileName = fields[BusinessLogic.DEPOSIT_COL_SOURCE_FILE];

        var yearKey = BusinessLogic.GetSearchYearKey(fields[BusinessLogic.DEPOSIT_COL_REG_DATE], eraMapping, config.DepositSearchFolders);

        var found = ShowCandidateDialog(config.DepositSearchFolders, null, yearKey, null, null, fileName,
            null, dialogTitle, primaryButtonText, null);
        if (found == null) return null;
        return new ResolvedInquiryFile { FullPath = found.FullPath, ExpectedSheet = null };
    }

    // 年度→フォルダの並びを順に、ファイル名完全一致の PDF を探す（最初の一致で打ち切る。
    // 並びは電子行と同じ BuildLabeledFolders の結果——年度指定時はその年度のみ・全年度は新しい年度から）
    private static string FindPaperPdf(List<KeyValuePair<string, string>> labeledFolders, string fileName)
    {
        foreach (var entry in labeledFolders)
        {
            if (!Directory.Exists(entry.Value)) continue;
            var found = FindPdfByFileName(entry.Value, fileName);
            if (found != null) return found;
        }
        return null;
    }

    // フォルダ配下を再帰走査し、ファイル名完全一致の PDF を探す（最初の一致で打ち切る）。
    // 各フォルダでは列挙でなく Path.Combine + File.Exists の直接確認で済ませる
    //（完全一致のため可能な最短経路 — ネットワークドライブ上の走査を最小にする）。
    // アクセス不能なサブフォルダで走査全体が中断しない規約は EnumerateXlsmFiles と同一
    private static string FindPdfByFileName(string root, string fileName)
    {
        var pending = new Stack<string>();
        pending.Push(root);
        while (pending.Count > 0)
        {
            var dir = pending.Pop();
            try
            {
                var direct = System.IO.Path.Combine(dir, fileName);
                if (File.Exists(direct)) return direct;
                foreach (var sub in Directory.GetDirectories(dir)) pending.Push(sub);
            }
            catch
            {
                // アクセス不能なフォルダは読み飛ばす
            }
        }
        return null;
    }

    // 照会結果フォルダの設定（年度→フォルダ）を（年度ラベル, パス）の並びに展開する
    // yearKey 指定時はその年度のみ（ラベルなし）、null なら設定済み全年度（新しい年度から）
    private static List<KeyValuePair<string, string>> BuildLabeledFolders(Dictionary<string, List<string>> folders, string yearKey)
    {
        var result = new List<KeyValuePair<string, string>>();
        if (yearKey != null)
        {
            foreach (var folder in folders[yearKey])
                result.Add(new KeyValuePair<string, string>("", folder));
            return result;
        }
        foreach (var pair in folders.OrderByDescending(p => p.Key, StringComparer.Ordinal))
            foreach (var folder in pair.Value)
                result.Add(new KeyValuePair<string, string>(pair.Key, folder));
        return result;
    }

    // フォルダ群を再帰走査し、ファイル名に宛名番号を含む .xlsm を候補として集める
    // 各候補は workbook.xml を覗いて期待シート名との照合結果を付ける（Excel は起動しない）
    private static List<InquiryFileCandidate> CollectInquiryFileCandidates(
        List<KeyValuePair<string, string>> labeledFolders, string addressNumber, string expectedSheet)
    {
        var result = new List<InquiryFileCandidate>();
        var seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in labeledFolders)
        {
            var root = entry.Value;
            if (!Directory.Exists(root)) continue;

            foreach (var path in EnumerateXlsmFiles(root))
            {
                var name = System.IO.Path.GetFileNameWithoutExtension(path);
                if (name.StartsWith("~$")) continue;   // Excel の一時ロックファイル
                if (name.IndexOf(addressNumber, StringComparison.Ordinal) < 0) continue;
                if (!seenPaths.Add(path)) continue;    // 複数年度が同一フォルダを指す場合の重複排除

                var candidate = new InquiryFileCandidate { FullPath = path };
                try { candidate.Modified = File.GetLastWriteTime(path); } catch { }   // 更新日時が取れないファイルも候補に残す（既定の日時で末尾に並ぶ）

                // 表示用の相対フォルダ（全年度検索時は年度ラベルを前置する）
                var relative = path.Length > root.Length
                    ? path.Substring(root.Length).TrimStart('\\') : System.IO.Path.GetFileName(path);
                var relativeDir = System.IO.Path.GetDirectoryName(relative);
                candidate.RelativePath =
                    (entry.Key != "" ? entry.Key + ": " : "") +
                    (string.IsNullOrEmpty(relativeDir) ? "" : relativeDir + "\\");

                candidate.HasExpectedSheet = expectedSheet != null && FileHasExpectedSheet(path, expectedSheet);
                result.Add(candidate);
            }
        }
        return result.OrderByDescending(c => c.Modified).ToList();
    }

    // フォルダ配下の .xlsm を再帰的に列挙する
    // アクセス不能なサブフォルダで走査全体が中断しないよう、フォルダ単位で読み飛ばす
    private static List<string> EnumerateXlsmFiles(string root)
    {
        var result = new List<string>();
        var pending = new Stack<string>();
        pending.Push(root);
        while (pending.Count > 0)
        {
            var dir = pending.Pop();
            try
            {
                result.AddRange(Directory.GetFiles(dir, "*.xlsm"));
                foreach (var sub in Directory.GetDirectories(dir)) pending.Push(sub);
            }
            catch
            {
                // アクセス不能フォルダは読み飛ばす
            }
        }
        return result;
    }

    // ファイルの workbook.xml を System.IO.Packaging で読み、期待シート名の有無を照合する（Excel は起動しない）
    private static bool FileHasExpectedSheet(string path, string expectedSheet)
    {
        try
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = Package.Open(fs, FileMode.Open, FileAccess.Read))
            {
                var wbUri = new Uri("/xl/workbook.xml", UriKind.Relative);
                if (!package.PartExists(wbUri)) return false;

                using (var stream = package.GetPart(wbUri).GetStream(FileMode.Open, FileAccess.Read))
                {
                    var doc = new XmlDocument();
                    doc.Load(stream);
                    var nsmgr = new XmlNamespaceManager(doc.NameTable);
                    nsmgr.AddNamespace("s", BusinessLogic.NS_SPREADSHEET);

                    foreach (XmlNode sheet in doc.SelectNodes("//s:sheets/s:sheet", nsmgr))
                    {
                        var nameAttr = sheet.Attributes["name"];
                        if (nameAttr != null && BusinessLogic.SheetNameMatchesExpected(nameAttr.Value, expectedSheet))
                            return true;
                    }
                }
            }
        }
        catch
        {
            // 読めないファイルは照合不一致として扱う（候補には残り、職員の判断に委ねる）
        }
        return false;
    }

    // 照会ファイルを Excel で開き、期待シートをアクティブ化して先頭へ移動する
    // ・起動中の Excel があればそれに開かせる（二重起動させない）。ただしウィンドウを
    //   持たない残留プロセス（自動処理の異常終了等）にアタッチしてしまうことがあるため、
    //   アタッチ先での操作に失敗した場合は新規起動でもう一度だけ開き直す
    // ・既に開かれているブックは再オープンせずシート移動のみ行う
    // ・開いた（または移動した）ブックのウィンドウは常に最大化して前面に出す
    //   （直前に使っていた Excel のウィンドウサイズを引き継がない）
    // ・マクロの実行可否は通常起動と同じ扱い（AutomationSecurity=ByUI）
    // ・Excel の所有権はユーザーにあるため Quit は呼ばない（残留プロセスも強制終了しない）
    // ・それでも失敗した場合は関連付けでの単純オープンに退行する
    private void OpenInquiryFileInExcel(string path, string expectedSheet)
    {
        object excelObj = null;
        try
        {
            try { excelObj = Marshal.GetActiveObject("Excel.Application"); }
            catch (COMException) { }   // 起動中の Excel がなければ新しく起動する

            if (excelObj != null)
            {
                try
                {
                    OpenWorkbookInInstance(excelObj, true, path, expectedSheet);
                    return;
                }
                catch (Exception)
                {
                    // アタッチ先が壊れている → 参照を手放し、新規起動で開き直す
                    if (Marshal.IsComObject(excelObj)) Marshal.ReleaseComObject(excelObj);
                    excelObj = null;
                }
            }

            excelObj = CreateExcelInstance();
            OpenWorkbookInInstance(excelObj, false, path, expectedSheet);
        }
        catch (Exception)
        {
            try
            {
                Process.Start(path);
                ShowTransientStatus("Excel 連携に失敗したため、通常の方法で開きました（シート移動なし）");
            }
            catch (Exception ex)
            {
                MessageBox.Show(window, "照会ファイルを開けませんでした。\n\n" + path + "\n" + ex.Message,
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        finally
        {
            // Excel はユーザーが使い続けるため、参照だけを確実に手放す
            if (excelObj != null && Marshal.IsComObject(excelObj)) Marshal.ReleaseComObject(excelObj);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    // Excel の新規インスタンスを起動する
    private static object CreateExcelInstance()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null) throw new InvalidOperationException("Excel が見つかりません。");
        return Activator.CreateInstance(excelType);
    }

    // 指定の Excel インスタンスでファイルを開き、期待シートへ移動して最大化・前面化する
    // インスタンスの操作に失敗した場合は例外を投げ、呼び出し側の開き直しに委ねる
    private void OpenWorkbookInInstance(object excelObj, bool attached, string path, string expectedSheet)
    {
        dynamic excel = excelObj;
        excel.Visible = true;

        // 既存インスタンスに接続した場合のみ、既に開かれているブックを探す
        //（フルパスの大文字小文字は区別しない・新規起動の Excel にブックはない）
        dynamic workbook = null;
        if (attached)
        {
            int bookCount = (int)excel.Workbooks.Count;
            for (int i = 1; i <= bookCount; i++)
            {
                dynamic wb = excel.Workbooks[i];
                if (string.Equals((string)wb.FullName, path, StringComparison.OrdinalIgnoreCase))
                {
                    workbook = wb;
                    break;
                }
            }
        }

        if (workbook == null)
        {
            // 開いている間だけマクロ設定を ByUI にし、開き終えたら元へ戻す
            // （バージョンにより未サポートの場合は無視・judge.ps1 と同じ扱い）
            object origSecurity = null;
            try { origSecurity = excel.AutomationSecurity; excel.AutomationSecurity = 2; } catch { }   // 未サポートなら設定せずに開く
            try { workbook = excel.Workbooks.Open(path); }
            finally
            {
                if (origSecurity != null)
                    try { excel.AutomationSecurity = origSecurity; } catch { }   // 戻せなくても開いた状態を優先する
            }
        }

        // 期待シート（重複時の連番形を含む）をアクティブ化して先頭へ移動する
        // ジャンプできない場合（非表示シート等）もファイルは開いたままにする
        bool sheetJumped = false;
        if (expectedSheet != null)
        {
            try
            {
                int sheetCount = (int)workbook.Worksheets.Count;
                for (int i = 1; i <= sheetCount; i++)
                {
                    dynamic ws = workbook.Worksheets[i];
                    if (BusinessLogic.SheetNameMatchesExpected((string)ws.Name, expectedSheet))
                    {
                        ws.Activate();
                        excel.Goto(ws.Range("A1"), true);
                        sheetJumped = true;
                        break;
                    }
                }
            }
            catch { }   // シートへ移動できなくてもファイルは開いたまま（フッターで知らせる）
        }

        // ウィンドウ状態によらず最大化して Excel を前面へ出す（新規起動・アタッチの
        // どちらでも、直前に使っていた Excel のウィンドウサイズを引き継がない）
        workbook.Activate();
        try
        {
            excel.WindowState = -4137;   // xlMaximized（最小化されていた場合もそのまま最大化される）
            SetForegroundWindow(new IntPtr(Convert.ToInt64(excel.Hwnd)));
        }
        catch { }   // 最大化・前面化できなくてもファイルは開いたまま

        if (expectedSheet != null && !sheetJumped)
            ShowTransientStatus("期待するシート（" + expectedSheet + "）が見つからないため、ファイルをそのまま開きました");
    }

    // 候補ダイアログを表示し、選ばれた候補を返す（キャンセル時は null）
    // 各候補にファイル名・相対フォルダ・更新日時・シート照合の結果を提示し、判別は職員が行う。
    // 見出しとプライマリボタンの文言は終端アクション（開く / 差押リストに追加）に合わせて受け取る。
    // candidates が null なら yearKey の年度（null は全年度）の候補をワーカーで集める。集め終わるのを
    // 短い猶予だけ待ち、間に合えば窓を出さずに分岐する（期待シートの一致がちょうど 1 件・remembered が
    // まだ在る、なら候補を見せずにそれを返す）。間に合わなければ「探している間」の状態で窓を開き、
    // 完了を窓の中で受ける——共有フォルダが遅くても PADDOCK は止まらず、「キャンセル」でやめられる。
    // 0 件なら案内を出す。paperFileName を渡すと紙照会の PDF をファイル名の完全一致で探し、
    // 見つかればそのパスの候補を返す（候補の一覧にはならない）
    private InquiryFileCandidate ShowCandidateDialog(Dictionary<string, List<string>> folders,
        List<InquiryFileCandidate> candidates, string yearKey, string addressNumber, string expectedSheet,
        string paperFileName, InquiryFileCandidate remembered, string dialogTitle, string primaryButtonText, string selectPrompt)
    {
        var dialog = new Window
        {
            Title = dialogTitle,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Width = 580,
            SizeToContent = SizeToContent.Height,
            MaxHeight = 620,
            ResizeMode = ResizeMode.NoResize,
            Background = Brushes.White,
            FontFamily = FontUi,
            ShowInTaskbar = false
        };

        var root = new DockPanel();
        dialog.Content = root;

        // --- ヘッダー（検索条件の提示）---
        var header = new Border
        {
            BorderBrush = BrushRailLine,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(20, 14, 20, 12)
        };
        DockPanel.SetDock(header, Dock.Top);
        var headerStack = new StackPanel();
        var title = new TextBlock
        {
            FontSize = 13.5,
            Foreground = BrushTextMain
        };
        headerStack.Children.Add(title);
        var context = new TextBlock { FontSize = 11, Foreground = BrushTextSub, Margin = new Thickness(0, 5, 0, 0) };
        if (paperFileName != null)
        {
            context.Inlines.Add(new Run("ファイル名 "));
            context.Inlines.Add(new Run(paperFileName) { FontFamily = FontMono });
        }
        else
        {
            context.Inlines.Add(new Run("宛名番号 "));
            context.Inlines.Add(new Run(addressNumber) { FontFamily = FontMono });
            if (expectedSheet != null)
            {
                context.Inlines.Add(new Run("　·　期待シート: "));
                context.Inlines.Add(new Run(expectedSheet) { FontFamily = FontMono });
            }
        }
        headerStack.Children.Add(context);
        header.Child = headerStack;
        root.Children.Add(header);

        // --- フッター（全年度から探す / キャンセル / プライマリボタン=終端アクション）---
        var footer = new Border
        {
            Background = BrushSurfaceAlt,
            BorderBrush = BrushLine,
            BorderThickness = new Thickness(0, 1, 0, 0),
            Padding = new Thickness(14, 10, 14, 10)
        };
        DockPanel.SetDock(footer, Dock.Bottom);
        var footerPanel = new DockPanel();

        var searchAllButton = new Button
        {
            Content = "全年度から探す",
            Style = (Style)window.FindResource("PanelGhostBtn"),
            Width = 130
        };
        DockPanel.SetDock(searchAllButton, Dock.Left);
        footerPanel.Children.Add(searchAllButton);

        // プライマリボタン（文言は終端アクションに応じて可変。幅は文言の実寸から）
        var openButton = new Button
        {
            Content = primaryButtonText,
            Style = (Style)window.FindResource("PanelPrimaryBtn"),
            Width = Math.Max(110, MeasureUiText(primaryButtonText, 12) + 26)
        };
        DockPanel.SetDock(openButton, Dock.Right);
        var cancelButton = new Button
        {
            Content = "キャンセル",
            Style = (Style)window.FindResource("PanelGhostBtn"),
            Width = 100,
            Margin = new Thickness(0, 0, 8, 0)
        };
        DockPanel.SetDock(cancelButton, Dock.Right);
        footerPanel.Children.Add(openButton);
        footerPanel.Children.Add(cancelButton);
        footerPanel.Children.Add(new Border());   // 中央の余白
        footer.Child = footerPanel;
        root.Children.Add(footer);

        // --- 一覧の領域（候補の行・探している間のリング・見つからないときの案内）---
        var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
        var listPanel = new StackPanel { Margin = new Thickness(14, 10, 14, 10) };
        scroll.Content = listPanel;
        root.Children.Add(scroll);

        InquiryFileCandidate selected = null;
        Border selectedRow = null;
        bool closed = false;
        dialog.Closed += delegate { closed = true; };   // 閉じたあとに終わった収集は捨てる

        Action<Border, InquiryFileCandidate> select = delegate(Border row, InquiryFileCandidate candidate)
        {
            if (selectedRow != null)
            {
                selectedRow.ClearValue(Border.BackgroundProperty);
                selectedRow.ClearValue(Border.BorderBrushProperty);
            }
            selected = candidate;
            selectedRow = row;
            row.Background = BrushSelectedCardBg;
            row.BorderBrush = BrushAccent;
            openButton.IsEnabled = true;
        };

        // 候補の行（キーで選ぶときの並び。行はフォーカスを受け、移った行を選ぶ）
        var rows = new List<KeyValuePair<Border, InquiryFileCandidate>>();

        // 一覧の領域を空にして選択を解く（状態を切り替える前の共通の手順）
        Action reset = delegate
        {
            listPanel.Children.Clear();
            rows.Clear();
            selected = null;
            selectedRow = null;
            openButton.IsEnabled = false;
        };

        // 候補の一覧（更新日時降順・選択は職員が行う）
        Action<List<InquiryFileCandidate>, bool> populate = delegate(List<InquiryFileCandidate> items, bool allYears)
        {
            reset();
            title.Text = "候補 " + items.Count + " 件 — " + selectPrompt;
            searchAllButton.Visibility = allYears ? Visibility.Collapsed : Visibility.Visible;
            cancelButton.Content = "キャンセル";
            foreach (var candidate in items)
            {
                var row = BuildCandidateRow(candidate, expectedSheet != null);
                var captured = candidate;
                var capturedRow = row;
                row.Focusable = true;
                row.FocusVisualStyle = keyFocusStyle;
                row.GotKeyboardFocus += delegate { select(capturedRow, captured); };
                rows.Add(new KeyValuePair<Border, InquiryFileCandidate>(row, candidate));
                row.MouseLeftButtonUp += delegate { select(capturedRow, captured); };
                row.MouseLeftButtonDown += delegate(object s, MouseButtonEventArgs e)
                {
                    if (e.ClickCount >= 2)
                    {
                        select(capturedRow, captured);
                        dialog.DialogResult = true;
                    }
                };
                listPanel.Children.Add(row);
            }
        };

        // 見つからなかった（年度限定なら「全年度から探す」を出す。全年度なら閉じるだけ）
        Action<string, bool> showNone = delegate(string text, bool canSearchAll)
        {
            reset();
            title.Text = "照会ファイルが見つかりませんでした";
            searchAllButton.Visibility = canSearchAll ? Visibility.Visible : Visibility.Collapsed;
            cancelButton.Content = canSearchAll ? "キャンセル" : "閉じる";
            listPanel.Children.Add(new TextBlock
            {
                Text = text, FontSize = 11.5, Foreground = BrushTextMuted, TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Center, Margin = new Thickness(10, 30, 10, 26)
            });
        };

        // 完了の分岐（猶予の中で受けても窓の中で受けても同じ）。直行（特定済みのファイル・紙の発見・期待シートの
        // 一致 1 件）なら selected を立てて true を返し、ほかは窓の状態（案内・一覧）を整えて false を返す
        Func<object, Exception, string, bool> settle = delegate(object result, Exception error, string scopeYearKey)
        {
            if (error != null)
            {
                WriteErrorLog("照会ファイルの検索", error);
                showNone("照会結果フォルダを読めませんでした。\n" + error.Message, false);
                return false;
            }
            var kept = result as InquiryFileCandidate;   // 特定済みのファイルがまだ在った
            if (kept != null)
            {
                selected = kept;
                return true;
            }
            bool limited = scopeYearKey != null;
            if (paperFileName != null)
            {
                var found = (string)result;
                if (found != null)
                {
                    selected = new InquiryFileCandidate { FullPath = found };
                    return true;
                }
                showNone(limited ? scopeYearKey + " のフォルダに " + paperFileName + " が見つかりませんでした。\n" + SEARCH_ALL_HINT : NOT_FOUND_ALL_YEARS, limited);
                return false;
            }
            var items = (List<InquiryFileCandidate>)result;
            if (items.Count == 0)
            {
                showNone(limited ? scopeYearKey + " のフォルダに宛名番号 " + addressNumber + " のファイルが見つかりませんでした。\n" + SEARCH_ALL_HINT : NOT_FOUND_ALL_YEARS, limited);
                return false;
            }
            // シート照合で 1 件に特定できた場合は候補を見せずに直行する
            var matched = items.Where(c => c.HasExpectedSheet).ToList();
            if (matched.Count == 1)
            {
                selected = matched[0];
                return true;
            }
            populate(items, !limited);
            return false;
        };

        // 探している間（リングと文言。押せるのは「キャンセル」だけ）。収集はワーカーで行い、graceMs の間だけ
        // その完了を待つ——間に合えば窓を出さずに分岐し（直行なら true を返す）、間に合わなければ窓を出して
        // 完了を窓の中で受ける。keep は特定済みのファイル（まだ在れば集めずにそれを返す。null なら集める）
        Func<string, int, InquiryFileCandidate, bool> startCollect = delegate(string scopeYearKey, int graceMs, InquiryFileCandidate keep)
        {
            reset();
            title.Text = "照会ファイルを探しています…";
            searchAllButton.Visibility = Visibility.Collapsed;
            cancelButton.Content = "キャンセル";
            RotateTransform busyRotation;
            var line = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 34, 0, 30) };
            line.Children.Add(MakeSmallSpinner(out busyRotation, new Thickness(0, 1, 8, 0)));
            line.Children.Add(new TextBlock
            {
                Text = (scopeYearKey != null ? scopeYearKey + " のフォルダ" : "すべての年度のフォルダ") + "を探しています",
                FontSize = 11, Foreground = BrushTextMuted
            });
            listPanel.Children.Add(line);

            var labeled = BuildLabeledFolders(folders, scopeYearKey);
            object collected = null;
            Exception failure = null;
            bool settled = false;   // 猶予の中で受けた完了は窓の中で二度受けない
            var finished = new System.Threading.ManualResetEvent(false);
            var worker = new BackgroundWorker();
            worker.DoWork += delegate
            {
                // UI には触らない（渡された年度→フォルダの並びと宛名番号・期待シート名・ファイル名・特定済みのファイルだけを読む）
                try
                {
                    if (paperFileName != null) collected = FindPaperPdf(labeled, paperFileName);
                    else if (keep != null && File.Exists(keep.FullPath)) collected = keep;
                    else collected = CollectInquiryFileCandidates(labeled, addressNumber, expectedSheet);
                }
                catch (Exception ex) { failure = ex; }
                finally { finished.Set(); }
            };
            worker.RunWorkerCompleted += delegate
            {
                finished.Close();
                StopSpinner(busyRotation);
                if (closed || settled) return;   // キャンセルで閉じたあと・猶予の中で受けたあとの完了は捨てる
                settled = true;
                if (settle(collected, failure, scopeYearKey)) dialog.DialogResult = true;
            };
            worker.RunWorkerAsync();
            if (graceMs > 0 && finished.WaitOne(graceMs))
            {
                StopSpinner(busyRotation);
                settled = true;
                return settle(collected, failure, scopeYearKey);
            }
            return false;
        };

        bool direct = false;
        if (candidates != null) populate(candidates, yearKey == null);
        else direct = startCollect(yearKey, COLLECT_GRACE_MS, remembered);

        searchAllButton.Click += delegate { startCollect(null, 0, null); };
        openButton.Click += delegate { if (selected != null) dialog.DialogResult = true; };
        cancelButton.Click += delegate { dialog.DialogResult = false; };
        dialog.PreviewKeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) { dialog.DialogResult = false; e.Handled = true; }
            if (e.Key == Key.Enter && selected != null) { dialog.DialogResult = true; e.Handled = true; }
            // ↑↓: 選択中の候補の前後へ（未選択なら先頭へ）移って選ぶ
            if ((e.Key == Key.Up || e.Key == Key.Down) && Keyboard.Modifiers == ModifierKeys.None && rows.Count > 0)
            {
                int index = rows.FindIndex(delegate(KeyValuePair<Border, InquiryFileCandidate> pair) { return ReferenceEquals(pair.Key, selectedRow); });
                int next = index < 0 ? 0 : index + (e.Key == Key.Down ? 1 : -1);
                if (next >= 0 && next < rows.Count) rows[next].Key.Focus();
                e.Handled = true;
            }
        };

        if (direct) return selected;   // 猶予の中で直行が決まった（窓は出さない）
        dialog.Owner = window;   // 窓を出すときだけ所有させる（出さずに戻る窓を本体の所有の一覧に残さない）
        return dialog.ShowDialog() == true ? selected : null;
    }

    // 候補1件の行（ファイル名+相対フォルダ / シート照合チップ / 更新日時）を組み立てる
    private Border BuildCandidateRow(InquiryFileCandidate candidate, bool showSheetChip)
    {
        var row = new Border
        {
            Style = (Style)window.FindResource("AccountCard"),
            Margin = new Thickness(0, 0, 0, 7),
            Cursor = Cursors.Hand
        };

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var left = new StackPanel();
        left.Children.Add(new TextBlock
        {
            Text = System.IO.Path.GetFileName(candidate.FullPath),
            FontSize = 12.5,
            Foreground = BrushTextMain,
            TextTrimming = TextTrimming.CharacterEllipsis
        });
        left.Children.Add(new TextBlock
        {
            Text = candidate.RelativePath,
            FontSize = 10.5,
            Foreground = BrushTextMuted,
            Margin = new Thickness(0, 3, 0, 0),
            TextTrimming = TextTrimming.CharacterEllipsis
        });
        grid.Children.Add(left);

        if (showSheetChip)
        {
            var chip = new Border
            {
                Background = candidate.HasExpectedSheet ? BrushChipTealBg : BrushChipNeutralBg,
                CornerRadius = new CornerRadius(9),
                Padding = new Thickness(9, 1, 9, 2),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 0, 0, 0)
            };
            chip.Child = new TextBlock
            {
                Text = candidate.HasExpectedSheet ? "期待シートあり" : "期待シートなし",
                FontSize = 10,
                Foreground = candidate.HasExpectedSheet ? BrushChipTealFg : BrushChipNeutralFg
            };
            Grid.SetColumn(chip, 1);
            grid.Children.Add(chip);
        }

        var date = new TextBlock
        {
            Text = candidate.Modified.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture),
            FontFamily = FontMono,
            FontSize = 11,
            Foreground = BrushTextSub,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(10, 0, 0, 0)
        };
        Grid.SetColumn(date, 2);
        grid.Children.Add(date);

        row.Child = grid;
        return row;
    }

    // ==============================================================
    // 差押リスト連携
    // ==============================================================
    //
    // 台帳で判断材料を確認した職員が、その場から差押予定一覧への登録に進める動線。
    // ファイル解決は「照会ファイルを開く」とパイプラインを共有し、解決済みファイルを
    // 預金差押予定一覧 作成ツール（entry_deposit.exe）に引き渡す（生命保険の入口は
    // InsuranceTab が持ち、生命保険差押予定一覧 作成ツールへ渡す。起動の処理は共用）。
    // 登録の成否・内容の記録（差押登録ログ）は作成ツール側が書き、本ツールは読むのみ。
    // ログは PADDOCK 起点の履歴として ZEKKEN のログとは独立に保管する
    //（横断ビューは STEWARD が担う）。

    // 差押リストの対象金融機関: config の depositSeizureTargetBanks に一致しない
    // 金融機関の行では「差押リストに追加」を無効化する（非表示にしない・無効時の
    // ツールチップも出さない）。照合結果はキー単位でメモ化する——判定は年度横断の
    // カードにも及ぶため年度の distinct には固定せず、パターンが変わらない限り
    // キャッシュは年度を切り替えても有効（絞り込み・一覧描画の経路では使わない）
    private void RebuildDepositSeizureTargets()
    {
        depositSeizureTargetPatterns = BusinessLogic.ParseSeizureTargetBanks(config.DepositSeizureTargetBanks);
        depositSeizureTargetMatchCache = new Dictionary<string, bool>();
    }

    // 金融機関・保険会社の正規化キーが差押リストの対象か（patterns が null＝キー未設定・空欄は
    // 制御なし=常に対象）。照合の結果は cache にキー単位で控える
    private static bool IsSeizureTargetKey(string institutionKey,
        List<SeizureTargetPattern> patterns, Dictionary<string, bool> cache)
    {
        if (patterns == null) return true;
        if (string.IsNullOrEmpty(institutionKey)) return false;
        bool match;
        if (!cache.TryGetValue(institutionKey, out match))
        {
            match = BusinessLogic.MatchesSeizureTarget(institutionKey, patterns);
            cache[institutionKey] = match;
        }
        return match;
    }

    // 一覧の行を差押リストに追加できるか: 連携先があり、ファイル解決（depositSearchFolders）が
    // 設定済みで、対象金融機関に一致する行（右クリックメニューと Ctrl+D で共用）
    private bool CanAddToDepositSeizureList(DepositLedgerRecord record)
    {
        return depositSeizureScriptPath != null && record != null && config.DepositSearchFolders != null &&
            IsSeizureTargetKey(record.SearchInstitution, depositSeizureTargetPatterns, depositSeizureTargetMatchCache);
    }

    // 差押登録ログファイルのパスを返す（history/{Windowsユーザー名}/seizure_log.json）。
    // ログパスは --log 引数で作成ツールへ渡すため、書き込み側もそのまま追従する
    //（親フォルダは書き込み側が作成する）
    private string GetSeizureLogPath()
    {
        return System.IO.Path.Combine(userHistDir, "seizure_log.json");
    }

    // 台帳の行から照会ファイルを解決し、預金差押予定一覧 作成ツールに引き渡す
    // （右クリックメニュー・パネルフッターの「差押リストに追加」の入口）
    private void AddDepositRowToSeizureList(string[] fields)
    {
        if (depositSeizureScriptPath == null) return;   // 連携系UIは非表示のため通常は到達しない

        var resolved = ResolveDepositInquiryFileForRow(fields,
            "差押リストに追加 — 照会ファイルの候補", "このファイルで追加", "追加するファイルを選んでください");
        if (resolved != null && LaunchSeizureListTool(depositSeizureScriptPath, "預金差押予定一覧 作成ツール", resolved.FullPath))
            RecordRecentDepositPerson(fields);   // 実際に引き渡した行のみ記録する（起動中断・失敗時は記録しない）
    }

    // 差押予定一覧 作成ツール（scriptPath・案内での呼び名 toolName）を起動して照会ファイルを引き渡す
    // （どちらのツールへ渡すかは呼び出し元の台帳の種別で決まり、ファイルの中身による判定は行わない）
    // 起動に至った場合は true を返す（単一起動ガードで中断・起動失敗時は false）
    private bool LaunchSeizureListTool(string scriptPath, string toolName, string path)
    {
        // 対象ツールが既に起動中の場合は競合防止のため中断（単一起動ガード）
        var exeName = System.IO.Path.GetFileNameWithoutExtension(scriptPath);
        if (Process.GetProcessesByName(exeName).Length > 0)
        {
            MessageBox.Show(window,
                toolName + "が起動中です。\n処理が完了してから再度実行してください。",
                "確認", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }

        try
        {
            // ログパスを引数で渡し、子プロセスの終了時にログを読み直して表示へ反映する
            var logArg = "--log=\"" + GetSeizureLogPath() + "\"";
            var psi = new ProcessStartInfo
            {
                FileName = scriptPath,
                Arguments = logArg + " \"" + path + "\"",
                UseShellExecute = false
            };

            var proc = Process.Start(psi);
            proc.EnableRaisingEvents = true;
            proc.Exited += delegate { Dispatcher.BeginInvoke(new Action(ReloadSeizureLog)); };

            ShowTransientStatus("差押リストに追加中: " + System.IO.Path.GetFileName(path));
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(window,
                toolName + "の呼び出しに失敗しました:\n" + ex.Message,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }
}