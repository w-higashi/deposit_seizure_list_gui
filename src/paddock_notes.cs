// ==============================================================================
// paddock_notes.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 人物メモ（読み書き・詳細パネルのメモの欄・一覧の印）。
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
    // 人物メモ
    // ==============================================================
    //
    // 宛名番号（人物）単位の自由テキストメモ。各ユーザーが自分のファイル
    //（history/{ユーザー名}/person_notes.json）に書き込み、詳細パネルでは
    // 全ユーザーの記入分を集約して作成日時の降順に表示する。

    // 詳細パネル1つ分のメモの欄（預金・生命保険のパネルがそれぞれ持つ）
    private class NotesHost
    {
        public StackPanel Panel;      // メモの欄（見出し・入力エリア・一覧を組み立てる入れ物）
        public Button AddButton;      // 「＋ 追加」（キーで書き終えた後のフォーカスの戻り先）
        public Action BeginAdd;       // メモを書き始める処理（Ctrl+N。追加できないときは null）
        public bool Expanded;         // 「さらに表示」の展開状態（人物の切替でリセット）
        public StackPanel ListHost;   // 一覧の入れ物（秒切れで「元に戻す」の行を消すとき、一覧だけを描き直す）
    }

    // 自分のメモファイルのパス（history/{ユーザー名}/person_notes.json）
    private string GetPersonNotesPath()
    {
        return System.IO.Path.Combine(userHistDir, "person_notes.json");
    }

    // 起動時に全ユーザーの person_notes をまとめて読み込む（一覧の印にも使うため台帳より先）。
    // 以降はキャッシュから引く（depositOtherYearCache と同じ思想）。
    // 記入者（職員番号 = Windows ユーザー名）はフォルダ名から取得する
    //（ユーザー名の大小は区別しない）
    private void EnsureNotesLoaded()
    {
        if (myPersonNotes != null) return;
        myPersonNotes = new Dictionary<string, List<PersonNoteEntry>>();
        otherPersonNotes = new Dictionary<string, List<PersonNoteEntry>>();

        var histDir = System.IO.Path.Combine(exeDir, "history");
        if (!Directory.Exists(histDir)) return;

        string[] dirs;
        try { dirs = Directory.GetDirectories(histDir); }
        catch { return; }   // 列挙できない場合はメモなしで続行する（落ちない）

        foreach (var dir in dirs)
        {
            var file = System.IO.Path.Combine(dir, "person_notes.json");
            if (!File.Exists(file)) continue;
            var author = System.IO.Path.GetFileName(dir);
            bool isMine = string.Equals(author, Environment.UserName, StringComparison.OrdinalIgnoreCase);
            try
            {
                LoadPersonNotesFile(file, author, isMine ? myPersonNotes : otherPersonNotes);
            }
            catch (Exception ex)
            {
                // 読めないファイルはそのファイルのみ読み飛ばす（他ユーザーの分は表示を続ける）。
                // 自分のファイルが読めない場合は保存もしない（壊れたファイルへの上書きで
                // 既存メモを失わせない — recent_people の persistRecent と同一規約）
                if (isMine) { myPersonNotes.Clear(); persistNotes = false; }
                WriteErrorLog("メモの読み込み: " + author, ex);
            }
        }
    }

    // メモファイル1本を読み込み、宛名番号ごとの辞書へ合流させる。
    // 形式: { "notes": { "<宛名番号>": [ { "createdAt": "...", "text": "..." }, ... ] } }
    // メモ本文は任意文字列（波括弧・改行を含みうる）のため、配列・オブジェクトの
    // 区切りは文字列リテラルを解する FindMatchingBracket で取る（recent_people と同じ理由）。
    // 空のファイルは「ない」のと同じ（読むものがなく、次の保存で普通に書く）。構造が読めない
    //（notes がない・括弧が閉じない＝書いている途中で切れたなど）ファイルは例外にして
    // 呼び出し元の読み飛ばし・保存の停止に委ねる（壊れたファイルに重ねて書かない）
    private void LoadPersonNotesFile(string path, string author,
        Dictionary<string, List<PersonNoteEntry>> target)
    {
        var json = File.ReadAllText(path, Encoding.UTF8);
        if (json.Trim().Length == 0) return;
        var notesIdx = json.IndexOf("\"notes\"");
        if (notesIdx < 0) throw new InvalidDataException("notes の構造が読めません");
        var objStart = json.IndexOf('{', notesIdx);
        if (objStart < 0) throw new InvalidDataException("notes の構造が読めません");
        var objEnd = JsonHelper.FindMatchingBracket(json, objStart, '{', '}');
        if (objEnd < 0) throw new InvalidDataException("notes の構造が読めません");

        var keyRegex = new Regex("\"((?:[^\"\\\\]|\\\\.)*)\"\\s*:\\s*\\[");
        int pos = objStart + 1;
        int skipped = 0;
        while (pos < objEnd)
        {
            var m = keyRegex.Match(json, pos);
            if (!m.Success || m.Index >= objEnd) break;
            var arrStart = m.Index + m.Length - 1;   // 配列の '[' の位置
            var arrEnd = JsonHelper.FindMatchingBracket(json, arrStart, '[', ']');
            if (arrEnd < 0) break;

            var addressNumber = JsonHelper.UnescapeJsonString(m.Groups[1].Value);
            if (addressNumber != "")
            {
                List<PersonNoteEntry> list;
                if (!target.TryGetValue(addressNumber, out list))
                {
                    list = new List<PersonNoteEntry>();
                    target[addressNumber] = list;
                }
                int entryPos = arrStart + 1;
                while (entryPos < arrEnd)
                {
                    var entryStart = json.IndexOf('{', entryPos);
                    if (entryStart < 0 || entryStart >= arrEnd) break;
                    var entryEnd = JsonHelper.FindMatchingBracket(json, entryStart, '{', '}');
                    if (entryEnd < 0) break;
                    var note = ParsePersonNote(
                        json.Substring(entryStart, entryEnd - entryStart + 1), author);
                    if (note != null) list.Add(note);   // 形式不正はそのメモのみ読み飛ばす（件数はログに）
                    else skipped++;
                    entryPos = entryEnd + 1;
                }
            }
            pos = arrEnd + 1;
        }
        if (skipped > 0) WriteErrorLog("メモの読み込み: " + author, "読み飛ばし " + skipped + " 件");
    }

    // メモ1件を解析する（本文キーがない場合は null）。本文は改行を含みうるため、
    // 値の取り出しは文字列リテラルの正規表現+UnescapeJsonString で行う
    private PersonNoteEntry ParsePersonNote(string objJson, string author)
    {
        var m = Regex.Match(objJson, "\"text\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
        if (!m.Success) return null;

        var note = new PersonNoteEntry
        {
            Author = author,
            Text = JsonHelper.UnescapeJsonString(m.Groups[1].Value),
            CreatedAtRaw = JsonHelper.GetString(objJson, "createdAt") ?? ""
        };
        DateTime createdAt;
        note.CreatedAt = DateTime.TryParse(note.CreatedAtRaw, CultureInfo.InvariantCulture, DateTimeStyles.None, out createdAt)
            ? createdAt : DateTime.MinValue;
        return note;
    }

    // 指定人物のメモを全ユーザー分まとめて作成日時の降順で返す
    //（日時がパース不能のメモは末尾に集まる — 生値の表示で提示する）
    private List<PersonNoteEntry> CollectPersonNotes(string addressNumber)
    {
        var notes = new List<PersonNoteEntry>();
        List<PersonNoteEntry> list;
        if (myPersonNotes.TryGetValue(addressNumber, out list)) notes.AddRange(list);
        if (otherPersonNotes.TryGetValue(addressNumber, out list)) notes.AddRange(list);
        return notes.OrderByDescending(n => n.CreatedAt).ToList();
    }

    // メモセクション（見出し+件数チップ+追加ボタン+入力エリア+一覧）を組み立てる。
    // 一覧の描画は RenderNotesList に分け、「さらに表示」の開閉では一覧のみを
    // 再描画する（入力エリアの入力途中の本文を作り直さない）
    private void BuildNotesSection(NotesHost host, string addressNumber)
    {
        EnsureNotesLoaded();
        host.Panel.Children.Clear();
        host.BeginAdd = null;
        if (ReferenceEquals(host, noteUndoHost) && noteUndoAddress != addressNumber) DiscardNoteUndo(host);   // 同じ欄で別の人物を開いたら戻せない

        var notes = CollectPersonNotes(addressNumber);

        // --- 見出し行: 「メモ」+件数チップ（表示状態に関わらず全件数）+「＋ 追加」---
        var header = new Grid();
        header.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        header.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var title = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
        title.Children.Add(new TextBlock
        {
            Text = "メモ",
            FontSize = 12,
            FontWeight = FontWeights.Bold,
            Foreground = BrushTextSub,
            VerticalAlignment = VerticalAlignment.Center
        });
        if (notes.Count > 0)
        {
            var chip = new Border
            {
                Background = BrushChipVioletBg,
                CornerRadius = new CornerRadius(9),
                Padding = new Thickness(8, 1, 8, 2),
                Margin = new Thickness(7, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            chip.Child = new TextBlock
            {
                Text = notes.Count.ToString(),
                FontSize = 10.5,
                Foreground = BrushChipVioletFg
            };
            title.Children.Add(chip);
        }
        header.Children.Add(title);

        var addButton = MakeMiniButton("＋ 追加");
        Grid.SetColumn(addButton, 1);
        header.Children.Add(addButton);
        host.Panel.Children.Add(header);
        host.AddButton = addButton;

        // --- 追加の入力エリア（「＋ 追加」で開閉・アニメは AnimateExpand/AnimateCollapse）---
        TextBox inputBox;
        Button submitButton, cancelButton;
        var inputCard = BuildNoteInputCard(host, out inputBox, out submitButton, out cancelButton);
        host.Panel.Children.Add(inputCard);

        if (!persistNotes)
        {
            // メモファイルに書き込めない間は追加を受け付けない
            //（保存されない追加・壊れたファイルへの上書きを防ぐ）
            addButton.IsEnabled = false;
            addButton.ToolTip = "メモファイルに書き込めないため追加できません";
            ToolTipService.SetShowOnDisabled(addButton, true);
        }
        else
        {
            addButton.Click += delegate
            {
                if (inputCard.Visibility == Visibility.Visible)
                {
                    AnimateCollapse(inputCard);
                }
                else
                {
                    AnimateExpand(inputCard);
                    inputBox.Focus();
                }
            };
            host.BeginAdd = delegate
            {
                AnimateExpand(inputCard);
                inputBox.Focus();
            };
            cancelButton.Click += delegate
            {
                inputBox.Text = "";
                AnimateCollapse(inputCard);
            };
            submitButton.Click += delegate
            {
                var text = inputBox.Text.Trim();
                if (text == "") return;
                inputBox.Text = "";   // 先に空へ（ダブルクリックの2発目を空判定で弾く）
                AddNote(host, addressNumber, text);   // セクションを再描画するため入力エリアも閉じる
            };
        }

        // --- 一覧 / 空状態（「さらに表示」の開閉はこの部分のみ再描画する）---
        var listHost = new StackPanel();
        host.ListHost = listHost;
        host.Panel.Children.Add(listHost);
        RenderNotesList(host, addressNumber, listHost, notes);
    }

    // メモ一覧（最新 NOTES_VISIBLE_MAX 件+「さらに N 件を表示」/「折りたたむ」）を
    // listHost へ描画する
    private void RenderNotesList(NotesHost host, string addressNumber, StackPanel listHost, List<PersonNoteEntry> notes)
    {
        listHost.Children.Clear();

        // 削除した直後（NOTE_UNDO_SECONDS の間）は、消した行の位置に「元に戻す」の行を出す。
        // 最後の 1 件を消したときも空状態は出さず、この行だけを出す
        bool undoHere = ReferenceEquals(host, noteUndoHost) && noteUndoAddress == addressNumber;
        if (notes.Count == 0 && !undoHere)
        {
            listHost.Children.Add(new TextBlock
            {
                Text = "メモはまだありません",
                FontSize = 11,
                Foreground = BrushTextMuted,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 12, 0, 4)
            });
            return;
        }

        var list = new StackPanel { Margin = new Thickness(0, 5, 0, 0) };
        int visibleCount = host.Expanded ? notes.Count : Math.Min(notes.Count, NOTES_VISIBLE_MAX);
        for (int i = 0; i < visibleCount; i++)
            list.Children.Add(BuildNoteRow(host, addressNumber, notes[i], i == visibleCount - 1));
        if (undoHere) list.Children.Insert(Math.Min(noteUndoIndex, list.Children.Count), BuildNoteUndoRow());   // 位置は表示している末尾までに丸める
        listHost.Children.Add(list);

        // --- 「さらに N 件を表示」/「折りたたむ」（超過があるときのみ）---
        if (notes.Count > NOTES_VISIBLE_MAX)
        {
            var more = new Border
            {
                Background = Brushes.Transparent,
                CornerRadius = new CornerRadius(7),
                Padding = new Thickness(0, 5, 0, 5),
                Cursor = Cursors.Hand
            };
            more.Child = new TextBlock
            {
                Text = host.Expanded ? "折りたたむ"
                    : "さらに " + (notes.Count - NOTES_VISIBLE_MAX) + " 件を表示",
                FontSize = 11,
                Foreground = BrushAccent,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            more.MouseEnter += delegate { more.Background = BrushSelectedCardBg; };
            more.MouseLeave += delegate { more.Background = Brushes.Transparent; };
            Action toggle = delegate
            {
                bool refocus = more.IsKeyboardFocused;
                host.Expanded = !host.Expanded;
                RenderNotesList(host, addressNumber, listHost, notes);
                // 組み直した「折りたたむ」／「さらに表示」（末尾）へフォーカスを移す
                if (refocus) FocusLater(delegate { ((UIElement)listHost.Children[listHost.Children.Count - 1]).Focus(); });
            };
            more.MouseLeftButtonUp += delegate { toggle(); };
            MakeKeyActivatable(more, toggle);
            listHost.Children.Add(more);
        }
    }

    // メモ1件の行（記入者+作成日+本文）を組み立てる。自分のメモは左端のバーで示し、
    // 行のホバーで ✎（編集）/✕（削除）を出す。他ユーザーのメモは読み取り専用
    private UIElement BuildNoteRow(NotesHost host, string addressNumber, PersonNoteEntry note, bool isLast)
    {
        bool isMine = string.Equals(note.Author, Environment.UserName, StringComparison.OrdinalIgnoreCase);

        var row = new Border
        {
            // 背景は透明を明示する（null だと文字のない余白がヒットテスト対象外になり、
            // ✎/✕ へ向かう途中で MouseLeave が発火してホバー操作が消えてしまう）
            Background = Brushes.Transparent,
            BorderBrush = BrushLineWeak,
            BorderThickness = new Thickness(0, 0, 0, isLast ? 0 : 1),
            Padding = new Thickness(0, 8, 0, 8)
        };
        var grid = new Grid();
        row.Child = grid;

        // 自分のメモの左バー（一覧の選択行の左アクセントバーと同じ文法）
        if (isMine)
        {
            grid.Children.Add(new Border
            {
                Width = 2.5,
                CornerRadius = new CornerRadius(1.25),
                Background = BrushAccent,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(-12, 1, 0, 1)
            });
        }

        var content = new StackPanel();
        grid.Children.Add(content);

        // --- 記入者（職員番号）+作成日（日時パース不能時は生値）+ホバー操作 ---
        var meta = new Grid { Margin = new Thickness(0, 0, 0, 3) };
        meta.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        meta.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var metaLeft = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
        metaLeft.Children.Add(new TextBlock
        {
            Text = note.Author,
            FontFamily = FontMono,
            FontSize = 10.5,
            FontWeight = FontWeights.Bold,
            Foreground = isMine ? BrushAccent : BrushTextSub
        });
        metaLeft.Children.Add(new TextBlock
        {
            Text = "·",
            FontSize = 10.5,
            Foreground = BrushTextMuted,
            Margin = new Thickness(6, 0, 6, 0)
        });
        metaLeft.Children.Add(new TextBlock
        {
            Text = note.CreatedAt != DateTime.MinValue
                ? note.CreatedAt.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture) : note.CreatedAtRaw,
            FontFamily = FontMono,
            FontSize = 10.5,
            Foreground = BrushTextMuted
        });
        meta.Children.Add(metaLeft);
        content.Children.Add(meta);

        // --- 本文（無加工・改行はそのまま表示する）---
        var body = new TextBlock
        {
            Text = note.Text,
            FontSize = 12,
            Foreground = BrushTextMain,
            TextWrapping = TextWrapping.Wrap
        };
        content.Children.Add(body);

        // 自分のメモのみ編集・削除できる（メモファイルに書き込めない間は出さない）
        if (isMine && persistNotes)
        {
            var actions = new StackPanel { Orientation = Orientation.Horizontal, Visibility = Visibility.Hidden };
            var editButton = MakeNoteActionButton("✎", "編集");
            var deleteButton = MakeNoteActionButton("✕", "削除");
            actions.Children.Add(editButton);
            actions.Children.Add(deleteButton);
            Grid.SetColumn(actions, 1);
            meta.Children.Add(actions);

            // 編集に入ったら操作は出さない（✎ の再クリックでエディタが重複しないように。
            // 編集の終了は保存/キャンセルによるセクション再描画なのでフラグは戻さない）
            // 行にキーでフォーカスがある間もホバーと同じく ✎ ✕ を出し、Enter で書き直す（✎ と同じ）。
            // 削除はマウスの ✕ だけ（キーでは消さない）
            bool editing = false;
            Action beginEdit = delegate
            {
                if (editing) return;
                editing = true;
                actions.Visibility = Visibility.Hidden;
                BeginEditNote(host, addressNumber, note, content, body);
            };
            row.MouseEnter += delegate { if (!editing) actions.Visibility = Visibility.Visible; };
            row.MouseLeave += delegate { if (!row.IsKeyboardFocused) actions.Visibility = Visibility.Hidden; };
            row.GotKeyboardFocus += delegate { if (!editing && row.IsKeyboardFocused) actions.Visibility = Visibility.Visible; };
            row.LostKeyboardFocus += delegate { if (!row.IsMouseOver) actions.Visibility = Visibility.Hidden; };
            MakeKeyActivatable(row, beginEdit);
            editButton.MouseLeftButtonUp += delegate { beginEdit(); };
            deleteButton.MouseLeftButtonUp += delegate { DeleteNote(host, addressNumber, note); };   // 確認なしの即削除（NOTE_UNDO_SECONDS の間は戻せる）
        }

        return row;
    }

    // 削除した行の位置に出す「削除しました　元に戻す　N」の行（淡アイリス面・残り秒は右端）。
    // 「元に戻す」はクリックとキー（Enter／Space）で同じ処理（「さらに表示」のリンクと同じ文法）。
    // 行と残り秒の TextBlock は、捨てるときに外す先・1 秒ごとの更新先として控える（一覧を描き直すたびに新しいものへ差し替わる）
    private UIElement BuildNoteUndoRow()
    {
        noteUndoRow = new Border
        {
            Background = BrushIrisPanelBg,
            CornerRadius = new CornerRadius(7),
            Padding = new Thickness(10, 7, 10, 7),
            Margin = new Thickness(0, 3, 0, 3)
        };
        var row = noteUndoRow;
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        row.Child = grid;

        var left = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
        left.Children.Add(new TextBlock
        {
            Text = "削除しました",
            FontSize = 11,
            Foreground = BrushTextSub,
            VerticalAlignment = VerticalAlignment.Center
        });
        var link = new Border
        {
            Background = Brushes.Transparent,
            CornerRadius = new CornerRadius(5),
            Padding = new Thickness(5, 1, 5, 1),
            Margin = new Thickness(8, 0, 0, 0),
            Cursor = Cursors.Hand,
            VerticalAlignment = VerticalAlignment.Center
        };
        link.Child = new TextBlock
        {
            Text = "元に戻す",
            FontSize = 11,
            Foreground = BrushAccent
        };
        link.MouseEnter += delegate { link.Background = BrushSelectedCardBg; };
        link.MouseLeave += delegate { link.Background = Brushes.Transparent; };
        link.MouseLeftButtonUp += delegate { UndoNoteDelete(); };
        MakeKeyActivatable(link, UndoNoteDelete);
        left.Children.Add(link);
        grid.Children.Add(left);

        noteUndoSecondsText = new TextBlock
        {
            Text = noteUndoSeconds.ToString(),
            FontFamily = FontMono,
            FontSize = 10.5,
            Foreground = BrushTextMuted,
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetColumn(noteUndoSecondsText, 1);
        grid.Children.Add(noteUndoSecondsText);
        return row;
    }

    // メモ行の操作ボタン（✎/✕。設定ビューの行削除 ✕ と同じ意匠の小型版）
    private static Border MakeNoteActionButton(string glyph, string tooltip)
    {
        var btn = new Border
        {
            Width = 18,
            Height = 18,
            CornerRadius = new CornerRadius(9),
            Background = Brushes.Transparent,
            Cursor = Cursors.Hand,
            Margin = new Thickness(2, 0, 0, 0),
            ToolTip = tooltip,
            Child = new TextBlock
            {
                Text = glyph,
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

    // メモの入力欄・書き直しの欄のキー: Ctrl+Enter=追加／保存（押せるときだけ）・Esc=キャンセル。
    // Enter は改行のまま。どちらもボタンと同じ処理を呼び、終わったら「＋ 追加」へ戻る
    //（欄は閉じるか、セクションの組み直しで消えるため）
    private void WireNoteEditorKeys(NotesHost host, TextBox box, Button commit, Button cancel)
    {
        box.PreviewKeyDown += delegate(object s, KeyEventArgs e)
        {
            Button target;
            if (e.Key == Key.Enter && Keyboard.Modifiers == ModifierKeys.Control) target = commit.IsEnabled ? commit : null;
            else if (e.Key == Key.Escape && Keyboard.Modifiers == ModifierKeys.None) target = cancel;
            else return;
            e.Handled = true;
            if (target == null) return;
            target.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            FocusLater(delegate { if (host.AddButton != null && host.AddButton.IsVisible) host.AddButton.Focus(); });
        };
    }

    // メモの追加入力エリア（淡アイリス面カード+複数行入力+キャンセル/追加）を組み立てる。
    // 生成時は非表示。「追加」は本文が空白のみの間は無効
    private Border BuildNoteInputCard(NotesHost host, out TextBox inputBox, out Button submitButton, out Button cancelButton)
    {
        TextBox box;
        var textBorder = MakeNoteTextBox("", out box);

        var cancel = MakeMiniButton("キャンセル");
        var submit = new Button
        {
            Content = "追加",
            Style = (Style)window.FindResource("NotesMiniPrimaryBtn"),
            Margin = new Thickness(8, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            IsEnabled = false
        };
        box.TextChanged += delegate { submit.IsEnabled = box.Text.Trim() != ""; };
        WireNoteEditorKeys(host, box, submit, cancel);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 8, 0, 0)
        };
        buttons.Children.Add(cancel);
        buttons.Children.Add(submit);

        var body = new StackPanel();
        body.Children.Add(textBorder);
        body.Children.Add(buttons);

        var card = new Border
        {
            Background = BrushIrisPanelBg,
            BorderBrush = BrushSelectedLine,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(10, 10, 10, 10),
            Margin = new Thickness(0, 8, 0, 2),
            Visibility = Visibility.Collapsed,
            Child = body
        };
        inputBox = box;
        submitButton = submit;
        cancelButton = cancel;
        return card;
    }

    // メモ本文の入力欄（複数行・フォーカスでアイリス枠）を組み立てる。
    // MakeSettingsInput と同じ意匠の複数行版（設定の変更検知に繋がないため別に組む）
    private Border MakeNoteTextBox(string initialText, out TextBox box)
    {
        var textBox = new TextBox
        {
            Text = initialText,
            BorderThickness = new Thickness(0),
            Background = Brushes.Transparent,
            FontSize = 12,
            Foreground = BrushTextMain,
            AcceptsReturn = true,
            TextWrapping = TextWrapping.Wrap,
            MinHeight = 40,
            MaxHeight = 120,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        };
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = BrushLine,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(7),
            Padding = new Thickness(9, 6, 9, 6),
            Child = textBox
        };
        textBox.GotKeyboardFocus += delegate { border.BorderBrush = BrushAccent; };
        textBox.LostKeyboardFocus += delegate { border.BorderBrush = BrushLine; };
        box = textBox;
        return border;
    }

    // メモをその場で編集に切り替える（本文 → 入力欄+キャンセル/保存）。
    // 保存は本文の上書きのみで作成日時は変えない。キャンセルはセクションの再描画で戻す
    private void BeginEditNote(NotesHost host, string addressNumber, PersonNoteEntry note, StackPanel content, TextBlock body)
    {
        TextBox box;
        var textBorder = MakeNoteTextBox(note.Text, out box);

        var cancel = MakeMiniButton("キャンセル");
        var save = new Button
        {
            Content = "保存",
            Style = (Style)window.FindResource("NotesMiniPrimaryBtn"),
            Margin = new Thickness(8, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        box.TextChanged += delegate { save.IsEnabled = box.Text.Trim() != ""; };
        WireNoteEditorKeys(host, box, save, cancel);
        cancel.Click += delegate { BuildNotesSection(host, addressNumber); };
        save.Click += delegate
        {
            var text = box.Text.Trim();
            if (text == "") return;
            note.Text = text;
            DiscardNoteUndo(host);   // 書き直しのあとは削除を戻せない
            SaveMyNotes();
            BuildNotesSection(host, addressNumber);
            RebuildOtherNotesHost(host, addressNumber);
        };

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 6, 0, 0)
        };
        buttons.Children.Add(cancel);
        buttons.Children.Add(save);

        var editor = new StackPanel { Margin = new Thickness(0, 2, 0, 0) };
        editor.Children.Add(textBorder);
        editor.Children.Add(buttons);

        content.Children.Remove(body);
        content.Children.Add(editor);
        box.Focus();
        box.CaretIndex = box.Text.Length;
    }

    // 宛名番号→メモ件数の表を全ユーザー分から組み直す（起動時とメモの追加・削除時）。
    // 読み込みワーカーからは読むだけなので、差し替えは新しい表への参照の付け替えで行う
    private void RebuildNoteCounts()
    {
        var counts = new Dictionary<string, int>();
        foreach (var pair in myPersonNotes) counts[pair.Key] = pair.Value.Count;
        foreach (var pair in otherPersonNotes)
        {
            int current;
            counts.TryGetValue(pair.Key, out current);
            counts[pair.Key] = current + pair.Value.Count;
        }
        noteCountByAddress = counts;
    }

    // メモの追加・保存・削除のあと、もう一方のパネルが同じ宛名番号を表示していれば（閉じていても）
    // その欄も組み直す。パネルは閉じても表示中の人物を保持し、同じ人物を開き直したときメモの欄を
    // 組み直さないため
    private void RebuildOtherNotesHost(NotesHost host, string addressNumber)
    {
        if (!ReferenceEquals(host, depositNotes) && depositPanelAddressNumberShown == addressNumber)
            BuildNotesSection(depositNotes, addressNumber);
        if (!ReferenceEquals(host, insuranceTab.insuranceNotes) && insuranceTab.insurancePanelAddressNumberShown == addressNumber)
            BuildNotesSection(insuranceTab.insuranceNotes, addressNumber);
    }

    // メモの追加・削除を一覧の印に反映する（同じ宛名番号の行の印だけを変更通知で描き直す。一覧のコンテナは
    // 作り直さないのでスクロール位置は動かない。生命保険の一覧は読み込み済みのときだけ——未読み込みなら
    // 読み込むときに反映される。最近見た人ビューは表示のたびに再構築されるためここでは触らない）
    private void RefreshNoteMarks(string addressNumber)
    {
        RebuildNoteCounts();
        int noteCount;
        noteCountByAddress.TryGetValue(addressNumber, out noteCount);
        foreach (var record in depositAllRecords)
        {
            if (record.AddressNumber != addressNumber) continue;
            record.HasNotes = noteCount > 0;
            record.NotesToolTip = noteCount > 0 ? "メモ " + noteCount + "件" : null;
        }
        if (insuranceTab.IsLoaded) insuranceTab.RefreshInsuranceNoteMarks(addressNumber, noteCount);
    }

    // メモを追加する（自分のファイルとキャッシュへ即時反映・セクションを再描画）
    private void AddNote(NotesHost host, string addressNumber, string text)
    {
        DiscardNoteUndo(host);   // 追加のあとは削除を戻せない
        var now = DateTime.Now;
        var note = new PersonNoteEntry
        {
            Author = Environment.UserName,
            Text = text,
            CreatedAt = now,
            CreatedAtRaw = now.ToString("o")
        };
        List<PersonNoteEntry> list;
        if (!myPersonNotes.TryGetValue(addressNumber, out list))
        {
            list = new List<PersonNoteEntry>();
            myPersonNotes[addressNumber] = list;
        }
        list.Add(note);
        SaveMyNotes();
        RefreshNoteMarks(addressNumber);
        BuildNotesSection(host, addressNumber);
        RebuildOtherNotesHost(host, addressNumber);
    }

    // メモを削除する（確認なしの即削除。NOTE_UNDO_SECONDS の間は「元に戻す」で同じ位置に戻せる。
    // 空になった宛名番号のキーは残さない）
    private void DeleteNote(NotesHost host, string addressNumber, PersonNoteEntry note)
    {
        List<PersonNoteEntry> list;
        if (!myPersonNotes.TryGetValue(addressNumber, out list)) return;
        int index = CollectPersonNotes(addressNumber).IndexOf(note);   // 削除前の並びでの位置（「元に戻す」の行を出す位置）
        if (index < 0) return;   // 一覧にないメモ（組み直し前の行からの呼び出し）は何もしない
        list.Remove(note);
        if (list.Count == 0) myPersonNotes.Remove(addressNumber);
        SaveMyNotes();
        RefreshNoteMarks(addressNumber);
        ShowNoteUndo(host, addressNumber, note, index);   // 欄を組み直す前に控える（一覧の描画が行を出す）
        BuildNotesSection(host, addressNumber);
        RebuildOtherNotesHost(host, addressNumber);
    }

    // 削除したメモを控え、残り秒の計時を始める（控えは全体で 1 つ。前の控えは捨てる）
    private void ShowNoteUndo(NotesHost host, string addressNumber, PersonNoteEntry note, int index)
    {
        DiscardNoteUndo(host);
        noteUndoHost = host;
        noteUndoAddress = addressNumber;
        noteUndoNote = note;
        noteUndoIndex = index;
        noteUndoSeconds = NOTE_UNDO_SECONDS;
        if (noteUndoTimer == null)
        {
            noteUndoTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            noteUndoTimer.Tick += delegate { OnNoteUndoTick(); };
        }
        noteUndoTimer.Start();
    }

    // 1 秒ごとに残り秒を減らし、0 になったら控えを捨てる（行はその欄から外れて消える）
    private void OnNoteUndoTick()
    {
        noteUndoSeconds--;
        if (noteUndoSeconds > 0)
        {
            noteUndoSecondsText.Text = noteUndoSeconds.ToString();
            return;
        }
        DiscardNoteUndo(null);
    }

    // 控えを捨てる（計時も止める）。「元に戻す」の行を出していた欄が rebuilding（呼び出し側がこのあと組み直す欄）
    // 以外なら、その欄から行だけを外す（一覧は描き直さない — 書き直し中の行や入力中の追加の本文を作り直さないため）。
    // 最後の 1 件を消していたときは一覧を描き直して空状態を出す。行にキーのフォーカスがあれば「＋ 追加」へ移す
    private void DiscardNoteUndo(NotesHost rebuilding)
    {
        var host = noteUndoHost;
        var addressNumber = noteUndoAddress;
        var row = noteUndoRow;
        if (noteUndoTimer != null) noteUndoTimer.Stop();
        noteUndoHost = null;
        noteUndoAddress = null;
        noteUndoNote = null;
        noteUndoRow = null;
        noteUndoSecondsText = null;
        if (host == null || ReferenceEquals(host, rebuilding)) return;

        bool refocus = row.IsKeyboardFocusWithin;
        var list = row.Parent as Panel;
        if (list != null) list.Children.Remove(row);
        var notes = CollectPersonNotes(addressNumber);
        if (notes.Count == 0) RenderNotesList(host, addressNumber, host.ListHost, notes);
        if (refocus) FocusLater(delegate { if (host.AddButton != null && host.AddButton.IsVisible) host.AddButton.Focus(); });
    }

    // 「元に戻す」: 控えたメモを同じ作成日時のまま戻して保存し、一覧の印と欄を組み直す（並びは作成日時の
    // 降順なので同じ位置に戻る）。一覧の中にキーのフォーカスがあれば、戻った行へ移す（「折りたたむ」と同じ）
    private void UndoNoteDelete()
    {
        if (noteUndoNote == null) return;
        var host = noteUndoHost;
        var addressNumber = noteUndoAddress;
        var note = noteUndoNote;
        int index = noteUndoIndex;
        bool refocus = host.ListHost.IsKeyboardFocusWithin;
        DiscardNoteUndo(host);

        List<PersonNoteEntry> list;
        if (!myPersonNotes.TryGetValue(addressNumber, out list))
        {
            list = new List<PersonNoteEntry>();
            myPersonNotes[addressNumber] = list;
        }
        list.Add(note);
        SaveMyNotes();
        RefreshNoteMarks(addressNumber);
        BuildNotesSection(host, addressNumber);
        RebuildOtherNotesHost(host, addressNumber);
        if (refocus) FocusLater(delegate
        {
            var rows = (StackPanel)host.ListHost.Children[0];
            var target = (UIElement)rows.Children[Math.Min(index, rows.Children.Count - 1)];
            if (!target.Focus() && host.AddButton != null && host.AddButton.IsVisible) host.AddButton.Focus();   // 折りたたみで戻った行が見えていない・ほかの人の行なら「＋ 追加」へ
        });
    }

    // 自分のメモファイルを保存する（recent_people の SaveRecentPeople と同一規約。
    // 操作の都度呼び出す — ファイルは小さく、異常終了しても直前までのメモが残る）。
    // createdAt を text より先に書く（本文に「"createdAt"」等の語が含まれても
    // 値の取り出しが最初の一致で正しく決まるように）
    private void SaveMyNotes()
    {
        if (!persistNotes) return;

        try
        {
            var historyDir = System.IO.Path.GetDirectoryName(GetPersonNotesPath());
            if (!Directory.Exists(historyDir))
                Directory.CreateDirectory(historyDir);

            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine("  \"notes\": {");
            var addrs = myPersonNotes.Keys.OrderBy(k => k, StringComparer.Ordinal).ToList();
            for (int a = 0; a < addrs.Count; a++)
            {
                var list = myPersonNotes[addrs[a]];
                sb.AppendLine("    \"" + EscapeJson(addrs[a]) + "\": [");
                for (int i = 0; i < list.Count; i++)
                {
                    var createdAt = list[i].CreatedAt == DateTime.MinValue
                        ? list[i].CreatedAtRaw : list[i].CreatedAt.ToString("o");
                    sb.Append("      { \"createdAt\": \"" + EscapeJson(createdAt) + "\"");
                    sb.Append(", \"text\": \"" + EscapeJson(list[i].Text) + "\" }");
                    sb.AppendLine(i < list.Count - 1 ? "," : "");
                }
                sb.AppendLine(a < addrs.Count - 1 ? "    ]," : "    ]");
            }
            sb.AppendLine("  }");
            sb.AppendLine("}");

            WriteUserFileAtomic(GetPersonNotesPath(), sb.ToString());
        }
        catch
        {
            // 書き込み失敗時は以降の保存をスキップし、状況をステータスバーで知らせる
            //（このセッションで追加・編集した分はファイルに残らない）
            persistNotes = false;
            ShowTransientStatus("メモを保存できませんでした（メモの変更はこのセッション限りになります）");
        }
    }
}