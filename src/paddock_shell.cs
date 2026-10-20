// ==============================================================================
// paddock_shell.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// ナビゲーション（ビューの切替・サイドバーの開閉）・イベントの配線・キー操作・
// 画面の要素の取得。
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
    // ナビゲーション（中央ペインのビュー切替）
    // ==============================================================

    // 預金の一覧ビューを表示する（起動時の初期化と、ナビ「預金」クリック）
    private void ShowDepositLedgerView()
    {
        bool switching = depositLedgerView.Visibility != Visibility.Visible;
        insuranceTab.insuranceLedgerView.Visibility = Visibility.Collapsed;
        recentView.Visibility = Visibility.Collapsed;
        seizureView.Visibility = Visibility.Collapsed;
        settingsView.Visibility = Visibility.Collapsed;
        depositLedgerView.Visibility = Visibility.Visible;
        SetNavSelection(depositNavLedger);
        isInsuranceTabCurrent = false;
        UpdateYearSelectorVisibility();

        // 他ビューで開いた詳細パネルは持ち込まない（一覧の行選択との不整合を避ける）
        if (switching && isDepositPanelOpen) CloseDepositDetailPanel();
        if (switching && insuranceTab.isInsurancePanelOpen) insuranceTab.CloseInsuranceDetailPanel();
    }

    // 生命保険の一覧ビューを表示する（ナビ「生命保険」クリック）。初めて表示したときに台帳を検出して読む
    private void ShowInsuranceLedgerView()
    {
        bool switching = insuranceTab.insuranceLedgerView.Visibility != Visibility.Visible;
        depositLedgerView.Visibility = Visibility.Collapsed;
        recentView.Visibility = Visibility.Collapsed;
        seizureView.Visibility = Visibility.Collapsed;
        settingsView.Visibility = Visibility.Collapsed;
        insuranceTab.insuranceLedgerView.Visibility = Visibility.Visible;
        SetNavSelection(insuranceNavLedger);
        isInsuranceTabCurrent = true;
        UpdateYearSelectorVisibility();

        // 他ビューで開いた詳細パネルは持ち込まない（預金のパネルは生命保険の一覧とは別の人物を指しうる）
        if (isDepositPanelOpen) CloseDepositDetailPanel();
        if (switching && insuranceTab.isInsurancePanelOpen) insuranceTab.CloseInsuranceDetailPanel();
        insuranceTab.EnsureInsuranceLedgerLoaded();
    }

    // 差押登録ビューを表示する（詳細パネルはこのビューでは使わないため閉じる）
    private void ShowSeizureView()
    {
        if (isDepositPanelOpen) CloseDepositDetailPanel();
        if (insuranceTab.isInsurancePanelOpen) insuranceTab.CloseInsuranceDetailPanel();
        depositLedgerView.Visibility = Visibility.Collapsed;
        insuranceTab.insuranceLedgerView.Visibility = Visibility.Collapsed;
        recentView.Visibility = Visibility.Collapsed;
        settingsView.Visibility = Visibility.Collapsed;
        seizureView.Visibility = Visibility.Visible;
        SetNavSelection(navSeizure);
        ReloadSeizureLog();   // 切替の都度読み直す（ログは小さく、他プロセスが随時追記するため）
    }

    // 最近見た人ビューを表示する。履歴の反映はビューへ入るたびの再構築で行う
    //（表示中の記録＝パネルからの照会ファイル操作では再構築しない。操作中に
    //  行が並び替わるのを避けるためで、次にこのビューへ入り直したとき反映される）
    private void ShowRecentView()
    {
        bool switching = recentView.Visibility != Visibility.Visible;
        depositLedgerView.Visibility = Visibility.Collapsed;
        insuranceTab.insuranceLedgerView.Visibility = Visibility.Collapsed;
        seizureView.Visibility = Visibility.Collapsed;
        settingsView.Visibility = Visibility.Collapsed;
        recentView.Visibility = Visibility.Visible;
        SetNavSelection(navRecent);

        // 他ビューで開いた詳細パネルは持ち込まない（履歴行の選択との不整合を避ける）
        if (switching && isDepositPanelOpen) CloseDepositDetailPanel();
        if (switching && insuranceTab.isInsurancePanelOpen) insuranceTab.CloseInsuranceDetailPanel();
        if (switching) BuildRecentView();
    }

    // 設定ビューを表示する（詳細パネルはこのビューでは使わないため閉じる）。
    // 中身はビューへ入るたびに再構築して現在の状態（設定ファイルの今の値・フォルダの有無・件数）を映すが、
    // 未保存の変更がある間は再構築せず、編集状態と保存バーをそのまま保持する
    private void ShowSettingsView()
    {
        if (isDepositPanelOpen) CloseDepositDetailPanel();
        if (insuranceTab.isInsurancePanelOpen) insuranceTab.CloseInsuranceDetailPanel();
        depositLedgerView.Visibility = Visibility.Collapsed;
        insuranceTab.insuranceLedgerView.Visibility = Visibility.Collapsed;
        recentView.Visibility = Visibility.Collapsed;
        seizureView.Visibility = Visibility.Collapsed;
        settingsView.Visibility = Visibility.Visible;
        SetNavSelection(navSettings);
        if (!settingsDirty) BuildSettingsView();
    }

    // ナビ項目の活性表示（左3pxアイリスバー+選択背景+文字色。重みは変えない）を切り替える
    private void SetNavActive(Border item, TextBlock text, TextBlock icon, bool active)
    {
        if (active)
        {
            item.BorderBrush = BrushAccent;
            item.Background = BrushSelectedCardBg;
            text.Foreground = BrushTextMain;
            icon.Foreground = BrushAccent;
        }
        else
        {
            item.ClearValue(Border.BorderBrushProperty);       // スタイル既定の Transparent に戻す
            item.Background = Brushes.Transparent;             // null に戻すと余白がクリックに反応しなくなる
            text.Foreground = BrushTextSub;
            icon.Foreground = BrushTextMuted;
        }
    }

    // レール側のナビ項目の活性表示（左3pxアイリスバー+選択面・展開パネル側と同じ文法）
    private void SetRailActive(Border item, TextBlock icon, bool active)
    {
        if (active)
        {
            item.BorderBrush = BrushAccent;
            item.Background = BrushSelectedCardBg;
            icon.Foreground = BrushAccent;
        }
        else
        {
            item.ClearValue(Border.BorderBrushProperty);       // スタイル既定の Transparent に戻す
            item.Background = Brushes.Transparent;             // null に戻すと余白がクリックに反応しなくなる
            icon.Foreground = BrushTextMuted;
        }
    }

    // ナビの活性表示を一括更新する（展開パネル5項目とレール5項目を同期させる）
    private void SetNavSelection(Border active)
    {
        SetNavActive(depositNavLedger, depositNavLedgerText, depositNavLedgerIcon, active == depositNavLedger);
        SetNavActive(insuranceNavLedger, insuranceNavLedgerText, insuranceNavLedgerIcon, active == insuranceNavLedger);
        SetNavActive(navRecent, navRecentText, navRecentIcon, active == navRecent);
        SetNavActive(navSeizure, navSeizureText, navSeizureIcon, active == navSeizure);
        SetNavActive(navSettings, navSettingsText, navSettingsIcon, active == navSettings);
        SetRailActive(depositRailLedger, depositRailLedgerIcon, active == depositNavLedger);
        SetRailActive(insuranceRailLedger, insuranceRailLedgerIcon, active == insuranceNavLedger);
        SetRailActive(railRecent, railRecentIcon, active == navRecent);
        SetRailActive(railSeizure, railSeizureIcon, active == navSeizure);
        SetRailActive(railSettings, railSettingsIcon, active == navSettings);
    }

    // サイドバーの展開（レールと同位置に展開パネルを被せ、48→208px にスライドする）
    private void ExpandSidebar()
    {
        sidebarExpandTimer.Stop();
        sidebarCollapseTimer.Stop();
        if (sidebarExpanded) return;
        sidebarExpanded = true;
        sidebarFlyout.Visibility = Visibility.Visible;
        AnimateSidebarWidth(SIDEBAR_FLYOUT_WIDTH, null);
    }

    // サイドバーの収納（→48px。完了した時点でレールのみに戻す）
    private void CollapseSidebar()
    {
        if (!sidebarExpanded) return;
        sidebarExpanded = false;
        AnimateSidebarWidth(SIDEBAR_RAIL_WIDTH, delegate
        {
            // 収納中に再展開された場合はアニメーションが差し替わりここへ来ないが、
            // 念のため論理状態を確かめてから隠す
            if (sidebarExpanded) return;
            sidebarFlyout.BeginAnimation(FrameworkElement.WidthProperty, null);
            sidebarFlyout.Width = SIDEBAR_RAIL_WIDTH;
            sidebarFlyout.Visibility = Visibility.Collapsed;
        });
    }

    // 展開パネル幅のアニメーション（150ms。中断時は現在幅から滑らかに継続する）
    private void AnimateSidebarWidth(double toWidth, EventHandler completed)
    {
        var animation = new DoubleAnimation
        {
            To = toWidth,
            Duration = TimeSpan.FromMilliseconds(150),
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };
        if (completed != null) animation.Completed += completed;
        sidebarFlyout.BeginAnimation(FrameworkElement.WidthProperty, animation);
    }

    // 表示中（直前に表示していた）タブの年度のコンボ（預金・生命保険）
    private ComboBox CurrentYearCombo()
    {
        return isInsuranceTabCurrent ? insuranceTab.insuranceYearCombo : depositYearCombo;
    }

    // 展開パネルの年度の欄を表示中のタブの分だけ見せる（もう一方は Collapsed）。レールの表示も合わせる
    private void UpdateYearSelectorVisibility()
    {
        depositYearPanel.Visibility = isInsuranceTabCurrent ? Visibility.Collapsed : Visibility.Visible;
        insuranceYearPanel.Visibility = isInsuranceTabCurrent ? Visibility.Visible : Visibility.Collapsed;
        UpdateRailYearLabel();
    }

    // レール下端の種別（預金／生保）と年度の短縮表示（'26 形式。折りたたみ中もどの台帳を見ているかを常時示す。
    // 年度は表示中のタブのコンボから読む。台帳なし・生命保険の未読み込みは —）
    private void UpdateRailYearLabel()
    {
        railKindText.Text = isInsuranceTabCurrent ? "生保" : "預金";
        var combo = CurrentYearCombo();
        var year = GetSelectedYear(combo);
        if (year != null) railYearText.Text = "'" + (year.Value % 100).ToString("00");
        else railYearText.Text = combo.SelectedItem != null ? ALL_YEARS_RAIL : "—";
    }

    // ==============================================================
    // イベントハンドラ
    // ==============================================================

    private void SetupEvents()
    {
        // --- 年度セレクタ ---
        depositYearCombo.SelectionChanged += delegate
        {
            UpdateRailYearLabel();   // レールの年度表示は抑止中（初期化）も追従させる
            if (suppressDepositYearChange) return;
            if (depositYearCombo.SelectedItem == null) return;

            // 年度切替でパネルを閉じる（一覧が総入れ替えになるため、前の年度で
            // 開いた人物の文脈を右側に残さない。ビュー切替時と同じ扱い）
            if (isDepositPanelOpen) CloseDepositDetailPanel();
            depositPanelAddressNumberShown = null;   // 開き直したら組み直す（捨てる前に組んだパネルを使い回さない）
            LoadDepositLedger(GetSelectedYear(depositYearCombo));
        };

        // --- 検索ボックス ---
        // プレースホルダ・クリアボタンの表示切替は即時、絞り込みの反映はデバウンス
        depositSearchBox.TextChanged += delegate
        {
            bool hasText = depositSearchBox.Text.Length > 0;
            depositSearchPlaceholder.Visibility = hasText ? Visibility.Collapsed : Visibility.Visible;
            depositClearButton.Visibility = hasText ? Visibility.Visible : Visibility.Collapsed;
            RestartDepositFilterDebounce();
        };

        // ↓ で一覧の先頭行へ移る（保留中の絞り込みは先に反映する）
        depositSearchBox.PreviewKeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.Key != Key.Down || Keyboard.Modifiers != ModifierKeys.None) return;
            if (depositFilterDebounceTimer != null && depositFilterDebounceTimer.IsEnabled) OnDepositAdvancedFilterChanged();
            if (depositLedgerList.Items.Count == 0) return;
            depositLedgerList.SelectedIndex = 0;
            depositLedgerList.ScrollIntoView(depositLedgerList.Items[0]);
            FocusLater(delegate { FocusDepositLedgerSelection(); });
            e.Handled = true;
        };

        // --- クリアボタン（検索ボックス内 ×。解除は待たずに即時反映する）---
        depositClearButton.Click += delegate { depositSearchBox.Text = ""; depositSearchBox.Focus(); OnDepositAdvancedFilterChanged(); };

        // --- 検索キーの切替（入力中のみ再絞り込み）---
        depositSearchKeyCombo.SelectionChanged += delegate
        {
            if (depositSearchBox.Text.Trim().Length > 0) ApplyDepositFilter();
        };

        // --- 詳細検索 ---
        depositAdvToggle.MouseLeftButtonUp += delegate { ToggleDepositAdvancedPanel(); };
        MakeKeyActivatable(depositAdvToggle, ToggleDepositAdvancedPanel);
        // ホバーは条件チップと同じ色遣い（トグルは常時 選択面のため）
        depositAdvToggle.MouseEnter += delegate { depositAdvToggle.Background = BrushChipVioletBg; };
        depositAdvToggle.MouseLeave += delegate { depositAdvToggle.Background = BrushSelectedCardBg; };

        // 条件の変化は OnDepositAdvancedFilterChanged に集約する（文字入力はデバウンス経由・トグルは即時）
        depositAdvAddressBox.TextChanged += delegate { RestartDepositFilterDebounce(); };
        depositAdvNameBox.TextChanged += delegate { RestartDepositFilterDebounce(); };
        depositAdvKanaBox.TextChanged += delegate { RestartDepositFilterDebounce(); };
        depositAdvInstitutionBox.TextChanged += delegate { RestartDepositFilterDebounce(); };
        depositAdvBalanceMinBox.TextChanged += delegate { RestartDepositFilterDebounce(); };
        depositAdvBalanceMaxBox.TextChanged += delegate { RestartDepositFilterDebounce(); };
        depositAdvBaseFromBox.TextChanged += delegate { RestartDepositFilterDebounce(); };
        depositAdvBaseToBox.TextChanged += delegate { RestartDepositFilterDebounce(); };
        depositAdvTxFromBox.TextChanged += delegate { RestartDepositFilterDebounce(); };
        depositAdvTxToBox.TextChanged += delegate { RestartDepositFilterDebounce(); };
        foreach (var chip in depositAdvDetectChips)
        {
            chip.Checked += delegate { OnDepositAdvancedFilterChanged(); };
            chip.Unchecked += delegate { OnDepositAdvancedFilterChanged(); };
        }
        depositAdvClassRep.Checked += delegate { OnDepositAdvancedFilterChanged(); };
        depositAdvClassRep.Unchecked += delegate { OnDepositAdvancedFilterChanged(); };
        depositAdvClassSpecial.Checked += delegate { OnDepositAdvancedFilterChanged(); };
        depositAdvClassSpecial.Unchecked += delegate { OnDepositAdvancedFilterChanged(); };
        depositAdvClassNoAccount.Checked += delegate { OnDepositAdvancedFilterChanged(); };
        depositAdvClassNoAccount.Unchecked += delegate { OnDepositAdvancedFilterChanged(); };

        // 「条件をすべてクリア」とプレースホルダ（配線後は参照しないためローカルで引く）
        var depositAdvClearLink = (TextBlock)window.FindName("DepositAdvClearLink");
        depositAdvClearLink.MouseLeftButtonUp += delegate { ClearDepositAdvancedFilter(); };
        MakeKeyActivatable(depositAdvClearLink, ClearDepositAdvancedFilter);
        depositAdvClearLink.MouseEnter += delegate { depositAdvClearLink.Foreground = BrushTextMain; };
        depositAdvClearLink.MouseLeave += delegate { depositAdvClearLink.Foreground = BrushTextSub; };
        WirePlaceholder(depositAdvAddressBox, (TextBlock)window.FindName("DepositAdvAddressPh"));
        WirePlaceholder(depositAdvNameBox, (TextBlock)window.FindName("DepositAdvNamePh"));
        WirePlaceholder(depositAdvKanaBox, (TextBlock)window.FindName("DepositAdvKanaPh"));
        WirePlaceholder(depositAdvInstitutionBox, (TextBlock)window.FindName("DepositAdvInstitutionPh"));
        WirePlaceholder(depositAdvBalanceMinBox, (TextBlock)window.FindName("DepositAdvBalanceMinPh"));
        WirePlaceholder(depositAdvBalanceMaxBox, (TextBlock)window.FindName("DepositAdvBalanceMaxPh"));
        WirePlaceholder(depositAdvBaseFromBox, (TextBlock)window.FindName("DepositAdvBaseFromPh"));
        WirePlaceholder(depositAdvBaseToBox, (TextBlock)window.FindName("DepositAdvBaseToPh"));
        WirePlaceholder(depositAdvTxFromBox, (TextBlock)window.FindName("DepositAdvTxFromPh"));
        WirePlaceholder(depositAdvTxToBox, (TextBlock)window.FindName("DepositAdvTxToPh"));

        // 残高欄の3桁区切り整形と、残高・日付欄のフォーカス喪失時の補正
        WireAmountFormatting(depositAdvBalanceMinBox, OnDepositAdvancedFilterChanged);
        WireAmountFormatting(depositAdvBalanceMaxBox, OnDepositAdvancedFilterChanged);
        WireDateNormalize(depositAdvBaseFromBox, OnDepositAdvancedFilterChanged);
        WireDateNormalize(depositAdvBaseToBox, OnDepositAdvancedFilterChanged);
        WireDateNormalize(depositAdvTxFromBox, OnDepositAdvancedFilterChanged);
        WireDateNormalize(depositAdvTxToBox, OnDepositAdvancedFilterChanged);

        // --- 日付欄のカレンダー ---
        foreach (var dateBox in new[] { depositAdvBaseFromBox, depositAdvBaseToBox, depositAdvTxFromBox, depositAdvTxToBox })
            WireAdvDateBox(dateBox, OnDepositAdvancedFilterChanged);
        advDateCalendar.SelectedDatesChanged += delegate
        {
            if (advDateSuppress || advDateCalendar.SelectedDate == null) return;
            CommitAdvCalendarDate(advDateCalendar.SelectedDate.Value);
        };
        // 選択済みと同じ日のクリックは SelectedDatesChanged が発火しないため、
        // 日ボタンを直接検出して確定する（同日再選択への対応。欄の日付を選択した状態で
        // 開くため、開くたびに選択をクリアする方式はとれない）
        advDateCalendar.PreviewMouseLeftButtonUp += delegate(object s, MouseButtonEventArgs me)
        {
            var element = me.OriginalSource as DependencyObject;
            while (element != null && !(element is CalendarDayButton))
            {
                if (element is System.Windows.Controls.Calendar) return;
                element = VisualTreeHelper.GetParent(element);
            }
            var day = element as CalendarDayButton;
            if (day == null || !(day.DataContext is DateTime)) return;
            var date = (DateTime)day.DataContext;
            if (advDateCalendar.SelectedDate != null && advDateCalendar.SelectedDate.Value.Date == date.Date)
                CommitAdvCalendarDate(date);
        };
        // Calendar はクリック後にマウスをキャプチャしたままにすることがあり、
        // 直後の1クリックが吸われる。操作の区切りごとに解放しておく
        advDateCalendar.PreviewMouseUp += delegate
        {
            if (Mouse.Captured is System.Windows.Controls.Calendar || Mouse.Captured is CalendarItem ||
                Mouse.Captured is CalendarDayButton || Mouse.Captured is CalendarButton)
                Mouse.Capture(null);
        };
        advDateTodayLink.MouseLeftButtonUp += delegate { CommitAdvCalendarDate(DateTime.Today); };
        advDateTodayLink.MouseEnter += delegate { advDateTodayLink.Foreground = BrushAccent; };
        advDateTodayLink.MouseLeave += delegate { advDateTodayLink.Foreground = BrushTextSub; };
        // 閉じる操作の明示配線（XAML 側コメント参照）。ルーティングイベントは Popup の
        // 境界を越えて本体まで届くため、ポップアップ内側で発生したクリックは除外する
        // （除外しないと日付選択・見出し・‹›送りのすべてが閉じる操作になってしまう）。
        // 欄自身のクリックは Down で一度閉じたあと Up の配線が開き直す
        window.PreviewMouseDown += delegate
        {
            if (!advDatePopup.IsOpen) return;
            if (advDatePopup.Child != null && advDatePopup.Child.IsMouseOver) return;
            advDatePopup.IsOpen = false;
        };
        window.Deactivated += delegate { if (advDatePopup.IsOpen) advDatePopup.IsOpen = false; };
        window.LocationChanged += delegate { if (advDatePopup.IsOpen) advDatePopup.IsOpen = false; };

        // --- 一覧の行クリック（シングル/ダブルの判別）---
        // パネルが閉じているときのマウスクリックのみ、ダブルクリック判定が確定するまで
        // パネルを開くのを保留する（ダブルクリック時にパネルがチラつくのを防ぐ）。
        // OSのダブルクリック判定は1回目の Down からの経過時間で行われるため、
        // 待機時間は Down 起点で計算し、Up までに消化した分を差し引いて最短化する。
        depositClickTimer = new DispatcherTimer();
        depositClickTimer.Tick += delegate
        {
            depositClickTimer.Stop();
            // 保留中に選択が移動した場合（キーボード操作・再絞り込み等）は破棄する
            if (!isDepositPanelOpen && depositPendingClickRecord != null &&
                ReferenceEquals(depositPendingClickRecord, depositLedgerList.SelectedItem))
                OpenDepositDetailPanel(depositPendingClickRecord);
            depositPendingClickRecord = null;
        };

        depositLedgerList.PreviewMouseLeftButtonDown += delegate(object s, MouseButtonEventArgs e)
        {
            depositLastClickCount = e.ClickCount;
            depositLastMouseDownTick = Environment.TickCount;
        };
        depositLedgerList.PreviewMouseLeftButtonUp += delegate(object s, MouseButtonEventArgs e)
        {
            if (depositLastClickCount >= 2) return;   // ダブルクリック2回目の Up では保留を再開しない
            var record = GetDepositRecordFromSource(e.OriginalSource);
            if (record == null) return;               // ヘッダー・余白・スクロールバーのクリックは無視
            if (isDepositPanelOpen) return;           // 開いているときは SelectionChanged が中身を切替済み

            // 残り待機時間 = ダブルクリック判定時間 − Down からの経過時間
            // （経過済みならダブルクリックはもう成立しないため即座に開く）
            int remaining = (int)GetDoubleClickTime() - (Environment.TickCount - depositLastMouseDownTick);
            if (remaining <= 0)
            {
                OpenDepositDetailPanel(record);
                return;
            }
            depositPendingClickRecord = record;
            depositClickTimer.Stop();
            depositClickTimer.Interval = TimeSpan.FromMilliseconds(remaining);
            depositClickTimer.Start();
        };
        depositLedgerList.MouseDoubleClick += delegate(object s, MouseButtonEventArgs e)
        {
            depositClickTimer.Stop();
            depositPendingClickRecord = null;
            var record = GetDepositRecordFromSource(e.OriginalSource);
            if (record != null) OpenDepositInquiryFileForRow(record.RawFields);
        };

        // --- Enter: 選択行の照会ファイルを開く（ダブルクリックと同じ動作）---
        depositLedgerList.KeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            var record = depositLedgerList.SelectedItem as DepositLedgerRecord;
            if (record != null)
            {
                OpenDepositInquiryFileForRow(record.RawFields);
                e.Handled = true;
            }
        };

        // --- Space=詳細パネルを開く（開いていればパネルのカードへ移る）・Ctrl+O=照会ファイルを開く・
        //     Ctrl+D=差押リストに追加（右クリックメニューと同じ条件）。Space は行コンテナが先に
        //     処理するため Preview で受ける ---
        depositLedgerList.PreviewKeyDown += delegate(object s, KeyEventArgs e)
        {
            var record = depositLedgerList.SelectedItem as DepositLedgerRecord;
            if (record == null) return;
            var modifiers = Keyboard.Modifiers;
            if (e.Key == Key.Space && modifiers == ModifierKeys.None)
            {
                if (!isDepositPanelOpen) OpenDepositDetailPanel(record);
                else FocusDepositPanelCard();
            }
            else if (e.Key == Key.O && modifiers == ModifierKeys.Control) OpenDepositInquiryFileForRow(record.RawFields);
            else if (e.Key == Key.D && modifiers == ModifierKeys.Control)
            {
                if (CanAddToDepositSeizureList(record)) AddDepositRowToSeizureList(record.RawFields);
            }
            else return;
            e.Handled = true;
        };

        // --- 右クリックメニュー（行の右クリックで選択された行が対象）---
        var openMenuItem = new MenuItem { Header = "照会ファイルを開く" };
        openMenuItem.Click += delegate
        {
            var record = depositLedgerList.SelectedItem as DepositLedgerRecord;
            if (record != null) OpenDepositInquiryFileForRow(record.RawFields);
        };
        var listMenu = new ContextMenu();
        listMenu.Items.Add(openMenuItem);
        listMenu.Opened += delegate { openMenuItem.IsEnabled = depositLedgerList.SelectedItem != null; };

        // 「差押リストに追加」は連携先設定時のみ生成・追加する（連携を使わない環境では項目ごと出さない）
        if (depositSeizureScriptPath != null)
        {
            var seizureMenuItem = new MenuItem { Header = "差押リストに追加" };
            seizureMenuItem.Click += delegate
            {
                var record = depositLedgerList.SelectedItem as DepositLedgerRecord;
                if (record != null) AddDepositRowToSeizureList(record.RawFields);
            };
            listMenu.Items.Add(seizureMenuItem);
            listMenu.Opened += delegate
            {
                seizureMenuItem.IsEnabled = CanAddToDepositSeizureList(depositLedgerList.SelectedItem as DepositLedgerRecord);
            };
        }
        depositLedgerList.ContextMenu = listMenu;

        // --- 行選択の変化（クリック・↑↓共通）: パネルが開いていれば中身を切替 ---
        // パネルが閉じているときは ↑↓ だけではパネルを開かない（開くのはクリックか Space）
        depositLedgerList.SelectionChanged += delegate
        {
            if (!isDepositPanelOpen) return;
            var record = depositLedgerList.SelectedItem as DepositLedgerRecord;
            if (record != null) SwitchDepositPanelPerson(record.AddressNumber, record.RawFields);
        };

        // --- フッター「照会ファイルを開く」「差押リストに追加」: 選択中カードが対象 ---
        depositPanelOpenFileButton.Click += delegate { OpenDepositInquiryFileForRow(depositSelectedCardFields); };
        depositPanelSeizureButton.Click += delegate { AddDepositRowToSeizureList(depositSelectedCardFields); };

        // --- ナビ（預金 / 生命保険 / 最近見た人 / 差押登録 / 設定。レール側も同じ動作）---
        depositNavLedger.MouseLeftButtonUp += delegate { ShowDepositLedgerView(); };
        insuranceNavLedger.MouseLeftButtonUp += delegate { ShowInsuranceLedgerView(); };
        navRecent.MouseLeftButtonUp += delegate { ShowRecentView(); };
        navSeizure.MouseLeftButtonUp += delegate { ShowSeizureView(); };
        navSettings.MouseLeftButtonUp += delegate { ShowSettingsView(); };
        depositRailLedger.MouseLeftButtonUp += delegate { ShowDepositLedgerView(); };
        insuranceRailLedger.MouseLeftButtonUp += delegate { ShowInsuranceLedgerView(); };
        railRecent.MouseLeftButtonUp += delegate { ShowRecentView(); };
        railSeizure.MouseLeftButtonUp += delegate { ShowSeizureView(); };
        railSettings.MouseLeftButtonUp += delegate { ShowSettingsView(); };
        railKindText.MouseLeftButtonUp += delegate { ExpandSidebar(); };   // 年度コンボへの導線（種別と年度のどちらを押しても展開する）
        railYearText.MouseLeftButtonUp += delegate { ExpandSidebar(); };
        // キーではビューの切替のあと、そのビューの最初の操作対象へ移る（Ctrl+1・2・3・4・, と同じ）。
        // 年度の短縮表示は Ctrl+Y と同じく年度の一覧まで開く（展開だけではキーで選べないため）
        MakeKeyActivatable(depositRailLedger, ShowDepositLedgerViewByKey, keyFocusInsetStyle);
        MakeKeyActivatable(insuranceRailLedger, ShowInsuranceLedgerViewByKey, keyFocusInsetStyle);
        MakeKeyActivatable(railRecent, ShowRecentViewByKey, keyFocusInsetStyle);
        MakeKeyActivatable(railSeizure, ShowSeizureViewByKey, keyFocusInsetStyle);
        MakeKeyActivatable(railSettings, ShowSettingsViewByKey, keyFocusInsetStyle);
        MakeKeyActivatable(railYearText, OpenYearSelectorByKey);

        // --- 最近見た人・差押登録: 年度の見出しと行を ↑↓ で辿る ---
        recentSectionsPanel.KeyDown += OnSectionListKeyDown;
        seizureSectionsPanel.KeyDown += OnSectionListKeyDown;

        // --- サイドバーの折りたたみ ---
        // レールにカーソルが0.2秒とどまったら展開し、展開パネルから離れて0.2秒で収納する
        //（パネル→一覧へ斜めに移動する途中のチラつき防止）。
        // 年度コンボ（預金・生命保険）のドロップダウンを開いている間は収納しない
        sidebarExpandTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
        sidebarExpandTimer.Tick += delegate { sidebarExpandTimer.Stop(); ExpandSidebar(); };
        sidebarCollapseTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
        sidebarCollapseTimer.Tick += delegate
        {
            sidebarCollapseTimer.Stop();
            if (depositYearCombo.IsDropDownOpen || insuranceTab.insuranceYearCombo.IsDropDownOpen) return;
            CollapseSidebar();
        };
        sidebarRail.MouseEnter += delegate { sidebarExpandTimer.Start(); };
        sidebarRail.MouseLeave += delegate { sidebarExpandTimer.Stop(); };
        sidebarFlyout.MouseEnter += delegate { sidebarCollapseTimer.Stop(); };
        sidebarFlyout.MouseLeave += delegate { sidebarCollapseTimer.Stop(); sidebarCollapseTimer.Start(); };
        // ドロップダウンを閉じた時点で改めて離脱判定を始める（カーソルがパネル上に
        // 戻っていれば MouseEnter が取り消す）。2つのコンボに同じ処理を配線する
        EventHandler yearDropDownClosed = delegate
        {
            if (sidebarExpanded) { sidebarCollapseTimer.Stop(); sidebarCollapseTimer.Start(); }
            if (yearKeyOpened) RestoreFocusAfterYearKey();
        };
        depositYearCombo.DropDownClosed += yearDropDownClosed;
        insuranceTab.insuranceYearCombo.DropDownClosed += yearDropDownClosed;

        // --- 設定ビューの保存バー ---
        settingsSaveButton.Click += delegate { SaveSettings(); };
        settingsDiscardButton.Click += delegate { BuildSettingsView(); };   // ファイルの今の値へ戻す（読み直す）

        // --- アップデート告知を閉じる（✕。閉じた時点で既読として保存する）---
        noticeCloseButton.Click += delegate { CloseNoticeOverlay(); };

        // --- 人物ヘッダーの各項目はダブルクリック（または Tab で移って Enter）でその文字列をクリップボードへ ---
        WireCopy(depositPanelName, delegate { return depositPanelName.Text == NO_NAME_TEXT ? "" : depositPanelName.Text; }, depositPanelCopyChip);
        WireCopy(depositPanelAddressNumber, delegate { return depositPanelAddressNumber.Text; }, depositPanelCopyChip);
        WireCopy(depositPanelKana, delegate { return depositPanelKana.Text; }, depositPanelCopyChip);
        WireCopy(depositPanelAddress, delegate { return depositPanelAddress.Text; }, depositPanelCopyChip);

        // --- パネルを閉じる（✕。Esc は本体のキー処理 OnWindowPreviewKeyDown）---
        depositPanelCloseButton.Click += delegate { CloseDepositDetailPanel(); };
        window.PreviewKeyDown += OnWindowPreviewKeyDown;
        depositDetailPanel.KeyDown += OnDepositDetailPanelKeyDown;

        // --- 年度横断読み込みのデバウンス（行を高速に移動している間は走査を保留する）---
        depositCrossYearTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
        depositCrossYearTimer.Tick += delegate
        {
            depositCrossYearTimer.Stop();
            StartDepositCrossYearLoad();
        };

        // --- 生命保険タブの配線（年度・検索欄・詳細検索）---
        insuranceTab.SetupInsuranceEvents();

        // --- 空白クリックでテキストボックスからフォーカスを外す ---
        // 入力欄の外（パネル・カード等の余白）をクリックしたらキャレットを解放する。
        // クリック先が Panel / Border / Window のときだけ外す（クリックを自前で処理する
        // 部品はそちらが勝つ）。
        // 解放は入力欄にフォーカスがある場合に限る（一覧のキーボード操作（↑↓・Enter）を
        // 担う行フォーカスまでは奪わない）
        window.MouseDown += delegate(object s, MouseButtonEventArgs me)
        {
            if (!(Keyboard.FocusedElement is TextBox)) return;
            if (me.OriginalSource is Panel ||
                me.OriginalSource is Border ||
                me.OriginalSource is Window)
            {
                FocusManager.SetFocusedElement(window, window);
                Keyboard.ClearFocus();
            }
        };
    }

    // ==============================================================
    // キー操作
    // ==============================================================
    //
    // マウスで行う操作はキーでも行える。キーはマウス操作と同じ処理を呼び、動作を
    // 二重に持たない。入力欄の中では文字入力と入力欄の既定のキー（Ctrl+A／C／V／X／Z、
    // やり直しの Ctrl+Y）を優先する。お知らせの表示中は閉じるキーだけを受け付ける。

    private bool yearKeyOpened;                    // 年度の一覧を Ctrl+Y で開いた（閉じたら元の場所へ戻す）
    private ComboBox yearKeyCombo;                 // 〃 開いたコンボ（預金・生命保険。閉じるまでにタブが変わっても取り違えない）
    private IInputElement yearKeyReturnFocus;      // 〃 開く前のフォーカス
    private int yearKeySelectedBefore;             // 〃 開く前の年度（変わったら検索欄へ移る）

    // 本体ウィンドウのキー。Esc は手前のものから1つずつ閉じる（コンボの一覧 → カレンダー →
    // メモの入力 → 詳細パネル）。ショートカットは Ctrl（＋Shift）との組み合わせだけを受け付ける
    private void OnWindowPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (noticeOverlay.Visibility == Visibility.Visible)
        {
            if (e.Key == Key.Escape || e.Key == Key.Enter)
            {
                if (!e.IsRepeat) CloseNoticeOverlay();
                e.Handled = true;
            }
            else if (e.Key == Key.Tab) e.Handled = true;   // 覆われた下の画面へフォーカスを移させない
            return;
        }

        if (e.Key == Key.Escape)
        {
            if (depositYearCombo.IsDropDownOpen || depositSearchKeyCombo.IsDropDownOpen ||
                insuranceTab.insuranceYearCombo.IsDropDownOpen || insuranceTab.insuranceSearchKeyCombo.IsDropDownOpen) return;   // 一覧はコンボ自身が閉じる
            if (advDatePopup.IsOpen)
            {
                advDatePopup.IsOpen = false;
                e.Handled = true;
            }
            else if ((depositNotes.Panel.IsKeyboardFocusWithin || insuranceTab.insuranceNotes.Panel.IsKeyboardFocusWithin) && Keyboard.FocusedElement is TextBox)
            {
                // メモの入力欄・書き直しの欄は欄自身が取り消す（パネルは閉じない）
            }
            else if (isDepositPanelOpen)
            {
                CloseDepositDetailPanel();
                e.Handled = true;
            }
            else if (insuranceTab.isInsurancePanelOpen)
            {
                insuranceTab.CloseInsuranceDetailPanel();
                e.Handled = true;
            }
            return;
        }

        // F5=台帳の再読み込み（預金・生命保険の一覧ビューの表示中のみ。表示中のタブの台帳を読み直す）
        if (e.Key == Key.F5)
        {
            if (e.IsRepeat) { e.Handled = true; return; }   // 長押しのリピートでは読み直さない（押した回数だけ検出と読み込みのワーカーが起きるため）
            if (depositLedgerView.Visibility == Visibility.Visible) ReloadDepositLedger();
            else if (insuranceTab.insuranceLedgerView.Visibility == Visibility.Visible) insuranceTab.ReloadInsuranceLedger();
            else return;
            e.Handled = true;
            return;
        }

        // Ctrl+F・Ctrl+Shift+F は表示中（直前に表示していた）タブに効く
        var modifiers = Keyboard.Modifiers;
        if (modifiers == (ModifierKeys.Control | ModifierKeys.Shift) && e.Key == Key.F)
        {
            // 詳細検索の開閉。開いたら最初の欄へ、閉じたら検索欄へ
            if (isInsuranceTabCurrent)
            {
                ShowInsuranceLedgerView();
                insuranceTab.ToggleInsuranceAdvancedPanelByKey();
            }
            else
            {
                ShowDepositLedgerView();
                ToggleDepositAdvancedPanel();
                if (depositAdvPanelOpen) FocusLater(delegate { depositAdvAddressBox.Focus(); });
                else FocusLater(delegate { depositSearchBox.Focus(); });
            }
            e.Handled = true;
            return;
        }
        if (modifiers != ModifierKeys.Control) return;

        switch (e.Key)
        {
            case Key.F:
                if (isInsuranceTabCurrent)
                {
                    ShowInsuranceLedgerView();
                    insuranceTab.FocusInsuranceSearchByKey();
                }
                else
                {
                    ShowDepositLedgerView();
                    FocusLater(delegate { depositSearchBox.Focus(); depositSearchBox.SelectAll(); });
                }
                break;
            case Key.D1:
                ShowDepositLedgerViewByKey();
                break;
            case Key.D2:
                ShowInsuranceLedgerViewByKey();
                break;
            case Key.D3:
                ShowRecentViewByKey();
                break;
            case Key.D4:
                if (navSeizure.Visibility != Visibility.Visible) return;   // 差押リスト連携を使わない環境
                ShowSeizureViewByKey();
                break;
            case Key.OemComma:
                ShowSettingsViewByKey();
                break;
            case Key.Y:
                if (Keyboard.FocusedElement is TextBoxBase) return;   // 入力欄ではやり直しを優先する
                OpenYearSelectorByKey();
                break;
            case Key.S:
                if (settingsView.Visibility != Visibility.Visible || !settingsSaveBarShown) return;
                SaveSettings();
                break;
            case Key.N:
                if (isDepositPanelOpen && depositNotes.BeginAdd != null) depositNotes.BeginAdd();
                else if (insuranceTab.isInsurancePanelOpen && insuranceTab.insuranceNotes.BeginAdd != null) insuranceTab.insuranceNotes.BeginAdd();
                else return;   // パネルが開いていない・メモファイルに書き込めない間は追加できない
                break;
            default:
                return;
        }
        e.Handled = true;
    }

    // 年度の一覧をキーで開く（表示中のタブの一覧へ切り替え、サイドバーを展開してそのタブの年度コンボの一覧を出す）。
    // 開いたコンボを控え、閉じたときの戻し先はそのコンボのタブで決める
    private void OpenYearSelectorByKey()
    {
        var combo = CurrentYearCombo();
        if (!combo.IsEnabled || combo.IsDropDownOpen) return;   // 台帳が1本もない・既に開いている
        if (isInsuranceTabCurrent) ShowInsuranceLedgerView();
        else ShowDepositLedgerView();
        var returnFocus = Keyboard.FocusedElement;
        int selectedBefore = combo.SelectedIndex;
        ExpandSidebar();
        FocusLater(delegate
        {
            if (!combo.IsVisible) return;
            yearKeyOpened = true;
            yearKeyCombo = combo;
            yearKeyReturnFocus = returnFocus;
            yearKeySelectedBefore = selectedBefore;
            combo.Focus();
            combo.IsDropDownOpen = true;
        });
    }

    // キーで開いた年度の一覧が閉じたら元の場所へ戻る。年度を変えたときはそのタブの検索欄へ
    //（一覧が入れ替わり、元の行が残らないため）。年度を変えずに一覧の外をクリックして
    // 閉じた場合は、クリック先にフォーカスが移っているので触らない（サイドバーは既存の
    // 離脱判定で畳まれる）
    private void RestoreFocusAfterYearKey()
    {
        yearKeyOpened = false;
        var combo = yearKeyCombo;
        bool insurance = combo == insuranceTab.insuranceYearCombo;
        var searchBox = insurance ? insuranceTab.insuranceSearchBox : depositSearchBox;
        var previous = yearKeyReturnFocus as UIElement;
        yearKeyReturnFocus = null;
        bool changed = combo.SelectedIndex != yearKeySelectedBefore;
        FocusLater(delegate
        {
            if (changed) searchBox.Focus();
            else if (!combo.IsKeyboardFocusWithin) return;
            else if (previous != null && previous.IsVisible) previous.Focus();
            else if (!(insurance ? insuranceTab.FocusInsuranceLedgerSelection() : FocusDepositLedgerSelection())) searchBox.Focus();
        });
    }

    // 表示を切り替えた直後の要素はまだ配置されていないため、フォーカスは配置の後に移す
    private void FocusLater(Action focus)
    {
        Dispatcher.BeginInvoke(focus, DispatcherPriority.Loaded);
    }

    // キーでビューを切り替える（Ctrl+1・2・3・4・, とレールのナビの Enter／Space）。ナビのクリックと
    // 同じ処理で切り替えた後、そのビューの最初の操作対象へフォーカスを移す
    private void ShowDepositLedgerViewByKey()
    {
        ShowDepositLedgerView();
        FocusLater(delegate { if (!FocusDepositLedgerSelection()) depositSearchBox.Focus(); });
    }

    private void ShowInsuranceLedgerViewByKey()
    {
        ShowInsuranceLedgerView();
        FocusLater(delegate { if (!insuranceTab.FocusInsuranceLedgerSelection()) insuranceTab.insuranceSearchBox.Focus(); });
    }

    private void ShowRecentViewByKey()
    {
        ShowRecentView();
        FocusLater(delegate
        {
            if (recentSelectedBorder != null && recentSelectedBorder.IsVisible) recentSelectedBorder.Focus();
            else FocusFirstKeyTarget(recentSectionsPanel);
        });
    }

    private void ShowSeizureViewByKey()
    {
        ShowSeizureView();
        FocusLater(delegate { FocusFirstKeyTarget(seizureSectionsPanel); });
    }

    private void ShowSettingsViewByKey()
    {
        ShowSettingsView();
        FocusLater(delegate { if (depositSettingsRows.LedgerBox != null) depositSettingsRows.LedgerBox.Focus(); });
    }

    // root の中でフォーカスを受けられる表示中の部品を、表示順（ビジュアルツリーの深さ優先）で集める。
    // 見つけた部品の中までは探さない
    private static List<UIElement> GetKeyTargets(DependencyObject root)
    {
        var targets = new List<UIElement>();
        CollectKeyTargets(root, targets);
        return targets;
    }

    private static void CollectKeyTargets(DependencyObject parent, List<UIElement> targets)
    {
        int count = VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < count; i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            var element = child as UIElement;
            if (element != null && !element.IsVisible) continue;
            if (element != null && element.Focusable) targets.Add(element);
            else CollectKeyTargets(child, targets);
        }
    }

    private static void FocusFirstKeyTarget(DependencyObject root)
    {
        var targets = GetKeyTargets(root);
        if (targets.Count > 0) targets[0].Focus();
    }

    // 最近見た人・差押登録の年度の見出しと行を ↑↓ で辿る（端ではスクロールに任せる）
    private static void OnSectionListKeyDown(object sender, KeyEventArgs e)
    {
        if ((e.Key != Key.Up && e.Key != Key.Down) || Keyboard.Modifiers != ModifierKeys.None) return;
        var targets = GetKeyTargets((DependencyObject)sender);
        int index = targets.IndexOf(e.OriginalSource as UIElement);
        int next = index + (e.Key == Key.Down ? 1 : -1);
        if (index < 0 || next < 0 || next >= targets.Count) return;
        targets[next].Focus();
        e.Handled = true;
    }

    // 詳細パネルの中のキー（カード・人物ヘッダー・ボタンなど、メモの入力欄以外）。
    // ← で元の一覧へ戻る・Ctrl+M で選択中のカードの明細ウィンドウ（⧉ が出ているとき）・
    // Ctrl+D で選択中のカードを差押リストへ（下部の「差押リストに追加」が押せるとき）
    private void OnDepositDetailPanelKeyDown(object sender, KeyEventArgs e)
    {
        if (e.OriginalSource is TextBox) return;
        var modifiers = Keyboard.Modifiers;
        if (e.Key == Key.Left && modifiers == ModifierKeys.None)
        {
            ReturnFocusFromDepositPanel();
        }
        else if (e.Key == Key.M && modifiers == ModifierKeys.Control)
        {
            DepositCardExpansion expansion;
            if (depositSelectedCardBorder != null && depositCardExpansions.TryGetValue(depositSelectedCardBorder, out expansion) &&
                expansion.OpenWindow != null)
                expansion.OpenWindow();
        }
        else if (e.Key == Key.D && modifiers == ModifierKeys.Control)
        {
            if (depositPanelSeizureButton.Visibility == Visibility.Visible && depositPanelSeizureButton.IsEnabled)
                AddDepositRowToSeizureList(depositSelectedCardFields);
        }
        else return;
        e.Handled = true;
    }

    // フォーカス枠（XAML の KeyFocus／KeyFocusPill／KeyFocusInset。コードで組む部品に付ける）
    private Style keyFocusStyle;
    private Style keyFocusPillStyle;
    private Style keyFocusInsetStyle;

    // クリックで操作する部品（Border・TextBlock など）をキーでも押せるようにする:
    // フォーカスを受け、Enter／Space で action を呼ぶ（クリックと同じ処理）。マウスで押しても
    // フォーカスは移らない（Border・TextBlock はクリックでフォーカスを取らない）。
    // focusStyle を省くと角丸用の枠（KeyFocus）
    private void MakeKeyActivatable(FrameworkElement element, Action action, Style focusStyle = null)
    {
        element.Focusable = true;
        element.FocusVisualStyle = focusStyle ?? keyFocusStyle;
        element.KeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.OriginalSource != element || Keyboard.Modifiers != ModifierKeys.None) return;
            if (e.Key != Key.Enter && e.Key != Key.Space) return;
            action();
            e.Handled = true;
        };
    }

    // ==============================================================
    // ウィンドウ構築（XAML）
    // ==============================================================

    private void FindUiElements()
    {
        depositNavLedger         = (Border)window.FindName("DepositNavLedger");
        depositNavLedgerText     = (TextBlock)window.FindName("DepositNavLedgerText");
        depositNavLedgerIcon     = (TextBlock)window.FindName("DepositNavLedgerIcon");
        insuranceNavLedger       = (Border)window.FindName("InsuranceNavLedger");
        insuranceNavLedgerText   = (TextBlock)window.FindName("InsuranceNavLedgerText");
        insuranceNavLedgerIcon   = (TextBlock)window.FindName("InsuranceNavLedgerIcon");
        navRecent                = (Border)window.FindName("NavRecent");
        navRecentText            = (TextBlock)window.FindName("NavRecentText");
        navRecentIcon            = (TextBlock)window.FindName("NavRecentIcon");
        navSeizure               = (Border)window.FindName("NavSeizure");
        navSeizureText           = (TextBlock)window.FindName("NavSeizureText");
        navSeizureIcon           = (TextBlock)window.FindName("NavSeizureIcon");
        navSettings              = (Border)window.FindName("NavSettings");
        navSettingsText          = (TextBlock)window.FindName("NavSettingsText");
        navSettingsIcon          = (TextBlock)window.FindName("NavSettingsIcon");
        sidebarRail              = (Border)window.FindName("SidebarRail");
        sidebarFlyout            = (Border)window.FindName("SidebarFlyout");
        railKindText             = (TextBlock)window.FindName("RailKindText");
        railYearText             = (TextBlock)window.FindName("RailYearText");
        depositRailLedger        = (Border)window.FindName("DepositRailLedger");
        depositRailLedgerIcon    = (TextBlock)window.FindName("DepositRailLedgerIcon");
        insuranceRailLedger      = (Border)window.FindName("InsuranceRailLedger");
        insuranceRailLedgerIcon  = (TextBlock)window.FindName("InsuranceRailLedgerIcon");
        railRecent               = (Border)window.FindName("RailRecent");
        railRecentIcon           = (TextBlock)window.FindName("RailRecentIcon");
        railSeizure              = (Border)window.FindName("RailSeizure");
        railSeizureIcon          = (TextBlock)window.FindName("RailSeizureIcon");
        railSettings             = (Border)window.FindName("RailSettings");
        railSettingsIcon         = (TextBlock)window.FindName("RailSettingsIcon");
        depositLedgerView        = (DockPanel)window.FindName("DepositLedgerView");
        recentView               = (DockPanel)window.FindName("RecentView");
        seizureView              = (DockPanel)window.FindName("SeizureView");
        settingsView             = (DockPanel)window.FindName("SettingsView");

        depositYearPanel         = (StackPanel)window.FindName("DepositYearPanel");
        insuranceYearPanel       = (StackPanel)window.FindName("InsuranceYearPanel");
        depositYearCombo         = (ComboBox)window.FindName("DepositYearCombo");
        depositSearchKeyCombo    = (ComboBox)window.FindName("DepositSearchKeyCombo");
        depositSearchBox         = (TextBox)window.FindName("DepositSearchBox");
        depositSearchPlaceholder = (TextBlock)window.FindName("DepositSearchPlaceholder");
        depositClearButton       = (Button)window.FindName("DepositClearButton");
        // 詳細検索
        depositAdvToggle         = (Border)window.FindName("DepositAdvToggle");
        depositAdvCountBadge     = (Border)window.FindName("DepositAdvCountBadge");
        depositAdvCountText      = (TextBlock)window.FindName("DepositAdvCountText");
        depositAdvToggleCaret    = (TextBlock)window.FindName("DepositAdvToggleCaret");
        depositAdvPanel          = (Border)window.FindName("DepositAdvPanel");
        depositFilterChipsRow    = (WrapPanel)window.FindName("DepositFilterChipsRow");
        depositAdvAddressBox     = (TextBox)window.FindName("DepositAdvAddressBox");
        depositAdvNameBox        = (TextBox)window.FindName("DepositAdvNameBox");
        depositAdvKanaBox        = (TextBox)window.FindName("DepositAdvKanaBox");
        depositAdvInstitutionBox = (TextBox)window.FindName("DepositAdvInstitutionBox");
        depositAdvBalanceMinBox  = (TextBox)window.FindName("DepositAdvBalanceMinBox");
        depositAdvBalanceMaxBox  = (TextBox)window.FindName("DepositAdvBalanceMaxBox");
        depositAdvBalanceMinBd   = (Border)window.FindName("DepositAdvBalanceMinBd");
        depositAdvBalanceMaxBd   = (Border)window.FindName("DepositAdvBalanceMaxBd");
        depositAdvBaseFromBox    = (TextBox)window.FindName("DepositAdvBaseFromBox");
        depositAdvBaseToBox      = (TextBox)window.FindName("DepositAdvBaseToBox");
        depositAdvBaseFromBd     = (Border)window.FindName("DepositAdvBaseFromBd");
        depositAdvBaseToBd       = (Border)window.FindName("DepositAdvBaseToBd");
        depositAdvTxFromBox      = (TextBox)window.FindName("DepositAdvTxFromBox");
        depositAdvTxToBox        = (TextBox)window.FindName("DepositAdvTxToBox");
        depositAdvTxFromBd       = (Border)window.FindName("DepositAdvTxFromBd");
        depositAdvTxToBd         = (Border)window.FindName("DepositAdvTxToBd");
        advDatePopup             = (Popup)window.FindName("AdvDatePopup");
        advDateCalendar          = (System.Windows.Controls.Calendar)window.FindName("AdvDateCalendar");
        advDateTodayLink         = (TextBlock)window.FindName("AdvDateTodayLink");
        advDateTodayValue        = (Run)window.FindName("AdvDateTodayValue");
        depositAdvDetectChips = new ToggleButton[6];   // DETECTION_NAMES と同順
        for (int i = 0; i < depositAdvDetectChips.Length; i++)
            depositAdvDetectChips[i] = (ToggleButton)window.FindName("DepositAdvDetect" + i);
        depositAdvClassRep       = (ToggleButton)window.FindName("DepositAdvClassRep");
        depositAdvClassSpecial   = (ToggleButton)window.FindName("DepositAdvClassSpecial");
        depositAdvClassNoAccount = (ToggleButton)window.FindName("DepositAdvClassNoAccount");
        depositAdvStaffPanel     = (WrapPanel)window.FindName("DepositAdvStaffPanel");
        depositAdvStaffEmptyText = (TextBlock)window.FindName("DepositAdvStaffEmptyText");
        depositLedgerList        = (ListView)window.FindName("DepositLedgerList");
        depositDetectColumn      = (GridViewColumn)window.FindName("DepositDetectColumn");
        depositLoadingOverlay    = (Grid)window.FindName("DepositLoadingOverlay");
        depositSpinnerRotation   = (RotateTransform)((System.Windows.Shapes.Path)window.FindName("DepositSpinnerPath")).RenderTransform;
        depositEmptyStateText    = (TextBlock)window.FindName("DepositEmptyStateText");
        depositStatusLeft        = (TextBlock)window.FindName("DepositStatusLeft");
        depositStatusRight       = (TextBlock)window.FindName("DepositStatusRight");

        // 詳細パネル
        depositDetailPanel        = (Border)window.FindName("DepositDetailPanel");
        depositPanelAvatarText    = (TextBlock)window.FindName("DepositPanelAvatarText");
        depositPanelName          = (TextBlock)window.FindName("DepositPanelName");
        depositPanelAddressNumber = (Run)window.FindName("DepositPanelAddressNumber");
        depositPanelKana          = (Run)window.FindName("DepositPanelKana");
        depositPanelAddress       = (TextBlock)window.FindName("DepositPanelAddress");
        depositPanelCopyChip      = (Border)window.FindName("DepositPanelCopyChip");
        depositPanelStatRounds    = (TextBlock)window.FindName("DepositPanelStatRounds");
        depositPanelStatAccounts  = (TextBlock)window.FindName("DepositPanelStatAccounts");
        depositPanelStatFirst     = (TextBlock)window.FindName("DepositPanelStatFirst");
        depositPanelScroll        = (ScrollViewer)window.FindName("DepositPanelScroll");
        depositNotes.Panel        = (StackPanel)window.FindName("DepositNotesPanel");
        depositTimelinePanel      = (StackPanel)window.FindName("DepositTimelinePanel");
        depositPanelFootnote      = (TextBlock)window.FindName("DepositPanelFootnote");
        depositPanelMiniSpinner   = (System.Windows.Shapes.Path)window.FindName("DepositPanelMiniSpinner");
        depositPanelSpinnerRotation = (RotateTransform)depositPanelMiniSpinner.RenderTransform;
        depositPanelCloseButton   = (Button)window.FindName("DepositPanelCloseButton");
        depositPanelOpenFileButton = (Button)window.FindName("DepositPanelOpenFileButton");
        depositPanelSeizureButton = (Button)window.FindName("DepositPanelSeizureButton");

        // 差押登録ビュー
        seizureCard                 = (Border)window.FindName("SeizureCard");
        seizureSectionsPanel        = (StackPanel)window.FindName("SeizureSectionsPanel");
        seizureEmptyText            = (TextBlock)window.FindName("SeizureEmptyText");
        seizureFooterText           = (TextBlock)window.FindName("SeizureFooterText");

        // 最近見た人ビュー
        recentCard                 = (Border)window.FindName("RecentCard");
        recentSectionsPanel        = (StackPanel)window.FindName("RecentSectionsPanel");
        recentSubtitleText         = (TextBlock)window.FindName("RecentSubtitleText");
        recentEmptyText            = (TextBlock)window.FindName("RecentEmptyText");
        recentFooterText           = (TextBlock)window.FindName("RecentFooterText");

        // 設定ビュー
        settingsPanel                = (StackPanel)window.FindName("SettingsPanel");
        settingsSaveBar              = (Border)window.FindName("SettingsSaveBar");
        settingsChangeSummary        = (TextBlock)window.FindName("SettingsChangeSummary");
        settingsSaveButton           = (Button)window.FindName("SettingsSaveButton");
        settingsDiscardButton        = (Button)window.FindName("SettingsDiscardButton");
        settingsFooterLeft           = (TextBlock)window.FindName("SettingsFooterLeft");
        settingsFooterRight          = (TextBlock)window.FindName("SettingsFooterRight");

        // アップデート告知オーバーレイ
        noticeOverlay            = (Grid)window.FindName("NoticeOverlay");
        noticeContent            = (StackPanel)window.FindName("NoticeContent");
        noticeCloseButton        = (Button)window.FindName("NoticeCloseButton");

        // 生命保険タブ
        insuranceTab.FindInsuranceUiElements();
    }
}