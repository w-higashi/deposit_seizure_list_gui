// ==============================================================================
// paddock_notice.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// アップデート告知（update_notice.json）と既読の保存。
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
    // アップデート告知（update_notice.json）
    // ==============================================================
    //
    // exe 同階層の update_notice.json をユーザーごとに初回のみオーバーレイ表示する。
    // 既読バージョンは history/{ユーザー名}/notice_state.json に保管する — config は全ユーザー共有かつ
    // 手書き前提のため、ユーザー固有の状態を混ぜない。

    // 既読情報ファイルのパス（history/{ユーザー名}/notice_state.json）
    private string GetNoticeStatePath()
    {
        return System.IO.Path.Combine(userHistDir, "notice_state.json");
    }

    // 既読バージョンを読み込む（読めない場合は未読扱い = 告知が再表示されるだけで実害なし）
    private void LoadNoticeState()
    {
        try
        {
            var path = GetNoticeStatePath();
            if (!File.Exists(path)) return;
            var json = File.ReadAllText(path, Encoding.UTF8);
            lastSeenVersion = JsonHelper.GetString(json, "lastSeenVersion") ?? "";
        }
        catch (Exception ex)
        {
            // 未読扱いで続行
            WriteErrorLog("既読の読み込み", ex);
        }
    }

    // 既読バージョンを保存する（書き込み失敗時はスキップ。次回起動で告知が再表示されるだけ）
    private void SaveNoticeState()
    {
        try
        {
            var path = GetNoticeStatePath();
            var histDir = System.IO.Path.GetDirectoryName(path);
            if (!Directory.Exists(histDir)) Directory.CreateDirectory(histDir);
            WriteUserFileAtomic(path,
                "{\n  \"lastSeenVersion\": \"" + EscapeJson(lastSeenVersion) + "\"\n}\n");
        }
        catch
        {
            // 保存できなくても動作に影響しない
        }
    }

    // 起動時にアップデート告知の表示要否をチェック
    private void CheckUpdateNotice()
    {
        try
        {
            var noticePath = System.IO.Path.Combine(exeDir, NOTICE_FILE);
            if (!File.Exists(noticePath)) return;

            var notice = LoadUpdateNotice(noticePath);
            if (notice == null || string.IsNullOrEmpty(notice.Version)) return;
            if (notice.Version == lastSeenVersion) return;
            if (notice.Features.Count == 0) return;

            pendingNoticeVersion = notice.Version;
            BuildNoticeUI(notice);
            ShowNoticeOverlay();
        }
        catch (Exception ex)
        {
            // 告知読み込み失敗時は通知をスキップ
            WriteErrorLog("告知の読み込み", ex);
        }
    }

    // 設定ビューの「お知らせを再表示」。最新の update_notice.json を読み直して表示する
    //（既読情報はここでは変えず、閉じたときに通常どおり保存される）
    private void ReplayUpdateNotice()
    {
        try
        {
            var noticePath = System.IO.Path.Combine(exeDir, NOTICE_FILE);
            var notice = File.Exists(noticePath) ? LoadUpdateNotice(noticePath) : null;
            if (notice == null || string.IsNullOrEmpty(notice.Version) || notice.Features.Count == 0)
            {
                ShowSettingsFooterTransient("update_notice.json を読み込めませんでした");
                return;
            }
            pendingNoticeVersion = notice.Version;
            BuildNoticeUI(notice);
            ShowNoticeOverlay();
        }
        catch
        {
            ShowSettingsFooterTransient("update_notice.json を読み込めませんでした");
        }
    }

    // update_notice.json を手動パースして UpdateNotice を返す
    //（括弧の対応は文字列リテラルを解する JsonHelper.FindMatchingBracket で取る —
    //  告知本文に { } [ ] を含んでも壊れない）
    private UpdateNotice LoadUpdateNotice(string path)
    {
        var json = File.ReadAllText(path, Encoding.UTF8);
        var notice = new UpdateNotice();
        notice.Version = JsonHelper.GetString(json, "version");

        // features 配列の範囲を取得
        var featuresIdx = json.IndexOf("\"features\"");
        if (featuresIdx < 0) return notice;
        var featArrStart = json.IndexOf('[', featuresIdx);
        if (featArrStart < 0) return notice;
        var featArrEnd = JsonHelper.FindMatchingBracket(json, featArrStart, '[', ']');
        if (featArrEnd < 0) return notice;

        // 各 feature オブジェクトを解析
        int pos = featArrStart + 1;
        while (pos < featArrEnd)
        {
            var objStart = json.IndexOf('{', pos);
            if (objStart < 0 || objStart >= featArrEnd) break;
            var objEnd = JsonHelper.FindMatchingBracket(json, objStart, '{', '}');
            if (objEnd < 0) break;
            var objContent = json.Substring(objStart, objEnd - objStart + 1);

            var feature = new NoticeFeature();
            feature.Title = JsonHelper.GetString(objContent, "title") ?? "";
            feature.Description = JsonHelper.GetString(objContent, "description") ?? "";

            // sections 配列
            var secIdx = objContent.IndexOf("\"sections\"");
            if (secIdx >= 0)
            {
                var secArrStart = objContent.IndexOf('[', secIdx);
                var secArrEnd = secArrStart < 0 ? -1
                    : JsonHelper.FindMatchingBracket(objContent, secArrStart, '[', ']');
                if (secArrStart >= 0 && secArrEnd >= 0)
                {
                    int sp = secArrStart + 1;
                    while (sp < secArrEnd)
                    {
                        var sObjStart = objContent.IndexOf('{', sp);
                        if (sObjStart < 0 || sObjStart >= secArrEnd) break;
                        var sObjEnd = JsonHelper.FindMatchingBracket(objContent, sObjStart, '{', '}');
                        if (sObjEnd < 0) break;
                        var sObj = objContent.Substring(sObjStart, sObjEnd - sObjStart + 1);

                        var section = new NoticeSection();
                        section.Heading = JsonHelper.GetString(sObj, "heading") ?? "";

                        // items 配列（文字列の並び）
                        var itemsIdx = sObj.IndexOf("\"items\"");
                        if (itemsIdx >= 0)
                        {
                            var iArrStart = sObj.IndexOf('[', itemsIdx);
                            var iArrEnd = iArrStart < 0 ? -1
                                : JsonHelper.FindMatchingBracket(sObj, iArrStart, '[', ']');
                            if (iArrStart >= 0 && iArrEnd >= 0)
                            {
                                var iContent = sObj.Substring(iArrStart + 1, iArrEnd - iArrStart - 1);
                                foreach (Match m in Regex.Matches(iContent, "\"((?:[^\"\\\\]|\\\\.)*)\""))
                                    section.Items.Add(JsonHelper.UnescapeJsonString(m.Groups[1].Value));
                            }
                        }

                        feature.Sections.Add(section);
                        sp = sObjEnd + 1;
                    }
                }
            }

            notice.Features.Add(feature);
            pos = objEnd + 1;
        }

        return notice;
    }

    // UpdateNotice の内容で告知カードの中身を動的構築する
    private void BuildNoticeUI(UpdateNotice notice)
    {
        noticeContent.Children.Clear();

        // バージョンバッジ
        var versionBadge = new Border
        {
            Background = BrushChipVioletBg,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10, 3, 10, 4),
            HorizontalAlignment = HorizontalAlignment.Left,
            Margin = new Thickness(0, 8, 0, 0)
        };
        versionBadge.Child = new TextBlock
        {
            Text = "ver " + notice.Version,
            FontSize = 11,
            Foreground = BrushChipVioletFg
        };
        noticeContent.Children.Add(versionBadge);

        // 区切り線
        noticeContent.Children.Add(new Border
        {
            BorderBrush = BrushLineWeak,
            BorderThickness = new Thickness(0, 1, 0, 0),
            Margin = new Thickness(0, 14, 0, 16)
        });

        // 各 feature ブロック
        for (int fi = 0; fi < notice.Features.Count; fi++)
        {
            var feature = notice.Features[fi];
            if (fi > 0)
                noticeContent.Children.Add(new Border { Height = 14 });

            // タイトル
            //（告知カードの見出し類は Bold。Meiryo UI に Medium 面はなく、WPF は 500 を
            //  Regular へ解決してしまうため、見出しの階層が消える）
            noticeContent.Children.Add(new TextBlock
            {
                Text = feature.Title,
                FontSize = 13,
                FontWeight = FontWeights.Bold,
                Foreground = BrushAccent,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            // 説明文
            if (!string.IsNullOrEmpty(feature.Description))
            {
                noticeContent.Children.Add(new TextBlock
                {
                    Text = feature.Description,
                    FontSize = 12,
                    Foreground = BrushTextSub,
                    TextWrapping = TextWrapping.Wrap,
                    LineHeight = 20,
                    Margin = new Thickness(0, 0, 0, 12)
                });
            }

            // セクション
            for (int si = 0; si < feature.Sections.Count; si++)
            {
                var section = feature.Sections[si];
                noticeContent.Children.Add(new TextBlock
                {
                    Text = section.Heading,
                    FontSize = 12,
                    FontWeight = FontWeights.Bold,
                    Foreground = BrushTextMain,
                    Margin = new Thickness(0, 0, 0, 4)
                });

                foreach (var item in section.Items)
                {
                    noticeContent.Children.Add(new TextBlock
                    {
                        Text = item,
                        FontSize = 12,
                        Foreground = BrushTextSub,
                        TextWrapping = TextWrapping.Wrap,
                        LineHeight = 20,
                        Margin = new Thickness(0, 0, 0, 2)
                    });
                }

                // セクション間のスペーサー（最後のセクションには追加しない）
                if (si < feature.Sections.Count - 1)
                    noticeContent.Children.Add(new Border { Height = 10 });
            }
        }
    }

    // 告知オーバーレイを表示する（フェードイン150ms）
    private void ShowNoticeOverlay()
    {
        noticeOverlay.BeginAnimation(UIElement.OpacityProperty, null);
        noticeOverlay.Opacity = 0;
        noticeOverlay.Visibility = Visibility.Visible;
        noticeOverlay.BeginAnimation(UIElement.OpacityProperty,
            new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150))));
        FocusLater(delegate { noticeCloseButton.Focus(); });   // Esc／Enter で閉じられる状態にする
    }

    // 告知オーバーレイを閉じる（フェードアウト → 既読バージョンを保存）
    private void CloseNoticeOverlay()
    {
        if (!string.IsNullOrEmpty(pendingNoticeVersion))
            lastSeenVersion = pendingNoticeVersion;
        SaveNoticeState();

        var anim = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(150)));
        anim.Completed += delegate
        {
            noticeOverlay.Visibility = Visibility.Collapsed;
            noticeOverlay.BeginAnimation(UIElement.OpacityProperty, null);
            noticeOverlay.Opacity = 1;
            if (depositLedgerView.Visibility == Visibility.Visible) depositSearchBox.Focus();
        };
        noticeOverlay.BeginAnimation(UIElement.OpacityProperty, anim);
    }
}