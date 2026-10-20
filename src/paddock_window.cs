// ==============================================================================
// paddock_window.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// ウィンドウの XAML（スタイルとレイアウトの全体。生命保険ビューの XAML 片は
// paddock_insurance_window.cs の定数を文字列としてつなぐ）。
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
    // ウィンドウの XAML をインライン文字列で定義し、XamlReader.Load で読み込む
    private Window BuildWindow()
    {
        var xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Title='PADDOCK'
        Width='1280' Height='800' MinWidth='1100' MinHeight='600'
        WindowStartupLocation='CenterScreen'
        FontFamily='Meiryo UI' FontSize='12'
        Background='#F7F7F9'>
    <Window.Resources>
        <!-- 画面の表示倍率。ルート Grid・カレンダーポップアップ・検出ツールチップが
             この1インスタンスを共有し、コード側の ScaleX/Y の書き換えで一括追従する -->
        <ScaleTransform x:Key='DisplayScale' ScaleX='1' ScaleY='1'/>

        <BooleanToVisibilityConverter x:Key='BoolVis'/>

        <!-- キーボードで移ったときのフォーカス枠（マウス操作では出ない）。角丸の部品用と、
             丸・ピル形の部品用は部品の外側 2px に描く -->
        <Style x:Key='KeyFocus'>
            <Setter Property='Control.Template'><Setter.Value>
                <ControlTemplate>
                    <Border Margin='-2' BorderBrush='#5B5BD6' BorderThickness='2' CornerRadius='9'
                            SnapsToDevicePixels='True'/>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>
        <Style x:Key='KeyFocusPill'>
            <Setter Property='Control.Template'><Setter.Value>
                <ControlTemplate>
                    <Border Margin='-2' BorderBrush='#5B5BD6' BorderThickness='2' CornerRadius='14'
                            SnapsToDevicePixels='True'/>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>
        <!-- 幅いっぱいの行・見出し・ナビ用（外側に余白がないため内側 2px に描く）-->
        <Style x:Key='KeyFocusInset'>
            <Setter Property='Control.Template'><Setter.Value>
                <ControlTemplate>
                    <Border BorderBrush='#5B5BD6' BorderThickness='2' SnapsToDevicePixels='True'/>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- テーブル行（左3pxアクセントバー方式・仮想化と両立するGrid重ね合わせ）-->
        <Style x:Key='LedgerItem' TargetType='ListViewItem'>
            <Setter Property='Foreground' Value='#1A1523'/>
            <Setter Property='HorizontalContentAlignment' Value='Stretch'/>
            <Setter Property='Padding' Value='0'/><Setter Property='Margin' Value='0'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='ListViewItem'>
                    <!-- 行高は固定32px（チップ有無で行高が揺れないように・最近見た人/差押登録の行と統一。
                         仮想化の見積りも安定する）-->
                    <Grid Height='32'>
                        <Border x:Name='rowBd' Background='White'
                                BorderBrush='#F1F0F4' BorderThickness='0,0,0,1'>
                            <GridViewRowPresenter VerticalAlignment='Center'/></Border>
                        <Border x:Name='accent' HorizontalAlignment='Left'
                                Width='3' Background='Transparent'/></Grid>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='rowBd' Property='Background' Value='#FAFAFC'/></Trigger>
                        <Trigger Property='IsSelected' Value='True'>
                            <Setter TargetName='accent' Property='Background' Value='#5B5BD6'/>
                            <Setter TargetName='rowBd' Property='Background' Value='#F6F5FE'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
            <Style.Triggers>
                <!-- 口座なし行・契約なし行・電子回答不可行は全体をミュート色で描画（預金・生命保険の行の型が同名のプロパティで示す）-->
                <DataTrigger Binding='{Binding IsMutedRow}' Value='True'>
                    <Setter Property='Foreground' Value='#B9B8C0'/>
                </DataTrigger>
            </Style.Triggers>
        </Style>

        <!-- テーブルヘッダー（フラット・リサイズなし）-->
        <Style TargetType='GridViewColumnHeader'>
            <Setter Property='Background' Value='White'/>
            <Setter Property='Foreground' Value='#9B9AA3'/>
            <Setter Property='FontSize' Value='11'/>
            <Setter Property='HorizontalContentAlignment' Value='Left'/>
            <!-- 左18px = セルの Margin 12px + GridViewRowPresenter がセルへ内蔵付与する 6px。
                 見出しを行の実文字位置に揃える -->
            <Setter Property='Padding' Value='18,7,10,7'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='GridViewColumnHeader'>
                    <Border Background='{TemplateBinding Background}'
                            BorderBrush='#ECEBF1' BorderThickness='0,0,0,1'
                            Padding='{TemplateBinding Padding}'>
                        <ContentPresenter HorizontalAlignment='{TemplateBinding HorizontalContentAlignment}'
                                          VerticalAlignment='Center'/>
                    </Border>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>
        <Style x:Key='ColHeadCenter' TargetType='GridViewColumnHeader'
               BasedOn='{StaticResource {x:Type GridViewColumnHeader}}'>
            <!-- 残高の見出し用（データは右揃えのまま・見出しのみ中央）。Padding を
                 データ数値の実領域（左18・右16 = Margin 10px + 内蔵 6px）に合わせ、
                 見出しがデータ域の中央に来るようにする -->
            <Setter Property='HorizontalContentAlignment' Value='Center'/>
            <Setter Property='Padding' Value='18,7,16,7'/>
        </Style>

        <!-- ナビ項目 -->
        <Style x:Key='NavItem' TargetType='Border'>
            <Setter Property='Padding' Value='0,7'/>
            <Setter Property='BorderThickness' Value='3,0,0,0'/>
            <Setter Property='BorderBrush' Value='Transparent'/>
        </Style>

        <!-- フィルタチップ（着脱可能なトグル）-->
        <!-- 詳細検索パネルのトグルチップ（選択=選択面+アイリス文字）-->
        <Style x:Key='AdvToggleChip' TargetType='ToggleButton'>
            <Setter Property='FocusVisualStyle' Value='{StaticResource KeyFocusPill}'/>
            <Setter Property='Cursor' Value='Hand'/>
            <Setter Property='FontSize' Value='11'/>
            <Setter Property='Foreground' Value='#6F6E77'/>
            <Setter Property='Margin' Value='0,0,6,4'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='ToggleButton'>
                    <Border x:Name='chipBd' Background='White' BorderBrush='#ECEBF1'
                            BorderThickness='1' CornerRadius='12' Padding='11,4'>
                        <ContentPresenter VerticalAlignment='Center'/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='chipBd' Property='Background' Value='#FAFAFC'/></Trigger>
                        <Trigger Property='IsChecked' Value='True'>
                            <Setter TargetName='chipBd' Property='Background' Value='#F6F5FE'/>
                            <Setter TargetName='chipBd' Property='BorderBrush' Value='#DCDAF6'/>
                            <Setter Property='Foreground' Value='#5B5BD6'/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- 年度セレクタ・検索キー選択（フラットなComboBox・キーの焦点の輪はほかの部品と同じ KeyFocus）-->
        <Style TargetType='ComboBox'>
            <Setter Property='FocusVisualStyle' Value='{StaticResource KeyFocus}'/>
            <Setter Property='FontSize' Value='12'/>
            <Setter Property='Foreground' Value='#1A1523'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='ComboBox'>
                    <Grid x:Name='comboRoot'>
                        <ToggleButton x:Name='toggle' Focusable='False' ClickMode='Press'
                                      IsChecked='{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}'>
                            <ToggleButton.Template>
                                <ControlTemplate TargetType='ToggleButton'>
                                    <Border x:Name='comboBd' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width='*'/>
                                                <ColumnDefinition Width='24'/>
                                            </Grid.ColumnDefinitions>
                                            <Path Grid.Column='1' Data='M 0,0 L 4,4 L 8,0'
                                                  Stroke='#9B9AA3' StrokeThickness='1.4'
                                                  HorizontalAlignment='Center' VerticalAlignment='Center'/>
                                        </Grid>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property='IsMouseOver' Value='True'>
                                            <Setter TargetName='comboBd' Property='BorderBrush' Value='#5B5BD6'/></Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </ToggleButton.Template>
                        </ToggleButton>
                        <ContentPresenter Content='{TemplateBinding SelectionBoxItem}'
                                          Margin='11,6,26,6' VerticalAlignment='Center' IsHitTestVisible='False'/>
                        <Popup IsOpen='{TemplateBinding IsDropDownOpen}' Placement='Bottom'
                               StaysOpen='False' AllowsTransparency='True'>
                            <Border Background='White' BorderBrush='#ECEBF1' BorderThickness='1'
                                    CornerRadius='7' Margin='0,3,8,8'
                                    MinWidth='{TemplateBinding ActualWidth}'>
                                <Border.Effect>
                                    <DropShadowEffect BlurRadius='12' ShadowDepth='2' Opacity='0.12' Color='#1A1523'/>
                                </Border.Effect>
                                <ScrollViewer MaxHeight='240' VerticalScrollBarVisibility='Auto'>
                                    <ItemsPresenter/>
                                </ScrollViewer>
                            </Border>
                        </Popup>
                    </Grid>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsEnabled' Value='False'>
                            <Setter TargetName='comboRoot' Property='Opacity' Value='0.55'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>
        <Style TargetType='ComboBoxItem'>
            <Setter Property='FontSize' Value='12'/>
            <Setter Property='Foreground' Value='#1A1523'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='ComboBoxItem'>
                    <Border x:Name='itemBd' Background='White' Padding='11,6'>
                        <ContentPresenter/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='itemBd' Property='Background' Value='#F6F5FE'/></Trigger>
                        <Trigger Property='IsSelected' Value='True'>
                            <Setter TargetName='itemBd' Property='Background' Value='#F6F5FE'/>
                            <Setter Property='Foreground' Value='#4C4AC0'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- 検出ツールチップ（暗色サーフェス・浮遊要素のためDropShadow可。
             別ビジュアルツリーのため表示倍率は各 ToolTip 要素への直接指定で追従する
             — Style の Setter に置くと Style の Seal 時に共有インスタンスが Freeze され、
             以後の倍率変更が不能になるため Setter は使わない）-->
        <Style x:Key='DetectTip' TargetType='ToolTip'>
            <Setter Property='Background' Value='Transparent'/>
            <Setter Property='BorderThickness' Value='0'/>
            <Setter Property='HasDropShadow' Value='False'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='ToolTip'>
                    <Border Background='#211C2F' CornerRadius='9' Padding='13,10' Margin='0,0,10,10'>
                        <Border.Effect>
                            <DropShadowEffect BlurRadius='16' ShadowDepth='3' Opacity='0.28' Color='#1A1523'/>
                        </Border.Effect>
                        <ContentPresenter/>
                    </Border>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>
        <!-- 詳細パネルのボタン（プライマリ=アイリス塗り / ゴースト=白+罫線）-->
        <Style x:Key='PanelPrimaryBtn' TargetType='Button'>
            <Setter Property='FocusVisualStyle' Value='{StaticResource KeyFocus}'/>
            <Setter Property='FontSize' Value='12'/>
            <Setter Property='Foreground' Value='White'/>
            <Setter Property='Cursor' Value='Hand'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='Button'>
                    <Border x:Name='bd' Background='#5B5BD6' CornerRadius='7' Padding='0,8'>
                        <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='bd' Property='Background' Value='#4C4AC0'/></Trigger>
                        <Trigger Property='IsEnabled' Value='False'>
                            <Setter TargetName='bd' Property='Opacity' Value='0.45'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>
        <Style x:Key='PanelGhostBtn' TargetType='Button'>
            <Setter Property='FocusVisualStyle' Value='{StaticResource KeyFocus}'/>
            <Setter Property='FontSize' Value='12'/>
            <Setter Property='Foreground' Value='#6F6E77'/>
            <Setter Property='Cursor' Value='Hand'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='Button'>
                    <Border x:Name='bd' Background='White' BorderBrush='#ECEBF1' BorderThickness='1'
                            CornerRadius='7' Padding='0,8'>
                        <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='bd' Property='Background' Value='#FAFAFC'/></Trigger>
                        <Trigger Property='IsEnabled' Value='False'>
                            <Setter TargetName='bd' Property='Opacity' Value='0.45'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- 表示倍率ステッパーの −/＋ ボタン（外枠は外側の Border が持つ）-->
        <Style x:Key='ScaleStepBtn' TargetType='Button'>
            <Setter Property='FocusVisualStyle' Value='{StaticResource KeyFocus}'/>
            <Setter Property='Width' Value='34'/>
            <Setter Property='Height' Value='30'/>
            <Setter Property='FontSize' Value='15'/>
            <Setter Property='Foreground' Value='#5B5BD6'/>
            <Setter Property='Cursor' Value='Hand'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='Button'>
                    <Border x:Name='bd' Background='White'>
                        <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='bd' Property='Background' Value='#EEEDFB'/></Trigger>
                        <Trigger Property='IsEnabled' Value='False'>
                            <Setter Property='Foreground' Value='#C9C8D0'/>
                            <Setter TargetName='bd' Property='Background' Value='White'/>
                            <Setter Property='Cursor' Value='Arrow'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- 設定ビューの小型ボタン（参照…・お知らせを再表示。入力欄と同じ高さ）-->
        <Style x:Key='SettingsMiniBtn' TargetType='Button'>
            <Setter Property='FocusVisualStyle' Value='{StaticResource KeyFocus}'/>
            <Setter Property='FontSize' Value='11'/>
            <Setter Property='Foreground' Value='#1A1523'/>
            <Setter Property='Cursor' Value='Hand'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='Button'>
                    <Border x:Name='bd' Background='White' BorderBrush='#ECEBF1' BorderThickness='1'
                            CornerRadius='7' Padding='12,5'>
                        <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='bd' Property='Background' Value='#FAFAFC'/></Trigger>
                        <Trigger Property='IsEnabled' Value='False'>
                            <Setter TargetName='bd' Property='Opacity' Value='0.45'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- メモの入力エリアの主ボタン（追加・保存。SettingsMiniBtn と同寸の主ボタン版）-->
        <Style x:Key='NotesMiniPrimaryBtn' TargetType='Button'>
            <Setter Property='FocusVisualStyle' Value='{StaticResource KeyFocus}'/>
            <Setter Property='FontSize' Value='11'/>
            <Setter Property='Foreground' Value='White'/>
            <Setter Property='Cursor' Value='Hand'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='Button'>
                    <Border x:Name='bd' Background='#5B5BD6' CornerRadius='7' Padding='12,5'>
                        <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='bd' Property='Background' Value='#4C4AC0'/></Trigger>
                        <Trigger Property='IsEnabled' Value='False'>
                            <Setter TargetName='bd' Property='Opacity' Value='0.45'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- 詳細パネルの口座行カード・候補ダイアログの候補行（選択時はコードから
             背景と枠をローカル値で上書きし、解除時にスタイルへ戻す）-->
        <Style x:Key='AccountCard' TargetType='Border'>
            <Setter Property='Background' Value='White'/>
            <Setter Property='BorderBrush' Value='#ECEBF1'/>
            <Setter Property='BorderThickness' Value='1'/>
            <Setter Property='CornerRadius' Value='8'/>
            <Setter Property='Padding' Value='10,8'/>
            <Style.Triggers>
                <Trigger Property='IsMouseOver' Value='True'>
                    <Setter Property='Background' Value='#FAFAFC'/></Trigger>
            </Style.Triggers>
        </Style>

        <!-- カレンダー（日付欄のポップアップ用・意匠一式は他ツールへ移植可能な独立部品）-->
        <!-- 日セル: 選択=アイリス塗り+白文字・今日=アイリス枠+文字（枠は1px固定・色のみ切替）・
             ホバーは条件チップと同色・隣接月は淡色。トリガーは後勝ちのため選択が最優先 -->
        <Style x:Key='AdvCalendarDay' TargetType='CalendarDayButton'>
            <Setter Property='Width' Value='30'/>
            <Setter Property='Height' Value='29'/>
            <Setter Property='FontFamily' Value='Consolas'/>
            <Setter Property='FontSize' Value='11.5'/>
            <Setter Property='Focusable' Value='False'/>
            <Setter Property='Template'>
                <Setter.Value>
                    <ControlTemplate TargetType='CalendarDayButton'>
                        <Border x:Name='dayBd' Background='Transparent' BorderBrush='Transparent'
                                BorderThickness='1' CornerRadius='7'>
                            <ContentPresenter x:Name='dayText' HorizontalAlignment='Center' VerticalAlignment='Center'
                                              TextElement.Foreground='#1A1523'/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property='IsInactive' Value='True'>
                                <Setter TargetName='dayText' Property='TextElement.Foreground' Value='#B9B8C0'/></Trigger>
                            <Trigger Property='IsMouseOver' Value='True'>
                                <Setter TargetName='dayBd' Property='Background' Value='#EEEDFB'/></Trigger>
                            <Trigger Property='IsToday' Value='True'>
                                <Setter TargetName='dayBd' Property='BorderBrush' Value='#5B5BD6'/>
                                <Setter TargetName='dayText' Property='TextElement.Foreground' Value='#5B5BD6'/></Trigger>
                            <Trigger Property='IsSelected' Value='True'>
                                <Setter TargetName='dayBd' Property='Background' Value='#5B5BD6'/>
                                <Setter TargetName='dayBd' Property='BorderBrush' Value='#5B5BD6'/>
                                <Setter TargetName='dayText' Property='TextElement.Foreground' Value='#FFFFFF'/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- 月/年セル: 現在値は選択面+アイリス文字（トグルの ON と同じ文法）。
             年選択（Decade）では数字のみのため Consolas に切り替える -->
        <Style x:Key='AdvCalendarCell' TargetType='CalendarButton'>
            <Setter Property='Width' Value='52'/>
            <Setter Property='Height' Value='38'/>
            <Setter Property='FontSize' Value='11.5'/>
            <Setter Property='Focusable' Value='False'/>
            <Setter Property='Template'>
                <Setter.Value>
                    <ControlTemplate TargetType='CalendarButton'>
                        <Border x:Name='cellBd' Background='Transparent' BorderBrush='Transparent'
                                BorderThickness='1' CornerRadius='7'>
                            <ContentPresenter x:Name='cellText' HorizontalAlignment='Center' VerticalAlignment='Center'
                                              TextElement.Foreground='#1A1523'/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <DataTrigger Binding='{Binding DisplayMode, RelativeSource={RelativeSource FindAncestor, AncestorType={x:Type Calendar}}}' Value='Decade'>
                                <Setter TargetName='cellText' Property='TextElement.FontFamily' Value='Consolas'/></DataTrigger>
                            <Trigger Property='IsInactive' Value='True'>
                                <Setter TargetName='cellText' Property='TextElement.Foreground' Value='#B9B8C0'/></Trigger>
                            <Trigger Property='IsMouseOver' Value='True'>
                                <Setter TargetName='cellBd' Property='Background' Value='#EEEDFB'/></Trigger>
                            <Trigger Property='HasSelectedDays' Value='True'>
                                <Setter TargetName='cellBd' Property='Background' Value='#F6F5FE'/>
                                <Setter TargetName='cellBd' Property='BorderBrush' Value='#DCDAF6'/>
                                <Setter TargetName='cellText' Property='TextElement.Foreground' Value='#5B5BD6'/>
                                <Setter TargetName='cellText' Property='TextElement.FontWeight' Value='Bold'/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- カレンダー本体: 見出し（月/年/年代の切替ボタン）+ ‹ › 送り + 日/月年ビュー。
             PART_* の名前と MonthView 7×7・YearView 4×3 の行列定義は Calendar の
             内部処理が前提とする構造のため変更しない -->
        <Style x:Key='AdvCalendarItem' TargetType='CalendarItem'>
            <Setter Property='Template'>
                <Setter.Value>
                    <ControlTemplate TargetType='CalendarItem'>
                        <ControlTemplate.Resources>
                            <DataTemplate x:Key='DayTitleTemplate'>
                                <TextBlock Text='{Binding}' FontSize='10' Foreground='#9B9AA3'
                                           HorizontalAlignment='Center' VerticalAlignment='Center' Margin='0,3,0,4'/>
                            </DataTemplate>
                        </ControlTemplate.Resources>
                        <Grid x:Name='PART_Root'>
                            <Grid.RowDefinitions>
                                <RowDefinition Height='Auto'/>
                                <RowDefinition Height='*'/>
                            </Grid.RowDefinitions>
                            <Grid Grid.Row='0' Margin='0,0,0,3'>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width='Auto'/>
                                    <ColumnDefinition Width='*'/>
                                    <ColumnDefinition Width='Auto'/>
                                </Grid.ColumnDefinitions>
                                <Button x:Name='PART_PreviousButton' Grid.Column='0' Width='24' Height='24' Focusable='False' Cursor='Hand'>
                                    <Button.Template>
                                        <ControlTemplate TargetType='Button'>
                                            <Border x:Name='navBd' Background='Transparent' CornerRadius='7'>
                                                <TextBlock Text='‹' FontSize='13' Foreground='#6F6E77'
                                                           HorizontalAlignment='Center' VerticalAlignment='Center' Margin='0,0,0,2'/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property='IsMouseOver' Value='True'>
                                                    <Setter TargetName='navBd' Property='Background' Value='#F4F4F6'/></Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Button.Template>
                                </Button>
                                <Button x:Name='PART_HeaderButton' Grid.Column='1' Height='24' Focusable='False' Cursor='Hand' Margin='4,0'>
                                    <Button.Template>
                                        <ControlTemplate TargetType='Button'>
                                            <Border x:Name='headBd' Background='Transparent' CornerRadius='7'>
                                                <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'
                                                                  TextElement.FontSize='12' TextElement.FontWeight='Bold'
                                                                  TextElement.Foreground='#1A1523'/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property='IsMouseOver' Value='True'>
                                                    <Setter TargetName='headBd' Property='Background' Value='#F4F4F6'/></Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Button.Template>
                                </Button>
                                <Button x:Name='PART_NextButton' Grid.Column='2' Width='24' Height='24' Focusable='False' Cursor='Hand'>
                                    <Button.Template>
                                        <ControlTemplate TargetType='Button'>
                                            <Border x:Name='navBd' Background='Transparent' CornerRadius='7'>
                                                <TextBlock Text='›' FontSize='13' Foreground='#6F6E77'
                                                           HorizontalAlignment='Center' VerticalAlignment='Center' Margin='0,0,0,2'/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property='IsMouseOver' Value='True'>
                                                    <Setter TargetName='navBd' Property='Background' Value='#F4F4F6'/></Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Button.Template>
                                </Button>
                            </Grid>
                            <Grid x:Name='PART_MonthView' Grid.Row='1' Visibility='Visible' HorizontalAlignment='Center'>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition/><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/>
                                    <ColumnDefinition/><ColumnDefinition/><ColumnDefinition/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/>
                                    <RowDefinition/><RowDefinition/><RowDefinition/>
                                </Grid.RowDefinitions>
                            </Grid>
                            <Grid x:Name='PART_YearView' Grid.Row='1' Visibility='Hidden' HorizontalAlignment='Center'>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition/><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition/><RowDefinition/><RowDefinition/>
                                </Grid.RowDefinitions>
                            </Grid>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <DataTrigger Binding='{Binding DisplayMode, RelativeSource={RelativeSource FindAncestor, AncestorType={x:Type Calendar}}}' Value='Year'>
                                <Setter TargetName='PART_MonthView' Property='Visibility' Value='Hidden'/>
                                <Setter TargetName='PART_YearView' Property='Visibility' Value='Visible'/></DataTrigger>
                            <DataTrigger Binding='{Binding DisplayMode, RelativeSource={RelativeSource FindAncestor, AncestorType={x:Type Calendar}}}' Value='Decade'>
                                <Setter TargetName='PART_MonthView' Property='Visibility' Value='Hidden'/>
                                <Setter TargetName='PART_YearView' Property='Visibility' Value='Visible'/></DataTrigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key='AdvCalendar' TargetType='Calendar'>
            <Setter Property='Focusable' Value='False'/>
            <Setter Property='Background' Value='White'/>
            <Setter Property='BorderThickness' Value='0'/>
            <Setter Property='Margin' Value='0'/>
            <Setter Property='SelectionMode' Value='SingleDate'/>
            <Setter Property='CalendarItemStyle' Value='{StaticResource AdvCalendarItem}'/>
            <Setter Property='CalendarDayButtonStyle' Value='{StaticResource AdvCalendarDay}'/>
            <Setter Property='CalendarButtonStyle' Value='{StaticResource AdvCalendarCell}'/>
        </Style>
    </Window.Resources>

    <Grid LayoutTransform='{StaticResource DisplayScale}'>
        <Grid.ColumnDefinitions>
            <!-- サイドバーはアイコンレール（48px）。展開はオーバーレイ式で列幅は変えない -->
            <ColumnDefinition Width='48'/>
            <ColumnDefinition Width='*'/>
            <!-- 詳細パネル用の列（パネル自身の幅 0⇔440px のアニメーションに追従する）-->
            <ColumnDefinition Width='Auto'/>
        </Grid.ColumnDefinitions>

        <!-- ============ サイドバー（アイコンレール）============ -->
        <!-- 既定は48pxのレール。カーソルが0.2秒とどまると SidebarFlyout（最下部に定義）が
             同位置に被さって展開する。アイコンは展開を待たずに直接クリックできる -->
        <Border x:Name='SidebarRail' Grid.Column='0' Background='#FCFCFD'
                BorderBrush='#ECEBF1' BorderThickness='0,0,1,0'>
            <DockPanel Margin='0,16,0,14'>

                <!-- ワードマークの円環マーク（レールでは記号のみ）-->
                <Grid DockPanel.Dock='Top' Width='26' Height='26' Margin='0,0,0,16'
                      HorizontalAlignment='Center'>
                    <Ellipse Stroke='#5B5BD6' StrokeThickness='1.8'
                             StrokeDashArray='1.9,1.7' StrokeDashCap='Round' Margin='1'/>
                    <Ellipse Fill='#5B5BD6' Width='7' Height='7'/>
                </Grid>

                <!-- 種別（預金／生保）と年度の短縮表示（下部固定。折りたたみ中もどの台帳かを常時示し、
                     クリックで展開して年度コンボへの導線とする。キーで押せるのは年度の短縮表示）-->
                <StackPanel DockPanel.Dock='Bottom' Margin='0,12,0,0'>
                    <TextBlock x:Name='RailKindText' FontSize='9' Foreground='#9B9AA3' Cursor='Hand'
                               HorizontalAlignment='Center' Margin='0,0,0,1'/>
                    <TextBlock x:Name='RailYearText' FontSize='11'
                               FontFamily='Consolas' Foreground='#6F6E77' Cursor='Hand'
                               HorizontalAlignment='Center'
                               ToolTip='表示中の年度（クリックで切替パネルを開く・Ctrl+Y）'/>
                </StackPanel>
                <Border DockPanel.Dock='Bottom' Height='1' Background='#F1F0F4' Margin='10,0'/>

                <!-- ナビ（アイコンのみ。活性表示は SetNavSelection がレール/展開パネルの両側を同期する。
                     アイコンは字形ごとの見かけの大きさを補正する（⚙のみ密な字形のため14px）。
                     差押登録は depositSeizureListScript 未設定時に非表示になる）-->
                <StackPanel>
                    <Border x:Name='DepositRailLedger' Style='{StaticResource NavItem}' Cursor='Hand'
                            Background='Transparent' ToolTip='預金 (Ctrl+1)'>
                        <TextBlock x:Name='DepositRailLedgerIcon' Text='≡' FontSize='15' HorizontalAlignment='Center'/>
                    </Border>
                    <Border x:Name='InsuranceRailLedger' Style='{StaticResource NavItem}' Cursor='Hand'
                            Background='Transparent' ToolTip='生命保険 (Ctrl+2)'>
                        <TextBlock x:Name='InsuranceRailLedgerIcon' Text='☂' FontSize='15' HorizontalAlignment='Center'/>
                    </Border>
                    <Border x:Name='RailRecent' Style='{StaticResource NavItem}' Cursor='Hand'
                            Background='Transparent' ToolTip='最近見た人 (Ctrl+3)'>
                        <TextBlock x:Name='RailRecentIcon' Text='◷' FontSize='15' HorizontalAlignment='Center'/>
                    </Border>
                    <Border x:Name='RailSeizure' Style='{StaticResource NavItem}' Cursor='Hand'
                            Background='Transparent' ToolTip='差押登録 (Ctrl+4)'>
                        <TextBlock x:Name='RailSeizureIcon' Text='⚑' FontSize='15' HorizontalAlignment='Center'/>
                    </Border>
                    <Border x:Name='RailSettings' Style='{StaticResource NavItem}' Cursor='Hand'
                            Background='Transparent' ToolTip='設定 (Ctrl+,)'>
                        <TextBlock x:Name='RailSettingsIcon' Text='⚙' FontSize='14' HorizontalAlignment='Center'/>
                    </Border>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- ============ 中央ペイン（預金 / 生命保険 / 最近見た人 / 差押登録 / 設定 をナビで切替）============ -->
        <Grid Grid.Column='1'>

        <!-- ============ 預金の一覧ビュー ============ -->
        <DockPanel x:Name='DepositLedgerView'>

            <!-- 検索バー+詳細検索パネル+条件チップ行 -->
            <StackPanel DockPanel.Dock='Top' Margin='14,12,14,8'>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width='128'/>
                        <ColumnDefinition Width='*'/>
                        <ColumnDefinition Width='Auto'/>
                    </Grid.ColumnDefinitions>
                    <ComboBox x:Name='DepositSearchKeyCombo' Grid.Column='0' Margin='0,0,8,0' SelectedIndex='0'>
                        <ComboBoxItem Content='宛名番号'/>
                        <ComboBoxItem Content='氏名'/>
                        <ComboBoxItem Content='カナ氏名'/>
                        <ComboBoxItem Content='金融機関名'/>
                    </ComboBox>
                    <Border Grid.Column='1' Background='White' BorderBrush='#ECEBF1'
                            BorderThickness='1' CornerRadius='7'>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width='Auto'/>
                                <ColumnDefinition Width='*'/>
                                <ColumnDefinition Width='Auto'/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column='0' Text='🔍' FontSize='12' Foreground='#9B9AA3'
                                       Margin='11,0,0,0' VerticalAlignment='Center'/>
                            <TextBox x:Name='DepositSearchBox' Grid.Column='1' BorderThickness='0'
                                     Background='Transparent' FontSize='12' Foreground='#1A1523'
                                     Padding='7,6' VerticalContentAlignment='Center'/>
                            <TextBlock x:Name='DepositSearchPlaceholder' Grid.Column='1'
                                       Text='選択中のキーで部分一致検索' FontSize='12' Foreground='#9B9AA3'
                                       Margin='11,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                            <Button x:Name='DepositClearButton' Grid.Column='2' Content='✕' FontSize='10'
                                    Width='20' Height='20' Margin='0,0,8,0' Visibility='Collapsed'
                                    Cursor='Hand' Foreground='#6F6E77' FocusVisualStyle='{StaticResource KeyFocusPill}'>
                                <Button.Template>
                                    <ControlTemplate TargetType='Button'>
                                        <Border x:Name='clearBd' Background='#F4F4F6' CornerRadius='10'>
                                            <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property='IsMouseOver' Value='True'>
                                                <Setter TargetName='clearBd' Property='Background' Value='#ECEBF1'/></Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Button.Template>
                            </Button>
                        </Grid>
                    </Border>

                    <!-- 「詳細検索」トグル（常時 選択面+アイリス文字・条件の有無は件数バッジで示す。
                         キャレットは開閉表示）-->
                    <Border x:Name='DepositAdvToggle' Grid.Column='2' Margin='8,0,0,0' Background='#F6F5FE'
                            BorderBrush='#DCDAF6' BorderThickness='1' CornerRadius='7'
                            Padding='12,0' Cursor='Hand'>
                        <StackPanel Orientation='Horizontal' VerticalAlignment='Center'>
                            <TextBlock Text='詳細検索' FontSize='12' Foreground='#5B5BD6'/>
                            <Border x:Name='DepositAdvCountBadge' Background='#5B5BD6' CornerRadius='8'
                                    Padding='6,1,6,2' Margin='6,0,0,0' Visibility='Collapsed'>
                                <TextBlock x:Name='DepositAdvCountText' FontSize='10' Foreground='White'/>
                            </Border>
                            <TextBlock x:Name='DepositAdvToggleCaret' Text='▼' FontSize='9' Foreground='#9B9AA3'
                                       Margin='6,1,0,0' VerticalAlignment='Center'/>
                        </StackPanel>
                    </Border>
                </Grid>

                <!-- 詳細検索パネル（淡アイリス面=特殊エリアの明示。
                     開閉は高さアニメーション・条件は閉じても持続する）-->
                <Border x:Name='DepositAdvPanel' Visibility='Collapsed' Background='#F9F8FD' ClipToBounds='True'
                        BorderBrush='#DCDAF6' BorderThickness='1' CornerRadius='10'
                        Margin='0,8,0,0' Padding='16,12,18,8'>
                    <StackPanel>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width='*'/>
                                <ColumnDefinition Width='36'/>
                                <ColumnDefinition Width='*'/>
                            </Grid.ColumnDefinitions>

                            <!-- 左列: 文字・選択の条件（右列と行位置を揃えない独立の縦積み）-->
                            <StackPanel Grid.Column='0'>
                                <TextBlock FontSize='11' Margin='0,0,0,5'><Run Text='項目検索' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  それぞれ部分一致・すべて該当（AND）' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width='*'/>
                                        <ColumnDefinition Width='10'/>
                                        <ColumnDefinition Width='*'/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition/>
                                        <RowDefinition Height='6'/>
                                        <RowDefinition/>
                                    </Grid.RowDefinitions>
                                    <Border Grid.Row='0' Grid.Column='0' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='DepositAdvAddressBox' BorderThickness='0' Background='Transparent'
                                                     FontSize='12' Foreground='#1A1523' Padding='7,4' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='DepositAdvAddressPh' Text='宛名番号' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <Border Grid.Row='0' Grid.Column='2' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='DepositAdvNameBox' BorderThickness='0' Background='Transparent'
                                                     FontSize='12' Foreground='#1A1523' Padding='7,4' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='DepositAdvNamePh' Text='氏名' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <Border Grid.Row='2' Grid.Column='0' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='DepositAdvKanaBox' BorderThickness='0' Background='Transparent'
                                                     FontSize='12' Foreground='#1A1523' Padding='7,4' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='DepositAdvKanaPh' Text='カナ氏名' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <Border Grid.Row='2' Grid.Column='2' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='DepositAdvInstitutionBox' BorderThickness='0' Background='Transparent'
                                                     FontSize='12' Foreground='#1A1523' Padding='7,4' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='DepositAdvInstitutionPh' Text='金融機関名' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                </Grid>

                                <!-- 検出と行の分類は横並び。左列の幅が不足する場合は
                                     WrapPanel の折り返しで縦積みになる -->
                                <WrapPanel Margin='0,13,0,0'>
                                    <StackPanel Margin='0,0,28,6'>
                                        <TextBlock FontSize='11' Margin='0,0,0,5'><Run Text='検出' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  選択したものを すべて含む 行（AND）' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                        <WrapPanel>
                                            <ToggleButton x:Name='DepositAdvDetect0' Style='{StaticResource AdvToggleChip}' Content='反対債権'/>
                                            <ToggleButton x:Name='DepositAdvDetect1' Style='{StaticResource AdvToggleChip}' Content='年金'/>
                                            <ToggleButton x:Name='DepositAdvDetect2' Style='{StaticResource AdvToggleChip}' Content='給与'/>
                                            <ToggleButton x:Name='DepositAdvDetect3' Style='{StaticResource AdvToggleChip}' Content='保険'/>
                                            <ToggleButton x:Name='DepositAdvDetect4' Style='{StaticResource AdvToggleChip}' Content='定期入金'/>
                                            <ToggleButton x:Name='DepositAdvDetect5' Style='{StaticResource AdvToggleChip}' Content='定期出金'/>
                                        </WrapPanel>
                                    </StackPanel>
                                    <StackPanel Margin='0,0,0,6'>
                                        <TextBlock FontSize='11' Margin='0,0,0,5'><Run Text='行の分類' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  機械的に自動判定される分類・いずれかに該当（OR）' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                        <WrapPanel>
                                            <ToggleButton x:Name='DepositAdvClassRep' Style='{StaticResource AdvToggleChip}' Content='代表口座'
                                                          ToolTip='同じ人・同じ金融機関の口座の中から1口だけ選ばれた代表の行'/>
                                            <ToggleButton x:Name='DepositAdvClassSpecial' Style='{StaticResource AdvToggleChip}' Content='特殊口座のみ'
                                                          ToolTip='解約済・休眠などの特殊な口座しかなかったことを記録する代表の行'/>
                                            <ToggleButton x:Name='DepositAdvClassNoAccount' Style='{StaticResource AdvToggleChip}' Content='口座なし'
                                                          ToolTip='口座なし回答の行を表示する（既定では非表示）'/>
                                        </WrapPanel>
                                    </StackPanel>
                                </WrapPanel>

                                <!-- 担当。トグルは選択中年度の台帳から動的生成する
                                    （RebuildDepositStaffToggles）。記録がない年度は案内文言に切り替える -->
                                <TextBlock FontSize='11' Margin='0,7,0,5'><Run Text='担当' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  この年度の台帳から列挙・いずれかに該当（OR）' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                <WrapPanel x:Name='DepositAdvStaffPanel'/>
                                <TextBlock x:Name='DepositAdvStaffEmptyText' Text='この年度の台帳に担当の記録はありません'
                                           FontSize='11' Foreground='#9B9AA3'/>
                            </StackPanel>

                            <!-- 右列: 数値・日付の条件 -->
                            <StackPanel Grid.Column='2'>
                                <TextBlock FontSize='11' Margin='0,0,0,5'><Run Text='残高' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  「対象外」（回答コード）の行と口座なし行は数値条件の対象外' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                <StackPanel Orientation='Horizontal'>
                                    <Border x:Name='DepositAdvBalanceMinBd' Width='118' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='DepositAdvBalanceMinBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     TextAlignment='Right' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='DepositAdvBalanceMinPh' Text='下限なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='0,0,10,0' HorizontalAlignment='Right' VerticalAlignment='Center'
                                                       IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <TextBlock Text='〜' FontSize='12' Foreground='#9B9AA3' Margin='6,0' VerticalAlignment='Center'/>
                                    <Border x:Name='DepositAdvBalanceMaxBd' Width='118' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='DepositAdvBalanceMaxBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     TextAlignment='Right' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='DepositAdvBalanceMaxPh' Text='上限なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='0,0,10,0' HorizontalAlignment='Right' VerticalAlignment='Center'
                                                       IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <TextBlock Text='円' FontSize='11.5' Foreground='#6F6E77' Margin='6,0,0,0' VerticalAlignment='Center'/>
                                </StackPanel>

                                <TextBlock FontSize='11' Margin='0,13,0,5'><Run Text='日付の範囲' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  yyyy/mm/dd・片側だけでも可' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                <StackPanel Orientation='Horizontal' Margin='0,0,0,6'>
                                    <TextBlock Text='回答基準日' Width='76' FontSize='11.5' Foreground='#6F6E77' VerticalAlignment='Center'/>
                                    <Border x:Name='DepositAdvBaseFromBd' Width='104' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='DepositAdvBaseFromBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='DepositAdvBaseFromPh' Text='指定なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <TextBlock Text='〜' FontSize='12' Foreground='#9B9AA3' Margin='6,0' VerticalAlignment='Center'/>
                                    <Border x:Name='DepositAdvBaseToBd' Width='104' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='DepositAdvBaseToBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='DepositAdvBaseToPh' Text='指定なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                </StackPanel>
                                <StackPanel Orientation='Horizontal'>
                                    <TextBlock Text='最終取引日' Width='76' FontSize='11.5' Foreground='#6F6E77' VerticalAlignment='Center'/>
                                    <Border x:Name='DepositAdvTxFromBd' Width='104' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='DepositAdvTxFromBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='DepositAdvTxFromPh' Text='指定なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <TextBlock Text='〜' FontSize='12' Foreground='#9B9AA3' Margin='6,0' VerticalAlignment='Center'/>
                                    <Border x:Name='DepositAdvTxToBd' Width='104' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='DepositAdvTxToBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='DepositAdvTxToPh' Text='指定なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                </StackPanel>
                            </StackPanel>
                        </Grid>

                        <!-- パネル下端: 案内と全クリア -->
                        <Border BorderBrush='#F1F0F4' BorderThickness='0,1,0,0' Margin='0,10,0,0' Padding='0,8,0,2'>
                            <DockPanel>
                                <TextBlock x:Name='DepositAdvClearLink' DockPanel.Dock='Right' Text='条件をすべてクリア'
                                           FontSize='11.5' Foreground='#6F6E77' Cursor='Hand' VerticalAlignment='Center'/>
                                <TextBlock Text='条件は入力と同時に一覧へ反映されます。設定した条件は下のチップから個別に外せます。'
                                           FontSize='10.5' Foreground='#9B9AA3' VerticalAlignment='Center'/>
                            </DockPanel>
                        </Border>
                    </StackPanel>
                </Border>

                <!-- 条件チップ行（中身はコードから動的生成する。条件がないときは行ごと非表示）-->
                <WrapPanel x:Name='DepositFilterChipsRow' Margin='0,8,0,0' Visibility='Collapsed'/>
            </StackPanel>

            <!-- フッター -->
            <Border DockPanel.Dock='Bottom' Background='#FCFCFD' BorderBrush='#ECEBF1'
                    BorderThickness='0,1,0,0' Padding='14,5'>
                <DockPanel>
                    <TextBlock x:Name='DepositStatusRight' DockPanel.Dock='Right' FontSize='11' Foreground='#9B9AA3'/>
                    <TextBlock x:Name='DepositStatusLeft' FontSize='11' Foreground='#9B9AA3'/>
                </DockPanel>
            </Border>

            <!-- 一覧テーブル -->
            <Grid>
                <Border Background='White' BorderBrush='#ECEBF1' BorderThickness='0,1,0,0'>
                    <ListView x:Name='DepositLedgerList' BorderThickness='0' Background='White' FontSize='12'
                              SelectionMode='Single'
                              ItemContainerStyle='{StaticResource LedgerItem}'
                              VirtualizingStackPanel.IsVirtualizing='True'
                              VirtualizingStackPanel.VirtualizationMode='Recycling'
                              ScrollViewer.HorizontalScrollBarVisibility='Auto'>
                        <ListView.View>
                            <GridView AllowsColumnReorder='False'>
                                <GridViewColumn Header='回答基準日' Width='104'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding DisplayBaseDate}' FontFamily='Consolas' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='宛名番号' Width='104'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding AddressNumber}' FontFamily='Consolas' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='氏名' Width='150'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <!-- メモのある人物は氏名の直後に印（✎）を置く。氏名の幅を印の分だけ
                                             詰めて（列幅150−セル内余白16−印19）、長い氏名でも印が切れないようにする -->
                                        <StackPanel Orientation='Horizontal' Margin='12,0,4,0'>
                                            <TextBlock Text='{Binding Name}' TextTrimming='CharacterEllipsis'
                                                       ToolTip='{Binding Name}'>
                                                <TextBlock.Style>
                                                    <Style TargetType='TextBlock'>
                                                        <Setter Property='MaxWidth' Value='122'/>
                                                        <Style.Triggers>
                                                            <DataTrigger Binding='{Binding HasNotes}' Value='True'>
                                                                <Setter Property='MaxWidth' Value='103'/>
                                                            </DataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </TextBlock.Style>
                                            </TextBlock>
                                            <TextBlock Text='✎' FontSize='10' Foreground='#9B9AA3' Margin='5,0,0,0'
                                                       VerticalAlignment='Center' ToolTip='{Binding NotesToolTip}'
                                                       Visibility='{Binding HasNotes, Converter={StaticResource BoolVis}}'/>
                                        </StackPanel>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='カナ氏名' Width='140'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding Kana}' TextTrimming='CharacterEllipsis'
                                                   ToolTip='{Binding Kana}' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='金融機関名' Width='150'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding InstitutionName}' TextTrimming='CharacterEllipsis'
                                                   ToolTip='{Binding InstitutionName}' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='支店名' Width='140'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding BranchName}' TextTrimming='CharacterEllipsis'
                                                   ToolTip='{Binding BranchName}' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='種別' Width='92'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding AccountType}' TextTrimming='CharacterEllipsis'
                                                   ToolTip='{Binding AccountType}' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='口座番号' Width='100'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding AccountNumber}' FontFamily='Consolas' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='残高' Width='104'
                                                HeaderContainerStyle='{StaticResource ColHeadCenter}'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding DisplayBalance}' FontFamily='Consolas'
                                                   HorizontalAlignment='Right' Margin='4,0,10,0'
                                                   ToolTip='{Binding BalanceToolTip}'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='最終取引日' Width='104'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding DisplayLastTransaction}' FontFamily='Consolas' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn x:Name='DepositDetectColumn' Header='検出' Width='330'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <Grid Background='Transparent' Margin='12,0,4,0'
                                              ToolTipService.InitialShowDelay='350'
                                              ToolTipService.ShowDuration='60000'
                                              ToolTipService.IsEnabled='{Binding HasDetection}'>
                                            <Grid.ToolTip>
                                                <ToolTip Style='{StaticResource DetectTip}' LayoutTransform='{StaticResource DisplayScale}'>
                                                    <ItemsControl ItemsSource='{Binding TooltipLines}' MaxWidth='460'>
                                                        <ItemsControl.ItemTemplate><DataTemplate>
                                                            <Grid Margin='0,3'>
                                                                <Grid.ColumnDefinitions>
                                                                    <ColumnDefinition Width='Auto'/>
                                                                    <ColumnDefinition Width='*'/>
                                                                </Grid.ColumnDefinitions>
                                                                <Border Grid.Column='0' Background='{Binding ChipBg}'
                                                                        CornerRadius='9' Padding='9,1,9,2' VerticalAlignment='Top'>
                                                                    <TextBlock Text='{Binding ChipText}' FontSize='11'
                                                                               Foreground='{Binding ChipFg}'/>
                                                                </Border>
                                                                <TextBlock Grid.Column='1' Text='{Binding Content}'
                                                                           FontSize='12' Foreground='#F4F3F8'
                                                                           FontFamily='{Binding ContentFont}'
                                                                           TextWrapping='Wrap' Margin='9,1,0,0'/>
                                                            </Grid>
                                                        </DataTemplate></ItemsControl.ItemTemplate>
                                                    </ItemsControl>
                                                </ToolTip>
                                            </Grid.ToolTip>
                                            <!-- チップ列（口座なし行は代わりにピルを表示）-->
                                            <ItemsControl ItemsSource='{Binding Chips}' IsHitTestVisible='False'>
                                                <ItemsControl.ItemsPanel><ItemsPanelTemplate>
                                                    <StackPanel Orientation='Horizontal'/>
                                                </ItemsPanelTemplate></ItemsControl.ItemsPanel>
                                                <ItemsControl.ItemTemplate><DataTemplate>
                                                    <Border Background='{Binding Bg}' CornerRadius='9'
                                                            Padding='8,1,8,2' Margin='0,0,4,0'>
                                                        <TextBlock Text='{Binding Text}' FontSize='11'
                                                                   Foreground='{Binding Fg}'/>
                                                    </Border>
                                                </DataTemplate></ItemsControl.ItemTemplate>
                                            </ItemsControl>
                                            <Border Background='#F4F4F6' CornerRadius='9' Padding='8,1,8,2'
                                                    HorizontalAlignment='Left' IsHitTestVisible='False'
                                                    Visibility='{Binding IsNoAccount, Converter={StaticResource BoolVis}}'>
                                                <TextBlock Text='口座なし' FontSize='11' Foreground='#9B9AA3'/>
                                            </Border>
                                        </Grid>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Border>

                <!-- 空状態メッセージ -->
                <TextBlock x:Name='DepositEmptyStateText' Visibility='Collapsed'
                           HorizontalAlignment='Center' VerticalAlignment='Center'
                           TextAlignment='Center' FontSize='12' Foreground='#9B9AA3'
                           LineHeight='20'/>

                <!-- 読み込みオーバーレイ（テーブル範囲のみを覆う。スピナーのみの意匠）-->
                <Grid x:Name='DepositLoadingOverlay' Background='#B3FFFFFF' Visibility='Collapsed'>
                    <Path x:Name='DepositSpinnerPath' Data='M 20,2 A 18,18 0 1 1 2,20'
                          Stroke='#5B5BD6' StrokeThickness='4'
                          StrokeStartLineCap='Round' StrokeEndLineCap='Round'
                          Width='40' Height='40' Stretch='None'
                          HorizontalAlignment='Center' VerticalAlignment='Center'
                          RenderTransformOrigin='0.5,0.5'>
                        <Path.RenderTransform><RotateTransform/></Path.RenderTransform>
                    </Path>
                </Grid>
            </Grid>
        </DockPanel>
" + INSURANCE_VIEW_XAML + @"
        <!-- ============ 最近見た人ビュー（ナビ「最近見た人」で表示）============ -->
        <!-- 照会ファイルを開いた行の履歴を年度セクション（開閉式・中身はコードから動的生成）で表示する -->
        <DockPanel x:Name='RecentView' Visibility='Collapsed'>

            <!-- 見出し -->
            <StackPanel DockPanel.Dock='Top' Margin='14,12,14,10'>
                <TextBlock Text='最近見た人' FontSize='13.5' FontWeight='Medium' Foreground='#1A1523'/>
                <TextBlock x:Name='RecentSubtitleText' FontSize='10.5' Foreground='#9B9AA3'
                           Margin='0,3,0,0' TextWrapping='Wrap'/>
            </StackPanel>

            <!-- フッター（件数・履歴ファイルのパス）-->
            <Border DockPanel.Dock='Bottom' Background='#FCFCFD' BorderBrush='#ECEBF1'
                    BorderThickness='0,1,0,0' Padding='14,5'>
                <TextBlock x:Name='RecentFooterText' FontSize='11' Foreground='#9B9AA3'/>
            </Border>

            <!-- 本文（年度セクションを収める白カード / 0件時の案内）-->
            <Grid>
                <ScrollViewer VerticalScrollBarVisibility='Auto' HorizontalScrollBarVisibility='Auto'>
                    <Border x:Name='RecentCard' Background='White' BorderBrush='#ECEBF1'
                            BorderThickness='1' CornerRadius='10' Margin='14,2,14,14'
                            VerticalAlignment='Top'>
                        <StackPanel x:Name='RecentSectionsPanel'/>
                    </Border>
                </ScrollViewer>
                <TextBlock x:Name='RecentEmptyText' Visibility='Collapsed'
                           HorizontalAlignment='Center' VerticalAlignment='Center'
                           TextAlignment='Center' FontSize='12' Foreground='#9B9AA3'
                           LineHeight='20'/>
            </Grid>
        </DockPanel>

        <!-- ============ 差押登録ビュー（ナビ「差押登録」で表示）============ -->
        <!-- 差押登録ログを年度セクション（開閉式・中身はコードから動的生成）で全件表示する -->
        <DockPanel x:Name='SeizureView' Visibility='Collapsed'>

            <!-- 見出し -->
            <StackPanel DockPanel.Dock='Top' Margin='14,12,14,10'>
                <TextBlock Text='差押登録' FontSize='13.5' FontWeight='Medium' Foreground='#1A1523'/>
                <TextBlock FontSize='10.5' Foreground='#9B9AA3' Margin='0,3,0,0' TextWrapping='Wrap'
                           Text='PADDOCK から差押リストに連携した履歴です（このユーザーの分のみ）。'/>
            </StackPanel>

            <!-- フッター（件数・ログファイルのパス）-->
            <Border DockPanel.Dock='Bottom' Background='#FCFCFD' BorderBrush='#ECEBF1'
                    BorderThickness='0,1,0,0' Padding='14,5'>
                <TextBlock x:Name='SeizureFooterText' FontSize='11' Foreground='#9B9AA3'/>
            </Border>

            <!-- 本文（年度セクションを収める白カード / 0件時の案内）-->
            <Grid>
                <ScrollViewer VerticalScrollBarVisibility='Auto' HorizontalScrollBarVisibility='Auto'>
                    <Border x:Name='SeizureCard' Background='White' BorderBrush='#ECEBF1'
                            BorderThickness='1' CornerRadius='10' Margin='14,2,14,14'
                            VerticalAlignment='Top'>
                        <StackPanel x:Name='SeizureSectionsPanel'/>
                    </Border>
                </ScrollViewer>
                <TextBlock x:Name='SeizureEmptyText' Visibility='Collapsed'
                           HorizontalAlignment='Center' VerticalAlignment='Center'
                           TextAlignment='Center' FontSize='12' Foreground='#9B9AA3'
                           LineHeight='20'/>
            </Grid>
        </DockPanel>

        <!-- ============ 設定ビュー（ナビ「設定」で表示）============ -->
        <!-- 本文（3枚のカード）はコードから動的生成する（BuildSettingsView）。
             保存バーは未保存の変更がある間のみ表示され、フッターの直上に出る -->
        <DockPanel x:Name='SettingsView' Visibility='Collapsed'>

            <!-- 見出し -->
            <StackPanel DockPanel.Dock='Top' Margin='14,12,14,10'>
                <TextBlock Text='設定' FontSize='13.5' FontWeight='Medium' Foreground='#1A1523'/>
                <TextBlock FontSize='10.5' Foreground='#9B9AA3' Margin='0,3,0,0' TextWrapping='Wrap'
                           Text='PADDOCK の設定です。保存すると paddock_config.json に書き込まれます。'/>
            </StackPanel>

            <!-- フッター（設定ファイルのパス・バージョン）-->
            <Border DockPanel.Dock='Bottom' Background='#FCFCFD' BorderBrush='#ECEBF1'
                    BorderThickness='0,1,0,0' Padding='14,5'>
                <DockPanel>
                    <TextBlock x:Name='SettingsFooterRight' DockPanel.Dock='Right' FontSize='11' Foreground='#9B9AA3'/>
                    <TextBlock x:Name='SettingsFooterLeft' FontSize='11' Foreground='#9B9AA3'/>
                </DockPanel>
            </Border>

            <!-- 保存バー（変更あり時のみ表示）-->
            <Border x:Name='SettingsSaveBar' DockPanel.Dock='Bottom' Background='#F6F5FE'
                    BorderBrush='#DCDAF6' BorderThickness='0,1,0,0' Padding='14,8'
                    Visibility='Collapsed' ClipToBounds='True'>
                <DockPanel>
                    <Button x:Name='SettingsSaveButton' DockPanel.Dock='Right' Content='保存'
                            Style='{StaticResource PanelPrimaryBtn}' Width='96'/>
                    <Button x:Name='SettingsDiscardButton' DockPanel.Dock='Right' Content='変更を破棄'
                            Style='{StaticResource PanelGhostBtn}' Width='110' Margin='0,0,8,0'/>
                    <StackPanel VerticalAlignment='Center'>
                        <TextBlock Text='保存されていない変更があります' FontSize='11.5' Foreground='#1A1523'/>
                        <TextBlock x:Name='SettingsChangeSummary' FontSize='10.5' Foreground='#9B9AA3'
                                   Margin='0,1,0,0' TextTrimming='CharacterEllipsis'/>
                    </StackPanel>
                </DockPanel>
            </Border>

            <!-- 本文（3枚のカード）-->
            <ScrollViewer VerticalScrollBarVisibility='Auto'>
                <StackPanel x:Name='SettingsPanel' Margin='14,2,14,14'/>
            </ScrollViewer>
        </DockPanel>
        </Grid>

        <!-- ============ 詳細パネル（年度横断の人物ビュー）============ -->
        <!-- 幅 0⇔440px をアニメーションで開閉する（DETAIL_PANEL_WIDTH と一致させる）。中身は
             右端に固定幅で置き、開閉中のリフロー（文字の折り返し変化）を防ぐ -->
        <Border x:Name='DepositDetailPanel' Grid.Column='2' Width='0' ClipToBounds='True'
                Background='White' BorderBrush='#ECEBF1' BorderThickness='1,0,0,0'>
            <DockPanel Width='439' HorizontalAlignment='Right'>

                <!-- 人物ヘッダー -->
                <Border DockPanel.Dock='Top' BorderBrush='#F1F0F4' BorderThickness='0,0,0,1' Padding='18,14,12,14'>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width='Auto'/>
                            <ColumnDefinition Width='*'/>
                            <ColumnDefinition Width='Auto'/>
                        </Grid.ColumnDefinitions>
                        <Border Grid.Column='0' Width='40' Height='40' CornerRadius='20'
                                Background='#EEEDFB' VerticalAlignment='Top'>
                            <TextBlock x:Name='DepositPanelAvatarText' FontSize='16' FontWeight='Medium'
                                       Foreground='#4C4AC0' HorizontalAlignment='Center' VerticalAlignment='Center'/>
                        </Border>
                        <StackPanel Grid.Column='1' Margin='12,1,6,0'>
                            <TextBlock x:Name='DepositPanelName' FontSize='14' FontWeight='Medium'
                                       Foreground='#1A1523' TextTrimming='CharacterEllipsis'/>
                            <!-- 宛名番号とカナは1つの TextBlock に Run で並べる（別々の TextBlock だと
                                 Consolas と Meiryo UI の行送り・アセント差でベースラインが揃わない） -->
                            <TextBlock FontSize='11' Foreground='#6F6E77' Margin='0,3,0,0'
                                       TextTrimming='CharacterEllipsis'><Run x:Name='DepositPanelAddressNumber' FontFamily='Consolas'/><InlineUIContainer BaselineAlignment='Baseline'><Border Width='8'/></InlineUIContainer><Run x:Name='DepositPanelKana' Foreground='#9B9AA3'/></TextBlock>
                            <TextBlock x:Name='DepositPanelAddress' FontSize='11' Foreground='#6F6E77'
                                       TextWrapping='Wrap' Margin='0,5,0,0'/>
                        </StackPanel>
                        <!-- 「コピーしました」チップ・年度横断走査中のミニスピナー（走査中のみ回転）・閉じるボタン -->
                        <StackPanel Grid.Column='2' Orientation='Horizontal' VerticalAlignment='Top'>
                            <Border x:Name='DepositPanelCopyChip' CornerRadius='9' Padding='8,1,8,2'
                                    VerticalAlignment='Top' Margin='0,3,8,0' Visibility='Collapsed'>
                                <TextBlock FontSize='10'/>
                            </Border>
                            <Path x:Name='DepositPanelMiniSpinner' Data='M 7,1 A 6,6 0 1 1 1,7'
                                  Stroke='#5B5BD6' StrokeThickness='2'
                                  StrokeStartLineCap='Round' StrokeEndLineCap='Round'
                                  Width='14' Height='14' Stretch='None' Visibility='Collapsed'
                                  Margin='0,4,8,0' RenderTransformOrigin='0.5,0.5'
                                  ToolTip='他年度の照会履歴を確認しています'>
                                <Path.RenderTransform><RotateTransform/></Path.RenderTransform>
                            </Path>
                            <Button x:Name='DepositPanelCloseButton' Content='✕' FontSize='10'
                                    Width='22' Height='22' Cursor='Hand' FocusVisualStyle='{StaticResource KeyFocusPill}'
                                    Foreground='#6F6E77' ToolTip='閉じる (Esc)'>
                                <Button.Template>
                                    <ControlTemplate TargetType='Button'>
                                        <Border x:Name='bd' Background='Transparent' CornerRadius='11'>
                                            <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property='IsMouseOver' Value='True'>
                                                <Setter TargetName='bd' Property='Background' Value='#F4F4F6'/></Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Button.Template>
                            </Button>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- ファクト統計（事実の集計のみ・評価値は置かない）-->
                <Border DockPanel.Dock='Top' BorderBrush='#F1F0F4' BorderThickness='0,0,0,1' Padding='0,11'>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width='*'/>
                            <ColumnDefinition Width='*'/>
                            <ColumnDefinition Width='*'/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column='0'>
                            <TextBlock x:Name='DepositPanelStatRounds' FontSize='15' FontFamily='Consolas'
                                       FontWeight='Medium' Foreground='#1A1523' HorizontalAlignment='Center'/>
                            <TextBlock Text='照会回数' FontSize='10' Foreground='#9B9AA3'
                                       HorizontalAlignment='Center' Margin='0,2,0,0'/>
                        </StackPanel>
                        <StackPanel Grid.Column='1'>
                            <TextBlock x:Name='DepositPanelStatAccounts' FontSize='15' FontFamily='Consolas'
                                       FontWeight='Medium' Foreground='#1A1523' HorizontalAlignment='Center'/>
                            <TextBlock Text='のべ口座' FontSize='10' Foreground='#9B9AA3'
                                       HorizontalAlignment='Center' Margin='0,2,0,0'/>
                        </StackPanel>
                        <StackPanel Grid.Column='2'>
                            <TextBlock x:Name='DepositPanelStatFirst' FontSize='14' FontFamily='Consolas'
                                       FontWeight='Medium' Foreground='#1A1523'
                                       HorizontalAlignment='Center' Margin='0,1,0,0'/>
                            <TextBlock Text='初回照会' FontSize='10' Foreground='#9B9AA3'
                                       HorizontalAlignment='Center' Margin='0,2,0,0'/>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- フッター（両ボタンとも選択中の口座行カードが対象。
                     「差押リストに追加」は depositSeizureListScript 未設定時に非表示となり、
                     「照会ファイルを開く」がフッター全幅に広がる）-->
                <Border DockPanel.Dock='Bottom' Background='#FCFCFD' BorderBrush='#ECEBF1'
                        BorderThickness='0,1,0,0' Padding='14,10'>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width='*'/>
                            <ColumnDefinition Width='8'/>
                            <ColumnDefinition Width='*'/>
                        </Grid.ColumnDefinitions>
                        <Button x:Name='DepositPanelOpenFileButton' Grid.Column='0' Content='照会ファイルを開く'
                                Style='{StaticResource PanelGhostBtn}' IsEnabled='False'
                                ToolTipService.ShowOnDisabled='True'/>
                        <Button x:Name='DepositPanelSeizureButton' Grid.Column='2' Content='差押リストに追加'
                                Style='{StaticResource PanelPrimaryBtn}' IsEnabled='False'
                                ToolTipService.ShowOnDisabled='True'/>
                    </Grid>
                </Border>

                <!-- メモ+照会履歴タイムライン（見出し以外はコードから動的生成する）-->
                <ScrollViewer x:Name='DepositPanelScroll' VerticalScrollBarVisibility='Auto'>
                    <StackPanel Margin='18,14,14,14'>
                        <StackPanel x:Name='DepositNotesPanel' Margin='0,0,0,12'/>
                        <Border Height='1' Background='#ECEBF1' Margin='0,0,0,12'/>
                        <TextBlock Text='タイムライン' FontSize='12' FontWeight='Bold'
                                   Foreground='#6F6E77' Margin='0,0,0,10'/>
                        <StackPanel x:Name='DepositTimelinePanel'/>
                        <TextBlock x:Name='DepositPanelFootnote' FontSize='10.5' Foreground='#9B9AA3'
                                   Visibility='Collapsed' Margin='0,4,0,0' TextWrapping='Wrap'/>
                    </StackPanel>
                </ScrollViewer>
            </DockPanel>
        </Border>
" + INSURANCE_PANEL_XAML + @"
        <!-- ============ サイドバーの展開パネル（オーバーレイ式）============ -->
        <!-- レールにカーソルが0.2秒とどまると 48→208px にスライドして一覧の上に被さる
            （一覧の列位置・スクロールは動かさない）。中身は左端に固定幅で置いて開閉中の
             リフローを防ぐ（詳細パネルと同じ手法）。はみ出しは内側の Grid でクリップする
            （外枠でクリップすると右辺の影まで切れるため、影は外枠・クリップは内側で行う）。
             円環マークとナビのアイコンはレールと同じ48px帯・同じ位置・同じ大きさに置き、
             展開・収納してもガタつかせない -->
        <Border x:Name='SidebarFlyout' Grid.ColumnSpan='2' Width='48' Visibility='Collapsed'
                HorizontalAlignment='Left'
                Background='#FCFCFD' BorderBrush='#ECEBF1' BorderThickness='0,0,1,0'>
            <Border.Effect>
                <DropShadowEffect BlurRadius='16' ShadowDepth='3' Direction='0'
                                  Opacity='0.25' Color='#1A1523'/>
            </Border.Effect>
            <Grid ClipToBounds='True'>
                <DockPanel Width='207' HorizontalAlignment='Left' Margin='0,16,0,14'>

                    <!-- ワードマーク（点線円環+中心点のロックアップ。円環はレールと同位置・
                         高さもレールの26pxに合わせてナビの縦位置を揃える）-->
                    <Grid DockPanel.Dock='Top' Height='26' Margin='0,0,0,16'>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width='47'/>
                            <ColumnDefinition Width='*'/>
                        </Grid.ColumnDefinitions>
                        <Grid Grid.Column='0' Width='26' Height='26' HorizontalAlignment='Center'>
                            <Ellipse Stroke='#5B5BD6' StrokeThickness='1.8'
                                     StrokeDashArray='1.9,1.7' StrokeDashCap='Round' Margin='1'/>
                            <Ellipse Fill='#5B5BD6' Width='7' Height='7'/>
                        </Grid>
                        <StackPanel Grid.Column='1' Margin='2,0,0,0' VerticalAlignment='Center'>
                            <TextBlock Text='P A D D O C K' FontSize='13' Foreground='#1A1523' FontWeight='Bold'
                                       LineHeight='14' LineStackingStrategy='BlockLineHeight'/>
                            <TextBlock Text='財産調査台帳' FontSize='11' Foreground='#6F6E77'
                                       LineHeight='12' LineStackingStrategy='BlockLineHeight'/>
                        </StackPanel>
                    </Grid>

                    <!-- 年度セレクタ（下部固定。預金用・生命保険用の2つを縦に並べ、表示中のタブの分だけを
                         見せる。生命保険のコンボは台帳を読むまで空で無効）-->
                    <StackPanel DockPanel.Dock='Bottom' Margin='16,12,16,0'>
                        <StackPanel x:Name='DepositYearPanel'>
                            <TextBlock Text='年度（預金）' FontSize='11' Foreground='#9B9AA3' Margin='0,0,0,5'/>
                            <ComboBox x:Name='DepositYearCombo'/>
                        </StackPanel>
                        <StackPanel x:Name='InsuranceYearPanel' Visibility='Collapsed'>
                            <TextBlock Text='年度（生命保険）' FontSize='11' Foreground='#9B9AA3' Margin='0,0,0,5'/>
                            <ComboBox x:Name='InsuranceYearCombo' IsEnabled='False'/>
                        </StackPanel>
                    </StackPanel>
                    <Border DockPanel.Dock='Bottom' Height='1' Background='#F1F0F4' Margin='16,0'/>

                    <!-- ナビ（5ビューを切替可能。活性表示は SetNavSelection がコードから制御する。
                         アイコン列はレールと同じ位置・大きさの固定列とし、ラベルの縦列も揃える
                        （サイズの字形補正もレールと同一——⚙のみ14px）。
                         差押登録は depositSeizureListScript 未設定時に非表示になる）-->
                    <StackPanel>
                        <Border x:Name='DepositNavLedger' Style='{StaticResource NavItem}' Cursor='Hand' Background='Transparent'>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width='44'/>
                                    <ColumnDefinition Width='*'/>
                                </Grid.ColumnDefinitions>
                                <TextBlock x:Name='DepositNavLedgerIcon' Text='≡' FontSize='15' HorizontalAlignment='Center'/>
                                <TextBlock x:Name='DepositNavLedgerText' Grid.Column='1' Text='預金' FontSize='12' VerticalAlignment='Center'/>
                            </Grid>
                        </Border>
                        <Border x:Name='InsuranceNavLedger' Style='{StaticResource NavItem}' Cursor='Hand' Background='Transparent'>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width='44'/>
                                    <ColumnDefinition Width='*'/>
                                </Grid.ColumnDefinitions>
                                <TextBlock x:Name='InsuranceNavLedgerIcon' Text='☂' FontSize='15' HorizontalAlignment='Center'/>
                                <TextBlock x:Name='InsuranceNavLedgerText' Grid.Column='1' Text='生命保険' FontSize='12' VerticalAlignment='Center'/>
                            </Grid>
                        </Border>
                        <Border x:Name='NavRecent' Style='{StaticResource NavItem}' Cursor='Hand' Background='Transparent'>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width='44'/>
                                    <ColumnDefinition Width='*'/>
                                </Grid.ColumnDefinitions>
                                <TextBlock x:Name='NavRecentIcon' Text='◷' FontSize='15' HorizontalAlignment='Center'/>
                                <TextBlock x:Name='NavRecentText' Grid.Column='1' Text='最近見た人' FontSize='12' VerticalAlignment='Center'/>
                            </Grid>
                        </Border>
                        <Border x:Name='NavSeizure' Style='{StaticResource NavItem}' Cursor='Hand' Background='Transparent'>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width='44'/>
                                    <ColumnDefinition Width='*'/>
                                </Grid.ColumnDefinitions>
                                <TextBlock x:Name='NavSeizureIcon' Text='⚑' FontSize='15' HorizontalAlignment='Center'/>
                                <TextBlock x:Name='NavSeizureText' Grid.Column='1' Text='差押登録' FontSize='12' VerticalAlignment='Center'/>
                            </Grid>
                        </Border>
                        <Border x:Name='NavSettings' Style='{StaticResource NavItem}' Cursor='Hand' Background='Transparent'>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width='44'/>
                                    <ColumnDefinition Width='*'/>
                                </Grid.ColumnDefinitions>
                                <TextBlock x:Name='NavSettingsIcon' Text='⚙' FontSize='14' HorizontalAlignment='Center'/>
                                <TextBlock x:Name='NavSettingsText' Grid.Column='1' Text='設定' FontSize='12' VerticalAlignment='Center'/>
                            </Grid>
                        </Border>
                    </StackPanel>
                </DockPanel>
            </Grid>
        </Border>

        <!-- ============ 日付欄のカレンダー ============ -->
        <!-- 日付欄の共有カレンダー。預金・生命保険のすべての日付欄で1個を使い回し、
             クリックされた欄の直下に開く（PlacementTarget はコードで切替。
             Margin='8' は影の描画余白・オフセットで見かけの位置を欄に合わせる）。
             StaysOpen=True: 外側クリックの自動クローズ（=Popup のマウスキャプチャ）は
             TextBox・Calendar とキャプチャを奪い合って効かないため、閉じる操作は
             本体ウィンドウ側で明示的に配線する（SetupEvents。内側クリックは除外）。
             Popup は別ビジュアルツリーのためルートの変換が届かず、中身の Border が
             共有の DisplayScale を参照して表示倍率に追従する -->
        <Popup x:Name='AdvDatePopup' StaysOpen='True' AllowsTransparency='True'
               Placement='Bottom' HorizontalOffset='-8' VerticalOffset='-4'>
            <Border Background='White' BorderBrush='#ECEBF1' BorderThickness='1'
                    CornerRadius='10' Margin='8' Padding='9,9,9,7'
                    LayoutTransform='{StaticResource DisplayScale}'>
                <Border.Effect>
                    <DropShadowEffect BlurRadius='16' ShadowDepth='3' Opacity='0.18' Color='#1A1523'/>
                </Border.Effect>
                <StackPanel>
                    <Calendar x:Name='AdvDateCalendar' Style='{StaticResource AdvCalendar}'/>
                    <Border Background='#F1F0F4' Height='1' Margin='2,6,2,0'/>
                    <TextBlock x:Name='AdvDateTodayLink' Margin='5,6,5,1' FontSize='11' Foreground='#6F6E77'
                               Cursor='Hand' HorizontalAlignment='Left'><Run Text='今日 '/><Run x:Name='AdvDateTodayValue' FontFamily='Consolas'/></TextBlock>
                </StackPanel>
            </Border>
        </Popup>

        <!-- ============ アップデート告知オーバーレイ ============ -->
        <!-- update_notice.json の内容をユーザーごとに初回のみ表示する。
             中身はコードから動的生成する（BuildNoticeUI）。閉じるのは ✕ のみ -->
        <Grid x:Name='NoticeOverlay' Grid.ColumnSpan='3' Visibility='Collapsed'>
            <Border Background='#80000000'/>
            <Border Background='White' CornerRadius='10' Width='440'
                    Padding='24,20,28,20' MaxHeight='620'
                    HorizontalAlignment='Center' VerticalAlignment='Center'>
                <Border.Effect>
                    <DropShadowEffect BlurRadius='24' ShadowDepth='4' Opacity='0.25' Color='#1A1523'/>
                </Border.Effect>
                <DockPanel>
                    <DockPanel DockPanel.Dock='Top'>
                        <Button x:Name='NoticeCloseButton' DockPanel.Dock='Right' Content='✕' FontSize='10'
                                Width='22' Height='22' VerticalAlignment='Top' Margin='0,-2,-6,0'
                                Cursor='Hand' FocusVisualStyle='{StaticResource KeyFocusPill}' Foreground='#6F6E77' ToolTip='閉じる (Esc)'>
                            <Button.Template>
                                <ControlTemplate TargetType='Button'>
                                    <Border x:Name='bd' Background='Transparent' CornerRadius='11'>
                                        <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property='IsMouseOver' Value='True'>
                                            <Setter TargetName='bd' Property='Background' Value='#F4F4F6'/></Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Button.Template>
                        </Button>
                        <TextBlock Text='お知らせ' FontSize='15' FontWeight='Bold' Foreground='#1A1523'/>
                    </DockPanel>
                    <ScrollViewer VerticalScrollBarVisibility='Auto' Margin='0,6,0,0'>
                        <StackPanel x:Name='NoticeContent'/>
                    </ScrollViewer>
                </DockPanel>
            </Border>
        </Grid>
    </Grid>
</Window>";

        using (var reader = new XmlTextReader(new StringReader(xaml)))
            return (Window)XamlReader.Load(reader);
    }
}