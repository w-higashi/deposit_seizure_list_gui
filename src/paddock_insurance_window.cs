// ==============================================================================
// paddock_insurance_window.cs
// PADDOCK — 財産調査台帳（WPF GUI版）
//
// 生命保険ビューと生命保険の詳細パネルの XAML（本体の XAML に文字列としてつなぐ。預金の写しで、
// 検索キー・項目検索の欄・注記・行の分類・解約返戻金と日付の範囲・一覧の列・人物ヘッダーの
// 生年月日・統計の「のべ契約」が生命保険の台帳の形）。
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
    // 生命保険ビュー（ナビ「生命保険」で表示）。BuildWindow が預金ビューの直後につなぎ、XamlReader は
    // 本体と合わせて1つの XAML として読むため、本体の Window.Resources のスタイルをそのまま参照できる
    private const string INSURANCE_VIEW_XAML = @"
        <!-- ============ 生命保険の一覧ビュー（ナビ「生命保険」で表示）============ -->
        <DockPanel x:Name='InsuranceLedgerView' Visibility='Collapsed'>

            <!-- 検索バー+詳細検索パネル+条件チップ行 -->
            <StackPanel DockPanel.Dock='Top' Margin='14,12,14,8'>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width='128'/>
                        <ColumnDefinition Width='*'/>
                        <ColumnDefinition Width='Auto'/>
                    </Grid.ColumnDefinitions>
                    <ComboBox x:Name='InsuranceSearchKeyCombo' Grid.Column='0' Margin='0,0,8,0' SelectedIndex='0'>
                        <ComboBoxItem Content='宛名番号'/>
                        <ComboBoxItem Content='氏名'/>
                        <ComboBoxItem Content='カナ氏名'/>
                        <ComboBoxItem Content='保険会社名'/>
                        <ComboBoxItem Content='証券番号'/>
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
                            <TextBox x:Name='InsuranceSearchBox' Grid.Column='1' BorderThickness='0'
                                     Background='Transparent' FontSize='12' Foreground='#1A1523'
                                     Padding='7,6' VerticalContentAlignment='Center'/>
                            <TextBlock x:Name='InsuranceSearchPlaceholder' Grid.Column='1'
                                       Text='選択中のキーで部分一致検索' FontSize='12' Foreground='#9B9AA3'
                                       Margin='11,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                            <Button x:Name='InsuranceClearButton' Grid.Column='2' Content='✕' FontSize='10'
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
                    <Border x:Name='InsuranceAdvToggle' Grid.Column='2' Margin='8,0,0,0' Background='#F6F5FE'
                            BorderBrush='#DCDAF6' BorderThickness='1' CornerRadius='7'
                            Padding='12,0' Cursor='Hand'>
                        <StackPanel Orientation='Horizontal' VerticalAlignment='Center'>
                            <TextBlock Text='詳細検索' FontSize='12' Foreground='#5B5BD6'/>
                            <Border x:Name='InsuranceAdvCountBadge' Background='#5B5BD6' CornerRadius='8'
                                    Padding='6,1,6,2' Margin='6,0,0,0' Visibility='Collapsed'>
                                <TextBlock x:Name='InsuranceAdvCountText' FontSize='10' Foreground='White'/>
                            </Border>
                            <TextBlock x:Name='InsuranceAdvToggleCaret' Text='▼' FontSize='9' Foreground='#9B9AA3'
                                       Margin='6,1,0,0' VerticalAlignment='Center'/>
                        </StackPanel>
                    </Border>
                </Grid>

                <!-- 詳細検索パネル（淡アイリス面=特殊エリアの明示。
                     開閉は高さアニメーション・条件は閉じても持続する）-->
                <Border x:Name='InsuranceAdvPanel' Visibility='Collapsed' Background='#F9F8FD' ClipToBounds='True'
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
                                        <RowDefinition Height='6'/>
                                        <RowDefinition/>
                                    </Grid.RowDefinitions>
                                    <Border Grid.Row='0' Grid.Column='0' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvAddressBox' BorderThickness='0' Background='Transparent'
                                                     FontSize='12' Foreground='#1A1523' Padding='7,4' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvAddressPh' Text='宛名番号' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <Border Grid.Row='0' Grid.Column='2' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvNameBox' BorderThickness='0' Background='Transparent'
                                                     FontSize='12' Foreground='#1A1523' Padding='7,4' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvNamePh' Text='氏名' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <Border Grid.Row='2' Grid.Column='0' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvKanaBox' BorderThickness='0' Background='Transparent'
                                                     FontSize='12' Foreground='#1A1523' Padding='7,4' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvKanaPh' Text='カナ氏名' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <Border Grid.Row='2' Grid.Column='2' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvCompanyBox' BorderThickness='0' Background='Transparent'
                                                     FontSize='12' Foreground='#1A1523' Padding='7,4' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvCompanyPh' Text='保険会社名' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <Border Grid.Row='4' Grid.Column='0' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvPolicyBox' BorderThickness='0' Background='Transparent'
                                                     FontSize='12' Foreground='#1A1523' Padding='7,4' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvPolicyPh' Text='証券番号' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                </Grid>

                                <!-- 注記と行の分類は横並び。左列の幅が不足する場合は
                                     WrapPanel の折り返しで縦積みになる -->
                                <WrapPanel Margin='0,13,0,0'>
                                    <StackPanel Margin='0,0,28,6'>
                                        <TextBlock FontSize='11' Margin='0,0,0,5'><Run Text='注記' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  選択したものを すべて含む 行（AND）' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                        <WrapPanel>
                                            <ToggleButton x:Name='InsuranceAdvNote0' Style='{StaticResource AdvToggleChip}' Content='差押'/>
                                            <ToggleButton x:Name='InsuranceAdvNote1' Style='{StaticResource AdvToggleChip}' Content='質権'/>
                                            <ToggleButton x:Name='InsuranceAdvNote2' Style='{StaticResource AdvToggleChip}' Content='貸付金'/>
                                            <ToggleButton x:Name='InsuranceAdvNote3' Style='{StaticResource AdvToggleChip}' Content='未払保険料'/>
                                            <ToggleButton x:Name='InsuranceAdvNote4' Style='{StaticResource AdvToggleChip}' Content='配当'/>
                                            <ToggleButton x:Name='InsuranceAdvNote5' Style='{StaticResource AdvToggleChip}' Content='年金'/>
                                            <ToggleButton x:Name='InsuranceAdvNote6' Style='{StaticResource AdvToggleChip}' Content='解約済'/>
                                            <ToggleButton x:Name='InsuranceAdvNote7' Style='{StaticResource AdvToggleChip}' Content='保険金支払済'/>
                                        </WrapPanel>
                                    </StackPanel>
                                    <StackPanel Margin='0,0,0,6'>
                                        <TextBlock FontSize='11' Margin='0,0,0,5'><Run Text='行の分類' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  オンにした行も表示する' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                        <WrapPanel>
                                            <ToggleButton x:Name='InsuranceAdvClassNoContract' Style='{StaticResource AdvToggleChip}' Content='契約なし'
                                                          ToolTip='契約なし回答の行を表示する（既定では非表示）'/>
                                            <ToggleButton x:Name='InsuranceAdvClassUnanswered' Style='{StaticResource AdvToggleChip}' Content='電子回答不可'
                                                          ToolTip='電子回答不可の行を表示する（既定では非表示）'/>
                                        </WrapPanel>
                                    </StackPanel>
                                </WrapPanel>

                                <!-- 担当。トグルは選択中年度の台帳から動的生成する
                                    （RebuildInsuranceStaffToggles）。記録がない年度は案内文言に切り替える -->
                                <TextBlock FontSize='11' Margin='0,7,0,5'><Run Text='担当' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  この年度の台帳から列挙・いずれかに該当（OR）' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                <WrapPanel x:Name='InsuranceAdvStaffPanel'/>
                                <TextBlock x:Name='InsuranceAdvStaffEmptyText' Text='この年度の台帳に担当の記録はありません'
                                           FontSize='11' Foreground='#9B9AA3'/>
                            </StackPanel>

                            <!-- 右列: 数値・日付の条件 -->
                            <StackPanel Grid.Column='2'>
                                <TextBlock FontSize='11' Margin='0,0,0,5'><Run Text='解約返戻金' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  数値として読めない値の行は対象外' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                <StackPanel Orientation='Horizontal'>
                                    <Border x:Name='InsuranceAdvSurrenderMinBd' Width='118' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvSurrenderMinBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     TextAlignment='Right' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvSurrenderMinPh' Text='下限なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='0,0,10,0' HorizontalAlignment='Right' VerticalAlignment='Center'
                                                       IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <TextBlock Text='〜' FontSize='12' Foreground='#9B9AA3' Margin='6,0' VerticalAlignment='Center'/>
                                    <Border x:Name='InsuranceAdvSurrenderMaxBd' Width='118' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvSurrenderMaxBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     TextAlignment='Right' VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvSurrenderMaxPh' Text='上限なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='0,0,10,0' HorizontalAlignment='Right' VerticalAlignment='Center'
                                                       IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <TextBlock Text='円' FontSize='11.5' Foreground='#6F6E77' Margin='6,0,0,0' VerticalAlignment='Center'/>
                                </StackPanel>

                                <TextBlock FontSize='11' Margin='0,13,0,5'><Run Text='日付の範囲' FontWeight='Bold' Foreground='#1A1523'/><Run Text='  yyyy/mm/dd・片側だけでも可・満期の 99999999（終身）は対象外' FontSize='10' Foreground='#9B9AA3'/></TextBlock>
                                <StackPanel Orientation='Horizontal' Margin='0,0,0,6'>
                                    <TextBlock Text='回答基準日' Width='76' FontSize='11.5' Foreground='#6F6E77' VerticalAlignment='Center'/>
                                    <Border x:Name='InsuranceAdvBaseFromBd' Width='104' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvBaseFromBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvBaseFromPh' Text='指定なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <TextBlock Text='〜' FontSize='12' Foreground='#9B9AA3' Margin='6,0' VerticalAlignment='Center'/>
                                    <Border x:Name='InsuranceAdvBaseToBd' Width='104' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvBaseToBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvBaseToPh' Text='指定なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                </StackPanel>
                                <StackPanel Orientation='Horizontal'>
                                    <TextBlock Text='満期年月日' Width='76' FontSize='11.5' Foreground='#6F6E77' VerticalAlignment='Center'/>
                                    <Border x:Name='InsuranceAdvMaturityFromBd' Width='104' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvMaturityFromBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvMaturityFromPh' Text='指定なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                    <TextBlock Text='〜' FontSize='12' Foreground='#9B9AA3' Margin='6,0' VerticalAlignment='Center'/>
                                    <Border x:Name='InsuranceAdvMaturityToBd' Width='104' Background='White' BorderBrush='#ECEBF1'
                                            BorderThickness='1' CornerRadius='7'>
                                        <Grid>
                                            <TextBox x:Name='InsuranceAdvMaturityToBox' BorderThickness='0' Background='Transparent'
                                                     FontFamily='Consolas' FontSize='12' Foreground='#1A1523' Padding='7,4'
                                                     VerticalContentAlignment='Center'/>
                                            <TextBlock x:Name='InsuranceAdvMaturityToPh' Text='指定なし' FontSize='11' Foreground='#9B9AA3'
                                                       Margin='10,0,0,0' VerticalAlignment='Center' IsHitTestVisible='False'/>
                                        </Grid>
                                    </Border>
                                </StackPanel>
                            </StackPanel>
                        </Grid>

                        <!-- パネル下端: 案内と全クリア -->
                        <Border BorderBrush='#F1F0F4' BorderThickness='0,1,0,0' Margin='0,10,0,0' Padding='0,8,0,2'>
                            <DockPanel>
                                <TextBlock x:Name='InsuranceAdvClearLink' DockPanel.Dock='Right' Text='条件をすべてクリア'
                                           FontSize='11.5' Foreground='#6F6E77' Cursor='Hand' VerticalAlignment='Center'/>
                                <TextBlock Text='条件は入力と同時に一覧へ反映されます。設定した条件は下のチップから個別に外せます。'
                                           FontSize='10.5' Foreground='#9B9AA3' VerticalAlignment='Center'/>
                            </DockPanel>
                        </Border>
                    </StackPanel>
                </Border>

                <!-- 条件チップ行（中身はコードから動的生成する。条件がないときは行ごと非表示）-->
                <WrapPanel x:Name='InsuranceFilterChipsRow' Margin='0,8,0,0' Visibility='Collapsed'/>
            </StackPanel>

            <!-- フッター -->
            <Border DockPanel.Dock='Bottom' Background='#FCFCFD' BorderBrush='#ECEBF1'
                    BorderThickness='0,1,0,0' Padding='14,5'>
                <DockPanel>
                    <TextBlock x:Name='InsuranceStatusRight' DockPanel.Dock='Right' FontSize='11' Foreground='#9B9AA3'/>
                    <TextBlock x:Name='InsuranceStatusLeft' FontSize='11' Foreground='#9B9AA3'/>
                </DockPanel>
            </Border>

            <!-- 一覧テーブル -->
            <Grid>
                <Border Background='White' BorderBrush='#ECEBF1' BorderThickness='0,1,0,0'>
                    <ListView x:Name='InsuranceLedgerList' BorderThickness='0' Background='White' FontSize='12'
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
                                <GridViewColumn Header='保険会社名' Width='170'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding CompanyName}' TextTrimming='CharacterEllipsis'
                                                   ToolTip='{Binding CompanyName}' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='契約種類' Width='220'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding ContractType}' TextTrimming='CharacterEllipsis'
                                                   ToolTip='{Binding ContractType}' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='証券番号' Width='120'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding PolicyNumber}' FontFamily='Consolas' TextTrimming='CharacterEllipsis'
                                                   ToolTip='{Binding PolicyNumber}' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='契約の状態' Width='100'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding ContractStatus}' TextTrimming='CharacterEllipsis'
                                                   ToolTip='{Binding ContractStatus}' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='解約返戻金' Width='110'
                                                HeaderContainerStyle='{StaticResource ColHeadCenter}'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <!-- 数値として読めない生値（「電子回答不可」など）は薄文字色で出す -->
                                        <TextBlock Text='{Binding DisplaySurrenderValue}' FontFamily='Consolas'
                                                   HorizontalAlignment='Right' Margin='4,0,10,0'>
                                            <TextBlock.Style>
                                                <Style TargetType='TextBlock'>
                                                    <Style.Triggers>
                                                        <DataTrigger Binding='{Binding IsSurrenderValueText}' Value='True'>
                                                            <Setter Property='Foreground' Value='#B9B8C0'/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </TextBlock.Style>
                                        </TextBlock>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header='満期年月日' Width='104'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <TextBlock Text='{Binding MaturityDate}' FontFamily='Consolas' Margin='12,0,4,0'/>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn x:Name='InsuranceNoteColumn' Header='注記' Width='260'>
                                    <GridViewColumn.CellTemplate><DataTemplate>
                                        <Grid Background='Transparent' Margin='12,0,4,0'
                                              ToolTipService.InitialShowDelay='350'
                                              ToolTipService.ShowDuration='60000'
                                              ToolTipService.IsEnabled='{Binding HasNote}'>
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
                                            <!-- チップ列（契約なし行・電子回答不可行は代わりにピルを表示）-->
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
                                                    Visibility='{Binding IsMutedRow, Converter={StaticResource BoolVis}}'>
                                                <TextBlock Text='{Binding MutedPillText}' FontSize='11' Foreground='#9B9AA3'/>
                                            </Border>
                                        </Grid>
                                    </DataTemplate></GridViewColumn.CellTemplate>
                                </GridViewColumn>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Border>

                <!-- 空状態メッセージ -->
                <TextBlock x:Name='InsuranceEmptyStateText' Visibility='Collapsed'
                           HorizontalAlignment='Center' VerticalAlignment='Center'
                           TextAlignment='Center' FontSize='12' Foreground='#9B9AA3'
                           LineHeight='20'/>

                <!-- 読み込みオーバーレイ（テーブル範囲のみを覆う。スピナーのみの意匠）-->
                <Grid x:Name='InsuranceLoadingOverlay' Background='#B3FFFFFF' Visibility='Collapsed'>
                    <Path x:Name='InsuranceSpinnerPath' Data='M 20,2 A 18,18 0 1 1 2,20'
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
";

    // 生命保険の詳細パネルの XAML（本体の XAML の預金のパネルの直後につなぐ。預金のパネルの写しで、
    // 人物ヘッダーに契約者の生年月日の行があり、統計の2つ目が「のべ契約」）
    private const string INSURANCE_PANEL_XAML = @"
        <!-- ============ 生命保険の詳細パネル（年度横断の人物ビュー。預金のパネルと同じ列に置き、開いている方だけが幅を持つ）============ -->
        <!-- 幅 0⇔440px をアニメーションで開閉する（DETAIL_PANEL_WIDTH と一致させる）。中身は
             右端に固定幅で置き、開閉中のリフロー（文字の折り返し変化）を防ぐ -->
        <Border x:Name='InsuranceDetailPanel' Grid.Column='2' Width='0' ClipToBounds='True'
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
                            <TextBlock x:Name='InsurancePanelAvatarText' FontSize='16' FontWeight='Medium'
                                       Foreground='#4C4AC0' HorizontalAlignment='Center' VerticalAlignment='Center'/>
                        </Border>
                        <StackPanel Grid.Column='1' Margin='12,1,6,0'>
                            <TextBlock x:Name='InsurancePanelName' FontSize='14' FontWeight='Medium'
                                       Foreground='#1A1523' TextTrimming='CharacterEllipsis'/>
                            <!-- 宛名番号とカナは1つの TextBlock に Run で並べる（別々の TextBlock だと
                                 Consolas と Meiryo UI の行送り・アセント差でベースラインが揃わない） -->
                            <TextBlock FontSize='11' Foreground='#6F6E77' Margin='0,3,0,0'
                                       TextTrimming='CharacterEllipsis'><Run x:Name='InsurancePanelAddressNumber' FontFamily='Consolas'/><InlineUIContainer BaselineAlignment='Baseline'><Border Width='8'/></InlineUIContainer><Run x:Name='InsurancePanelKana' Foreground='#9B9AA3'/></TextBlock>
                            <TextBlock x:Name='InsurancePanelAddress' FontSize='11' Foreground='#6F6E77'
                                       TextWrapping='Wrap' Margin='0,5,0,0'/>
                            <!-- 契約者の生年月日（見出しの字は付けず日付だけ。値がなければ行ごと出さない）-->
                            <TextBlock x:Name='InsurancePanelBirthDate' FontSize='11' FontFamily='Consolas' Foreground='#6F6E77'
                                       Margin='0,5,0,0' Visibility='Collapsed'/>
                        </StackPanel>
                        <!-- 「コピーしました」チップ・年度横断走査中のミニスピナー（走査中のみ回転）・閉じるボタン -->
                        <StackPanel Grid.Column='2' Orientation='Horizontal' VerticalAlignment='Top'>
                            <Border x:Name='InsurancePanelCopyChip' CornerRadius='9' Padding='8,1,8,2'
                                    VerticalAlignment='Top' Margin='0,3,8,0' Visibility='Collapsed'>
                                <TextBlock FontSize='10'/>
                            </Border>
                            <Path x:Name='InsurancePanelMiniSpinner' Data='M 7,1 A 6,6 0 1 1 1,7'
                                  Stroke='#5B5BD6' StrokeThickness='2'
                                  StrokeStartLineCap='Round' StrokeEndLineCap='Round'
                                  Width='14' Height='14' Stretch='None' Visibility='Collapsed'
                                  Margin='0,4,8,0' RenderTransformOrigin='0.5,0.5'
                                  ToolTip='他年度の照会履歴を確認しています'>
                                <Path.RenderTransform><RotateTransform/></Path.RenderTransform>
                            </Path>
                            <Button x:Name='InsurancePanelCloseButton' Content='✕' FontSize='10'
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
                            <TextBlock x:Name='InsurancePanelStatRounds' FontSize='15' FontFamily='Consolas'
                                       FontWeight='Medium' Foreground='#1A1523' HorizontalAlignment='Center'/>
                            <TextBlock Text='照会回数' FontSize='10' Foreground='#9B9AA3'
                                       HorizontalAlignment='Center' Margin='0,2,0,0'/>
                        </StackPanel>
                        <StackPanel Grid.Column='1'>
                            <TextBlock x:Name='InsurancePanelStatContracts' FontSize='15' FontFamily='Consolas'
                                       FontWeight='Medium' Foreground='#1A1523' HorizontalAlignment='Center'/>
                            <TextBlock Text='のべ契約' FontSize='10' Foreground='#9B9AA3'
                                       HorizontalAlignment='Center' Margin='0,2,0,0'/>
                        </StackPanel>
                        <StackPanel Grid.Column='2'>
                            <TextBlock x:Name='InsurancePanelStatFirst' FontSize='14' FontFamily='Consolas'
                                       FontWeight='Medium' Foreground='#1A1523'
                                       HorizontalAlignment='Center' Margin='0,1,0,0'/>
                            <TextBlock Text='初回照会' FontSize='10' Foreground='#9B9AA3'
                                       HorizontalAlignment='Center' Margin='0,2,0,0'/>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- フッター（両ボタンとも選択中の契約カードが対象。
                     「差押リストに追加」は insuranceSeizureListScript 未設定時に非表示となり、
                     「照会ファイルを開く」がフッター全幅に広がる）-->
                <Border DockPanel.Dock='Bottom' Background='#FCFCFD' BorderBrush='#ECEBF1'
                        BorderThickness='0,1,0,0' Padding='14,10'>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width='*'/>
                            <ColumnDefinition Width='8'/>
                            <ColumnDefinition Width='*'/>
                        </Grid.ColumnDefinitions>
                        <Button x:Name='InsurancePanelOpenFileButton' Grid.Column='0' Content='照会ファイルを開く'
                                Style='{StaticResource PanelGhostBtn}' IsEnabled='False'
                                ToolTipService.ShowOnDisabled='True'/>
                        <Button x:Name='InsurancePanelSeizureButton' Grid.Column='2' Content='差押リストに追加'
                                Style='{StaticResource PanelPrimaryBtn}' IsEnabled='False'
                                ToolTipService.ShowOnDisabled='True'/>
                    </Grid>
                </Border>

                <!-- メモ+照会履歴タイムライン（見出し以外はコードから動的生成する）-->
                <ScrollViewer x:Name='InsurancePanelScroll' VerticalScrollBarVisibility='Auto'>
                    <StackPanel Margin='18,14,14,14'>
                        <StackPanel x:Name='InsuranceNotesPanel' Margin='0,0,0,12'/>
                        <Border Height='1' Background='#ECEBF1' Margin='0,0,0,12'/>
                        <TextBlock Text='タイムライン' FontSize='12' FontWeight='Bold'
                                   Foreground='#6F6E77' Margin='0,0,0,10'/>
                        <StackPanel x:Name='InsuranceTimelinePanel'/>
                        <TextBlock x:Name='InsurancePanelFootnote' FontSize='10.5' Foreground='#9B9AA3'
                                   Visibility='Collapsed' Margin='0,4,0,0' TextWrapping='Wrap'/>
                    </StackPanel>
                </ScrollViewer>
            </DockPanel>
        </Border>
";
}