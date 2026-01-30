Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase

# -------------------------
# データ保持用
# -------------------------
$FileItems = New-Object System.Collections.ObjectModel.ObservableCollection[object]

# -------------------------
# XAML（WindowChrome + DataGrid）
# -------------------------
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
        Title="Hash Checker"
        Width="1000" Height="600"
        WindowStyle="None"
        ResizeMode="CanResizeWithGrip"
        Background="#222"
        Foreground="White">

        <Window.Resources>
            <Style TargetType="CheckBox">
                <Setter Property="Foreground" Value="White"/>
            </Style>
        </Window.Resources>

    <shell:WindowChrome.WindowChrome>
        <shell:WindowChrome
            CaptionHeight="32"
            CornerRadius="0"
            GlassFrameThickness="0"
            ResizeBorderThickness="5"
            UseAeroCaptionButtons="False" />
    </shell:WindowChrome.WindowChrome>

    <Border BorderBrush="#444" BorderThickness="1" Background="#222">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="32"/>   <!-- タイトルバー -->
                <RowDefinition Height="Auto"/> <!-- 上部操作パネル -->
                <RowDefinition Height="*"/>    <!-- DataGrid -->
                <RowDefinition Height="24"/>   <!-- ステータスバー -->
            </Grid.RowDefinitions>

            <!-- タイトルバー -->
            <Grid Grid.Row="0" Background="#333" IsHitTestVisible="True">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="Hash Checker"
                           VerticalAlignment="Center"
                           Margin="8,0,0,0"
                           FontWeight="Bold"/>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <ToggleButton x:Name="BtnTopMost"
                                  Width="32" Height="32"
                                  Margin="0,0,4,0"
                                  ToolTip="最前面固定"
                                  Focusable="False"
                                  shell:WindowChrome.IsHitTestVisibleInChrome="True">
                    
                        <ToggleButton.Template>
                            <ControlTemplate TargetType="ToggleButton">
                                <Border Background="Transparent">
                                    <ContentPresenter HorizontalAlignment="Center"
                                                      VerticalAlignment="Center"/>
                                </Border>
                            </ControlTemplate>
                        </ToggleButton.Template>
                    
                        <ToggleButton.Style>
                            <Style TargetType="ToggleButton" BasedOn="{x:Null}">
                                <Setter Property="Foreground" Value="White"/>
                                <Setter Property="Content" Value="📌"/>   <!-- OFF：斜め -->
                    
                                <Style.Triggers>
                                    <Trigger Property="IsChecked" Value="True">
                                        <Setter Property="Content" Value="📍"/>  <!-- ON：縦 -->
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </ToggleButton.Style>
                    
                    </ToggleButton>
                    <Button x:Name="BtnMin"
                            Width="32" Height="22"
                            Margin="0,0,4,0"
                            Content="─"
                            ToolTip="最小化"
                            Focusable="False"
                            shell:WindowChrome.IsHitTestVisibleInChrome="True"/>

                    <Button x:Name="BtnMax"
                            Width="32" Height="22"
                            Margin="0,0,4,0"
                            Content="□"
                            ToolTip="最大化 / 元に戻す"
                            Focusable="False"
                            shell:WindowChrome.IsHitTestVisibleInChrome="True"/>

                    <Button x:Name="BtnClose"
                            Width="32" Height="22"
                            Content="✕"
                            Background="#933"
                            ToolTip="閉じる"
                            Focusable="False"
                            shell:WindowChrome.IsHitTestVisibleInChrome="True"/>

                </StackPanel>
            </Grid>

            <!-- 上部操作パネル -->
            <Grid Grid.Row="1" Margin="4" Background="#222">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>   <!-- チェックボックス群 -->
                    <ColumnDefinition Width="Auto"/>   <!-- フィルタ -->
                    <ColumnDefinition Width="*"/>      <!-- ボタン群（右寄せ） -->
                </Grid.ColumnDefinitions>

                <!-- チェックボックス群 -->
                <StackPanel Orientation="Horizontal" Margin="0,2,12,2" Grid.Column="0">
                    <CheckBox x:Name="ChkSize"    Content="ファイルサイズ" Margin="0,0,8,0" IsChecked="False"/>
                    <CheckBox x:Name="ChkCRC32"   Content="CRC32"           Margin="0,0,8,0" IsChecked="True"/>
                    <CheckBox x:Name="ChkMD5"     Content="MD5"                     Margin="0,0,8,0" IsChecked="True"/>
                    <CheckBox x:Name="ChkSHA1"    Content="SHA1"                    Margin="0,0,8,0" IsChecked="False"/>
                    <CheckBox x:Name="ChkSHA256"  Content="SHA256"          Margin="0,0,8,0" IsChecked="False"/>
                </StackPanel>

                <!-- フィルタ -->
                <StackPanel Orientation="Horizontal" Margin="0,2,12,2" Grid.Column="1">
                    <TextBlock Text="拡張子フィルタ：" VerticalAlignment="Center" Margin="0,0,4,0"/>

                    <Grid Width="150">
                        <TextBox x:Name="TxtExtFilter"
                                 VerticalAlignment="Center"
                                 ToolTip="スペース区切りで複数指定可 例：jpg jpeg"
                                 ToolTipService.InitialShowDelay="0"/>
                        <Button x:Name="BtnClearExt"
                                Content="✕"
                                Width="18"
                                Height="18"
                                HorizontalAlignment="Right"
                                VerticalAlignment="Center"
                                Margin="0,0,2,0"
                                Background="#444"
                                Foreground="White"
                                BorderThickness="0"
                                Padding="0"
                                ToolTip="フィルタをクリア"/>
                    </Grid>
                </StackPanel>

                <!-- ボタン群（右寄せ） -->
                <StackPanel Orientation="Horizontal" Margin="0,2,0,2" Grid.Column="2" HorizontalAlignment="Right">
                    <Button x:Name="BtnRecalc"   Content="リスト再計算" Margin="0,0,4,0" Padding="8,2"/>
                    <Button x:Name="BtnClear"    Content="リスト初期化" Margin="0,0,4,0" Padding="8,2"/>
                    <Button x:Name="BtnSaveCsv"  Content="CSV保存"      Margin="0,0,4,0" Padding="8,2"/>
                </StackPanel>
            </Grid>
            <!-- DataGrid 本体 -->
            <Border Grid.Row="2" Margin="4" BorderBrush="#555" BorderThickness="1" MinHeight="200">
                <DataGrid x:Name="DgFiles"

                          AutoGenerateColumns="False"
                          CanUserAddRows="False"
                          CanUserDeleteRows="False"
                          IsReadOnly="True"
                          SelectionMode="Extended"
                          SelectionUnit="FullRow"
                          HeadersVisibility="Column"
                          GridLinesVisibility="Horizontal"
                          Background="#222"
                          Foreground="Black"
                          AlternatingRowBackground="#262626"
                          BorderThickness="0">

                    <DataGrid.Resources>
                        <Style TargetType="DataGridRow">
                            <Setter Property="Background" Value="{Binding Tag.Background}"/>
                            <Setter Property="Foreground" Value="{Binding Tag.Foreground}"/>
                        </Style>

                        <Style TargetType="DataGridCell">
                            <Setter Property="Foreground" Value="{Binding Tag.Foreground}"/>
                            <Setter Property="Background" Value="{Binding Tag.Background}"/>
                        </Style>
                    </DataGrid.Resources>
                    <DataGrid.Columns>
                        <DataGridTextColumn x:Name="ColFileName" Header="ファイル名" Binding="{Binding FileName}" Width="200" SortMemberPath="FileName"/>
                        <DataGridTextColumn x:Name="ColSizeRaw"
                                            Header="サイズ"
                                            Binding="{Binding SizeDisplay}"
                                            Width="80"
                                            SortMemberPath="SizeRaw">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="TextAlignment" Value="Right"/>
                                    <Setter Property="Padding" Value="0,0,4,0"/> <!-- 少し右に余白 -->
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>
                        <DataGridTextColumn x:Name="ColCRC32"    Header="CRC32"    Binding="{Binding CRC32}" Width="62" SortMemberPath="CRC32"/>
                        <DataGridTextColumn x:Name="ColMD5"      Header="MD5"      Binding="{Binding MD5}" Width="222" SortMemberPath="MD5"/>
                        <DataGridTextColumn x:Name="ColSHA1"     Header="SHA1"     Binding="{Binding SHA1}" Width="275" SortMemberPath="SHA1"/>
                        <DataGridTextColumn x:Name="ColSHA256"   Header="SHA256"   Binding="{Binding SHA256}" Width="432" SortMemberPath="SHA256"/>
                    </DataGrid.Columns>
                    <DataGrid.ContextMenu>
                        <ContextMenu>
                            <MenuItem Header="コピー（選択行）" x:Name="CtxCopy"/>
                            <Separator/>
                            <MenuItem Header="全選択" x:Name="CtxSelectAll"/>
                            <MenuItem Header="選択解除" x:Name="CtxClearSelection"/>
                        </ContextMenu>
                    </DataGrid.ContextMenu>
                </DataGrid>
            </Border>

            <!-- ステータスバー -->
            <Grid Grid.Row="3" Background="#333">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>      <!-- 左側（メッセージ） -->
                    <ColumnDefinition Width="Auto"/>   <!-- 右側（ステータス） -->
                </Grid.ColumnDefinitions>

                <!-- 左側メッセージ -->
                <TextBlock Text=" ここにファイルをドラッグ＆ドロップしてください"
                           Margin="8,0,0,0"
                           VerticalAlignment="Center"
                           Foreground="#AAA"/>

                <!-- 右側ステータス -->
                <TextBlock x:Name="TxtStatus"
                           Grid.Column="1"
                           Margin="4,0,16,0"
                           VerticalAlignment="Center"
                           Foreground="White"
                           Text="処理：0 / 入力：0"/>
            </Grid>
        </Grid>
    </Border>
</Window>
'@

# XAML 読み込み
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

if ($null -eq $window) {
    Write-Host "❌ XAML の読み込みに失敗しました"
    exit
}

# コントロール取得
$DgFiles      = $window.FindName("DgFiles")
$ChkSize      = $window.FindName("ChkSize")
$ChkCRC32     = $window.FindName("ChkCRC32")
$ChkMD5       = $window.FindName("ChkMD5")
$ChkSHA1      = $window.FindName("ChkSHA1")
$ChkSHA256    = $window.FindName("ChkSHA256")
$TxtExtFilter = $window.FindName("TxtExtFilter")
$BtnClearExt = $window.FindName("BtnClearExt")
$BtnRecalc    = $window.FindName("BtnRecalc")
$BtnClear     = $window.FindName("BtnClear")
$BtnSaveCsv   = $window.FindName("BtnSaveCsv")
$BtnTopMost   = $window.FindName("BtnTopMost")
$BtnMin       = $window.FindName("BtnMin")
$BtnMax       = $window.FindName("BtnMax")
$BtnClose     = $window.FindName("BtnClose")
$TxtStatus    = $window.FindName("TxtStatus")
$CtxCopy          = $window.FindName("CtxCopy")
$CtxSelectAll     = $window.FindName("CtxSelectAll")
$CtxClearSelection = $window.FindName("CtxClearSelection")

# DataGrid にバインド
$DgFiles.ItemsSource = $FileItems

# --- Shift + ホイールで横スクロール（高速版） ---
$DgFiles.Add_PreviewMouseWheel({
    param($s, $e)

    if ([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftShift) -or
        [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::RightShift)) {

        $scrollViewer = $DgFiles.Template.FindName("DG_ScrollViewer", $DgFiles)

        if ($scrollViewer -ne $null) {

            $speed = 5   # ← ここを変えるだけでスクロール量を調整できる

            for ($i = 0; $i -lt $speed; $i++) {
                if ($e.Delta -gt 0) {
                    $scrollViewer.LineLeft()
                } else {
                    $scrollViewer.LineRight()
                }
            }
            $e.Handled = $true
        }
    }
})

# 拡張子フィルタのクリアボタン
$BtnClearExt.Add_Click({
    $TxtExtFilter.Text = ""
})

function Update-ColumnVisibility {
    $DgFiles.Columns[1].Visibility = $(if ($ChkSize.IsChecked)  { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed })
    $DgFiles.Columns[2].Visibility = $(if ($ChkCRC32.IsChecked) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed })
    $DgFiles.Columns[3].Visibility = $(if ($ChkMD5.IsChecked)   { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed })
    $DgFiles.Columns[4].Visibility = $(if ($ChkSHA1.IsChecked)  { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed })
    $DgFiles.Columns[5].Visibility = $(if ($ChkSHA256.IsChecked){ [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed })
}
# ★ 起動直後に列の表示/非表示を反映する
$window.Add_Loaded({
    Update-ColumnVisibility
})

# -------------------------
# タイトルバーのボタンイベント
# -------------------------
# 起動時の TopMost ボタンの見た目を OFF に統一
$BtnTopMost.IsChecked = $false
$BtnTopMost.Background = "#333333"
$window.Topmost = $false

# WindowState 変更時に最大化ボタンのアイコンを同期
$window.Add_StateChanged({
    if ($window.WindowState -eq 'Maximized') {
        $BtnMax.Content = "❐"   # 最大化中の表示
    }
    else {
        $BtnMax.Content = "□"   # 通常表示
    }
})

# 最前面固定（ToggleButton）
$BtnTopMost.Add_Click({
    if ($BtnTopMost.IsChecked) {
        $window.Topmost = $true
        $BtnTopMost.Background = "#5555FF"   # ON の見た目（任意）
    }
    else {
        $window.Topmost = $false
        $BtnTopMost.Background = "#333333"   # OFF の見た目（任意）
    }
})

# 最小化
$BtnMin.Add_Click({
    $window.WindowState = 'Minimized'
})

# 最大化 / 元に戻す
$BtnMax.Add_Click({
    if ($window.WindowState -eq 'Maximized') {
        $window.WindowState = 'Normal'
        $BtnMax.Content = "□"   # 通常表示
    }
    else {
        $window.WindowState = 'Maximized'
        $BtnMax.Content = "❐"   # 最大化中の表示
    }
})

# 閉じる
$BtnClose.Add_Click({
    $window.Close()
})

# -------------------------
# ユーティリティ
# -------------------------
function New-SoftColorPalette {
    @(
        "#FFCCE5FF", "#FFE0F5FF", "#FFD6E5FF", "#FFCFE8FF", "#FFCCE8F0",
        "#FFFFE0CC", "#FFFFF0CC", "#FFFFE8D1", "#FFFFE8C6", "#FFFFF2D6",
        "#FFE0FFCC", "#FFE8FFD6", "#FFE5FFE0", "#FFDFFFE8", "#FFE8FFF0",
        "#FFF2CCFF", "#FFE8CCFF", "#FFE0CCFF", "#FFDACCFF", "#FFE8D6FF",
        "#FFE0FFE8", "#FFD6FFF2", "#FFD6F0FF", "#FFD6E8FF", "#FFD6FFFF",
        "#FFE8F0CC", "#FFE8FFD1", "#FFE8F5CC", "#FFE8F0D6", "#FFE8F0E0"
    )
}

$ColorPalette = New-SoftColorPalette

# -------------------------
# ファイル追加・ハッシュ計算
# -------------------------
function Add-Files {
    param([string[]]$Paths)

    foreach ($file in $Paths) {

        # ★ LiteralPath を使わないと [] を含むパスが壊れる
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { continue }

        $fi = Get-Item -LiteralPath $file
        $full = $fi.FullName

        # 重複チェック
        if ($FileItems | Where-Object { $_.FullPath -eq $full }) {
            continue
        }

        # 拡張子フィルタ
        $extFilter = $TxtExtFilter.Text
        if ($extFilter) {
            $rawExts = $extFilter.Split(' ') | Where-Object { $_ -ne "" }
            $exts = foreach ($e in $rawExts) {
                $e = $e.Trim().ToLower()
                if ($e.StartsWith(".")) { $e } else { "." + $e }
            }
            if ($exts.Count -gt 0) {
                $ext = [IO.Path]::GetExtension($file).ToLower()
                if ($exts -notcontains $ext) { continue }
            }
        }

        # 情報作成
        $sizeRaw = $fi.Length
        $sizeDisplay = "{0:N0}" -f $sizeRaw

        $hashAlgo = [System.Security.Cryptography.MD5]::Create()
        $bytes    = [System.Text.Encoding]::UTF8.GetBytes($full)
        $hashID   = $hashAlgo.ComputeHash($bytes)
        $hashAlgo.Dispose()
        $hexID = ([System.BitConverter]::ToString($hashID)).Replace("-", "")

        $item = [PSCustomObject]@{
            FileName    = $fi.Name
            FullPath    = $full
            HashID      = $hexID
            SizeDisplay = $sizeDisplay
            SizeRaw     = $sizeRaw
            CRC32       = ""
            MD5         = ""
            SHA1        = ""
            SHA256      = ""
            Tag         = $null
        }

        $FileItems.Add($item)
    }

    Update-Status
    Update-RowColors
}

# --- C# CRC32 を最初に読み込む（switch の外に置くのが重要） ---
if (-not ("CRC32" -as [type])) {

$code = @"
using System;
using System.IO;

public class CRC32 {
    private static uint[] table;

    static CRC32() {
        uint poly = 0xedb88320;
        table = new uint[256];
        for (uint i = 0; i < 256; i++) {
            uint r = i;
            for (int j = 0; j < 8; j++) {
                r = (r & 1) != 0 ? (r >> 1) ^ poly : (r >> 1);
            }
            table[i] = r;
        }
    }

    public static string Calculate(Stream stream) {
        uint crc = 0xffffffff;
        int b;
        while ((b = stream.ReadByte()) != -1) {
            crc = (crc >> 8) ^ table[(crc ^ (byte)b) & 0xff];
        }
        return string.Format("{0:X8}", crc ^ 0xffffffff);
    }
}
"@

Add-Type -TypeDefinition $code
}

# --- ハッシュ関数本体 ---
function Get-Hash {
    param([string]$Path, [string]$Algorithm)

    $stream = [System.IO.File]::OpenRead($Path)

    try {
        switch ($Algorithm) {

            "CRC32" {
                return [CRC32]::Calculate($stream)
            }

            "MD5" {
                $hash = [System.Security.Cryptography.MD5]::Create().ComputeHash($stream)
            }

            "SHA1" {
                $hash = [System.Security.Cryptography.SHA1]::Create().ComputeHash($stream)
            }

            "SHA256" {
                $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($stream)
            }
        }
    }
    finally {
        $stream.Dispose()
    }

    return ([System.BitConverter]::ToString($hash)).Replace("-", "")
}

function Recalc-Hashes {
    param(
        [bool]$OnlyMissing = $true
    )

    # ★ FileItems を固定コピー（これが最重要）
    $items = @($FileItems)

    $total = $items.Count
    $i = 0

    foreach ($item in $items) {

        $path = $item.FullPath

        if ($ChkCRC32.IsChecked -and (-not $OnlyMissing -or [string]::IsNullOrEmpty($item.CRC32))) {
            $item.CRC32 = Get-Hash -Path $path -Algo "CRC32"
        }

        if ($ChkMD5.IsChecked -and (-not $OnlyMissing -or [string]::IsNullOrEmpty($item.MD5))) {
            $item.MD5 = Get-Hash -Path $path -Algo "MD5"
        }
        if ($ChkSHA1.IsChecked -and (-not $OnlyMissing -or [string]::IsNullOrEmpty($item.SHA1))) {
            $item.SHA1 = Get-Hash -Path $path -Algo "SHA1"
        }
        if ($ChkSHA256.IsChecked -and (-not $OnlyMissing -or [string]::IsNullOrEmpty($item.SHA256))) {
            $item.SHA256 = Get-Hash -Path $path -Algo "SHA256"
        }

        $i++

        # ★ UI 更新は 10 件ごとに限定（高速 & 安定）
        if ($i % 10 -eq 0) {
            $window.Dispatcher.Invoke([Action]{
                $TxtStatus.Text = "処理：$i / 入力：$total"
            }, "Background")
        }
    }

    # ★ 最後に 1 回だけ UI 更新
    $window.Dispatcher.Invoke([Action]{
        $TxtStatus.Text = "処理：$total / 入力：$total"
    }, "Background")

    Update-RowColors
}

# -------------------------
# 部分一致グループ＋色分け
# -------------------------
function Update-RowColors {

    if ($FileItems.Count -eq 0) { return }

    # グループID管理
    $groups = @{}
    $colorIndex = 0

    for ($i = 0; $i -lt $FileItems.Count; $i++) {

        if ($groups.ContainsKey($i)) { continue }

        $groups[$i] = $colorIndex
        $colorIndex++

        for ($j = $i + 1; $j -lt $FileItems.Count; $j++) {

            if ($groups.ContainsKey($j)) { continue }

            $a = $FileItems[$i]
            $b = $FileItems[$j]

            # ★ ハッシュ一致だけでグループ化（サイズは使わない）
            $match = $false

            if ($ChkCRC32.IsChecked -and $a.CRC32 -ne "" -and $a.CRC32 -eq $b.CRC32) { $match = $true }
            if ($ChkMD5.IsChecked   -and $a.MD5   -ne "" -and $a.MD5   -eq $b.MD5)   { $match = $true }
            if ($ChkSHA1.IsChecked  -and $a.SHA1  -ne "" -and $a.SHA1  -eq $b.SHA1)  { $match = $true }
            if ($ChkSHA256.IsChecked -and $a.SHA256 -ne "" -and $a.SHA256 -eq $b.SHA256) { $match = $true }

            if ($match) {
                $groups[$j] = $groups[$i]
            }
        }
    }

    # グループごとに色付け
    foreach ($index in 0..($FileItems.Count - 1)) {

        $row = $FileItems[$index]
        $groupId = $groups[$index]

        # 同じグループの行を取得
        $sameGroup = foreach ($k in $groups.Keys) {
            if ($groups[$k] -eq $groupId) { $FileItems[$k] }
        }

        $brushConv = New-Object System.Windows.Media.BrushConverter

        if ($sameGroup.Count -eq 1) {
            # グループなし → ダーク背景
            $bg = $brushConv.ConvertFromString("#222")
            $fg = $brushConv.ConvertFromString("White")
        }
        else {
            # グループ背景色
            $bg = $brushConv.ConvertFromString($ColorPalette[$groupId % $ColorPalette.Count])

            # ★ グループ内の不一致チェック（赤文字条件）
            $sizeUnique   = ($sameGroup | Select-Object -ExpandProperty SizeRaw   -Unique).Count
            $crc32Unique  = ($sameGroup | Select-Object -ExpandProperty CRC32     -Unique).Count
            $md5Unique    = ($sameGroup | Select-Object -ExpandProperty MD5       -Unique).Count
            $sha1Unique   = ($sameGroup | Select-Object -ExpandProperty SHA1      -Unique).Count
            $sha256Unique = ($sameGroup | Select-Object -ExpandProperty SHA256    -Unique).Count

            # ★ 赤文字条件（2つだけ）
            if ($sizeUnique -gt 1 -or
                $crc32Unique -gt 1 -or
                $md5Unique   -gt 1 -or
                $sha1Unique  -gt 1 -or
                $sha256Unique -gt 1) {

                $fg = $brushConv.ConvertFromString("Red")
            }
            else {
                $fg = $brushConv.ConvertFromString("Black")
            }
        }

        $row.Tag = [PSCustomObject]@{
            Background = $bg
            Foreground = $fg
        }
    }

    $DgFiles.Items.Refresh()
}

# -------------------------
# ステータス更新
# -------------------------
function Update-Status {
    $TxtStatus.Text = "処理：{0} / 入力：{1}" -f $FileItems.Count, $FileItems.Count
}

# -------------------------
# CSV 保存
# -------------------------
function Save-Csv {

    # ① 対象データ（選択行 or 全行）
    $items =
        if ($DgFiles.SelectedItems.Count -gt 0) {
            $DgFiles.SelectedItems
        }
        else {
            $FileItems
        }

    if ($items.Count -eq 0) { return }

    # ② 保存ダイアログ
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = @(
        "CSV UTF-8 (BOMなし)|*.csv",
        "CSV UTF-8 (BOMあり)|*.csv",
        "CSV Shift-JIS|*.csv",
        "TSV UTF-8 (BOMなし)|*.tsv",
        "TSV UTF-8 (BOMあり)|*.tsv",
        "TSV Shift-JIS|*.tsv",
        "すべてのファイル (*.*)|*.*"
    ) -join "|"

    $dlg.FilterIndex = 1
    $dlg.FileName = "hash_list.csv"

    if (-not $dlg.ShowDialog()) { return }

    # ③ 出力する列をチェックボックスに基づいて決定
    $columns = @(
        @{ Name = "ファイル名"; Key = "FileName" }
    )

    if ($ChkSize.IsChecked)  { $columns += @{ Name = "サイズ"; Key = "SizeRaw" } }
    if ($ChkCRC32.IsChecked) { $columns += @{ Name = "CRC32";  Key = "CRC32" } }
    if ($ChkMD5.IsChecked)   { $columns += @{ Name = "MD5";    Key = "MD5" } }
    if ($ChkSHA1.IsChecked)  { $columns += @{ Name = "SHA1";   Key = "SHA1" } }
    if ($ChkSHA256.IsChecked){ $columns += @{ Name = "SHA256"; Key = "SHA256" } }

    # ④ 区切り文字の決定
    $delimiter =
        if ($dlg.FilterIndex -le 3) { "," }   # CSV
        else { "`t" }                         # TSV

    # ⑤ ヘッダー行
    $header = ($columns | ForEach-Object { $_.Name }) -join $delimiter

    # ⑥ データ行
    $lines = foreach ($item in $items) {
        ($columns | ForEach-Object { $item.($_.Key) }) -join $delimiter
    }

    $text = $header + "`r`n" + ($lines -join "`r`n")

    # ⑦ エンコーディングの決定
    switch ($dlg.FilterIndex) {
        1 { $enc = [Text.UTF8Encoding]::new($false) }  # CSV UTF-8 BOMなし
        2 { $enc = [Text.UTF8Encoding]::new($true) }   # CSV UTF-8 BOMあり
        3 { $enc = [Text.Encoding]::GetEncoding("Shift-JIS") } # CSV SJIS
        4 { $enc = [Text.UTF8Encoding]::new($false) }  # TSV UTF-8 BOMなし
        5 { $enc = [Text.UTF8Encoding]::new($true) }   # TSV UTF-8 BOMあり
        6 { $enc = [Text.Encoding]::GetEncoding("Shift-JIS") } # TSV SJIS
        default { $enc = [Text.UTF8Encoding]::new($false) }
    }

    # ⑧ 保存
    [System.IO.File]::WriteAllText($dlg.FileName, $text, $enc)
}

# -------------------------
# イベントハンドラ
# -------------------------
$window.Add_MouseLeftButtonDown({
    param($sender,$e)

    # クリックされた要素がボタン／トグルボタンなら何もしない
    if ($e.OriginalSource -is [System.Windows.Controls.Primitives.ButtonBase]) {
        return
    }

    if ($e.GetPosition($window).Y -le 32) {
        $window.DragMove()
    }
})

# 再計算
$BtnRecalc.Add_Click({
    Recalc-Hashes -OnlyMissing:$true
})

# 初期化
$BtnClear.Add_Click({
    $FileItems.Clear()
    Update-Status
    Update-RowColors
})

# CSV 保存
$BtnSaveCsv.Add_Click({
    Save-Csv
})

# チェックボックス変更時 → 列表示/非表示 & 色更新
$ChkSize.Add_Click({
    $DgFiles.Columns[1].Visibility = $(if ($ChkSize.IsChecked) { "Visible" } else { "Collapsed" })
    Update-RowColors
})
$ChkCRC32.Add_Click({
    $DgFiles.Columns[2].Visibility = $(if ($ChkCRC32.IsChecked) { "Visible" } else { "Collapsed" })
    Update-RowColors
})
$ChkMD5.Add_Click({
    $DgFiles.Columns[3].Visibility = $(if ($ChkMD5.IsChecked) { "Visible" } else { "Collapsed" })
    Update-RowColors
})
$ChkSHA1.Add_Click({
    $DgFiles.Columns[4].Visibility = $(if ($ChkSHA1.IsChecked) { "Visible" } else { "Collapsed" })
    Update-RowColors
})
$ChkSHA256.Add_Click({
    $DgFiles.Columns[5].Visibility = $(if ($ChkSHA256.IsChecked) { "Visible" } else { "Collapsed" })
    Update-RowColors
})

# コンテキストメニュー
$CtxCopy.Add_Click({
    if ($DgFiles.SelectedItems.Count -eq 0) { return }
    $text = $DgFiles.SelectedItems | ForEach-Object {
        "{0}`t{1}`t{2}`t{3}`t{4}`t{5}" -f $_.FileName,$_.SizeRaw,$_.CRC32,$_.MD5,$_.SHA1,$_.SHA256
    } | Out-String
    [System.Windows.Clipboard]::SetText($text)
})
$CtxSelectAll.Add_Click({ $DgFiles.SelectAll() })
$CtxClearSelection.Add_Click({ $DgFiles.UnselectAll() })

$DgFiles.AllowDrop = $true
$DgFiles.Add_Drop({
    param($sender,$e)

    if ($e.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) {

        $paths = $e.Data.GetData([Windows.DataFormats]::FileDrop)

        $expanded = foreach ($path in $paths) {

            if (Test-Path -LiteralPath $path -PathType Leaf) {
                $path
            }
            elseif (Test-Path -LiteralPath $path -PathType Container) {
                Get-ChildItem -LiteralPath $path -File -Recurse | ForEach-Object {
                    $_.FullName
                }
            }
        }

        Add-Files -Paths $expanded
        Recalc-Hashes -OnlyMissing:$false
    }
})

# -------------------------
# 表示
# -------------------------
$window.ShowDialog() | Out-Null