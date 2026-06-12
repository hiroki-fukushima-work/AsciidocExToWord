# PlantUML簡易編集ツール
# 前提:
#   - この .ps1 と同じフォルダ配下に lib\plantuml-1.2025.4.jar を置く
#   - Java が実行できること: java -version
#
# 使い方:
#   powershell -ExecutionPolicy Bypass -File .\PlantUmlEditor.ps1
#
# ショートカット:
#   Ctrl + Enter : プレビュー更新
#   Ctrl + O     : 開く
#   Ctrl + N     : 新規
#   Ctrl + S     : 保存
#   Ctrl + Shift + S : 名前を付けて保存

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$ErrorActionPreference = "Stop"

$script:BaseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:PlantUmlJar = Join-Path $script:BaseDir "lib\plantuml-1.2025.4.jar"
$script:CurrentFile = $null
$script:IsDirty = $false
$script:TempDir = Join-Path ([System.IO.Path]::GetTempPath()) "PlantUmlSimpleEditor"
$script:LastImagePath = $null

if (-not (Test-Path $script:TempDir)) {
    New-Item -ItemType Directory -Path $script:TempDir | Out-Null
}

function Test-JavaAvailable {
    try {
        $p = Start-Process -FilePath "java" -ArgumentList "-version" -NoNewWindow -PassThru -Wait -RedirectStandardError (Join-Path $script:TempDir "java_version.err") -RedirectStandardOutput (Join-Path $script:TempDir "java_version.out")
        return $true
    } catch {
        return $false
    }
}

function Get-EditorText {
    return $TextEditor.Text
}

function Set-WindowTitle {
    $name = if ($script:CurrentFile) { [System.IO.Path]::GetFileName($script:CurrentFile) } else { "無題" }
    $dirty = if ($script:IsDirty) { " *" } else { "" }
    $Window.Title = "PlantUML簡易編集ツール - $name$dirty"
}

function Set-Status([string]$message) {
    $StatusText.Text = $message
}

function Confirm-DiscardChanges {
    if (-not $script:IsDirty) { return $true }

    $result = [System.Windows.MessageBox]::Show(
        "未保存の変更があります。保存しますか？",
        "確認",
        [System.Windows.MessageBoxButton]::YesNoCancel,
        [System.Windows.MessageBoxImage]::Question
    )

    switch ($result) {
        "Yes" { return Save-CurrentFile }
        "No" { return $true }
        default { return $false }
    }
}

function Save-CurrentFile {
    if (-not $script:CurrentFile) {
        return Save-CurrentFileAs
    }

    try {
        [System.IO.File]::WriteAllText($script:CurrentFile, (Get-EditorText), [System.Text.UTF8Encoding]::new($false))
        $script:IsDirty = $false
        Set-WindowTitle
        Set-Status "保存しました: $script:CurrentFile"
        return $true
    } catch {
        [System.Windows.MessageBox]::Show("保存に失敗しました。`n$($_.Exception.Message)", "エラー", "OK", "Error") | Out-Null
        return $false
    }
}

function Save-CurrentFileAs {
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Title = "PlantUMLファイルを保存"
    $dialog.Filter = "PlantUML / AsciiDoc (*.puml;*.plantuml;*.adoc;*.asciidoc;*.txt)|*.puml;*.plantuml;*.adoc;*.asciidoc;*.txt|All files (*.*)|*.*"
    $dialog.FileName = if ($script:CurrentFile) { [System.IO.Path]::GetFileName($script:CurrentFile) } else { "diagram.puml" }

    if ($dialog.ShowDialog() -eq $true) {
        $script:CurrentFile = $dialog.FileName
        return Save-CurrentFile
    }
    return $false
}

function New-Document {
    if (-not (Confirm-DiscardChanges)) { return }

    $script:CurrentFile = $null
    $TextEditor.Text = @"
@startuml
Alice -> Bob: Hello
Bob --> Alice: Hi!
@enduml
"@.Trim()
    $script:IsDirty = $false
    Set-WindowTitle
    Set-Status "新規作成しました。Ctrl+Enterでプレビュー更新できます。"
    Update-Preview
}

function Open-Document {
    if (-not (Confirm-DiscardChanges)) { return }

    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Title = "PlantUMLファイルを開く"
    $dialog.Filter = "PlantUML / AsciiDoc (*.puml;*.plantuml;*.adoc;*.asciidoc;*.txt)|*.puml;*.plantuml;*.adoc;*.asciidoc;*.txt|All files (*.*)|*.*"

    if ($dialog.ShowDialog() -eq $true) {
        try {
            $script:CurrentFile = $dialog.FileName
            $TextEditor.Text = [System.IO.File]::ReadAllText($script:CurrentFile, [System.Text.Encoding]::UTF8)
            $script:IsDirty = $false
            Set-WindowTitle
            Set-Status "開きました: $script:CurrentFile"
            Update-Preview
        } catch {
            [System.Windows.MessageBox]::Show("ファイルを開けませんでした。`n$($_.Exception.Message)", "エラー", "OK", "Error") | Out-Null
        }
    }
}

function Extract-PlantUmlFromAsciiDoc([string]$text) {
    # AsciiDoc の [plantuml] ---- ... ---- を優先して抽出。
    # 見つからなければ全文を PlantUML として扱う。
    $pattern = '(?ms)^\s*\[plantuml[^\]]*\]\s*\r?\n-{4,}\s*\r?\n(?<body>.*?)\r?\n-{4,}\s*$'
    $match = [regex]::Match($text, $pattern)
    if ($match.Success) {
        return $match.Groups['body'].Value.Trim()
    }
    return $text.Trim()
}

function Show-TextInPreview([string]$text) {
    $PreviewImage.Visibility = "Collapsed"
    $PreviewText.Visibility = "Visible"
    $PreviewText.Text = $text
}

function Show-ImageInPreview([string]$path) {
    # 同じファイルパスの画像を差し替えるとWPF側のキャッシュで古い画像が残ることがあるため、
    # キャッシュ無視 + Sourceクリアを明示する。
    $PreviewImage.Source = $null

    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $bitmap.BeginInit()
    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bitmap.CreateOptions = [System.Windows.Media.Imaging.BitmapCreateOptions]::IgnoreImageCache
    $bitmap.UriSource = [Uri]::new($path)
    $bitmap.EndInit()
    $bitmap.Freeze()

    $PreviewImage.Source = $bitmap
    $PreviewText.Visibility = "Collapsed"
    $PreviewImage.Visibility = "Visible"
}

function Update-Preview {
    try {
        if (-not (Test-Path $script:PlantUmlJar)) {
            Show-TextInPreview "PlantUML jar が見つかりません。`n`n期待パス:`n$script:PlantUmlJar"
            Set-Status "PlantUML jar が見つかりません。"
            return
        }

        if (-not (Test-JavaAvailable)) {
            Show-TextInPreview "Java が見つかりません。`njava コマンドを実行できるようにしてください。"
            Set-Status "Java が見つかりません。"
            return
        }

        $source = Extract-PlantUmlFromAsciiDoc (Get-EditorText)
        if ([string]::IsNullOrWhiteSpace($source)) {
            Show-TextInPreview "PlantUMLコードが空です。"
            Set-Status "プレビュー対象が空です。"
            return
        }
        # 更新ごとに別フォルダへ出力する。
        # これにより、WPFの画像キャッシュやPlantUML出力ファイルの取り違えを避ける。
        $renderId = [DateTimeOffset]::Now.ToUnixTimeMilliseconds().ToString()
        $renderDir = Join-Path $script:TempDir $renderId
        New-Item -ItemType Directory -Path $renderDir -Force | Out-Null

        $inputFile = Join-Path $renderDir "preview.puml"
        $stdoutFile = Join-Path $renderDir "plantuml_stdout.txt"
        $stderrFile = Join-Path $renderDir "plantuml_stderr.txt"

        [System.IO.File]::WriteAllText($inputFile, $source, [System.Text.UTF8Encoding]::new($false))
        $arguments = @(
            "-Djava.awt.headless=true",
            "-jar", "`"$script:PlantUmlJar`"",
            "-tpng",
            "`"$inputFile`""
        ) -join " "

        $process = Start-Process -FilePath "java" -ArgumentList $arguments -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile

        $generated = Join-Path $renderDir "preview.png"
        if (Test-Path $generated) {
            Show-ImageInPreview $generated
            if ($process.ExitCode -eq 0) {
                Set-Status "プレビューを更新しました: $([DateTime]::Now.ToString('HH:mm:ss'))"
            } else {
                Set-Status "PlantUMLがエラー画像を生成しました: $([DateTime]::Now.ToString('HH:mm:ss')) ExitCode=$($process.ExitCode)"
            }
            return
        }

        $stdout = if (Test-Path $stdoutFile) { [System.IO.File]::ReadAllText($stdoutFile) } else { "" }
        $stderr = if (Test-Path $stderrFile) { [System.IO.File]::ReadAllText($stderrFile) } else { "" }
        Show-TextInPreview "画像を生成できませんでした。`n`nExitCode: $($process.ExitCode)`n`nSTDOUT:`n$stdout`n`nSTDERR:`n$stderr"
        Set-Status "画像生成に失敗しました。"
    } catch {
        Show-TextInPreview "プレビュー更新中に例外が発生しました。`n`n$($_.Exception.Message)"
        Set-Status "プレビュー更新中に例外が発生しました。"
    }
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PlantUML簡易編集ツール"
        Width="1200"
        Height="800"
        WindowStartupLocation="CenterScreen">
    <DockPanel>
        <Menu DockPanel.Dock="Top">
            <MenuItem Header="_ファイル">
                <MenuItem x:Name="NewMenu" Header="_新規" InputGestureText="Ctrl+N" />
                <MenuItem x:Name="OpenMenu" Header="_開く..." InputGestureText="Ctrl+O" />
                <Separator />
                <MenuItem x:Name="SaveMenu" Header="_保存" InputGestureText="Ctrl+S" />
                <MenuItem x:Name="SaveAsMenu" Header="名前を付けて保存..." InputGestureText="Ctrl+Shift+S" />
                <Separator />
                <MenuItem x:Name="ExitMenu" Header="終了" />
            </MenuItem>
        </Menu>

        <StatusBar DockPanel.Dock="Bottom">
            <StatusBarItem>
                <TextBlock x:Name="StatusText" Text="準備完了" />
            </StatusBarItem>
        </StatusBar>

        <ToolBarTray DockPanel.Dock="Top">
            <ToolBar>
                <Button x:Name="NewButton" Content="新規" Margin="2" />
                <Button x:Name="OpenButton" Content="開く" Margin="2" />
                <Button x:Name="SaveButton" Content="保存" Margin="2" />
                <Separator />
                <Button x:Name="PreviewButton" Content="プレビュー更新  Ctrl+Enter" Margin="2" FontWeight="SemiBold" />
            </ToolBar>
        </ToolBarTray>

        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="5" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>

            <Grid Grid.Column="0" Margin="8">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="*" />
                </Grid.RowDefinitions>
                <TextBlock Text="エディター（AsciiDoc / PlantUML）" FontWeight="Bold" Margin="0,0,0,6" />
                <TextBox x:Name="TextEditor"
                         Grid.Row="1"
                         AcceptsReturn="True"
                         AcceptsTab="True"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Auto"
                         FontFamily="Consolas, MS Gothic"
                         FontSize="14"
                         TextWrapping="NoWrap" />
            </Grid>

            <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" />

            <Grid Grid.Column="2" Margin="8">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="*" />
                </Grid.RowDefinitions>
                <TextBlock Text="プレビュー" FontWeight="Bold" Margin="0,0,0,6" />
                <Border Grid.Row="1" BorderBrush="#CCCCCC" BorderThickness="1" Background="White">
                    <ScrollViewer HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto">
                        <Grid>
                            <Image x:Name="PreviewImage" Stretch="None" HorizontalAlignment="Left" VerticalAlignment="Top" />
                            <TextBox x:Name="PreviewText"
                                     Visibility="Collapsed"
                                     IsReadOnly="True"
                                     TextWrapping="Wrap"
                                     BorderThickness="0"
                                     Background="White"
                                     Foreground="DarkRed"
                                     FontFamily="Consolas, MS Gothic"
                                     FontSize="13"
                                     Padding="10" />
                        </Grid>
                    </ScrollViewer>
                </Border>
            </Grid>
        </Grid>
    </DockPanel>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

$TextEditor = $Window.FindName("TextEditor")
$PreviewImage = $Window.FindName("PreviewImage")
$PreviewText = $Window.FindName("PreviewText")
$StatusText = $Window.FindName("StatusText")

$NewMenu = $Window.FindName("NewMenu")
$OpenMenu = $Window.FindName("OpenMenu")
$SaveMenu = $Window.FindName("SaveMenu")
$SaveAsMenu = $Window.FindName("SaveAsMenu")
$ExitMenu = $Window.FindName("ExitMenu")
$NewButton = $Window.FindName("NewButton")
$OpenButton = $Window.FindName("OpenButton")
$SaveButton = $Window.FindName("SaveButton")
$PreviewButton = $Window.FindName("PreviewButton")

$NewMenu.Add_Click({ New-Document })
$OpenMenu.Add_Click({ Open-Document })
$SaveMenu.Add_Click({ Save-CurrentFile | Out-Null })
$SaveAsMenu.Add_Click({ Save-CurrentFileAs | Out-Null })
$ExitMenu.Add_Click({ $Window.Close() })

$NewButton.Add_Click({ New-Document })
$OpenButton.Add_Click({ Open-Document })
$SaveButton.Add_Click({ Save-CurrentFile | Out-Null })
$PreviewButton.Add_Click({ Update-Preview })

$TextEditor.Add_TextChanged({
    $script:IsDirty = $true
    Set-WindowTitle
})

$Window.Add_KeyDown({
    param($sender, $e)

    $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -eq [System.Windows.Input.ModifierKeys]::Control
    $shift = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift) -eq [System.Windows.Input.ModifierKeys]::Shift

    if ($ctrl -and $e.Key -eq [System.Windows.Input.Key]::Enter) {
        Update-Preview
        $e.Handled = $true
    } elseif ($ctrl -and -not $shift -and $e.Key -eq [System.Windows.Input.Key]::O) {
        Open-Document
        $e.Handled = $true
    } elseif ($ctrl -and -not $shift -and $e.Key -eq [System.Windows.Input.Key]::N) {
        New-Document
        $e.Handled = $true
    } elseif ($ctrl -and -not $shift -and $e.Key -eq [System.Windows.Input.Key]::S) {
        Save-CurrentFile | Out-Null
        $e.Handled = $true
    } elseif ($ctrl -and $shift -and $e.Key -eq [System.Windows.Input.Key]::S) {
        Save-CurrentFileAs | Out-Null
        $e.Handled = $true
    }
})

$Window.Add_Closing({
    param($sender, $e)
    if (-not (Confirm-DiscardChanges)) {
        $e.Cancel = $true
    }
})

New-Document
$Window.ShowDialog() | Out-Null
