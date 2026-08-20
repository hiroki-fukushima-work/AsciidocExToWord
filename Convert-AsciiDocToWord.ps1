[CmdletBinding()]

param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$AdocFullPath,
    [string]$OutputDir,
    [string]$ConfigFullPath = ".\conf\word-style.sample.json",
    [string]$TemplateFullPath = ".\Template\cover-template.docx",
    [switch]$Overwrite = $true,
    [switch]$TestMode
)

$script:DebugLogPath = Join-Path $PSScriptRoot 'convert-debug.log'
Remove-Item $script:DebugLogPath -ErrorAction SilentlyContinue

function Write-DebugLog {
    param([string]$Message)
    Add-Content -LiteralPath $script:DebugLogPath -Value $Message -Encoding UTF8
}

# 初期化
$context = @{
    Type        = $null          # 'number' / 'bullet' / 'text'
    BulletLevel = 0
}

function Set-HangingIndent {
    param(
        $Range,
        [float]$IndentWidth = 12
    )

    $pf = $Range.ParagraphFormat

    # 全体を右へ
    $pf.LeftIndent = $IndentWidth

    # 1行目だけ左へ戻す（ぶら下げ）
    $pf.FirstLineIndent = - $IndentWidth
}

function Resolve-PathFromScript {
    param (
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$BaseDir
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        # 絶対パス → そのまま
        return $Path
    }
    else {
        # 相対パス → スクリプト基準で結合
        return (Join-Path $BaseDir $Path)
    }
}

$baseDir = if ($PSScriptRoot) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

# デバッグ用: 動作確認時にフルパスを設定する（通常は空文字のまま）
$DebugAdocFullPath = ""

# インプットチェック
if (-not $TestMode) {
    if (-not [string]::IsNullOrWhiteSpace($DebugAdocFullPath)) {
        $AdocFullPath = @($DebugAdocFullPath)
    }

    if (-not $AdocFullPath -or $AdocFullPath.Count -eq 0) {
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = 'AsciiDoc/Markdown files (*.adoc;*.md;*.markdown)|*.adoc;*.md;*.markdown|All files (*.*)|*.*'
        $dialog.Title = '変換元ファイルを選択してください（複数選択可）'
        $dialog.Multiselect = $true
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $AdocFullPath = $dialog.FileNames
        }
        else {
            Write-Output 'ファイルが選択されませんでした。処理を中止します。'
            exit 0
        }
    }
}

if (-not [string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Resolve-PathFromScript -Path $OutputDir -BaseDir $baseDir
}

$ConfigFullPath = Resolve-PathFromScript -Path $ConfigFullPath -BaseDir $baseDir

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Script:ConverterVersion = 'ps51-v2-20260416'
$TIMESTAMP = Get-Date -Format "yyyy/MM/dd HH:mm:ss"

Write-Output "$TIMESTAMP Convert-AsciiDocToWord version: $Script:ConverterVersion"

function Convert-RgbToWordColor {
    param([string]$RgbHex)

    # RRGGBB → BBGGRR
    $r = $RgbHex.Substring(0, 2)
    $g = $RgbHex.Substring(2, 2)
    $b = $RgbHex.Substring(4, 2)

    return "$b$g$r"
}

function Remove-WrappingQuotes {
    param([string]$Value)

    if ($null -eq $Value) {
        return $Value
    }

    $result = ([string]$Value).Trim()
    while ($result.Length -ge 2) {
        $first = $result.Substring(0, 1)
        $last = $result.Substring($result.Length - 1, 1)
        $isDoubleQuoted = ($first -eq '"' -and $last -eq '"')
        $isSingleQuoted = ($first -eq "'" -and $last -eq "'")
        if (-not ($isDoubleQuoted -or $isSingleQuoted)) {
            break
        }
        $result = $result.Substring(1, $result.Length - 2).Trim()
    }

    return $result
}

function Get-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$BaseDirectory
    )

    if ($null -eq $Path) {
        throw 'パスが null です。'
    }

    $normalizedPath = Remove-WrappingQuotes -Value ([string]$Path)

    if ([string]::IsNullOrWhiteSpace($normalizedPath)) {
        throw 'パスが空文字です。'
    }

    try {
        if ([System.IO.Path]::IsPathRooted($normalizedPath)) {
            return [System.IO.Path]::GetFullPath($normalizedPath)
        }

        $normalizedBaseDirectory = $BaseDirectory
        if ([string]::IsNullOrWhiteSpace($normalizedBaseDirectory)) {
            $normalizedBaseDirectory = (Get-Location).Path
        }
        else {
            $normalizedBaseDirectory = Remove-WrappingQuotes -Value ([string]$normalizedBaseDirectory)
        }

        return [System.IO.Path]::GetFullPath((Join-Path $normalizedBaseDirectory $normalizedPath))
    }
    catch {
        throw "パスを解決できませんでした: '$Path' -> '$normalizedPath' (BaseDirectory='$BaseDirectory'). $($_.Exception.Message)"
    }
}

function Join-CommandLineArguments {
    param(
        [string[]]$Arguments
    )

    $escaped = foreach ($arg in $Arguments) {
        if ($null -eq $arg) {
            '""'
            continue
        }

        $s = [string]$arg

        if ($s -match '[\s"]') {
            '"' + ($s -replace '"', '\"') + '"'
        }
        else {
            $s
        }
    }

    return ($escaped -join ' ')
}

function Load-JsonConfig {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        $message = "設定ファイルが見つかりません: $($Path)"
        throw $message
    }

    return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

function New-Element {
    param(
        [string]$Type,
        [hashtable]$Data
    )

    $merged = @{ Type = $Type }
    foreach ($key in $Data.Keys) {
        $merged[$key] = $Data[$key]
    }
    return [pscustomobject]$merged
}

function Resolve-IncludePath {
    param(
        [string]$DirectivePath,
        [string]$CurrentFileDirectory
    )

    $normalized = $DirectivePath.Replace('/', [System.IO.Path]::DirectorySeparatorChar)
    return Get-AbsolutePath -Path $normalized -BaseDirectory $CurrentFileDirectory
}

function Merge-AttributeMaps {
    param(
        [hashtable]$Base,
        [hashtable]$Overlay
    )

    $result = @{}
    foreach ($key in $Base.Keys) { $result[$key] = $Base[$key] }
    foreach ($key in $Overlay.Keys) { $result[$key] = $Overlay[$key] }
    return $result
}

function Normalize-InlineText {
    param(
        [string]$Text,
        [hashtable]$Attributes
    )

    if ($null -eq $Text) { return '' }

    $value = $Text
    if ($Attributes) {
        foreach ($k in $Attributes.Keys) {
            $token = '{' + $k + '}'
            $value = $value.Replace($token, [string]$Attributes[$k])
        }
    }

    $value = [regex]::Replace($value, 'image::([^\[]+)\[(.*?)\]', '[画像: $1]')
    $value = [regex]::Replace($value, 'icon:[^\[]+\[(.*?)\]', '$1')
    $value = [regex]::Replace($value, '\[(?<role>[^\]]+)\]#(?<body>.*?)#', '${body}')
    #$value = [regex]::Replace($value, '\*([^*]+)\*', '$1')
    #$value = [regex]::Replace($value, '_([^_]+)_', '$1')
    $value = [regex]::Replace($value, '`([^`]+)`', '$1')
    $value = $value -replace '\s+\+$', ''
    $value = $value -replace '\+$', ''
    $value = $value -replace '&#169;', '©'
    return $value.TrimEnd()
}

function Convert-AnchorIdToBookmarkName {
    param(
        [string]$AnchorId,
        [hashtable]$UsedNames
    )

    $raw = if ([string]::IsNullOrWhiteSpace($AnchorId)) { 'anchor' } else { $AnchorId.Trim() }
    $normalized = [regex]::Replace($raw, '[^A-Za-z0-9_]', '_')

    if ([string]::IsNullOrWhiteSpace($normalized)) {
        $normalized = 'anchor'
    }

    if ($normalized -notmatch '^[A-Za-z]') {
        $normalized = 'a_' + $normalized
    }

    $base = 'adoc_' + $normalized
    if ($base.Length -gt 40) {
        $base = $base.Substring(0, 40)
    }

    if ($null -eq $UsedNames) {
        return $base
    }

    $candidate = $base
    $suffix = 1
    while ($UsedNames.ContainsKey($candidate)) {
        $suffixText = '_' + [string]$suffix
        $maxBaseLength = 40 - $suffixText.Length
        $trimmedBase = if ($base.Length -gt $maxBaseLength) { $base.Substring(0, $maxBaseLength) } else { $base }
        $candidate = $trimmedBase + $suffixText
        $suffix++
    }

    $UsedNames[$candidate] = $true
    return $candidate
}

function Add-WordBookmarkSafe {
    param(
        $Document,
        $Range,
        [string]$BookmarkName
    )

    if (-not $Document -or -not $Range -or [string]::IsNullOrWhiteSpace($BookmarkName)) {
        return $false
    }

    try {
        if ($Document.Bookmarks.Exists($BookmarkName)) {
            $Document.Bookmarks.Item($BookmarkName).Delete()
        }

        $Document.Bookmarks.Add($BookmarkName, $Range) | Out-Null
        return $true
    }
    catch {
        Write-DebugLog "BOOKMARK-ADD-FAILED name=[$BookmarkName] error=[$($_.Exception.Message)]"
        return $false
    }
}

function Resolve-InternalLinkBookmark {
    param(
        [string]$TargetId,
        [hashtable]$InternalLinkMap
    )

    if ([string]::IsNullOrWhiteSpace($TargetId) -or -not $InternalLinkMap) {
        return $null
    }

    if ($InternalLinkMap.ContainsKey($TargetId)) {
        return [string]$InternalLinkMap[$TargetId]
    }

    return $null
}

function Parse-InlineLinkSegments {
    param([string]$Text)

    $segments = New-Object System.Collections.Generic.List[object]

    if ($null -eq $Text) {
        return $segments
    }

    $pattern = 'link:(?<linkurl>[^\s\[]+)\[(?<linklabel>[^\]]*)\]|xref:(?<xrefid>[^\[\s]+)\[(?<xreflabel>[^\]]*)\]|<<(?<xid>[^,>]+)(?:,(?<xlabel>[^>]+))?>>|(?<rawurl>https?://[^\s\[]+)(?:\[(?<rawlabel>[^\]]+)\])?'
    $linkMatches = [regex]::Matches($Text, $pattern)
    $pos = 0

    foreach ($m in $linkMatches) {
        if ($m.Index -gt $pos) {
            $segments.Add([pscustomobject]@{
                    Type = 'text'
                    Text = $Text.Substring($pos, $m.Index - $pos)
                })
        }

        if ($m.Groups['linkurl'].Success) {
            $url = [string]$m.Groups['linkurl'].Value
            $label = [string]$m.Groups['linklabel'].Value
            if ([string]::IsNullOrWhiteSpace($label)) { $label = $url }

            $segments.Add([pscustomobject]@{
                    Type = 'external'
                    Url  = $url
                    Text = $label
                })
        }
        elseif ($m.Groups['xrefid'].Success) {
            $targetId = [string]$m.Groups['xrefid'].Value
            $label = [string]$m.Groups['xreflabel'].Value
            if ([string]::IsNullOrWhiteSpace($label)) { $label = $targetId }

            $segments.Add([pscustomobject]@{
                    Type     = 'internal'
                    TargetId = $targetId.Trim()
                    Text     = $label
                })
        }
        elseif ($m.Groups['xid'].Success) {
            $targetId = [string]$m.Groups['xid'].Value
            $label = [string]$m.Groups['xlabel'].Value
            if ([string]::IsNullOrWhiteSpace($label)) { $label = $targetId }

            $segments.Add([pscustomobject]@{
                    Type     = 'internal'
                    TargetId = $targetId.Trim()
                    Text     = $label
                })
        }
        elseif ($m.Groups['rawurl'].Success) {
            $url = [string]$m.Groups['rawurl'].Value
            $label = [string]$m.Groups['rawlabel'].Value
            if ([string]::IsNullOrWhiteSpace($label)) { $label = $url }

            $segments.Add([pscustomobject]@{
                    Type = 'external'
                    Url  = $url
                    Text = $label
                })
        }

        $pos = $m.Index + $m.Length
    }

    if ($pos -lt $Text.Length) {
        $segments.Add([pscustomobject]@{
                Type = 'text'
                Text = $Text.Substring($pos)
            })
    }

    return $segments
}

function Get-WordConstant {
    param([string]$Name)

    $map = @{
        wdCollapseEnd                    = 0
        wdPageBreak                      = 7
        wdHeaderFooterPrimary            = 1
        wdAlignParagraphLeft             = 0
        wdAlignParagraphCenter           = 1
        wdAlignParagraphRight            = 2
        wdAutoFitContent                 = 1
        wdCellAlignVerticalCenter        = 1
        wdLineStyleNone                  = 0
        wdListApplyToWholeList           = 0
        wdColorAutomatic                 = -16777216
        wdSaveFormatDocumentDefault      = 16
        wdWord9ListBehavior              = 1
        msoShapeRoundedRectangle         = 5
        msoTextOrientationHorizontal     = 1
        wdSectionBreakNextPage           = 2
        wdFieldPage                      = 33
        wdFieldSectionPages              = 87
        wdVerticalPositionRelativeToPage = 3
    }

    if (-not $map.ContainsKey($Name)) { 
        throw "未定義のWord定数です: $Name"
    }
    return $map[$Name]
}

function Apply-FontStyle {
    param(
        $Range,
        $StyleConfig
    )

    if ($null -eq $StyleConfig) { return }

    if ($StyleConfig.PSObject.Properties.Name -contains 'FontName' -and $StyleConfig.FontName) {
        $fontName = [string]$StyleConfig.FontName

        if (-not [string]::IsNullOrWhiteSpace($fontName)) {
            try { $Range.Font.Name = $fontName } catch {}
            try { $Range.Font.NameFarEast = $fontName } catch {}
            try { $Range.Font.NameAscii = $fontName } catch {}
            try { $Range.Font.NameOther = $fontName } catch {}
        }
    }
    
    if ($StyleConfig.PSObject.Properties.Name -contains 'Size' -and $StyleConfig.Size) {
        $Range.Font.Size = [double]$StyleConfig.Size
    }
    if ($StyleConfig.PSObject.Properties.Name -contains 'Bold') {
        $Range.Font.Bold = [int]([bool]$StyleConfig.Bold)
    }
    if ($StyleConfig.PSObject.Properties.Name -contains 'Italic') {
        $Range.Font.Italic = [int]([bool]$StyleConfig.Italic)
    }
    if ($StyleConfig.PSObject.Properties.Name -contains 'Underline' -and [bool]$StyleConfig.Underline) {
        $Range.Font.Underline = 1
    }
    if ($StyleConfig.PSObject.Properties.Name -contains 'Color') {
        try { $Range.Font.Color = [int]$StyleConfig.Color } catch {}
    }

    if ($StyleConfig.PSObject.Properties.Name -contains 'Alignment' -and $StyleConfig.Alignment) {
        switch (([string]$StyleConfig.Alignment).ToLowerInvariant()) {
            'left' { $Range.ParagraphFormat.Alignment = Get-WordConstant 'wdAlignParagraphLeft' }
            'center' { $Range.ParagraphFormat.Alignment = Get-WordConstant 'wdAlignParagraphCenter' }
            'right' { $Range.ParagraphFormat.Alignment = Get-WordConstant 'wdAlignParagraphRight' }
        }
    }

    if ($StyleConfig.PSObject.Properties.Name -contains 'SpaceBefore') {
        $Range.ParagraphFormat.SpaceBefore = [double]$StyleConfig.SpaceBefore
    }
    if ($StyleConfig.PSObject.Properties.Name -contains 'SpaceAfter') {
        $Range.ParagraphFormat.SpaceAfter = [double]$StyleConfig.SpaceAfter
    }
    if ($StyleConfig.PSObject.Properties.Name -contains 'LeftIndent') {
        $Range.ParagraphFormat.LeftIndent = [double]$StyleConfig.LeftIndent
    }
    if ($StyleConfig.PSObject.Properties.Name -contains 'FirstLineIndent') {
        $Range.ParagraphFormat.FirstLineIndent = [double]$StyleConfig.FirstLineIndent
    }

    if ($StyleConfig.PSObject.Properties.Name -contains 'BackgroundColor' -and
        $StyleConfig.BackgroundColor) {

        # "E6E6E6" → 0x00E6E6E6
        $color = '0x00' + $StyleConfig.BackgroundColor
        $Range.Shading.BackgroundPatternColor = [int]$color
    }
}

function Apply-ParagraphStyle {
    param(
        $Range,
        $StyleConfig
    )

    if (-not $Range -or -not $StyleConfig) {
        return
    }

    try {
        if ($StyleConfig.PSObject.Properties.Name -contains 'Alignment' -and $StyleConfig.Alignment) {
            switch ([string]$StyleConfig.Alignment) {
                'Left' { $Range.ParagraphFormat.Alignment = 0 }
                'Center' { $Range.ParagraphFormat.Alignment = 1 }
                'Right' { $Range.ParagraphFormat.Alignment = 2 }
                'Justify' { $Range.ParagraphFormat.Alignment = 3 }
            }
        }
    }
    catch {
    }

    try {
        if ($StyleConfig.PSObject.Properties.Name -contains 'KeepWithNext') {
            $Range.ParagraphFormat.KeepWithNext = [bool]$StyleConfig.KeepWithNext
        }
    }
    catch {
    }

    try {
        if ($StyleConfig.PSObject.Properties.Name -contains 'KeepTogether') {
            $Range.ParagraphFormat.KeepTogether = [bool]$StyleConfig.KeepTogether
        }
    }
    catch {
    }

    try {
        if ($StyleConfig.PSObject.Properties.Name -contains 'PageBreakBefore') {
            $Range.ParagraphFormat.PageBreakBefore = [bool]$StyleConfig.PageBreakBefore
        }
    }
    catch {
    }
    
    try {
        if ($StyleConfig.PSObject.Properties.Name -contains 'LineSpacingRule' -and
            $StyleConfig.LineSpacingRule -eq 'Exactly') {

            # Exactly（固定行間）
            $Range.ParagraphFormat.LineSpacingRule = 2   # wdLineSpaceExactly
            $Range.ParagraphFormat.LineSpacing = [float]$StyleConfig.LineSpacing
        }
    }
    catch {
    }

    try {
        # 段落前後の余白（省略時は 0）
        if ($StyleConfig.PSObject.Properties.Name -contains 'SpaceBefore') {
            $Range.ParagraphFormat.SpaceBefore = [float]$StyleConfig.SpaceBefore
        }
        else {
            $Range.ParagraphFormat.SpaceBefore = 0
        }

        if ($StyleConfig.PSObject.Properties.Name -contains 'SpaceAfter') {
            $Range.ParagraphFormat.SpaceAfter = [float]$StyleConfig.SpaceAfter
        }
        else {
            $Range.ParagraphFormat.SpaceAfter = 0
        }
    }
    catch {
    }

}
function Format-AdmonitionText {
    param(
        [string]$Kind,
        [string]$Text,
        $LabelConfig
    )

    switch ($Kind) {
        'IMPORTANT' { return "$($LabelConfig.Important) $Text" }
        'WARNING' { return "$($LabelConfig.Warning) $Text" }
        'CAUTION' { return "$($LabelConfig.Caution) $Text" }
        'NOTE' { return "$($LabelConfig.Note) $Text" }
        'TIP' { return "$($LabelConfig.Tip) $Text" }
        default { return "[$Kind] $Text" }
    }
}
function Parse-InlineText {
    param([string]$Text)

    $runs = @()

    $pattern = '(\*\*_.*?_\*\*|\*\*.*?\*\*|_.*?_)'

    $parts = [regex]::Split($Text, $pattern)

    foreach ($p in $parts) {

        if ($p -match '^\*\*_(.+)_\*\*$') {
            $runs += @{
                Text   = $matches[1]
                Bold   = $true
                Italic = $true
            }
        }
        elseif ($p -match '^\*\*(.+)\*\*$') {
            $runs += @{
                Text   = $matches[1]
                Bold   = $true
                Italic = $false
            }
        }
        elseif ($p -match '^_(.+)_$') {
            $runs += @{
                Text   = $matches[1]
                Bold   = $false
                Italic = $true
            }
        }
        else {
            $runs += @{
                Text   = $p
                Bold   = $false
                Italic = $false
            }
        }
    }

    return $runs
}

function Test-InlineStartBoundary {
    param(
        [string]$Text,
        [int]$Index
    )

    if ($Index -eq 0) { return $true }

    $prev = $Text.Substring($Index - 1, 1)
    return [char]::IsWhiteSpace($prev[0])
}

function Convert-InlineEmphasis {
    param(
        [string]$Text
    )

    $runs = New-Object System.Collections.Generic.List[object]

    $patterns = @(
        @{ Open = '**__'; Close = '__**'; Bold = $true; Italic = $true; NeedBoundary = $false },
        @{ Open = '__**'; Close = '**__'; Bold = $true; Italic = $true; NeedBoundary = $false },
        @{ Open = '*_'; Close = '_*'; Bold = $true; Italic = $true; NeedBoundary = $true },
        @{ Open = '_*'; Close = '*_'; Bold = $true; Italic = $true; NeedBoundary = $true },
        @{ Open = '**'; Close = '**'; Bold = $true; Italic = $false; NeedBoundary = $false },
        @{ Open = '__'; Close = '__'; Bold = $false; Italic = $true; NeedBoundary = $false },
        @{ Open = '*'; Close = '*'; Bold = $true; Italic = $false; NeedBoundary = $true },
        @{ Open = '_'; Close = '_'; Bold = $false; Italic = $true; NeedBoundary = $true }
    )

    function Add-Run {
        param(
            [string]$Value,
            [bool]$Bold,
            [bool]$Italic
        )

        if ([string]::IsNullOrEmpty($Value)) { return }

        $runs.Add([pscustomobject]@{
                Text   = $Value
                Bold   = $Bold
                Italic = $Italic
            })
    }

    $plain = New-Object System.Text.StringBuilder
    $i = 0

    while ($i -lt $Text.Length) {

        $matched = $null

        foreach ($p in $patterns) {
            $open = [string]$p.Open
            $close = [string]$p.Close

            if (($i + $open.Length) -gt $Text.Length) {
                continue
            }

            if ($Text.Substring($i, $open.Length) -ne $open) {
                continue
            }

            if ($p.NeedBoundary -and
                -not (Test-InlineStartBoundary -Text $Text -Index $i)) {
                continue
            }

            $closeIndex = $Text.IndexOf($close, $i + $open.Length)

            if ($closeIndex -lt 0) {
                continue
            }

            $matched = @{
                Pattern    = $p
                CloseIndex = $closeIndex
            }
            break
        }

        if ($matched) {
            Add-Run -Value $plain.ToString() -Bold $false -Italic $false
            [void]$plain.Clear()

            $p = $matched.Pattern
            $open = [string]$p.Open
            $close = [string]$p.Close
            $closeIndex = [int]$matched.CloseIndex

            $bodyStart = $i + $open.Length
            $bodyLength = $closeIndex - $bodyStart
            $body = $Text.Substring($bodyStart, $bodyLength)

            Add-Run `
                -Value $body `
                -Bold ([bool]$p.Bold) `
                -Italic ([bool]$p.Italic)

            $i = $closeIndex + $close.Length
            continue
        }

        [void]$plain.Append($Text.Substring($i, 1))
        $i++
    }

    Add-Run -Value $plain.ToString() -Bold $false -Italic $false

    return $runs
}

function Append-BlankParagraph {
    param($Document)
    
    Append-TextParagraph -Document $Document -Text '' -StyleConfig $null | Out-Null
}
function Append-TextParagraph {
    param(
        $Document,
        [string]$Text,
        $StyleConfig,
        [hashtable]$InternalLinkMap,
        [switch]$NoTrailingParagraph
    )

    $cursor = $Document.Content
    $cursor.Collapse((Get-WordConstant 'wdCollapseEnd'))

    $paragraphStart = $cursor.Start
    $inlineRanges = @()

    $segments = Parse-InlineLinkSegments -Text $Text

    foreach ($segment in $segments) {

        if ($segment.Type -eq 'text') {
            $runs = Convert-InlineEmphasis -Text ([string]$segment.Text)

            foreach ($run in $runs) {
                $runText = [string]$run.Text
                if ([string]::IsNullOrEmpty($runText)) { continue }

                $runRange = $Document.Range($cursor.End, $cursor.End)
                $runRange.Text = $runText

                $inlineRanges += [pscustomobject]@{
                    Range  = $runRange
                    Bold   = [bool]$run.Bold
                    Italic = [bool]$run.Italic
                }

                $cursor.SetRange($runRange.End, $runRange.End)
            }

            continue
        }

        if ($segment.Type -eq 'external') {
            $url = [string]$segment.Url
            $label = [string]$segment.Text

            try {
                $linkRange = $Document.Range($cursor.End, $cursor.End)

                $hyperlink = $Document.Hyperlinks.Add(
                    $linkRange,
                    $url,
                    '',
                    '',
                    $label
                )

                $cursor.SetRange($hyperlink.Range.End, $hyperlink.Range.End)
            }
            catch {
                $linkRange = $Document.Range($cursor.End, $cursor.End)
                $linkRange.Text = $label
                $cursor.SetRange($linkRange.End, $linkRange.End)
            }

            continue
        }

        if ($segment.Type -eq 'internal') {
            $targetId = [string]$segment.TargetId
            $label = [string]$segment.Text
            $bookmark = Resolve-InternalLinkBookmark -TargetId $targetId -InternalLinkMap $InternalLinkMap

            if ([string]::IsNullOrWhiteSpace($bookmark)) {
                Write-DebugLog "INTERNAL-LINK-NOT-FOUND id=[$targetId]"
                $plainRange = $Document.Range($cursor.End, $cursor.End)
                $plainRange.Text = $label
                $cursor.SetRange($plainRange.End, $plainRange.End)
                continue
            }

            try {
                $linkRange = $Document.Range($cursor.End, $cursor.End)

                $hyperlink = $Document.Hyperlinks.Add(
                    $linkRange,
                    '',
                    $bookmark,
                    '',
                    $label
                )

                $cursor.SetRange($hyperlink.Range.End, $hyperlink.Range.End)
            }
            catch {
                Write-DebugLog "INTERNAL-LINK-ADD-FAILED id=[$targetId] bookmark=[$bookmark] error=[$($_.Exception.Message)]"
                $plainRange = $Document.Range($cursor.End, $cursor.End)
                $plainRange.Text = $label
                $cursor.SetRange($plainRange.End, $plainRange.End)
            }

            continue
        }
    }

    if (-not $NoTrailingParagraph) {
        $cursor.InsertParagraphAfter()
        $cursor.Collapse((Get-WordConstant 'wdCollapseEnd'))
    }

    $paragraphRange = $Document.Range($paragraphStart, $cursor.End)

    # 段落全体にベーススタイルを適用
    Apply-FontStyle -Range $paragraphRange -StyleConfig $StyleConfig
    Apply-ParagraphStyle -Range $paragraphRange -StyleConfig $StyleConfig

    # Bold / Italic を後から再適用
    foreach ($item in $inlineRanges) {
        if ($item.Bold) {
            $item.Range.Font.Bold = 1
        }

        if ($item.Italic) {
            $item.Range.Font.Italic = 1
        }
    }

    return $paragraphRange
}

function Set-TableCellTextWithHyperlinks {
    param(
        $Document,
        $Cell,
        [string]$Text,
        $StyleConfig,
        [hashtable]$InternalLinkMap
    )

    $Text = $Text -replace "\\n", "`r"

    $cellTextRange = $Cell.Range.Duplicate
    $cellTextRange.MoveEnd(1, -1) | Out-Null
    $cellTextRange.Text = ''

    $cursor = $Document.Range($cellTextRange.Start, $cellTextRange.Start)
    $segments = Parse-InlineLinkSegments -Text $Text

    foreach ($segment in $segments) {
        if ($segment.Type -eq 'text') {
            $plain = [string]$segment.Text
            if (-not [string]::IsNullOrEmpty($plain)) {
                $r = $Document.Range($cursor.End, $cursor.End)
                $r.Text = $plain
                $cursor.SetRange($r.End, $r.End)
            }
            continue
        }

        if ($segment.Type -eq 'external') {
            $url = [string]$segment.Url
            $label = [string]$segment.Text

            try {
                $linkRange = $Document.Range($cursor.End, $cursor.End)

                $hyperlink = $Document.Hyperlinks.Add(
                    $linkRange,
                    $url,
                    '',
                    '',
                    $label
                )

                $cursor.SetRange($hyperlink.Range.End, $hyperlink.Range.End)
            }
            catch {
                $r = $Document.Range($cursor.End, $cursor.End)
                $r.Text = $label
                $cursor.SetRange($r.End, $r.End)
            }

            continue
        }

        if ($segment.Type -eq 'internal') {
            $targetId = [string]$segment.TargetId
            $label = [string]$segment.Text
            $bookmark = Resolve-InternalLinkBookmark -TargetId $targetId -InternalLinkMap $InternalLinkMap

            if ([string]::IsNullOrWhiteSpace($bookmark)) {
                Write-DebugLog "INTERNAL-LINK-NOT-FOUND id=[$targetId]"
                $r = $Document.Range($cursor.End, $cursor.End)
                $r.Text = $label
                $cursor.SetRange($r.End, $r.End)
                continue
            }

            try {
                $linkRange = $Document.Range($cursor.End, $cursor.End)

                $hyperlink = $Document.Hyperlinks.Add(
                    $linkRange,
                    '',
                    $bookmark,
                    '',
                    $label
                )

                $cursor.SetRange($hyperlink.Range.End, $hyperlink.Range.End)
            }
            catch {
                Write-DebugLog "INTERNAL-LINK-ADD-FAILED id=[$targetId] bookmark=[$bookmark] error=[$($_.Exception.Message)]"
                $r = $Document.Range($cursor.End, $cursor.End)
                $r.Text = $label
                $cursor.SetRange($r.End, $r.End)
            }

            continue
        }
    }

    $cellRange = $Cell.Range
    $cellRange.End -= 1

    Apply-FontStyle -Range $cellRange -StyleConfig $StyleConfig
    Apply-ParagraphStyle -Range $cellRange -StyleConfig $StyleConfig
}


function Edit-CoverPage {
    param(
        $Document,
        $Metadata,
        $Config
    )

    $replaceMap = @{
        '%%タイトル%%'   = [string]$Metadata.Title
        '%%サブタイトル%%' = [string]$Metadata.Subtitle
        '%%版数%%'     = [string]$Metadata.RevNumber
        '%%改定日%%'    = [string]$Metadata.RevDate
        '%%作成者%%'    = [string]$Metadata.Author
        '%%著作権%%'    = [string]$Metadata.Copyright
    }

    $ranges = @()

    # 本文
    $ranges += $Document.Content


    foreach ($section in $Document.Sections) {

        foreach ($hf in $section.Headers) {
            if ($hf.Exists -and -not $hf.LinkToPrevious) {
                $ranges += $hf.Range
            }
        }

        foreach ($hf in $section.Footers) {
            if ($hf.Exists -and -not $hf.LinkToPrevious) {
                $ranges += $hf.Range
            }
        }
    }

    foreach ($range in $ranges) {
        foreach ($findText in $replaceMap.Keys) {

            $replaceText = $replaceMap[$findText]

            $find = $range.Find

            $find.ClearFormatting()
            $find.Replacement.ClearFormatting()

            $find.Text = $findText
            $find.Replacement.Text = $replaceText

            $find.Forward = $true
            $find.Wrap = 1
            $find.Format = $false
            $find.MatchWildcards = $false
            $find.MatchCase = $false
            $find.MatchWholeWord = $false

            $find.Execute(
                [ref]$findText,
                [ref]$false,
                [ref]$false,
                [ref]$false,
                [ref]$false,
                [ref]$false,
                [ref]$true,
                [ref]1,
                [ref]$false,
                [ref]$replaceText,
                [ref]2
            ) | Out-Null
        }
    }
}

function Add-CoverPage {
    param(
        $Document,
        $Config,
        $Metadata
    )

    $app = $Document.Application
    $pageWidth = $Document.PageSetup.PageWidth
    $pageHeight = $Document.PageSetup.PageHeight
    $leftMargin = $Document.PageSetup.LeftMargin
    $rightMargin = $Document.PageSetup.RightMargin

    # タイトル用 角丸テキストボックス
    $boxWidth = $app.MillimetersToPoints(140)
    $boxHeight = $app.MillimetersToPoints(45)
    $boxLeft = ($pageWidth - $boxWidth) / 2
    $boxTop = $app.MillimetersToPoints(90)

    $titleBox = $Document.Shapes.AddShape(
        5,          # msoShapeRoundedRectangle
        $boxLeft,
        $boxTop,
        $boxWidth,
        $boxHeight
    )

    $titleBox.Fill.Visible = $false
    $titleBox.Line.Visible = $true
    # 作成者情報をタイトル枠の右下に配置
    if ($Metadata.Author) {

        $authorWidth = $app.MillimetersToPoints(70)
        $authorHeight = $app.MillimetersToPoints(25)

        $authorLeft = $boxLeft + $boxWidth - $authorWidth
        $authorTop = $boxTop + $boxHeight + $app.MillimetersToPoints(5)

        $authorBox = $Document.Shapes.AddTextbox(
            1,
            $authorLeft,
            $authorTop,
            $authorWidth,
            $authorHeight
        )

        $authorBox.Line.Visible = $false
        $authorBox.Fill.Visible = $false

        $authorBox.TextFrame.MarginTop = 0
        $authorBox.TextFrame.MarginBottom = 0
        $authorBox.TextFrame.MarginLeft = 0
        $authorBox.TextFrame.MarginRight = 0

        $authorRange = $authorBox.TextFrame.TextRange
        $authorRange.Text = $Metadata.Author

        Apply-FontStyle -Range $authorRange -StyleConfig $Config.Styles.Owner
        Apply-ParagraphStyle -Range $authorRange -StyleConfig $Config.Styles.Owner

        $authorRange.ParagraphFormat.Alignment = 2  # 右寄せ
    }
    
    $textRange = $titleBox.TextFrame.TextRange
    $textRange.Text = $Metadata.Title + "`r" + $Metadata.Subtitle

    Apply-FontStyle -Range $textRange -StyleConfig $Config.Styles.Title
    Apply-ParagraphStyle -Range $textRange -StyleConfig $Config.Styles.Title

    $titleLine = $textRange.Paragraphs(1).Range
    Apply-FontStyle -Range $titleLine -StyleConfig $Config.Styles.Title
    Apply-ParagraphStyle -Range $titleLine -StyleConfig $Config.Styles.Title
    $titleLine.Font.Underline = 1

    $subLine = $textRange.Paragraphs(2).Range
    Apply-FontStyle -Range $subLine -StyleConfig $Config.Styles.Subtitle
    Apply-ParagraphStyle -Range $subLine -StyleConfig $Config.Styles.Subtitle
    $subLine.Font.Underline = 0

    # 右下情報ボックス
    $infoWidth = $app.MillimetersToPoints(70)
    $infoHeight = $app.MillimetersToPoints(25)

    $infoLeft = $pageWidth - $rightMargin - $infoWidth
    $infoTop = $pageHeight - $app.MillimetersToPoints(55)

    $infoBox = $Document.Shapes.AddTextbox(
        1,
        $infoLeft,
        $infoTop,
        $infoWidth,
        $infoHeight
    )

    # 枠なし
    $infoBox.Line.Visible = $false
    $infoBox.Fill.Visible = $false

    # 内部余白少し削る
    $infoBox.TextFrame.MarginTop = 0
    $infoBox.TextFrame.MarginBottom = 0
    $infoBox.TextFrame.MarginLeft = 0
    $infoBox.TextFrame.MarginRight = 0

    $r = $infoBox.TextFrame.TextRange
    $r.Text = ""

    # テーブル追加
    $tbl = $Document.Tables.Add($r, 2, 2)

    # 罫線
    $wdBorderLeft = 1
    $wdBorderTop = 2
    $wdBorderBottom = 4
    $wdBorderRight = 3
    
    $tbl.Borders.Enable = 0
    
    $tbl.Borders.Item($wdBorderLeft).LineStyle = 1
    $tbl.Borders.Item($wdBorderTop).LineStyle = 1
    $tbl.Borders.Item($wdBorderBottom).LineStyle = 1
    $tbl.Borders.Item($wdBorderRight).LineStyle = 1

    # 列幅調整
    $tbl.Columns.Item(1).Width = $app.MillimetersToPoints(18)
    $tbl.Columns.Item(2).Width = $app.MillimetersToPoints(40)

    # 値設定
    $tbl.Cell(1, 1).Range.Text = "版数"
    $tbl.Cell(1, 2).Range.Text = $Metadata.RevNumber

    $tbl.Cell(2, 1).Range.Text = "改定日"
    $tbl.Cell(2, 2).Range.Text = $Metadata.RevDate

    # スタイル適用
    for ($row = 1; $row -le 2; $row++) {

        Apply-FontStyle `
            -Range $tbl.Cell($row, 1).Range `
            -StyleConfig $Config.Styles.Owner

        Apply-ParagraphStyle `
            -Range $tbl.Cell($row, 1).Range `
            -StyleConfig $Config.Styles.Owner

        Apply-FontStyle `
            -Range $tbl.Cell($row, 2).Range `
            -StyleConfig $Config.Styles.Revision

        Apply-ParagraphStyle `
            -Range $tbl.Cell($row, 2).Range `
            -StyleConfig $Config.Styles.Revision
    }
}

function Add-SectionBreakToDocument {
    param($Document)

    $range = $Document.Content
    $range.Collapse((Get-WordConstant 'wdCollapseEnd'))
    $range.InsertBreak((Get-WordConstant 'wdSectionBreakNextPage'))
}

function Set-BodyFooterPageNumber {
    param($Document)

    $section = $Document.Sections.Item($Document.Sections.Count)

    $footer = $section.Footers.Item((Get-WordConstant 'wdHeaderFooterPrimary'))
    $footer.LinkToPrevious = $false

    $section.PageSetup.DifferentFirstPageHeaderFooter = $false
    $footer.PageNumbers.RestartNumberingAtSection = $true
    $footer.PageNumbers.StartingNumber = 1

    $range = $footer.Range
    $range.Text = ''
    $range.ParagraphFormat.Alignment = 1

    # ✅ まずプレースホルダ文字列を作る
    $range.Text = "0/0"

    # ✅ 1文字目をPAGEに
    $char1 = $range.Characters.Item(1)
    $Document.Fields.Add($char1, -1, 'PAGE', $true) | Out-Null

    # ✅ 3文字目をSECTIONPAGESに
    $char3 = $range.Characters.Item(3)
    $Document.Fields.Add($char3, -1, 'SECTIONPAGES', $true) | Out-Null

    # 更新
    $Document.Repaginate() | Out-Null
    $Document.Fields.Update() | Out-Null
}



function Add-PageBreakToDocument {
    param($Document)

    try {
        $range = $Document.Range()
        $range.Collapse((Get-WordConstant 'wdCollapseEnd'))
        # [void]$range.InsertParagraphAfter()
        # $range.Collapse((Get-WordConstant 'wdCollapseEnd'))
        [void]$range.InsertBreak((Get-WordConstant 'wdPageBreak'))
        # [void]$range.InsertParagraphAfter()
    }
    catch {
        Append-TextParagraph -Document $Document -Text '[改ページ挿入失敗]' -StyleConfig $null | Out-Null
    }
}

function Add-PageBreakToDocument {
    param($Document)

    try {
        $range = $Document.Range()
        $range.Collapse((Get-WordConstant 'wdCollapseEnd'))

        # ★これだけでOK
        [void]$range.InsertBreak((Get-WordConstant 'wdPageBreak'))
    }
    catch {
        Append-TextParagraph -Document $Document -Text '[改ページ挿入失敗]' -StyleConfig $null | Out-Null
    }
}

function Append-HeadingParagraph {
    param(
        $Document,
        [string]$Text,
        [int]$Level,
        $StyleConfig
    )
    

    $range = $Document.Content
    $range.Collapse((Get-WordConstant 'wdCollapseEnd'))
    $range.Text = $Text + [Environment]::NewLine

    try {
        $styleNameJa = '見出し ' + [string]$Level
        $range.Style = $styleNameJa
    }
    catch {
        try {
            $styleNameEn = 'Heading ' + [string]$Level
            $range.Style = $styleNameEn
        }
        catch {
        }
    }

    Apply-FontStyle -Range $range -StyleConfig $StyleConfig
    Apply-ParagraphStyle -Range $range -StyleConfig $StyleConfig

    return $range
}

function Add-TableOfContents {
    param($Document, $Config)

    Append-HeadingParagraph -Document $Document -Text '目次' -Level 1 -StyleConfig $Config.Styles.Heading1 | Out-Null

    $range = $Document.Content
    $range.Collapse((Get-WordConstant 'wdCollapseEnd'))

    $Document.TablesOfContents.Add($range, $true, 1, 3) | Out-Null
            
    $section = $Document.Sections[$Document.Sections.Count]

    foreach ($hf in $section.Headers + $section.Footers) {
        $hf.LinkToPrevious = $false
    }

    # 明示的にフッダーを空にする
    foreach ($footer in $section.Footers) {
        $footer.Range.Text = ""
    }

}

function Get-ColumnCountFromColsAttribute {
    param(
        [hashtable]$Attributes
    )

    if (-not $Attributes -or -not $Attributes.ContainsKey('cols')) {
        return 0
    }

    $cols = [string]$Attributes['cols']

    if ([string]::IsNullOrWhiteSpace($cols)) {
        return 0
    }

    $parts = $cols -split '\s*,\s*'
    return $parts.Count
}

function Convert-TableRows {
    param(
        [string[]]$Lines,
        [int]$ExpectedColumns = 0
    )

    $rows = @()
    $maxColumns = 0
    $currentRow = @()
    $currentCols = 0

    function Get-TableCellsFromLine {
        param([string]$Line)

        $cells = @()

        # |セル1|セル2|セル3 のような1行複数セルに対応
        $pattern = '(?<rs>\d+\.?)?(?<cs>\d+\+)?(?<header>h)?\|(?<text>[^|]*)'

        foreach ($m in [regex]::Matches($Line, $pattern)) {
            if (-not $m.Success) { continue }

            $rowSpan = 1
            if ($m.Groups['rs'].Success -and $m.Groups['rs'].Value) {
                $rowSpan = [int]($m.Groups['rs'].Value -replace '\.', '')
            }

            $colSpan = 1
            if ($m.Groups['cs'].Success -and $m.Groups['cs'].Value) {
                $colSpan = [int]($m.Groups['cs'].Value -replace '\+', '')
            }

            $cells += @{
                Text     = $m.Groups['text'].Value.TrimEnd()
                RowSpan  = $rowSpan
                ColSpan  = $colSpan
                IsHeader = ($m.Groups['header'].Value -eq 'h')
            }
        }

        return $cells
    }

    function Flush-CurrentRow {
        param(
            [ref]$Rows,
            [ref]$CurrentRow,
            [ref]$CurrentCols,
            [ref]$MaxColumns
        )

        if ($CurrentRow.Value.Count -eq 0) {
            return
        }

        $Rows.Value += , $CurrentRow.Value

        if ($CurrentCols.Value -gt $MaxColumns.Value) {
            $MaxColumns.Value = $CurrentCols.Value
        }

        $CurrentRow.Value = @()
        $CurrentCols.Value = 0
    }

    foreach ($line in $Lines) {
        $trimmed = ([string]$line).Trim()

        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            Flush-CurrentRow `
                -Rows ([ref]$rows) `
                -CurrentRow ([ref]$currentRow) `
                -CurrentCols ([ref]$currentCols) `
                -MaxColumns ([ref]$maxColumns)
            continue
        }

        $cells = Get-TableCellsFromLine -Line $trimmed

        foreach ($cell in $cells) {
            $currentRow += $cell
            $currentCols += [int]$cell.ColSpan

            if ($ExpectedColumns -gt 0 -and $currentCols -ge $ExpectedColumns) {
                Flush-CurrentRow `
                    -Rows ([ref]$rows) `
                    -CurrentRow ([ref]$currentRow) `
                    -CurrentCols ([ref]$currentCols) `
                    -MaxColumns ([ref]$maxColumns)
            }
        }
    }

    Flush-CurrentRow `
        -Rows ([ref]$rows) `
        -CurrentRow ([ref]$currentRow) `
        -CurrentCols ([ref]$currentCols) `
        -MaxColumns ([ref]$maxColumns)

    if ($ExpectedColumns -gt 0) {
        $maxColumns = $ExpectedColumns
    }

    return [pscustomobject]@{
        Rows       = $rows
        MaxColumns = $maxColumns
    }
}

function Replace-WordText {
    param(
        $Document,
        [string]$FindText,
        [string]$ReplaceText
    )

    $find = $Document.Content.Find
    $find.ClearFormatting()
    $find.Replacement.ClearFormatting()

    $find.Text = $FindText
    $find.Replacement.Text = $ReplaceText
    $find.Forward = $true
    $find.Wrap = 1 # wdFindContinue
    $find.Format = $false
    $find.MatchCase = $false
    $find.MatchWholeWord = $false
    $find.MatchWildcards = $false

    $find.Execute(
        [ref]$FindText,
        [ref]$false,
        [ref]$false,
        [ref]$false,
        [ref]$false,
        [ref]$false,
        [ref]$true,
        [ref]1,
        [ref]$false,
        [ref]$ReplaceText,
        [ref]2
    ) | Out-Null
}

function Add-WordTable {
    param(
        $Document,
        [object[]]$Rows,
        [int]$ColumnCount,
        $Config,
        [string]$Caption,
        $Attributes,
        [hashtable]$InternalLinkMap
    )

    if ($Caption) {
        $captionParagraph = Append-TextParagraph -Document $Document -Text $Caption -StyleConfig $Config.Styles.FigureCaption -InternalLinkMap $InternalLinkMap
        try {
            $captionParagraph.Range.ParagraphFormat.KeepWithNext = $true
        }
        catch {}
    }
    # Write-Host "Add-WordTable Caption:$($Caption)"

    if (-not $Rows -or $Rows.Count -eq 0) {
        Append-TextParagraph -Document $Document -Text '[空テーブル]' -StyleConfig $Config.Styles.Body -InternalLinkMap $InternalLinkMap | Out-Null
        return
    }

    # =========================
    # ① テーブル生成
    # =========================

    # グローバル設定の AutoNumberTables: ソース側で未指定の場合のみ空列を先頭挿入
    $sourceHasAutoNumber = $Attributes -and $Attributes.ContainsKey('options') -and
    ([string]$Attributes['options']) -match 'autonumber'
    $globalAutoNumberColumn = $false   # vertical-header でスキップする列か否かを記録
    if (-not $sourceHasAutoNumber -and
        $Config.Document -and
        $Config.Document.PSObject.Properties.Name -contains 'AutoNumberTables' -and
        [bool]$Config.Document.AutoNumberTables) {

        # ヘッダー行の検出（Markdown: IsHeader、AsciiDoc: options=header）
        $detectedHeader = ($Attributes -and $Attributes.ContainsKey('options') -and
            ([string]$Attributes['options']) -match 'header') -or
        ($Rows.Count -gt 0 -and $Rows[0].Count -gt 0 -and $Rows[0][0].IsHeader)

        # Attributes に autonumber / header を追加して既存の per-cell ロジックに乗せる
        if ($null -eq $Attributes) { $Attributes = @{} }
        $existingOpt = if ($Attributes.ContainsKey('options')) { [string]$Attributes['options'] } else { '' }
        $newOpt = if ([string]::IsNullOrWhiteSpace($existingOpt)) { 'autonumber' } else { $existingOpt + ',autonumber' }
        if ($detectedHeader -and $newOpt -notmatch 'header') { $newOpt += ',header' }
        $Attributes['options'] = $newOpt

        # 各行の先頭に空の連番セルを挿入
        $ColumnCount++
        $globalAutoNumberColumn = $true
        $newRows = @()
        foreach ($row in $Rows) {
            $isHdr = $detectedHeader -and ($row.Count -gt 0) -and $row[0].IsHeader
            $emptyCell = @{ Text = ''; RowSpan = 1; ColSpan = 1; IsHeader = $isHdr }
            $newRows += , (@($emptyCell) + $row)
        }
        $Rows = $newRows
    }

    $range = $Document.Range($Document.Content.End - 1, $Document.Content.End - 1)

    $table = $Document.Tables.Add($range, $Rows.Count, $ColumnCount)
    $table.Borders.Enable = 1
    $table.Rows.AllowBreakAcrossPages = $false

    # ★ 初期はAutoFit（後でOFF）
    try { $table.AutoFitBehavior((Get-WordConstant 'wdAutoFitContent')) | Out-Null } catch {}

    # =========================
    # ② セル内容設定（そのまま）
    # =========================
    $grid = @{}
    $occupied = @{}

    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $col = 0

        foreach ($cell in $Rows[$r]) {
            while ($occupied["$r,$col"]) { $col++ }
            if ($col -ge $ColumnCount) { break }

            $grid["$r,$col"] = $cell

            if ($cell.RowSpan -gt 1) {
                for ($rr = 1; $rr -lt $cell.RowSpan; $rr++) {
                    for ($cc = 0; $cc -lt $cell.ColSpan; $cc++) {
                        $occupied["$($r+$rr),$($col+$cc)"] = $true
                    }
                }
            }

            $col += $cell.ColSpan
            while ($occupied["$r,$col"]) { $col++ }
        }
    }
    $table.Rows(1).HeadingFormat = -1

    # ---- テキスト投入
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        for ($c = 0; $c -lt $ColumnCount; $c++) {

            $key = "$r,$c"
            if (-not $grid.ContainsKey($key)) { continue }

            try {
                $cellRange = $table.Cell($r + 1, $c + 1)
            }
            catch { continue }

            $cell = $grid[$key]
            $text = [string]$cell.Text

            # =========================
            # ★ autonumber 復活（ここ）
            # =========================
            $hasHeader = $false
            $autoNumber = $false
            if ($Attributes -and
                $Attributes.ContainsKey('options')) {

                $opt = [string]$Attributes['options']

                if ($opt -match 'autonumber') {
                    $autoNumber = $true
                }
                if ($opt -match 'header') {
                    $hasHeader = $true
                }
            }

            if ($autoNumber -and $c -eq 0) {

                # ヘッダー行は除外
                if (-not ($hasHeader -and $r -eq 0)) {

                    if ([string]::IsNullOrWhiteSpace($text)) {

                        $startRow = if ($hasHeader) { 1 } else { 0 }
                        $text = [string]($r - $startRow + 1)
                    }
                }
            }

            # =========================
            # スタイル選択
            # =========================
            $isHeaderCell = $cell.IsHeader -or ($hasHeader -and $r -eq 0)

            $style = if ($isHeaderCell) {
                $Config.Styles.TableHeader
            }
            else {
                $Config.Styles.TableBody
            }

            Set-TableCellTextWithHyperlinks `
                -Document $Document `
                -Cell $cellRange `
                -Text $text `
                -StyleConfig $style `
                -InternalLinkMap $InternalLinkMap
            if ($autoNumber -and $c -eq 0) {
                try {
                    $cellRange.ParagraphFormat.Alignment = 2 # 右寄せ
                    $cellRange.ParagraphFormat.WordWrap = $false
                    $cellRange.ParagraphFormat.Hyphenation = 0
                }
                catch {}
            }
        }
    }


    try {
        $pageWidth = $Document.PageSetup.PageWidth
        $leftMargin = $Document.PageSetup.LeftMargin
        $rightMargin = $Document.PageSetup.RightMargin

        $range = $Document.Content
        $range.Collapse((Get-WordConstant 'wdCollapseEnd'))

        $leftIndent = $range.ParagraphFormat.LeftIndent

        $availableWidth = $pageWidth - $leftMargin - $rightMargin - $leftIndent

        # ★ AutoFit OFF（絶対必要）
        $table.AutoFitBehavior(0)

        $table.PreferredWidthType = 3
        $table.PreferredWidth = $availableWidth

        # ---- cols属性
        if ($Attributes -and $Attributes.ContainsKey('cols')) {

            $colDefs = $Attributes['cols'] -split ',' | ForEach-Object {

                $s = $_.Trim()

                $ratio = 1
                $align = $null
                $isAsciiDoc = $false

                # 数値
                if ($s -match '^\d+') {
                    $ratio = [double]$matches[0]
                }

                # 属性
                if ($s -match '[lcr]') {
                    $align = $matches[0]
                }

                if ($s -match 'a') {
                    $isAsciiDoc = $true
                }

                [PSCustomObject]@{
                    Ratio      = $ratio
                    Align      = $align
                    IsAsciiDoc = $isAsciiDoc
                }
            }
        }
        else {
            $colDefs = @(for ($i = 0; $i -lt $ColumnCount; $i++) { 11 }) | ForEach-Object {
                [PSCustomObject]@{
                    Ratio      = $_
                    Align      = $null
                    IsAsciiDoc = $false
                }
            }
            
        }
        $ratios = $colDefs | ForEach-Object { $_.Ratio }

        if ($autoNumber) {
            $remainRatios = $ratios[1..($ratios.Count - 1)]
            $availableWidth -= $Config.Styles.Table.NumberWidth
        }
        else {
            $remainRatios = $ratios
        }

        $sum = ($remainRatios | Measure-Object -Sum).Sum

        $table.AllowAutoFit = $false

        if ($ratios.Count -eq $ColumnCount -and $sum -gt 0) {

            for ($i = 1; $i -le $ColumnCount; $i++) {
                $width = 0.0
                if ($autoNumber -and $i -eq 1) {
                    $width = [float]$Config.Styles.Table.NumberWidth
                }
                else {
                    $width = [float]($availableWidth * ($ratios[$i - 1] / $sum))
                    # remainRatiosを使わない理由は ループのインデックスはNumber列も含んでいる為
                    # remainRatiosは列を除いた比率を計算するために使っているだけ
                }
                #Write-Host "i:($i) Widht:$($width)"
                #$table.Columns.Item($i).Width = [float]$width

                # テーブルを追加してすぐにWidthを設定するとエラーになるのでリトライを入れる
                for ($retry = 0; $retry -lt 5; $retry++) {
                    try {
                        $table.Columns.Item($i).Width = [float]$width
                        break
                    }
                    catch {
                        Start-Sleep -Milliseconds 1000
                        if ($retry -eq 4) { 
                            throw 
                        }
                    }
                }

            }
        }
        else {
            # フォールバック（均等）
            $w = $availableWidth / $ColumnCount
            for ($i = 1; $i -le $ColumnCount; $i++) {
                $table.Columns.Item($i).Width = $w
            }
        }
    }
    catch {
        Write-Host "Error"
        Write-Host "WARN failed to set table width: NumberWidth:$($Config.Styles.Table.NumberWidth) $($_.InvocationInfo.Line):($($_.InvocationInfo.ScriptLineNumber)) $($_.Exception.Message) $($_.Exception.StackTrace)"
    }

    # ---- Merge
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        for ($c = 0; $c -lt $ColumnCount; $c++) {

            $key = "$r,$c"
            if (-not $grid.ContainsKey($key)) { continue }

            $cell = $grid[$key]
            if ($cell.RowSpan -le 1 -and $cell.ColSpan -le 1) { continue }

            try {
                $cellRange = $table.Cell($r + 1, $c + 1)
                $cellRange.Merge($table.Cell($r + $cell.RowSpan, $c + $cell.ColSpan))
            }
            catch {}
        }
    }

    # ---- 縦書きヘッダー: options="...,vertical-header" または [vertical-header] で指定
    $hasVerticalHeader = $false
    if ($Attributes) {
        if ($Attributes.ContainsKey('options') -and ([string]$Attributes['options']) -match 'vertical-header') {
            $hasVerticalHeader = $true
        }
        elseif ($Attributes.Values -contains 'vertical-header') {
            $hasVerticalHeader = $true
        }
    }

    if ($hasVerticalHeader) {
        $isHeaderOpt = $Attributes.ContainsKey('options') -and ([string]$Attributes['options']) -match 'header'
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            for ($c = 0; $c -lt $ColumnCount; $c++) {
                # グローバル連番列（c=0）は縦書き対象外
                if ($globalAutoNumberColumn -and $c -eq 0) { continue }
                $key = "$r,$c"
                if (-not $grid.ContainsKey($key)) { continue }
                $cell = $grid[$key]
                $isHdrCell = $cell.IsHeader -or ($isHeaderOpt -and $r -eq 0)
                if ($isHdrCell) {
                    try {
                        $tc = $table.Cell($r + 1, $c + 1)
                        $tc.Range.Orientation = 3   # wdTextOrientationUpward (下→上, 90度)
                        $tc.VerticalAlignment = 1   # wdCellAlignVerticalCenter
                    }
                    catch {}
                }
            }
        }
    }

    # ★ 最後に余計な段落は作らない
}

function Get-ImageFullPath {
    param(
        [Parameter(Mandatory = $true)][string]$ImageReference,
        [Parameter(Mandatory = $true)][string]$CurrentFileDirectory,
        [hashtable]$Attributes
    )

    $imagePath = $ImageReference

    if ($Attributes -and $Attributes.ContainsKey('imagesdir')) {
        $imagesDir = [string]$Attributes['imagesdir']
        if (-not [string]::IsNullOrWhiteSpace($imagesDir)) {
            if (-not [System.IO.Path]::IsPathRooted($imagePath)) {
                $imagePath = Join-Path $imagesDir $imagePath
            }
        }
    }

    if (-not [System.IO.Path]::IsPathRooted($imagePath)) {
        $imagePath = Join-Path $CurrentFileDirectory $imagePath
    }

    return [System.IO.Path]::GetFullPath($imagePath)
}

function Add-ImageToDocument {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][string]$ImagePath,
        [string]$Caption,
        $Config
    )

    $app = $Document.Application

    # -----------------------------
    # ① キャプション追加
    # -----------------------------
    $captionRange = $null
    if ($Caption) {
        $captionRange = Append-TextParagraph `
            -Document $Document `
            -Text $Caption `
            -StyleConfig $Config.Styles.FigureCaption

        try {
            $captionRange.ParagraphFormat.KeepWithNext = $true
        }
        catch {}
    }

    # -----------------------------
    # ② 挿入位置取得
    # -----------------------------
    $range = $Document.Content
    $range.Collapse(0)

    # 現在Y位置
    $currentY = $range.Information(3)

    # ページ情報
    $pageHeight = $Document.PageSetup.PageHeight
    $topMargin = $Document.PageSetup.TopMargin
    $bottomMargin = $Document.PageSetup.BottomMargin

    # 残り高さ
    $remainingHeight = $pageHeight - $bottomMargin - $currentY

    # -----------------------------
    # ③ 画像挿入（仮）
    # -----------------------------
    if (Test-Path -LiteralPath $ImagePath) {
        try {
            $shape = $Document.InlineShapes.AddPicture($ImagePath, $false, $true, $range)
            $shape.Range.ParagraphFormat.KeepTogether = $true

            # -----------------------------
            # ④ サイズ制御（ここがガチ）
            # -----------------------------
            $originalWidth = $shape.Width
            $originalHeight = $shape.Height

            # 設定値（あれば）
            $maxWidth = $null
            if ($Config.Image.MaxWidthMm) {
                $maxWidth = $app.MillimetersToPoints([double]$Config.Image.MaxWidthMm)
            }

            # --- 横方向制限 ---
            if ($maxWidth -and $originalWidth -gt $maxWidth) {
                $shape.LockAspectRatio = $true
                $shape.Width = $maxWidth
            }

            # 再取得（縮んだ可能性）
            $currentHeight = $shape.Height

            # キャプション分を考慮（適当だが効果あり）
            $captionReserve = 150  # pt（微調整可）

            $availableHeight = $remainingHeight - $captionReserve

            # -----------------------------
            # ⑤ はみ出る場合の処理
            # -----------------------------
            if ($currentHeight -gt $availableHeight) {

                # --- 縮小可能なら縮小 ---
                if ($availableHeight -gt 50) {

                    $ratio = $availableHeight / $currentHeight

                    if ($ratio -lt 1.0) {
                        $shape.LockAspectRatio = $true
                        $shape.Height = $shape.Height * $ratio
                    }

                }
                else {
                    # --- どうしようもない → 改ページ ---
                    $shape.Delete()

                    # 改ページ
                    $range = $Document.Content
                    $range.Collapse(0)
                    $range.InsertBreak(7) # wdPageBreak

                    # 再挿入
                    $range = $Document.Content
                    $range.Collapse(0)

                    $shape = $Document.InlineShapes.AddPicture($ImagePath, $false, $true, $range)
                }
            }
        }
        catch {}
    }
    else {
        $text = "【画像が見つかりません】:$ImagePath"
        Write-Host $text -ForegroundColor Red
        $warningConfig = $Config.Styles.Body | ConvertTo-Json -Depth 10 | ConvertFrom-Json

        $warningConfig.Color = 'FF0000' # 赤
        $warningConfig.Bold = $true

        Append-TextParagraph `
            -Document $Document `
            -Text $text `
            -StyleConfig $Config.Styles.Body | Out-Null
    }
    
    # -----------------------------
    # ⑥ 後処理
    # -----------------------------
    $Document.Content.InsertParagraphAfter() | Out-Null
}

function Set-HeaderFooter {
    param(
        $Document,
        $Config,
        $Metadata,
        $IsCovePage = $false
    )

    $section = $Document.Sections.Item(1)
    $section.PageSetup.DifferentFirstPageHeaderFooter = $true
    $headerRange = $section.Headers.Item((Get-WordConstant 'wdHeaderFooterPrimary')).Range
    $footerRange = $section.Footers.Item((Get-WordConstant 'wdHeaderFooterPrimary')).Range

    $headerText = ''
    if ((-not $IsCovePage) -and $Config.HeaderFooter.Header -and $Config.HeaderFooter.Header.Text) { $headerText = [string]$Config.HeaderFooter.Header.Text }
    $footerText = ''
    if ($Config.HeaderFooter.Footer -and $Config.HeaderFooter.Footer.Text) { $footerText = [string]$Config.HeaderFooter.Footer.Text }

    $replacements = @{
        '{title}'     = [string]($Metadata.Title)
        '{subtitle}'  = [string]($Metadata.Subtitle)
        '{author}'    = [string]($Metadata.Author)
        '{revnumber}' = [string]($Metadata.RevNumber)
        '{revdate}'   = [string]($Metadata.RevDate)
        '{copyright}' = [string]($Metadata.Copyright)
    }

    foreach ($key in $replacements.Keys) {
        $headerText = $headerText.Replace($key, $replacements[$key])
        $footerText = $footerText.Replace($key, $replacements[$key])
    }

    $headerRange.Text = $headerText
    Apply-FontStyle -Range $headerRange -StyleConfig $Config.HeaderFooter.Header.Style

    $footerRange.Text = $footerText
    Apply-FontStyle -Range $footerRange -StyleConfig $Config.HeaderFooter.Footer.Style
}

function Get-PlantUmlOutputExtension {
    param([string]$Format)
    $f = ([string]$Format).ToLowerInvariant()
    switch ($f) {
        'svg' { return 'svg' }
        'png' { return 'png' }
        default { return $f }
    }
}

function Parse-AsciiDocAttributeList {
    param([string]$Text)

    $result = @{}
    if ([string]::IsNullOrWhiteSpace($Text)) { return $result }

    $segments = New-Object System.Collections.Generic.List[string]
    $current = New-Object System.Text.StringBuilder
    $inQuote = $false
    $quoteChar = ''

    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]

        if (($ch -eq '"' -or $ch -eq "'") -and ($i -eq 0 -or $Text[$i - 1] -ne '\')) {
            if (-not $inQuote) {
                $inQuote = $true
                $quoteChar = $ch
            }
            elseif ($quoteChar -eq $ch) {
                $inQuote = $false
                $quoteChar = ''
            }

            [void]$current.Append($ch)
            continue
        }

        if ($ch -eq ',' -and -not $inQuote) {
            $segments.Add($current.ToString().Trim())
            [void]$current.Clear()
            continue
        }

        [void]$current.Append($ch)
    }

    if ($current.Length -gt 0) {
        $segments.Add($current.ToString().Trim())
    }

    foreach ($segment in $segments) {
        if ([string]::IsNullOrWhiteSpace($segment)) { continue }

        if ($segment -match '^(?<key>[A-Za-z0-9_\-]+)\s*=\s*"(?<value>.*)"$') {
            $result[$matches['key']] = $matches['value']
            continue
        }

        if ($segment -match "^(?<key>[A-Za-z0-9_\-]+)\s*=\s*'(?<value>.*)'$") {
            $result[$matches['key']] = $matches['value']
            continue
        }

        if ($segment -match '^(?<key>[A-Za-z0-9_\-]+)\s*=\s*(?<value>.+)$') {
            $result[$matches['key']] = $matches['value']
            continue
        }

        $index = $result.Count
        $result[[string]$index] = $segment
    }

    return $result
}

function Invoke-PlantUmlRender {
    param(
        [string]$PlantUmlSource,
        [string]$SourceFilePath,
        [hashtable]$Attributes,
        [hashtable]$Options,
        $Config,
        [int]$Sequence
    )

    $result = @{
        Success      = $false
        ImagePath    = $null
        ErrorMessage = $null
        Format       = $null
    }

    if (-not $Config.PlantUml -or -not [bool]$Config.PlantUml.Enabled) {
        $result.ErrorMessage = 'PlantUML が設定で無効化されています。'
        return [pscustomobject]$result
    }

    $sourceDir = Split-Path -Parent $SourceFilePath
    $javaPath = if ($Config.PlantUml.JavaPath) { [string]$Config.PlantUml.JavaPath } else { 'java' }
    $jarPath = if ($Config.PlantUml.JarPath) { Get-AbsolutePath -Path ([string]$Config.PlantUml.JarPath) -BaseDirectory $PSScriptRoot } else { $null }
    if ([string]::IsNullOrWhiteSpace($jarPath) -or -not (Test-Path $jarPath)) {
        $result.ErrorMessage = "PlantUML jar が見つかりません: $jarPath"
        return [pscustomobject]$result
    }

    $format = $null
    if ($Options.ContainsKey('generated-image-format')) {
        $format = [string]$Options['generated-image-format']
    }
    elseif ($Options.ContainsKey('format')) {
        $format = [string]$Options['format']
    }
    elseif ($Config.PlantUml.DefaultFormat) {
        $format = [string]$Config.PlantUml.DefaultFormat
    }
    else {
        $format = 'png'
    }
    $result.Format = $format

    $targetName = $null
    if ($Options.ContainsKey('target')) {
        $targetName = [string]$Options['target']
    }
    if ([string]::IsNullOrWhiteSpace($targetName)) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourceFilePath)
        $targetName = '{0}-plantuml-{1:D4}' -f $baseName, $Sequence
    }

    $outputRoot = if ($Config.PlantUml.OutputDir) { [string]$Config.PlantUml.OutputDir } else { '.\\generated-images' }
    $outputDir = Get-AbsolutePath -Path $outputRoot -BaseDirectory $sourceDir
    if (-not (Test-Path $outputDir)) {
        [void](New-Item -ItemType Directory -Path $outputDir -Force)
    }

    $tempDir = Join-Path $outputDir '_tmp'
    if (-not (Test-Path $tempDir)) {
        [void](New-Item -ItemType Directory -Path $tempDir -Force)
    }

    $tempSourcePath = Join-Path $tempDir ($targetName + '.puml')
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($tempSourcePath, $PlantUmlSource, $encoding)

    $outputExt = Get-PlantUmlOutputExtension -Format $format
    $OutputFullPath = Join-Path $outputDir ($targetName + '.' + $outputExt)
    if (Test-Path $OutputFullPath) { Remove-Item -LiteralPath $OutputFullPath -Force }

    $arguments = @('-jar', $jarPath, ('-t' + $format), '-charset', 'UTF-8', '-o', $outputDir, $tempSourcePath)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $javaPath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.Arguments = Join-CommandLineArguments -Arguments $arguments

    $process = [System.Diagnostics.Process]::Start($psi)
    $stdOut = $process.StandardOutput.ReadToEnd()
    $stdErr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    if ($process.ExitCode -ne 0) {
        $message = ($stdErr, $stdOut | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' '
        if ([string]::IsNullOrWhiteSpace($message)) { $message = 'PlantUML 実行エラー' }
        $result.ErrorMessage = $message.Trim()
        return [pscustomobject]$result
    }

    if (-not (Test-Path $OutputFullPath)) {
        $generated = Get-ChildItem -LiteralPath $outputDir -Filter ($targetName + '.*') -File | Select-Object -First 1
        if ($generated) {
            $OutputFullPath = $generated.FullName
        }
    }

    if (-not (Test-Path $OutputFullPath)) {
        $result.ErrorMessage = "PlantUML 画像が生成されませんでした: $targetName"
        return [pscustomobject]$result
    }

    $result.Success = $true
    $result.ImagePath = $OutputFullPath
    return [pscustomobject]$result
}

function Invoke-DrawIoRender {
    param(
        [string]$SourceFilePath,
        [hashtable]$Options,
        $Config
    )

    $result = @{
        Success      = $false
        ImagePath    = $null
        ErrorMessage = $null
        Format       = $null
    }

    if (-not $Config.DrawIo -or -not [bool]$Config.DrawIo.Enabled) {
        $result.ErrorMessage = 'draw.io が設定で無効化されています。'
        return [pscustomobject]$result
    }

    if (-not $Options -or -not $Options.ContainsKey('file')) {
        $result.ErrorMessage = '[drawio] に file 属性がありません。'
        return [pscustomobject]$result
    }

    $sourceDir = Split-Path -Parent $SourceFilePath
    $drawioPath = Get-AbsolutePath -Path ([string]$Options['file']) -BaseDirectory $sourceDir

    if (-not (Test-Path -LiteralPath $drawioPath)) {
        $result.ErrorMessage = "draw.io ファイルが見つかりません: $drawioPath"
        return [pscustomobject]$result
    }

    $format = 'svg'
    if ($Options.ContainsKey('generated-image-format')) {
        $format = [string]$Options['generated-image-format']
    }
    elseif ($Options.ContainsKey('format')) {
        $format = [string]$Options['format']
    }
    elseif ($Config.DrawIo.DefaultFormat) {
        $format = [string]$Config.DrawIo.DefaultFormat
    }
    $svgTheme = 'light'

    $format = $format.ToLowerInvariant()
    $result.Format = $format

    $cliPath = [string]$Config.DrawIo.CliPath
    if ([string]::IsNullOrWhiteSpace($cliPath) -or -not (Test-Path -LiteralPath $cliPath)) {
        $result.ErrorMessage = "draw.io CLI が見つかりません: $cliPath"
        return [pscustomobject]$result
    }

    $outputRoot = if ($Config.DrawIo.OutputDir) { [string]$Config.DrawIo.OutputDir } else { '.\generated-images' }
    $outputDir = Get-AbsolutePath -Path $outputRoot -BaseDirectory $sourceDir

    if (-not (Test-Path -LiteralPath $outputDir)) {
        [void](New-Item -ItemType Directory -Path $outputDir -Force)
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($drawioPath)
    $outputPath = Join-Path $outputDir ($baseName + '.' + $format)

    if (Test-Path -LiteralPath $outputPath) {
        Remove-Item -LiteralPath $outputPath -Force
    }

    $arguments = @(
        '--export',
        $drawioPath,
        '--output',
        $outputPath,
        '--format',
        $format ,
        '--svg-theme',
        $svgTheme
        
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $cliPath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.Arguments = Join-CommandLineArguments -Arguments $arguments

    $process = [System.Diagnostics.Process]::Start($psi)
    $stdOut = $process.StandardOutput.ReadToEnd()
    $stdErr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    if ($process.ExitCode -ne 0) {
        $message = ($stdErr, $stdOut | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' '
        if ([string]::IsNullOrWhiteSpace($message)) {
            $message = 'draw.io 変換エラー'
        }

        $result.ErrorMessage = $message.Trim()
        return [pscustomobject]$result
    }

    if (-not (Test-Path -LiteralPath $outputPath)) {
        $result.ErrorMessage = "draw.io 画像が生成されませんでした: $outputPath"
        return [pscustomobject]$result
    }

    $result.Success = $true
    $result.ImagePath = $outputPath
    return [pscustomobject]$result
}

function Get-NextNonEmptyTrimmedLine {
    param(
        [string[]]$Lines,
        [int]$StartIndex
    )

    for ($i = $StartIndex; $i -lt $Lines.Count; $i++) {
        $candidate = [string]$Lines[$i]
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            return $candidate.Trim()
        }
    }

    return $null
}

function Parse-AsciiDocFile {
    param(
        [string]$Path,
        [hashtable]$Attributes,
        [System.Collections.Generic.HashSet[string]]$Visited,
        $Config
    )

    $absolutePath = Get-AbsolutePath -Path $Path
    if ($Visited.Contains($absolutePath)) {
        return [pscustomobject]@{
            Metadata   = @{}
            Elements   = @()
            Attributes = $Attributes
        }
    }
    [void]$Visited.Add($absolutePath)

    $fileDir = Split-Path -Parent $absolutePath
    $text = Get-Content -LiteralPath $absolutePath -Encoding UTF8 -Raw
    # 改行コード統一
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"

    $lines = @($text -split "`n")

    $elements = New-Object System.Collections.Generic.List[object]
    $metadata = @{
        Title     = $null
        Subtitle  = $null
        Author    = $null
        RevNumber = $null
        RevDate   = $null
        Copyright = $null
        ImagesDir = $null
    }
    $paragraphBuffer = New-Object System.Collections.Generic.List[string]
    $pendingCaption = $null
    $pendingBlockAttributes = $null
    $inFence = $false
    $fenceDelimiter = $null
    $fenceLines = New-Object System.Collections.Generic.List[string]
    $pendingBlockType = $null
    $lineIndex = 0
    $plantUmlSequence = 1
    

    function Flush-ParagraphBuffer {
        if ($paragraphBuffer.Count -eq 0) { return }

        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($p in $paragraphBuffer) {
            $t = [string]$p
            if ($t -eq '__LINEBREAK__') {
                $parts.Add([Environment]::NewLine)
            }
            else {
                $parts.Add($t.TrimEnd())
            }
        }

        $joined = ''
        foreach ($part in $parts) {
            if ($part -eq [Environment]::NewLine) {
                $joined += [Environment]::NewLine
            }
            elseif ([string]::IsNullOrWhiteSpace($joined)) {
                $joined = $part
            }
            elseif ($joined.EndsWith([Environment]::NewLine)) {
                $joined += $part
            }
            else {
                $joined += ' ' + $part
            }
        }

        $normalized = Normalize-InlineText -Text $joined -Attributes $Attributes
        if (-not [string]::IsNullOrWhiteSpace($normalized)) {
            $elements.Add((New-Element -Type 'paragraph' -Data @{ Text = $normalized }))
        }
        $paragraphBuffer.Clear()
    }
    
    while ($lineIndex -lt $lines.Count) {
        $line = [string]$lines[$lineIndex]
        $trimmed = $line.Trim()
        
        if ($inFence) {
            if ($trimmed -eq $fenceDelimiter) {
                $blockText = ($fenceLines -join [Environment]::NewLine)
                if ($pendingBlockType -eq 'plantuml') {
                    $render = Invoke-PlantUmlRender -PlantUmlSource $blockText -SourceFilePath $absolutePath -Attributes $Attributes -Options $pendingBlockAttributes -Config $Config -Sequence $plantUmlSequence
                    $plantUmlSequence++
                    if ($render.Success) {
                        $elements.Add((New-Element -Type 'image' -Data @{ Path = $render.ImagePath; Caption = $pendingCaption; GeneratedBy = 'plantuml'; Source = $blockText }))
                    }
                    else {
                        $fallback = "[PlantUML 画像生成失敗] $($render.ErrorMessage)"
                        $elements.Add((New-Element -Type 'admonition' -Data @{ Kind = 'WARNING'; Text = $fallback }))
                        $elements.Add((New-Element -Type 'code' -Data @{
                                    Text    = $blockText
                                    Caption = $pendingCaption
                                }))
                    }
                }
                else {
                    $elements.Add((New-Element -Type 'code' -Data @{
                                Text    = $blockText
                                Caption = $pendingCaption
                            }))
                }

                $fenceLines.Clear()
                $inFence = $false
                $fenceDelimiter = $null
                $pendingBlockType = $null
                $pendingBlockAttributes = $null
                $pendingCaption = $null
                $lineIndex++
                continue
            }
            $fenceLines.Add($line)
            $lineIndex++
            continue
        }

        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            Flush-ParagraphBuffer
            $lineIndex++
            continue
        }

        if ($trimmed.StartsWith('//')) {
            $lineIndex++
            continue
        }

        if ($trimmed -match '^(ifeval|ifdef|ifndef|endif)::') {
            $lineIndex++
            continue
        }

        if ($trimmed -match '^:([^:]+):\s*(.*)$') {
            Flush-ParagraphBuffer
            $attrName = $matches[1].Trim()
            $attrValue = $matches[2]
            if ($attrName -eq 'author') {
                $parts = New-Object System.Collections.Generic.List[string]
                if ($attrValue) { $parts.Add($attrValue.Trim()) }
                while ($lineIndex + 1 -lt $lines.Count) {
                    $current = [string]$lines[$lineIndex]
                    $peek = [string]$lines[$lineIndex + 1]
                    if ($current.TrimEnd().EndsWith('+')) {
                        if ($parts.Count -gt 0) {
                            $parts[$parts.Count - 1] = $parts[$parts.Count - 1].TrimEnd(' ', '+')
                        }
                        $lineIndex++
                        $parts.Add($peek.Trim())
                        continue
                    }
                    break
                }
                $attrValue = ($parts -join [Environment]::NewLine)
            }
            $Attributes[$attrName] = $attrValue
            switch ($attrName) {
                'title' { $metadata.Title = Normalize-InlineText -Text $attrValue -Attributes $Attributes }
                'subtitle' { $metadata.Subtitle = Normalize-InlineText -Text $attrValue -Attributes $Attributes }
                'author' { $metadata.Author = $attrValue }
                'revnumber' { $metadata.RevNumber = $attrValue }
                'revdate' { $metadata.RevDate = $attrValue }
                'copyright' { $metadata.Copyright = Normalize-InlineText -Text $attrValue -Attributes $Attributes }
                'imagesdir' { $metadata.ImagesDir = $attrValue }
                default { }
            }
            $lineIndex++
            continue
        }

        if ($trimmed -match '^include::([^\[]+)\[(.*?)\]$') {
            Flush-ParagraphBuffer
            $includePath = Resolve-IncludePath -DirectivePath $matches[1] -CurrentFileDirectory $fileDir
            if (Test-Path $includePath) {
                $childAttributes = Merge-AttributeMaps -Base $Attributes -Overlay @{}
                $included = Parse-AsciiDocFile -Path $includePath -Attributes $childAttributes -Visited $Visited -Config $Config
                foreach ($child in $included.Elements) { $elements.Add($child) }
            }
            else {
                $elements.Add((New-Element -Type 'paragraph' -Data @{ Text = "[include ファイルが見つかりません: $includePath]" }))
            }
            $lineIndex++
            continue
        }

        if ($trimmed -eq '<<<') {
            Flush-ParagraphBuffer
            $elements.Add((New-Element -Type 'pagebreak' -Data @{}))
            $lineIndex++
            # 改ページの後の空行を捨てる
            while ($lineIndex -lt $lines.Count) {
                $tLine = [string]$lines[$lineIndex]
                $tTrim = $tLine.Trim()
                if ($tTrim) {
                    break
                }
                $lineIndex++
            }
            continue
        }

        if ($trimmed -match '^=\s+(.+?)(?:\s*:\s*(.+))?$') {
            Flush-ParagraphBuffer
            $metadata.Title = Normalize-InlineText -Text $matches[1] -Attributes $Attributes
            if (-not $metadata.Subtitle) {
                $metadata.Subtitle = $null
            }
            if ($matches[2]) { 
                $metadata.Subtitle = Normalize-InlineText -Text $matches[2] -Attributes $Attributes 
            }
            $Attributes['title'] = $metadata.Title
            if ($metadata.Subtitle) { $Attributes['subtitle'] = $metadata.Subtitle }
            $elements.Add((New-Element -Type 'title' -Data @{ Text = $metadata.Title; Subtitle = $metadata.Subtitle }))
            $lineIndex++
            continue
        }

        if ($trimmed -match '^(==+)\s+(.+)$') {
            Flush-ParagraphBuffer
            $level = $matches[1].Length - 1
            $text = Normalize-InlineText -Text $matches[2] -Attributes $Attributes
            $elements.Add((New-Element -Type 'heading' -Data @{ Level = $level; Text = $text }))
            $lineIndex++
            continue
        }

        if ($trimmed -match '^\[\[([^\]]+)\]\]$') {
            Flush-ParagraphBuffer
            $anchorId = [string]$matches[1]
            if (-not [string]::IsNullOrWhiteSpace($anchorId)) {
                $elements.Add((New-Element -Type 'anchor' -Data @{ Id = $anchorId.Trim() }))
            }
            $lineIndex++
            continue
        }

        if ($trimmed -match '^\[#([^\]]+)\]$') {
            Flush-ParagraphBuffer
            $anchorId = [string]$matches[1]
            if (-not [string]::IsNullOrWhiteSpace($anchorId)) {
                $elements.Add((New-Element -Type 'anchor' -Data @{ Id = $anchorId.Trim() }))
            }
            $lineIndex++
            continue
        }

        if ($trimmed -match '^\.(\S.*)$' -and -not $trimmed.StartsWith('..')) {
            $captionText = $matches[1]
            $nextTrimmed = Get-NextNonEmptyTrimmedLine -Lines $lines -StartIndex ($lineIndex + 1)

            Write-DebugLog "CAPTION-CANDIDATE line=$lineIndex text=[$trimmed] next=[$nextTrimmed]"

            $isCaptionTarget = $false
            if ($nextTrimmed) {
                if ($nextTrimmed -match '^\[.*\]$' -or
                    $nextTrimmed -match '^image::' -or
                    $nextTrimmed -match '^\|={3,}$' -or
                    $nextTrimmed -eq '----' -or
                    $nextTrimmed -eq '....' -or
                    $nextTrimmed -eq '```') {
                    $isCaptionTarget = $true
                }
            }

            if ($isCaptionTarget) {
                Flush-ParagraphBuffer
                $pendingCaption = Normalize-InlineText -Text $captionText -Attributes $Attributes
                $lineIndex++
                continue
            }
        }

        if ($trimmed -match '^\[(.+)\]$') {
            Flush-ParagraphBuffer
            $inside = $matches[1]
            

            if ($trimmed -match '^(NOTE|TIP|IMPORTANT|WARNING|CAUTION):\s+(.+)$') {
                Flush-ParagraphBuffer
                $elements.Add((New-Element -Type 'admonition' -Data @{
                            Kind = $matches[1]
                            Text = (Normalize-InlineText -Text $matches[2] -Attributes $Attributes)
                        }))
                $lineIndex++
                continue
            }

            $attrList = Parse-AsciiDocAttributeList -Text $inside
            if ($attrList.Values -contains 'plantuml' -or $inside -match '^plantuml(?:,|$)') {
                $pendingBlockType = 'plantuml'
                $pendingBlockAttributes = $attrList
                $lineIndex++
                continue
            }

            if ($attrList.Values -contains 'drawio' -or $inside -match '^drawio(?:,|$)') {
                Flush-ParagraphBuffer

                $caption = $pendingCaption
                if ($attrList.ContainsKey('caption')) {
                    $caption = [string]$attrList['caption']
                }

                $render = Invoke-DrawIoRender `
                    -SourceFilePath $absolutePath `
                    -Options $attrList `
                    -Config $Config

                if ($render.Success) {
                    $elements.Add((New-Element -Type 'image' -Data @{
                                Path        = $render.ImagePath
                                Caption     = $caption
                                GeneratedBy = 'drawio'
                            }))
                }
                else {
                    $elements.Add((New-Element -Type 'admonition' -Data @{
                                Kind = 'WARNING'
                                Text = "[draw.io 画像生成失敗] $($render.ErrorMessage)"
                            }))
                }

                $pendingCaption = $null
                $pendingBlockAttributes = $null
                $lineIndex++
                continue
            }

            if ($attrList.Values -contains 'source' -or $inside -match '^source(?:,|$)') {
                $pendingBlockType = 'source'
                $pendingBlockAttributes = $attrList
                $lineIndex++
                continue
            }
            if ($inside -match 'cols=' -or $inside -match 'options=') {
                $pendingBlockType = 'table'
                $pendingBlockAttributes = $attrList
                $lineIndex++
                continue
            }

            $pendingBlockAttributes = $attrList
            $lineIndex++
            continue
        }

        if ($trimmed -eq '```' -or $trimmed -eq '----' -or $trimmed -eq '....') {
            Flush-ParagraphBuffer
            $inFence = $true
            $fenceDelimiter = $trimmed
            $fenceLines.Clear()
            if (-not $pendingBlockType) { $pendingBlockType = 'code' }
            $lineIndex++
            continue
        }

        if ($trimmed -match '^image::([^\[]+)\[(.*?)\]$') {
            Flush-ParagraphBuffer
            $imageRef = $matches[1]
            $imagePath = Get-ImageFullPath -ImageReference $imageRef -CurrentFileDirectory $fileDir -Attributes $Attributes
            $elements.Add((New-Element -Type 'image' -Data @{ Path = $imagePath; Caption = $pendingCaption }))
            $pendingCaption = $null
            $lineIndex++
            continue
        }

        if ($trimmed -match '^\|={3,}$') {
            Flush-ParagraphBuffer
             
            Write-DebugLog "TABLE-START line=$lineIndex pendingCaption=[$pendingCaption]"

            $tableLines = New-Object System.Collections.Generic.List[string]
            $lineIndex++
            while ($lineIndex -lt $lines.Count) {
                $tLine = [string]$lines[$lineIndex]
                $tTrim = $tLine.Trim()
                #Write-Output "Table $tTrim"
                if ($tTrim -match '^\|={3,}$') { break }
                # 空行は捨てる
                #if ($tTrim) {
                $tableLines.Add($tLine)
                #}
                $lineIndex++
            }
            $expectedColumns = Get-ColumnCountFromColsAttribute -Attributes $pendingBlockAttributes
            $tableInfo = Convert-TableRows -Lines $tableLines -ExpectedColumns $expectedColumns

            if (-not [string]::IsNullOrWhiteSpace($pendingCaption)) {
                $elements.Add((New-Element -Type 'tablecaption' -Data @{
                            Text = $pendingCaption
                        }))
            }

            Write-DebugLog "TABLE-ADD rows=$($tableInfo.Rows.Count) cols=$($tableInfo.MaxColumns) caption=[$pendingCaption]"

            $elements.Add((New-Element -Type 'table' -Data @{
                        tableInfo  = $tableInfo
                        Caption    = $null
                        Attributes = $pendingBlockAttributes
                    }))

            $pendingCaption = $null
            $pendingBlockAttributes = $null
            $pendingBlockType = $null
            $lineIndex++
            continue
        }

        if ($trimmed -match '^([\*\-]+)\s+(.+)$') {
            Flush-ParagraphBuffer
            $indentLength = ([regex]::Match($line, '^\s*')).Value.Length
            $level = [Math]::Max(1, [int][Math]::Floor($indentLength / 2) + 1)
            $elements.Add((New-Element -Type 'bullet' -Data @{ Text = (Normalize-InlineText -Text $matches[2] -Attributes $Attributes); Level = $level; MarkerCount = $matches[1].Length }))
            $lineIndex++
            continue
        }

        if ($trimmed -match '^(\d+)\.\s+(.+)$') {
            Flush-ParagraphBuffer
            $indentLength = ([regex]::Match($line, '^\s*')).Value.Length
            $level = [Math]::Max(1, [int][Math]::Floor($indentLength / 2) + 1)
            $elements.Add((New-Element -Type 'numbered' -Data @{ Text = (Normalize-InlineText -Text $matches[2] -Attributes $Attributes); Level = $level; MarkerCount = $matches[1].Length }))
            $lineIndex++
            continue
        }

        if ($trimmed -match '^\.+\s+(.+)$') {
            $markerCount = $matches[1].Length
            Flush-ParagraphBuffer

            $dotPrefix = ([regex]::Match($trimmed, '^\.+')).Value.Length
            $text = Normalize-InlineText -Text $matches[1] -Attributes $Attributes

            $lineIndex++

            # AsciiDoc のリスト継続:
            # +
            # 続き
            # を複数回処理する
            while (($lineIndex + 1) -lt $lines.Count -and
                ([string]$lines[$lineIndex]).Trim() -eq '+') {

                $nextTrimmed = ([string]$lines[$lineIndex + 1]).Trim()

                if ($nextTrimmed -eq '```' -or
                    $nextTrimmed -eq '----' -or
                    $nextTrimmed -eq '....' -or
                    $nextTrimmed -match '^\|={3,}$' -or
                    $nextTrimmed -match '^image::') {
                    break
                }

                $continueLine = [string]$lines[$lineIndex + 1]
                $continueText = Normalize-InlineText -Text $continueLine -Attributes $Attributes

                if (-not [string]::IsNullOrWhiteSpace($continueText)) {
                    $text = $text + [char]11 + $continueText
                }

                $lineIndex += 2
            }

            $elements.Add((New-Element -Type 'numbered' -Data @{
                        Text        = $text
                        Level       = $dotPrefix
                        MarkerCount = $dotPrefix
                    }))

            continue
        }

        if ($trimmed -match '^(NOTE|TIP|IMPORTANT|WARNING|CAUTION):\s+(.+)$') {
            Flush-ParagraphBuffer
            $elements.Add((New-Element -Type 'admonition' -Data @{ Kind = $matches[1]; Text = (Normalize-InlineText -Text $matches[2] -Attributes $Attributes) }))
            $lineIndex++
            continue
        }

        if ($trimmed -match '^\[\[([^\]]+)\]\]$') {
            $lineIndex++
            continue
        }

        if ($trimmed -eq '+') {
            $nextTrimmed = Get-NextNonEmptyTrimmedLine -Lines $lines -StartIndex ($lineIndex + 1)

            if ($nextTrimmed -eq '```' -or
                $nextTrimmed -eq '----' -or
                $nextTrimmed -eq '....' -or
                $nextTrimmed -match '^\|={3,}$' -or
                $nextTrimmed -match '^image::') {

                $elements.Add((New-Element -Type 'continuation' -Data @{}))
                $lineIndex++
                continue
            }

            if ($paragraphBuffer.Count -gt 0) {
                $paragraphBuffer.Add('__LINEBREAK__')
            }

            $lineIndex++
            continue
        }

        # 行解析中
        if ($trimmed -match '^:sectnums:\s*$') {
            if (-not $Attributes) {
                $Attributes = @{}
            }

            $Attributes['sectnums'] = $true
            continue
        }

        if ($line.TrimEnd().EndsWith('+')) {
            # 行末の + を除去
            $withoutPlus = $line.TrimEnd()
            $withoutPlus = $withoutPlus.Substring(0, $withoutPlus.Length - 1)

            $paragraphBuffer.Add($withoutPlus)
            $paragraphBuffer.Add('__LINEBREAK__')
            $lineIndex++
            continue
        }

        $paragraphBuffer.Add($line)
        $lineIndex++
    }

    Flush-ParagraphBuffer

    return [pscustomobject]@{
        Metadata   = $metadata
        Elements   = $elements
        Attributes = $Attributes
    }
}

function Add-ListParagraph {
    param(
        $Document,
        [string]$Text,
        [int]$Level,
        [int]$ElementLevel,
        $StyleConfig,
        [hashtable]$InternalLinkMap,
        [switch]$Numbered,
        [int]$ListIndex = 1,
        [int]$ParentIndex = 1
    )

    if ($Numbered) {
        if ($Level -eq 1) {
            $prefix = "$($ListIndex)) "
        }
        else {
            $nums = New-Object System.Collections.Generic.List[string]
            for ($i = 0; $i -lt $Level; $i++) {
                if ($listCounters[$i] -gt 0) {
                    $nums.Add([string]$listCounters[$i])
                }
            }

            $prefix = ($nums -join '-') + ')'
        }
    }
    else {
        $prefix = '•' * $ElementLevel + ' '
    }

    $style = $StyleConfig.PSObject.Copy()
    $leftIndent = (($Level - 1) * 18)

    if ($style.PSObject.Properties.Name -contains 'LeftIndent') {
        $style.LeftIndent = $leftIndent
    }
    else {
        $style | Add-Member -NotePropertyName LeftIndent -NotePropertyValue $leftIndent
    }

    Append-TextParagraph `
        -Document $Document `
        -Text ($prefix + $Text) `
        -StyleConfig $style `
        -InternalLinkMap $InternalLinkMap | Out-Null
}

function Build-WordDocument {
    param(
        [object]$Parsed,
        $Config,
        [string]$OutputFullPath
    )

    $justAfterPageBreak = $false
    $lastElementType = $null
    $lastHeadingLevel = 0
    $currentListLevel = 0
    
    # 番号付き箇条書きの番号
    $listCounters = @(0, 0, 0, 0, 0, 0)

    # 章番号
    $headingCounters = @(0, 0, 0, 0, 0, 0)
    $isFirstHeading = $true

    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $adocDir = Split-Path -Parent $inputFullPath

        $metadata = [pscustomobject]@{
            Title     = [string]($Parsed.Metadata.Title)
            Subtitle  = [string]($Parsed.Metadata.Subtitle)
            Author    = [string]($Parsed.Metadata.Author)
            RevNumber = [string]($Parsed.Metadata.RevNumber)
            RevDate   = [string]($Parsed.Metadata.RevDate)
            Copyright = [string]($Parsed.Metadata.Copyright)
        }

        $absoluteOutput = Get-AbsolutePath -Path $OutputFullPath
        $outputDir = Split-Path -Parent $absoluteOutput
        if (-not (Test-Path $outputDir)) {
            [void](New-Item -ItemType Directory -Path $outputDir -Force)
        }
        
        $templatePath = ""
        if (Test-Path $TemplateFullPath) {
            $templatePath = $TemplateFullPath
        }
        else {
            $templatePath = $Config.CoverPage.TemplatePath
        }

        if (Test-Path $templatePath) {
            #
            # 表紙テンプレートをベース文書として開く
            #
            # $templatePath = Get-AbsolutePath `
            #     -Path $templatePath `
            #     -BaseDirectory $adocDir
            $templatePath = Get-AbsolutePath `
                -Path $templatePath 
            
            $document = $word.Documents.Open($templatePath, $false, $true)
            $document.SaveAs($absoluteOutput)

            #
            # プレースホルダー置換
            #
            Edit-CoverPage `
                -Document $document `
                -Metadata $metadata `
                -Config $Config
            
            $usingCoverTemplate = $true
        }
        else {

            #
            # 新規文書
            #
            $document = $word.Documents.Add()

            #
            # 旧来の動的表紙生成
            #
            Add-CoverPage `
                -Document $document `
                -Config $Config `
                -Metadata $metadata
        }
        $hasTitlePage = $true

        if ($Config.Document -and $Config.Document.PageSetup) {
            if ($Config.Document.PageSetup.TopMarginMm) { $document.PageSetup.TopMargin = $word.MillimetersToPoints([double]$Config.Document.PageSetup.TopMarginMm) }
            if ($Config.Document.PageSetup.BottomMarginMm) { $document.PageSetup.BottomMargin = $word.MillimetersToPoints([double]$Config.Document.PageSetup.BottomMarginMm) }
            if ($Config.Document.PageSetup.LeftMarginMm) { $document.PageSetup.LeftMargin = $word.MillimetersToPoints([double]$Config.Document.PageSetup.LeftMarginMm) }
            if ($Config.Document.PageSetup.RightMarginMm) { $document.PageSetup.RightMargin = $word.MillimetersToPoints([double]$Config.Document.PageSetup.RightMarginMm) }
        }

        $ownerInserted = $false
        $hasTitlePage = $false
        $tocInserted = $false
        $insideListContinuation = $false
        $pendingAnchorBookmark = $null

        $internalLinkMap = @{}
        $bookmarkNames = @{}
        foreach ($e in $Parsed.Elements) {
            if ($e.Type -ne 'anchor') { continue }

            $anchorId = [string]$e.Id
            if ([string]::IsNullOrWhiteSpace($anchorId)) { continue }
            if ($internalLinkMap.ContainsKey($anchorId)) { continue }

            $internalLinkMap[$anchorId] = Convert-AnchorIdToBookmarkName -AnchorId $anchorId -UsedNames $bookmarkNames
        }
        
        foreach ($element in $Parsed.Elements) {
            if ($element.Type -eq 'anchor') {
                $anchorId = [string]$element.Id
                if (-not [string]::IsNullOrWhiteSpace($anchorId) -and $internalLinkMap.ContainsKey($anchorId)) {
                    $pendingAnchorBookmark = [string]$internalLinkMap[$anchorId]
                }
                continue
            }

            if (-not $tocInserted -and $hasTitlePage -and $element.Type -notin @('title', 'anchor')) {
                Add-SectionBreakToDocument -Document $document
                $justAfterPageBreak = $true

                Add-TableOfContents -Document $document -Config $Config
                Add-SectionBreakToDocument -Document $document
                Set-BodyFooterPageNumber -Document $document                
                $justAfterPageBreak = $true
                $tocInserted = $true
            }

            # 箇条書きの連番リセット
            $resetList = $false

            switch ($element.Type) {

                'heading' { $resetList = $true }
                #'pagebreak'  { $resetList = $true }
                'title' { $resetList = $true }
                #'admonition' { $resetList = $true }
                #'table'      { $resetList = $true }
                #'image'      { $resetList = $true }

                default { $resetList = $false }
            }

            if ($resetList) {
                $listCounters = @(0, 0, 0, 0, 0, 0)
                $currentListLevel = 0
            }

            if (-not [string]::IsNullOrWhiteSpace($pendingAnchorBookmark)) {
                $anchorRange = $document.Content
                $anchorRange.Collapse((Get-WordConstant 'wdCollapseEnd'))
                [void](Add-WordBookmarkSafe -Document $document -Range $anchorRange -BookmarkName $pendingAnchorBookmark)
                $pendingAnchorBookmark = $null
            }

            switch ($element.Type) {
                'title' {
                    
                    #Set-HeaderFooter -Document $document -Config $Config -Metadata $metadata -IsCovePage $true
                    
                    $hasTitlePage = $true
                }
                'heading' {

                    $level = [int]$element.Level
                    if ($level -lt 1) { $level = 1 }
                    if ($level -gt 6) { $level = 6 }

                    # =============================
                    # ① 章（Level1）
                    # =============================
                    if ($level -eq 1) {

                        # 直前が改ページでなければ改ページ
                        if (-not $justAfterPageBreak -and -not $isFirstHeading) {
                            Add-PageBreakToDocument -Document $document
                            $justAfterPageBreak = $true
                        }
                    }

                    # =============================
                    # ② 節（Level2）以降
                    # =============================
                    else {

                        # 条件で空行挿入
                        if (
                            -not $justAfterPageBreak -and
                            -not ($lastElementType -eq 'heading')
                            #-not ($lastElementType -eq 'heading' -and $lastHeadingLevel -eq 1)
                        ) {
                            Append-BlankParagraph -Document $document
                        }
                    }

                    # =============================
                    # 出力（既存そのまま）
                    # =============================
                    $isFirstHeading = $false

                    $headingCounters[$level - 1]++
                    for ($i = $level; $i -lt $headingCounters.Length; $i++) {
                        $headingCounters[$i] = 0
                    }

                    $nums = New-Object System.Collections.Generic.List[string]
                    for ($i = 0; $i -lt $level; $i++) {
                        if ($headingCounters[$i] -gt 0) {
                            $nums.Add([string]$headingCounters[$i])
                        }
                    }

                    $useSectionNums = $Parsed.Attributes.ContainsKey('sectnums') -or
                    ($Config.Document -and
                    $Config.Document.PSObject.Properties.Name -contains 'SectionNumbers' -and
                    [bool]$Config.Document.SectionNumbers)
                    if ($useSectionNums -and $nums.Count -gt 0) {
                        $headingText = ($nums -join '.') + ' ' + $element.Text.Trim()
                    }
                    else {
                        $headingText = $element.Text.Trim()
                    }

                    $styleName = 'Heading' + [string]$level
                    $styleConfig = $Config.Styles.$styleName
                    if (-not $styleConfig) { $styleConfig = $Config.Styles.HeadingDefault }

                    Append-HeadingParagraph `
                        -Document $document `
                        -Text $headingText `
                        -Level $level `
                        -StyleConfig $styleConfig | Out-Null

                    # =============================
                    # 状態更新
                    # =============================
                    $lastElementType = 'heading'
                    $lastHeadingLevel = $level
                    $justAfterPageBreak = $false
                }

                'paragraph' {
                    $style = $Config.Styles.Body
                    $text = [string]$element.Text

                    if ($currentListLevel -gt 0) {
                        $style = $style.PSObject.Copy()
                        $leftIndent = 18 + (($currentListLevel - 1) * 18)

                        if ($style.PSObject.Properties.Name -contains 'LeftIndent') {
                            $style.LeftIndent = $leftIndent
                        }
                        else {
                            $style | Add-Member -NotePropertyName LeftIndent -NotePropertyValue $leftIndent
                        }

                        # リスト配下の継続段落だけ、AsciiDoc上の見た目インデントを除去する。
                        $text = $text -replace '(^|[\r\n])[ \t]+', '$1'
                    }

                    Append-TextParagraph `
                        -Document $document `
                        -Text $text `
                        -StyleConfig $style `
                        -InternalLinkMap $internalLinkMap | Out-Null
                }
                'bullet' {
                    $level = [int]$element.Level
                    if ($currentListLevel -gt 0) {
                        $level = $level + $currentListLevel
                    }

                    Add-ListParagraph `
                        -Document $document `
                        -Text $element.Text `
                        -Level $level `
                        -ElementLevel $element.MarkerCount `
                        -StyleConfig $Config.Styles.Bullet `
                        -InternalLinkMap $internalLinkMap
                }
                'numbered' {
                    $level = [int]$element.Level
                    if ($level -lt 1) { $level = 1 }
                    if ($level -gt 6) { $level = 6 }

                    $currentListLevel = $level

                    $listCounters[$level - 1]++

                    for ($i = $level; $i -lt $listCounters.Length; $i++) {
                        $listCounters[$i] = 0
                    }

                    $parentNo = $listCounters[0]
                    $currentNo = $listCounters[$level - 1]

                    Add-ListParagraph `
                        -Document $document `
                        -Text $element.Text `
                        -Level $level `
                        -ElementLevel $element.MarkerCount `
                        -StyleConfig $Config.Styles.Numbered `
                        -InternalLinkMap $internalLinkMap `
                        -Numbered `
                        -ParentIndex $parentNo `
                        -ListIndex $currentNo
                }
                'admonition' { 

                    $formattedText = Format-AdmonitionText -Kind $element.Kind -Text $element.Text -LabelConfig $Config.Styles.Admonition.Label

                    $range = Append-TextParagraph `
                        -Document $document `
                        -Text $formattedText `
                        -StyleConfig $Config.Styles.Admonition `
                        -InternalLinkMap $internalLinkMap

                    # 種類別 背景色
                    $colorHex = $Config.Styles.Admonition.Colors.$($element.Kind)
                    if ($colorHex) {
                        try {
                            $wordColor = Convert-RgbToWordColor $colorHex
                            $range.Paragraphs(1).Shading.BackgroundPatternColor =
                            [int]("0x00$wordColor")
                        }
                        catch {
                            # 失敗しても致命的ではないので握りつぶし
                        }
                    }
                }
                'code' {
                    $style = $Config.Styles.Code

                    if ($element.Caption) {
                        Append-TextParagraph `
                            -Document $document `
                            -Text $element.Caption `
                            -StyleConfig $Config.Styles.CodeCaption `
                            -InternalLinkMap $internalLinkMap | Out-Null
                    }

                    if (($insideListContinuation -or $currentListLevel -gt 0) -and $currentListLevel -gt 0) {
                        $style = $style.PSObject.Copy()
                        $leftIndent = 36 + (($currentListLevel - 1) * 18)
                        if ($style.PSObject.Properties.Name -contains 'LeftIndent') {
                            $style.LeftIndent = $leftIndent
                        }
                        else {
                            $style | Add-Member -NotePropertyName LeftIndent -NotePropertyValue $leftIndent
                        }
                    }

                    $codeText = [string]$element.Text
                    $codeLines = $codeText -split "`r?`n", -1

                    foreach ($codeLine in $codeLines) {
                        Append-TextParagraph `
                            -Document $document `
                            -Text $codeLine `
                            -StyleConfig $style `
                            -InternalLinkMap $internalLinkMap | Out-Null
                    }

                    Append-BlankParagraph -Document $document
                    $insideListContinuation = $false
                }
                'pagebreak' { 
                    Add-PageBreakToDocument -Document $document 
                    $justAfterPageBreak = $true
                }
                'image' { 
                    Add-ImageToDocument -Document $document -ImagePath $element.Path -Caption $element.Caption -Config $Config 
                }
                'table' { 
                    Write-DebugLog "BUILD-TABLE caption=[$($element.Caption)] rows=$($element.tableInfo.Rows.Count)"
                    
                    Add-WordTable -Document $document -Rows $element.tableInfo.Rows -ColumnCount $element.tableInfo.MaxColumns  -Config $Config -Caption $element.Caption -Attributes $element.Attributes -InternalLinkMap $internalLinkMap
                }
                'continuation' {
                    $insideListContinuation = $true
                }
                'tablecaption' {
                    Append-TextParagraph `
                        -Document $document `
                        -Text $element.Text `
                        -StyleConfig $Config.Styles.TableCaption `
                        -InternalLinkMap $internalLinkMap | Out-Null
                }                
                'include-source' {
                    $srcStyle = if ($element.StyleConfig) { $element.StyleConfig } else { $Config.Styles.IncludeSource }
                    if (-not $srcStyle) {
                        $srcStyle = [pscustomobject]@{ FontName = '游ゴシック'; Size = 9; Bold = $false; Italic = $true; Color = 8421504 }
                    }
                    Append-TextParagraph `
                        -Document $document `
                        -Text "[ソース: $($element.Text)]" `
                        -StyleConfig $srcStyle `
                        -InternalLinkMap $internalLinkMap | Out-Null
                }
                default { 
                    Append-TextParagraph -Document $document -Text ('[未対応要素: ' + $element.Type + ']') -StyleConfig $Config.Styles.Body -InternalLinkMap $internalLinkMap | Out-Null 
                }
            }
            # if ($element.Type -ne 'pagebreak') { 
            #     $script:pageBreaked = $false
            # }

            switch ($element.Type) {

                'heading' {
                    # すでにheading内で更新してるので何もしない
                }

                'pagebreak' {
                    $justAfterPageBreak = $true
                    # lastElementTypeは変更しない
                }

                { $_ -in @('paragraph', 'bullet', 'numbered', 'image', 'table', 'admonition', 'tablecaption') } {
                    $lastElementType = 'paragraph'
                    $justAfterPageBreak = $false
                }

                default {
                    # 基本は本文扱い
                    $lastElementType = 'paragraph'
                    $justAfterPageBreak = $false
                }
            }            
        }

        if (-not $tocInserted -and $hasTitlePage) {
            Add-PageBreakToDocument -Document $document
            Add-TableOfContents -Document $document -Config $Config

            # 目次の直後に本文開始用の改ページを1つだけ入れる
            $range = $document.Content
            $range.Collapse((Get-WordConstant 'wdCollapseEnd'))
            $range.InsertBreak((Get-WordConstant 'wdPageBreak'))

            $tocInserted = $true
        }

        #Set-HeaderFooter -Document $document -Config $Config -Metadata $metadata

        try {
            foreach ($toc in $document.TablesOfContents) {
                $toc.Update() | Out-Null
            }
        }
        catch {
        }

        $document.SaveAs([ref]$absoluteOutput, [ref](Get-WordConstant 'wdSaveFormatDocumentDefault'))

        # PDF出力
        try {
            $pdfPath = [System.IO.Path]::ChangeExtension($absoluteOutput, '.pdf')

            # 17 = wdExportFormatPDF
            $document.ExportAsFixedFormat(
                $pdfPath,
                17
            )

            Write-Output "PDF出力完了: $pdfPath"
        }
        catch {
            Write-Warning "PDF出力失敗: $($_.Exception.Message)"
        }

        $document.Close()
        $word.Quit()
    }
    finally {
        if ($null -ne $document) {
            try { 
                $document.Close() | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null 
            }
            catch {}
        }
        if ($null -ne $word) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Normalize-MarkdownInlineText {
    param(
        [string]$Text,
        [hashtable]$Attributes
    )

    if ($null -eq $Text) { return '' }

    $value = $Text

    if ($Attributes) {
        foreach ($k in $Attributes.Keys) {
            $token = '{' + $k + '}'
            $value = $value.Replace($token, [string]$Attributes[$k])
        }
    }

    # 画像リンク ![alt](path) → テキスト表記
    $value = [regex]::Replace($value, '!\[([^\]]*)\]\([^\)]+\)', '[画像: $1]')

    # Markdown リンク [label](url) → link:url[label] (Parse-InlineLinkSegments 互換)
    $value = [regex]::Replace($value, '\[([^\]]+)\]\(([^\)]+)\)', 'link:$2[$1]')

    # *italic* (単一アスタリスク) → _italic_
    $value = [regex]::Replace($value, '(?<!\*)\*(?!\*)([^*\n]+?)(?<!\*)\*(?!\*)', '_$1_')

    # インラインコードのバッククォートを除去
    $value = [regex]::Replace($value, '`([^`]+)`', '$1')

    # 打消し線を除去
    $value = [regex]::Replace($value, '~~([^~]+)~~', '$1')

    return $value.TrimEnd()
}

function Get-JsonIncludeErrorText {
    param(
        [string]$ErrorCode,
        [string]$FilePath
    )
    return "[ERROR]`n$ErrorCode`n$FilePath"
}

function Get-JsonTreeChildren {
    param($Node)

    $children = New-Object System.Collections.Generic.List[object]

    if ($Node -is [System.Management.Automation.PSCustomObject]) {
        foreach ($prop in $Node.PSObject.Properties) {
            $children.Add([pscustomobject]@{
                    Label = [string]$prop.Name
                    Value = $prop.Value
                })
        }
    }
    elseif ($Node -is [System.Object[]]) {
        $priorityKeys = @('id', 'name', 'screenId', 'itemId', 'label')
        $i = 0
        foreach ($item in $Node) {
            $label = "[$i]"
            if ($item -is [System.Management.Automation.PSCustomObject]) {
                foreach ($key in $priorityKeys) {
                    $prop = $item.PSObject.Properties[$key]
                    if ($null -ne $prop -and
                        $null -ne $prop.Value -and
                        $prop.Value -isnot [System.Object[]] -and
                        $prop.Value -isnot [System.Management.Automation.PSCustomObject]) {
                        $label = [string]$prop.Value
                        break
                    }
                }
            }
            elseif ($null -ne $item -and
                $item -isnot [System.Object[]] -and
                $item -isnot [System.Management.Automation.PSCustomObject]) {
                $label = [string]$item
            }
            $children.Add([pscustomobject]@{
                    Label = $label
                    Value = $item
                })
            $i++
        }
    }

    return $children
}

function Build-JsonTreeLines {
    param(
        $Node,
        [string]$Prefix = '',
        [bool]$IsTopLevel = $true
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $children = @(Get-JsonTreeChildren -Node $Node)

    for ($i = 0; $i -lt $children.Count; $i++) {
        $child = $children[$i]
        $isLast = ($i -eq $children.Count - 1)

        $hasChildren = ($child.Value -is [System.Management.Automation.PSCustomObject]) -or
        ($child.Value -is [System.Object[]])

        # leaf node: append scalar value as "key: value"
        $displayLabel = $child.Label
        if (-not $hasChildren -and $null -ne $child.Value) {
            $valStr = if ($child.Value -is [bool]) { if ($child.Value) { 'true' } else { 'false' } }
            else { [string]$child.Value }
            if (-not [string]::IsNullOrWhiteSpace($valStr)) {
                $displayLabel = "$($child.Label): $valStr"
            }
        }

        if ($IsTopLevel) {
            $lines.Add($displayLabel)
            $childPrefix = ''
        }
        else {
            $connector = if ($isLast) { '└─ ' } else { '├─ ' }
            $lines.Add($Prefix + $connector + $displayLabel)
            $childPrefix = if ($isLast) { $Prefix + '   ' } else { $Prefix + '│  ' }
        }

        if ($hasChildren) {
            $subLines = @(Build-JsonTreeLines -Node $child.Value -Prefix $childPrefix -IsTopLevel $false)
            foreach ($sl in $subLines) {
                $lines.Add($sl)
            }
        }
    }

    return $lines
}

function Invoke-JsonTableDirective {
    param(
        [string]$FilePath,
        [string]$BaseDirectory
    )

    $result = @{
        Success   = $false
        TableInfo = $null
        ErrorCode = $null
        FilePath  = $FilePath
        Metadata  = $null
    }

    $absPath = $null
    try {
        $absPath = Get-AbsolutePath -Path $FilePath -BaseDirectory $BaseDirectory
    }
    catch {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    if (-not (Test-Path -LiteralPath $absPath -PathType Leaf)) {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    $rawJson = $null
    try {
        $rawJson = Get-Content -LiteralPath $absPath -Raw -Encoding UTF8
    }
    catch {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    if ([string]::IsNullOrWhiteSpace($rawJson)) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    $jsonData = $null
    try {
        $jsonData = $rawJson | ConvertFrom-Json
    }
    catch {
        $lineInfo = ''
        if ($_.Exception.Message -match '\((\d+)\):') {
            $charPos = [int]$matches[1]
            $before = if ($charPos -le $rawJson.Length) { $rawJson.Substring(0, $charPos) } else { $rawJson }
            $lineNum = ($before -split "`n").Count
            $lineInfo = " (line $lineNum)"
        }
        $result.ErrorCode = "Invalid JSON$lineInfo"
        return [pscustomobject]$result
    }

    # _* プロパティを除外しメタデータ・tableinfo を収集してデータ配列を特定する
    $tableInfoDef = $null
    if ($jsonData -is [System.Management.Automation.PSCustomObject]) {
        $metaDesc = $null
        $metaRules = $null
        $found = $null

        foreach ($prop in $jsonData.PSObject.Properties) {
            $pname = [string]$prop.Name
            $val = $prop.Value

            if ($pname.StartsWith('_')) { continue }

            if ($pname -eq 'description' -and $val -is [string]) {
                $metaDesc = $val
                continue
            }
            if ($pname -eq 'rules') {
                if ($val -is [System.Object[]]) { $metaRules = $val }
                elseif ($val -is [string]) { $metaRules = @($val) }
                continue
            }
            if ($pname -eq 'columns' -and $val -is [System.Management.Automation.PSCustomObject]) {
                $tableInfoDef = $val
                continue
            }
            if ($null -eq $found -and $val -is [System.Object[]] -and $val.Count -gt 0 -and
                $val[0] -is [System.Management.Automation.PSCustomObject]) {
                $found = $val
            }
        }

        if ($null -eq $found) {
            $result.ErrorCode = 'Unsupported JSON structure'
            return [pscustomobject]$result
        }
        $jsonData = $found

        if ($null -ne $metaDesc -or $null -ne $metaRules) {
            $result.Metadata = [pscustomobject]@{
                Description = $metaDesc
                Rules       = $metaRules
            }
        }
    }

    if ($jsonData -isnot [System.Object[]]) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    if ($jsonData.Count -eq 0) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    $columns = @()
    $colHeaders = @{}

    if ($null -ne $tableInfoDef) {
        # tableinfo が定義されている場合はその順序・タイトルで列を確定する
        foreach ($tiProp in $tableInfoDef.PSObject.Properties) {
            $cname = [string]$tiProp.Name
            $columns += $cname
            $colHeaders[$cname] = [string]$tiProp.Value
        }
        foreach ($item in $jsonData) {
            if ($item -isnot [System.Management.Automation.PSCustomObject]) {
                $result.ErrorCode = 'Unsupported JSON structure'
                return [pscustomobject]$result
            }
        }
    }
    else {
        # 全要素を走査してスカラープロパティ名を収集（_* およびネスト型は除外）
        $colOrder = New-Object System.Collections.Specialized.OrderedDictionary
        $nestedCols = New-Object System.Collections.Generic.HashSet[string]

        foreach ($item in $jsonData) {
            if ($item -isnot [System.Management.Automation.PSCustomObject]) {
                $result.ErrorCode = 'Unsupported JSON structure'
                return [pscustomobject]$result
            }
            foreach ($prop in $item.PSObject.Properties) {
                $pname = [string]$prop.Name
                if ($pname.StartsWith('_')) { continue }
                $val = $prop.Value
                if ($val -is [System.Object[]] -or $val -is [System.Management.Automation.PSCustomObject]) {
                    [void]$nestedCols.Add($pname)
                }
                elseif (-not $colOrder.Contains($pname)) {
                    $colOrder[$pname] = $true
                }
            }
        }

        foreach ($n in $nestedCols) { $colOrder.Remove($n) }
        foreach ($k in $colOrder.Keys) { $columns += [string]$k }
    }

    if ($columns.Count -eq 0) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    $rows = New-Object System.Collections.Generic.List[object]

    $headerCells = @()
    foreach ($col in $columns) {
        $hdr = if ($colHeaders.ContainsKey($col)) { $colHeaders[$col] } else { $col }
        $headerCells += @{ Text = $hdr; RowSpan = 1; ColSpan = 1; IsHeader = $true }
    }
    $rows.Add($headerCells)

    foreach ($item in $jsonData) {
        $cells = @()
        foreach ($col in $columns) {
            $prop = $item.PSObject.Properties[$col]
            $val = if ($null -ne $prop) { $prop.Value } else { $null }
            if ($val -is [System.Object[]]) {
                # スカラー要素の配列はカンマ区切りで出力
                $parts = @()
                foreach ($v in $val) {
                    if ($v -isnot [System.Object[]] -and $v -isnot [System.Management.Automation.PSCustomObject]) {
                        $parts += [string]$v
                    }
                }
                $val = if ($parts.Count -gt 0) { $parts -join ', ' } else { $null }
            }
            elseif ($val -is [System.Management.Automation.PSCustomObject]) {
                $val = $null
            }
            $text = if ($null -eq $val) { '' }
            elseif ($val -is [bool]) { if ($val) { 'true' } else { 'false' } }
            else { [string]$val }
            $cells += @{ Text = $text; RowSpan = 1; ColSpan = 1; IsHeader = $false }
        }
        $rows.Add($cells)
    }

    $result.Success = $true
    $result.TableInfo = [pscustomobject]@{
        Rows       = $rows.ToArray()
        MaxColumns = $columns.Count
    }
    return [pscustomobject]$result
}

function Invoke-JsonTreeDirective {
    param(
        [string]$FilePath,
        [string]$BaseDirectory
    )

    $result = @{
        Success   = $false
        TreeText  = $null
        ErrorCode = $null
        FilePath  = $FilePath
    }

    $absPath = $null
    try {
        $absPath = Get-AbsolutePath -Path $FilePath -BaseDirectory $BaseDirectory
    }
    catch {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    if (-not (Test-Path -LiteralPath $absPath -PathType Leaf)) {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    $rawJson = $null
    try {
        $rawJson = Get-Content -LiteralPath $absPath -Raw -Encoding UTF8
    }
    catch {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    if ([string]::IsNullOrWhiteSpace($rawJson)) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    $jsonData = $null
    try {
        $jsonData = $rawJson | ConvertFrom-Json
    }
    catch {
        $lineInfo = ''
        if ($_.Exception.Message -match '\((\d+)\):') {
            $charPos = [int]$matches[1]
            $before = if ($charPos -le $rawJson.Length) { $rawJson.Substring(0, $charPos) } else { $rawJson }
            $lineNum = ($before -split "`n").Count
            $lineInfo = " (line $lineNum)"
        }
        $result.ErrorCode = "Invalid JSON$lineInfo"
        return [pscustomobject]$result
    }

    if ($null -eq $jsonData) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    try {
        $lines = Build-JsonTreeLines -Node $jsonData -Prefix '' -IsTopLevel $true
        $result.TreeText = ($lines -join "`n")
        $result.Success = $true
    }
    catch {
        $result.ErrorCode = 'Unsupported JSON structure'
    }

    return [pscustomobject]$result
}

function Invoke-JsonConfigDirective {
    param(
        [string]$FilePath,
        [string]$BaseDirectory
    )

    $result = @{
        Success     = $false
        ErrorCode   = $null
        FilePath    = $FilePath
        Description = $null
        Rules       = $null
        Categories  = @()
    }

    $absPath = $null
    try {
        $absPath = Get-AbsolutePath -Path $FilePath -BaseDirectory $BaseDirectory
    }
    catch {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    if (-not (Test-Path -LiteralPath $absPath -PathType Leaf)) {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    $rawJson = $null
    try {
        $rawJson = Get-Content -LiteralPath $absPath -Raw -Encoding UTF8
    }
    catch {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    if ([string]::IsNullOrWhiteSpace($rawJson)) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    $jsonData = $null
    try {
        $jsonData = $rawJson | ConvertFrom-Json
    }
    catch {
        $lineInfo = ''
        if ($_.Exception.Message -match '\((\d+)\):') {
            $charPos = [int]$matches[1]
            $before = if ($charPos -le $rawJson.Length) { $rawJson.Substring(0, $charPos) } else { $rawJson }
            $lineNum = ($before -split "`n").Count
            $lineInfo = " (line $lineNum)"
        }
        $result.ErrorCode = "Invalid JSON$lineInfo"
        return [pscustomobject]$result
    }

    if ($jsonData -isnot [System.Management.Automation.PSCustomObject]) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    $rootDesc = $null
    $rootRules = $null
    $categories = New-Object System.Collections.Generic.List[object]

    foreach ($prop in $jsonData.PSObject.Properties) {
        $pname = [string]$prop.Name
        $val = $prop.Value

        if ($pname.StartsWith('_')) { continue }

        if ($pname -eq 'description' -and $val -is [string]) {
            $rootDesc = $val
            continue
        }
        if ($pname -eq 'rules') {
            if ($val -is [System.Object[]]) { $rootRules = $val }
            elseif ($val -is [string]) { $rootRules = @($val) }
            continue
        }

        if ($val -is [System.Management.Automation.PSCustomObject]) {
            $catDesc = $null
            $constants = New-Object System.Collections.Generic.List[object]

            foreach ($cProp in $val.PSObject.Properties) {
                $cname = [string]$cProp.Name
                $cval = $cProp.Value

                if ($cname -eq '_description') {
                    if ($cval -is [string]) { $catDesc = $cval }
                    continue
                }
                if ($cname.StartsWith('_')) { continue }

                if ($cval -is [System.Management.Automation.PSCustomObject]) {
                    $vProp = $cval.PSObject.Properties['value']
                    $dProp = $cval.PSObject.Properties['description']

                    $constValue = ''
                    if ($null -ne $vProp) {
                        $v = $vProp.Value
                        if ($v -is [System.Object[]]) {
                            $parts = @()
                            foreach ($item in $v) {
                                if ($item -isnot [System.Object[]] -and
                                    $item -isnot [System.Management.Automation.PSCustomObject]) {
                                    $parts += [string]$item
                                }
                            }
                            $constValue = $parts -join ', '
                        }
                        elseif ($v -is [bool]) {
                            $constValue = if ($v) { 'true' } else { 'false' }
                        }
                        elseif ($null -ne $v) {
                            $constValue = [string]$v
                        }
                    }

                    $constDesc = ''
                    if ($null -ne $dProp -and $dProp.Value -is [string]) {
                        $constDesc = [string]$dProp.Value
                    }

                    $constants.Add([pscustomobject]@{
                            Name        = $cname
                            Value       = $constValue
                            Description = $constDesc
                        })
                }
            }

            $categories.Add([pscustomobject]@{
                    Key         = $pname
                    Description = $catDesc
                    Constants   = $constants.ToArray()
                })
        }
    }

    $result.Success = $true
    $result.Description = $rootDesc
    $result.Rules = $rootRules
    $result.Categories = $categories.ToArray()
    return [pscustomobject]$result
}

function Invoke-JsonEnumDirective {
    param(
        [string]$FilePath,
        [string]$BaseDirectory
    )

    $result = @{
        Success     = $false
        ErrorCode   = $null
        FilePath    = $FilePath
        Description = $null
        Rules       = $null
        Enums       = @()
    }

    $absPath = $null
    try {
        $absPath = Get-AbsolutePath -Path $FilePath -BaseDirectory $BaseDirectory
    }
    catch {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    if (-not (Test-Path -LiteralPath $absPath -PathType Leaf)) {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    $rawJson = $null
    try {
        $rawJson = Get-Content -LiteralPath $absPath -Raw -Encoding UTF8
    }
    catch {
        $result.ErrorCode = 'File not found'
        return [pscustomobject]$result
    }

    if ([string]::IsNullOrWhiteSpace($rawJson)) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    $jsonData = $null
    try {
        $jsonData = $rawJson | ConvertFrom-Json
    }
    catch {
        $lineInfo = ''
        if ($_.Exception.Message -match '\((\d+)\):') {
            $charPos = [int]$matches[1]
            $before = if ($charPos -le $rawJson.Length) { $rawJson.Substring(0, $charPos) } else { $rawJson }
            $lineNum = ($before -split "`n").Count
            $lineInfo = " (line $lineNum)"
        }
        $result.ErrorCode = "Invalid JSON$lineInfo"
        return [pscustomobject]$result
    }

    if ($jsonData -isnot [System.Management.Automation.PSCustomObject]) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    $rootDesc = $null
    $rootRules = $null
    $columnDefs = $null
    $enumsRaw = $null

    foreach ($prop in $jsonData.PSObject.Properties) {
        $pname = [string]$prop.Name
        $val = $prop.Value
        if ($pname.StartsWith('_')) { continue }
        switch ($pname) {
            'description' { if ($val -is [string]) { $rootDesc = $val } }
            'rules' {
                if ($val -is [System.Object[]]) { $rootRules = $val }
                elseif ($val -is [string]) { $rootRules = @($val) }
            }
            'columns' { if ($val -is [System.Management.Automation.PSCustomObject]) { $columnDefs = $val } }
            'enums' {
                if ($val -is [System.Object[]]) { $enumsRaw = $val }
                elseif ($val -is [System.Management.Automation.PSCustomObject]) { $enumsRaw = @($val) }
            }
        }
    }

    if ($null -eq $enumsRaw) {
        $result.ErrorCode = 'Unsupported JSON structure'
        return [pscustomobject]$result
    }

    # Build base column name→header map from 'columns' definition
    $baseColNames = @()
    $baseColHeaders = @{}
    if ($null -ne $columnDefs) {
        foreach ($cProp in $columnDefs.PSObject.Properties) {
            $cname = [string]$cProp.Name
            $baseColNames += $cname
            $baseColHeaders[$cname] = [string]$cProp.Value
        }
    }

    $enumList = New-Object System.Collections.Generic.List[object]

    foreach ($enumItem in $enumsRaw) {
        if ($enumItem -isnot [System.Management.Automation.PSCustomObject]) { continue }

        $enumNameProp = $enumItem.PSObject.Properties['enumName']
        if ($null -eq $enumNameProp) {
            $enumList.Add([pscustomobject]@{
                    EnumName    = '[enumNameなし]'
                    Description = $null
                    HasValues   = $false
                    TableInfo   = $null
                    IsError     = $true
                })
            continue
        }

        $enumName = [string]$enumNameProp.Value
        $enumDesc = $null
        $enumDescProp = $enumItem.PSObject.Properties['description']
        if ($null -ne $enumDescProp -and $enumDescProp.Value -is [string]) {
            $enumDesc = [string]$enumDescProp.Value
        }

        $valuesProp = $enumItem.PSObject.Properties['values']
        $valuesRaw = if ($null -ne $valuesProp) { $valuesProp.Value } else { $null }
        if ($valuesRaw -is [System.Management.Automation.PSCustomObject]) { $valuesRaw = @($valuesRaw) }
        if ($null -eq $valuesRaw -or $valuesRaw -isnot [System.Object[]]) {
            $enumList.Add([pscustomobject]@{
                    EnumName    = $enumName
                    Description = $enumDesc
                    HasValues   = $false
                    TableInfo   = $null
                    IsError     = $false
                })
            continue
        }

        # Collect columns: base columns first, then extra columns found in values
        $colOrder = New-Object System.Collections.Specialized.OrderedDictionary
        foreach ($c in $baseColNames) { $colOrder[$c] = $true }
        foreach ($vItem in $valuesRaw) {
            if ($vItem -isnot [System.Management.Automation.PSCustomObject]) { continue }
            foreach ($vProp in $vItem.PSObject.Properties) {
                $vpname = [string]$vProp.Name
                if ($vpname.StartsWith('_')) { continue }
                if (-not $colOrder.Contains($vpname)) { $colOrder[$vpname] = $true }
            }
        }
        $columns = @()
        foreach ($k in $colOrder.Keys) { $columns += [string]$k }

        # Build table
        $rows = New-Object System.Collections.Generic.List[object]
        $headerCells = @()
        foreach ($col in $columns) {
            $hdr = if ($baseColHeaders.ContainsKey($col)) { $baseColHeaders[$col] } else { $col }
            $headerCells += @{ Text = $hdr; RowSpan = 1; ColSpan = 1; IsHeader = $true }
        }
        $rows.Add($headerCells)

        foreach ($vItem in $valuesRaw) {
            if ($vItem -isnot [System.Management.Automation.PSCustomObject]) { continue }
            $dataCells = @()
            foreach ($col in $columns) {
                $vp = $vItem.PSObject.Properties[$col]
                $val = if ($null -ne $vp) { $vp.Value } else { $null }
                if ($val -is [System.Object[]]) {
                    $parts = @()
                    foreach ($item in $val) {
                        if ($item -isnot [System.Object[]] -and $item -isnot [System.Management.Automation.PSCustomObject]) {
                            $parts += [string]$item
                        }
                    }
                    $val = $parts -join ', '
                }
                elseif ($val -is [System.Management.Automation.PSCustomObject]) { $val = $null }
                $text = if ($null -eq $val) { '' }
                elseif ($val -is [bool]) { if ($val) { 'true' } else { 'false' } }
                else { [string]$val }
                $dataCells += @{ Text = $text; RowSpan = 1; ColSpan = 1; IsHeader = $false }
            }
            $rows.Add($dataCells)
        }

        $enumList.Add([pscustomobject]@{
                EnumName    = $enumName
                Description = $enumDesc
                HasValues   = $true
                TableInfo   = [pscustomobject]@{ Rows = $rows.ToArray(); MaxColumns = $columns.Count }
                IsError     = $false
            })
    }

    $result.Success = $true
    $result.Description = $rootDesc
    $result.Rules = $rootRules
    $result.Enums = $enumList.ToArray()
    return [pscustomobject]$result
}

function Build-WorkflowPlantUml {
    param(
        $Workflow,
        $Actors,
        [string]$Filter   # $null=all, 'PREPAID', 'POSTPAID'
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('@startuml')
    [void]$sb.AppendLine('')

    foreach ($sid in $Workflow.States.Keys) {
        $lbl = ([string]$Workflow.States[$sid].Label) -replace '"', "'"
        [void]$sb.AppendLine("state `"$lbl`" as $sid")
    }
    [void]$sb.AppendLine('')

    foreach ($t in $Workflow.Transitions) {
        $cond = $t.Condition
        if ($null -ne $Filter) {
            if (-not [string]::IsNullOrWhiteSpace($cond)) {
                if ($cond -notmatch $Filter) { continue }
            }
        }

        $actorLabel = if (-not [string]::IsNullOrWhiteSpace($t.Actor) -and $Actors.ContainsKey($t.Actor)) {
            $Actors[$t.Actor]
        }
        else { $t.Actor }

        $label = "$($t.Trigger)\n$actorLabel"
        if (-not [string]::IsNullOrWhiteSpace($cond)) { $label += "\n[$cond]" }

        foreach ($fromVal in $t.FromList) {
            $arrow = if ($fromVal -eq '[*]') { "[*] --> $($t.To)" } else { "$fromVal --> $($t.To)" }
            [void]$sb.AppendLine("$arrow : $label")
        }
    }

    [void]$sb.AppendLine('')
    [void]$sb.Append('@enduml')
    return $sb.ToString()
}

function Invoke-JsonWorkflowDirective {
    param(
        [string]$FilePath,
        [string]$BaseDirectory
    )

    $result = @{
        Success   = $false
        ErrorCode = $null
        FilePath  = $FilePath
        Actors    = @{}
        Workflows = @()
    }

    if (-not (Get-Command ConvertFrom-Yaml -ErrorAction SilentlyContinue)) {
        try { Import-Module powershell-yaml -ErrorAction Stop }
        catch {
            $result.ErrorCode = 'YAML module required. Run: Install-Module powershell-yaml'
            return [pscustomobject]$result
        }
    }

    $absPath = $null
    try { $absPath = Get-AbsolutePath -Path $FilePath -BaseDirectory $BaseDirectory }
    catch { $result.ErrorCode = 'File not found'; return [pscustomobject]$result }
    if (-not (Test-Path -LiteralPath $absPath -PathType Leaf)) {
        $result.ErrorCode = 'File not found'; return [pscustomobject]$result
    }

    $rawContent = $null
    try { $rawContent = Get-Content -LiteralPath $absPath -Raw -Encoding UTF8 }
    catch { $result.ErrorCode = 'File not found'; return [pscustomobject]$result }

    $yaml = $null
    try { $yaml = $rawContent | ConvertFrom-Yaml }
    catch { $result.ErrorCode = "Invalid YAML: $($_.Exception.Message)"; return [pscustomobject]$result }

    if ($null -eq $yaml -or $yaml -isnot [System.Collections.IDictionary]) {
        $result.ErrorCode = 'Unsupported YAML structure'; return [pscustomobject]$result
    }

    # Actors
    $actors = @{}
    if ($yaml.Keys -contains 'actors' -and $null -ne $yaml['actors']) {
        $ad = $yaml['actors']
        if ($ad -is [System.Collections.IDictionary]) {
            foreach ($rid in $ad.Keys) {
                $ao = $ad[$rid]
                $lbl = if ($null -ne $ao -and $ao -is [System.Collections.IDictionary] -and $ao.Keys -contains 'label') {
                    [string]$ao['label']
                }
                else { [string]$rid }
                $actors[[string]$rid] = $lbl
            }
        }
    }

    # Workflows: all root keys except 'actors'
    $workflows = New-Object System.Collections.Generic.List[object]

    foreach ($wfKey in $yaml.Keys) {
        if ([string]$wfKey -eq 'actors') { continue }
        $wf = $yaml[$wfKey]
        if ($null -eq $wf -or $wf -isnot [System.Collections.IDictionary]) { continue }

        $wfDesc = if ($wf.Keys -contains 'description') { [string]$wf['description'] } else { $null }

        # States
        $statesMap = [ordered]@{}
        if ($wf.Keys -contains 'states' -and $null -ne $wf['states'] -and $wf['states'] -is [System.Collections.IDictionary]) {
            foreach ($sid in $wf['states'].Keys) {
                $so = $wf['states'][$sid]
                $sl = if ($null -ne $so -and $so -is [System.Collections.IDictionary] -and $so.Keys -contains 'label') { [string]$so['label'] } else { [string]$sid }
                $sd = if ($null -ne $so -and $so -is [System.Collections.IDictionary] -and $so.Keys -contains 'description') { [string]$so['description'] } else { '' }
                $statesMap[[string]$sid] = @{ Label = $sl; Description = $sd }
            }
        }

        # TransitionColumns
        $colDefs = [ordered]@{}
        if ($wf.Keys -contains 'transitionColumns' -and $null -ne $wf['transitionColumns'] -and $wf['transitionColumns'] -is [System.Collections.IDictionary]) {
            foreach ($ck in $wf['transitionColumns'].Keys) { $colDefs[[string]$ck] = [string]$wf['transitionColumns'][$ck] }
        }

        # Transitions
        $transList = New-Object System.Collections.Generic.List[object]
        $hasConditions = $false

        if ($wf.Keys -contains 'transitions' -and $null -ne $wf['transitions']) {
            $td = $wf['transitions']
            if ($td -isnot [System.Object[]]) { $td = @($td) }

            foreach ($t in $td) {
                if ($null -eq $t -or $t -isnot [System.Collections.IDictionary]) { continue }

                $fromRaw = if ($t.Keys -contains 'from') { $t['from'] } else { $null }
                $fromList = @()
                if ($null -eq $fromRaw) {
                    $fromList = @('[*]')
                }
                elseif ($fromRaw -is [System.Object[]]) {
                    foreach ($f in $fromRaw) { $fromList += [string]$f }
                }
                else {
                    $fs = [string]$fromRaw
                    $fromList = if ([string]::IsNullOrWhiteSpace($fs) -or $fs -eq 'null') { @('[*]') } else { @($fs) }
                }

                $toRaw = if ($t.Keys -contains 'to') { $t['to'] } else { $null }
                $toVal = if ($null -eq $toRaw) { '[*]' } else { [string]$toRaw }

                $cond = if ($t.Keys -contains 'condition') { [string]$t['condition'] } else { $null }
                if (-not [string]::IsNullOrWhiteSpace($cond)) { $hasConditions = $true }

                $std = @('from', 'to', 'trigger', 'actor', 'screen', 'notes', 'condition')
                $extras = @{}
                foreach ($ek in $t.Keys) {
                    if ($std -notcontains [string]$ek) { $extras[[string]$ek] = [string]$t[$ek] }
                }

                $transList.Add([pscustomobject]@{
                        FromList  = $fromList
                        To        = $toVal
                        Trigger   = if ($t.Keys -contains 'trigger') { [string]$t['trigger'] } else { '' }
                        Actor     = if ($t.Keys -contains 'actor') { [string]$t['actor'] }   else { '' }
                        Screen    = if ($t.Keys -contains 'screen') { [string]$t['screen'] }  else { '' }
                        Notes     = if ($t.Keys -contains 'notes') { [string]$t['notes'] }   else { '' }
                        Condition = $cond
                        Extras    = $extras
                    })
            }
        }

        $workflows.Add([pscustomobject]@{
                Key           = [string]$wfKey
                Description   = $wfDesc
                States        = $statesMap
                ColDefs       = $colDefs
                Transitions   = $transList.ToArray()
                HasConditions = $hasConditions
            })
    }

    $result.Success = $true
    $result.Actors = $actors
    $result.Workflows = $workflows.ToArray()
    return [pscustomobject]$result
}

function Parse-MarkdownFile {
    param(
        [string]$Path,
        [hashtable]$Attributes,
        [System.Collections.Generic.HashSet[string]]$Visited,
        $Config,
        [int]$LevelOffset = 0
    )

    $absolutePath = Get-AbsolutePath -Path $Path
    if ($Visited.Contains($absolutePath)) {
        return [pscustomobject]@{
            Metadata   = @{}
            Elements   = @()
            Attributes = $Attributes
        }
    }
    [void]$Visited.Add($absolutePath)

    $fileDir = Split-Path -Parent $absolutePath
    $rawText = Get-Content -LiteralPath $absolutePath -Encoding UTF8 -Raw
    if ($null -eq $rawText) { $rawText = '' }
    $rawText = $rawText -replace "`r`n", "`n"
    $rawText = $rawText -replace "`r", "`n"
    $lines = @($rawText -split "`n")

    $elements = New-Object System.Collections.Generic.List[object]
    $metadata = @{
        Title     = $null
        Subtitle  = $null
        Author    = $null
        RevNumber = $null
        RevDate   = $null
        Copyright = $null
        ImagesDir = $null
    }

    $lineIndex = 0
    $paragraphBuffer = New-Object System.Collections.Generic.List[string]

    function Flush-MdParagraph {
        if ($paragraphBuffer.Count -eq 0) { return }
        $joined = ($paragraphBuffer | ForEach-Object { ([string]$_).Trim() }) -join ' '
        $normalized = Normalize-MarkdownInlineText -Text $joined -Attributes $Attributes
        if (-not [string]::IsNullOrWhiteSpace($normalized)) {
            $elements.Add((New-Element -Type 'paragraph' -Data @{ Text = $normalized }))
        }
        $paragraphBuffer.Clear()
    }

    # YAML フロントマター解析
    if ($lines.Count -gt 0 -and ([string]$lines[0]).Trim() -eq '---') {
        $lineIndex = 1
        while ($lineIndex -lt $lines.Count) {
            $fmLine = [string]$lines[$lineIndex]
            $fmTrimmed = $fmLine.Trim()
            if ($fmTrimmed -eq '---' -or $fmTrimmed -eq '...') {
                $lineIndex++
                break
            }
            if ($fmLine -match '^(\w[\w-]*):\s*(.*)$') {
                $key = $matches[1].ToLowerInvariant()
                $val = $matches[2].Trim().Trim('"').Trim("'")
                $Attributes[$key] = $val
                switch ($key) {
                    'title' { $metadata.Title = $val }
                    'subtitle' { $metadata.Subtitle = $val }
                    'author' { $metadata.Author = $val }
                    'revnumber' { $metadata.RevNumber = $val }
                    'revdate' { $metadata.RevDate = $val }
                    'date' { if (-not $metadata.RevDate) { $metadata.RevDate = $val } }
                    'copyright' { $metadata.Copyright = $val }
                }
            }
            $lineIndex++
        }
    }

    $inFence = $false
    $fenceChar = $null
    $fenceLines = New-Object System.Collections.Generic.List[string]
    $pendingCaption = $null
    $pendingMdTableOptions = $null

    while ($lineIndex -lt $lines.Count) {
        $line = [string]$lines[$lineIndex]
        $trimmed = $line.Trim()

        # フェンスコードブロック内
        if ($inFence) {
            if ($trimmed -match "^$([regex]::Escape($fenceChar)){3,}\s*$") {
                $blockText = ($fenceLines -join "`n")
                $elements.Add((New-Element -Type 'code' -Data @{
                            Text    = $blockText
                            Caption = $pendingCaption
                        }))
                $pendingCaption = $null
                $fenceLines.Clear()
                $inFence = $false
                $fenceChar = $null
            }
            else {
                $fenceLines.Add($line)
            }
            $lineIndex++
            continue
        }

        # 空行
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            Flush-MdParagraph
            $lineIndex++
            continue
        }

        # ATX 見出し: # H1 ～ ###### H6
        if ($trimmed -match '^(#{1,6})\s+(.+)') {
            Flush-MdParagraph
            $level = $matches[1].Length
            # 末尾の  ## を除去してテキストを取得
            $rawHeadingText = ([string]$matches[2]) -replace '\s+#+\s*$', ''
            $headingText = Normalize-MarkdownInlineText -Text $rawHeadingText.Trim() -Attributes $Attributes
            if ($level -eq 1 -and -not $metadata.Title -and $LevelOffset -eq 0) {
                $metadata.Title = $headingText
                $elements.Add((New-Element -Type 'title' -Data @{ Text = $headingText; Subtitle = $metadata.Subtitle }))
            }
            else {
                # ## → Level=1、### → Level=2 とAsciiDocの==、===に封対応; includeのオフセット適用
                $elements.Add((New-Element -Type 'heading' -Data @{ Level = [Math]::Max(1, $level - 1 + $LevelOffset); Text = $headingText }))
            }
            $lineIndex++
            continue
        }

        # Setext 見出し (次行が === または --- の場合)
        if ($lineIndex + 1 -lt $lines.Count) {
            $nextLine = [string]$lines[$lineIndex + 1]
            if ($nextLine -match '^=+\s*$' -and -not [string]::IsNullOrWhiteSpace($trimmed)) {
                Flush-MdParagraph
                $headingText = Normalize-MarkdownInlineText -Text $trimmed -Attributes $Attributes
                if ($LevelOffset -eq 0 -and -not $metadata.Title) {
                    $metadata.Title = $headingText
                    $elements.Add((New-Element -Type 'title' -Data @{ Text = $headingText; Subtitle = $null }))
                }
                else {
                    $elements.Add((New-Element -Type 'heading' -Data @{ Level = [Math]::Max(1, 0 + $LevelOffset); Text = $headingText }))
                }
                $lineIndex += 2
                continue
            }
            if ($nextLine -match '^-+\s*$' -and
                -not [string]::IsNullOrWhiteSpace($trimmed) -and
                $trimmed -notmatch '^[-*_\s]+$') {
                Flush-MdParagraph
                $headingText = Normalize-MarkdownInlineText -Text $trimmed -Attributes $Attributes
                $elements.Add((New-Element -Type 'heading' -Data @{ Level = [Math]::Max(1, 1 + $LevelOffset); Text = $headingText }))
                $lineIndex += 2
                continue
            }
        }

        # フェンスコードブロック開始
        if ($trimmed -match '^(`{3,}|~{3,})') {
            Flush-MdParagraph
            $fenceChar = $matches[1].Substring(0, 1)
            $inFence = $true
            $fenceLines.Clear()
            $lineIndex++
            continue
        }

        # 水平線 (→ 改ページ扱い)
        if ($trimmed -match '^(-{3,}|\*{3,}|_{3,})\s*$') {
            Flush-MdParagraph
            $elements.Add((New-Element -Type 'pagebreak' -Data @{}))
            $lineIndex++
            continue
        }

        # Markdown拡張: <!-- include filename --> でファイルをインクルード
        if ($trimmed -match '^<!--\s*include\s+(.+?)\s*-->$') {
            Flush-MdParagraph
            $includedFile = $matches[1].Trim()
            $resolvedInclude = if ([System.IO.Path]::IsPathRooted($includedFile)) {
                $includedFile
            }
            else {
                Get-AbsolutePath -Path $includedFile -BaseDirectory $fileDir
            }
            if (Test-Path -LiteralPath $resolvedInclude -PathType Leaf) {
                $childAttrs = @{}
                foreach ($k in $Attributes.Keys) { $childAttrs[$k] = $Attributes[$k] }

                # インクルードファイル名表示（設定で有効化）
                if ($Config.Markdown -and $Config.Markdown.ShowIncludeSource) {
                    $displayName = [System.IO.Path]::GetFileName($resolvedInclude)
                    $labelStyle = if ($Config.Markdown.IncludeSourceStyle) { $Config.Markdown.IncludeSourceStyle } else { $null }
                    $elements.Add((New-Element -Type 'include-source' -Data @{ Text = $displayName; StyleConfig = $labelStyle }))
                }

                $included = Parse-MarkdownFile -Path $resolvedInclude -Attributes $childAttrs -Visited $Visited -Config $Config -LevelOffset ($LevelOffset + 1)
                foreach ($child in $included.Elements) { $elements.Add($child) }
            }
            else {
                $elements.Add((New-Element -Type 'paragraph' -Data @{ Text = "[include ファイルが見つかりません: $resolvedInclude]" }))
            }
            $lineIndex++
            continue
        }

        # テーブル属性コメント: <!-- options: autonumber -->
        if ($trimmed -match '^<!--\s*options:\s*(.+?)\s*-->$') {
            Flush-MdParagraph
            $pendingMdTableOptions = $matches[1].Trim()
            $lineIndex++
            continue
        }

        # テーブル
        if ($trimmed -match '^\|') {
            Flush-MdParagraph
            $mdTableLines = New-Object System.Collections.Generic.List[string]
            $mdSepColCount = 0
            while ($lineIndex -lt $lines.Count) {
                $tLine = [string]$lines[$lineIndex]
                $tTrim = $tLine.Trim()
                if (-not ($tTrim -match '^\|')) { break }
                # セパレータ行: 列数を取得してスキップ
                if ([string]::IsNullOrEmpty(($tTrim -replace '[|\-: ]', ''))) {
                    $sepCols = ([regex]::Matches($tTrim, '\|')).Count - 1
                    if ($sepCols -gt $mdSepColCount) { $mdSepColCount = $sepCols }
                    $lineIndex++
                    continue
                }
                $mdTableLines.Add($tTrim)
                $lineIndex++
            }

            if ($mdTableLines.Count -gt 0) {
                $mdRows = @()
                $colCount = $mdSepColCount
                $occupied = @{}
                $rowIndex = 0
                $isFirstRow = $true

                foreach ($tl in $mdTableLines) {
                    $cellTexts = @()
                    foreach ($cm in [regex]::Matches($tl, '\|([^|]*)')) {
                        $cellTexts += $cm.Groups[1].Value.Trim()
                    }
                    # 末尾の空セル (trailing |) を除去
                    while ($cellTexts.Count -gt 0 -and $cellTexts[-1] -eq '') {
                        $cellTexts = $cellTexts[0..($cellTexts.Count - 2)]
                    }

                    $cells = @()
                    $physCol = 0
                    $ctIndex = 0

                    while ($ctIndex -lt $cellTexts.Count) {
                        # rowspan で occupied なカラムと対応するプレースホルダーをスキップ
                        while ($occupied["$rowIndex,$physCol"]) {
                            $physCol++
                            if ($ctIndex -lt $cellTexts.Count) { $ctIndex++ }
                        }
                        if ($ctIndex -ge $cellTexts.Count) { break }

                        $ct = $cellTexts[$ctIndex]
                        $rowSpan = 1
                        $colSpan = 1
                        $cellText = $ct

                        # スパン指定解析: 2.3+text → rowspan=2,colspan=3
                        if ($ct -match '^(\d+)\.(\d+)\+(.*)$') {
                            $rowSpan = [int]$matches[1]
                            $colSpan = [int]$matches[2]
                            $cellText = $matches[3].Trim()
                        }
                        # rowspan のみ: 2.text (ドット直後が非数字)
                        elseif ($ct -match '^(\d+)\.([^\d].*)$' -or $ct -match '^(\d+)\.$') {
                            $rowSpan = [int]$matches[1]
                            $cellText = if ($ct -match '^(\d+)\.(.*)$') { $matches[2].Trim() } else { '' }
                        }
                        # colspan のみ: 2+text
                        elseif ($ct -match '^(\d+)\+(.*)$') {
                            $colSpan = [int]$matches[1]
                            $cellText = $matches[2].Trim()
                        }

                        $cells += @{
                            Text     = (Normalize-MarkdownInlineText -Text $cellText -Attributes $Attributes)
                            RowSpan  = $rowSpan
                            ColSpan  = $colSpan
                            IsHeader = $isFirstRow
                        }

                        # rowspan による後続行の occupied をマーク
                        for ($rr = 1; $rr -lt $rowSpan; $rr++) {
                            for ($cc = 0; $cc -lt $colSpan; $cc++) {
                                $occupied["$($rowIndex+$rr),$($physCol+$cc)"] = $true
                            }
                        }

                        $physCol += $colSpan
                        $ctIndex += $colSpan  # colspan 分のプレースホルダーをスキップ
                    }

                    # 有効列数を span を考慮して計算
                    $effectiveCols = 0
                    foreach ($c in $cells) { $effectiveCols += [int]$c.ColSpan }
                    if ($effectiveCols -gt $colCount) { $colCount = $effectiveCols }
                    $mdRows += , $cells
                    $isFirstRow = $false
                    $rowIndex++
                }

                if ($mdRows.Count -gt 0 -and $colCount -gt 0) {
                    $tableAttributes = @{}
                    if (-not [string]::IsNullOrWhiteSpace($pendingMdTableOptions)) {
                        $tableAttributes['options'] = $pendingMdTableOptions
                    }
                    $tableInfo = [pscustomobject]@{
                        Rows       = $mdRows
                        MaxColumns = $colCount
                    }
                    $elements.Add((New-Element -Type 'table' -Data @{
                                tableInfo  = $tableInfo
                                Caption    = $pendingCaption
                                Attributes = $tableAttributes
                            }))
                    $pendingCaption = $null
                    $pendingMdTableOptions = $null
                }
            }
            continue
        }

        # 引用ブロック (→ NOTE アドモニション)
        if ($trimmed -match '^>\s*(.*)$') {
            Flush-MdParagraph
            $quoteLines = New-Object System.Collections.Generic.List[string]
            $quoteLines.Add([string]$matches[1])
            while ($lineIndex + 1 -lt $lines.Count) {
                $nextLine = [string]$lines[$lineIndex + 1]
                if ($nextLine.Trim() -match '^>\s*(.*)$') {
                    $quoteLines.Add([string]$matches[1])
                    $lineIndex++
                }
                else { break }
            }
            $fullQuote = ($quoteLines | ForEach-Object { ([string]$_).Trim() }) -join ' '
            $elements.Add((New-Element -Type 'admonition' -Data @{
                        Kind = 'NOTE'
                        Text = (Normalize-MarkdownInlineText -Text $fullQuote -Attributes $Attributes)
                    }))
            $lineIndex++
            continue
        }

        # 箇条書き (unordered: -, *, +)
        if ($trimmed -match '^[-*+]\s+(.+)$') {
            Flush-MdParagraph
            $indentLen = ([regex]::Match($line, '^\s*')).Value.Length
            $level = [Math]::Max(1, [int][Math]::Floor($indentLen / 2) + 1)
            $elements.Add((New-Element -Type 'bullet' -Data @{
                        Text        = (Normalize-MarkdownInlineText -Text $matches[1] -Attributes $Attributes)
                        Level       = $level
                        MarkerCount = 1
                    }))
            $lineIndex++
            continue
        }

        # 番号付きリスト
        if ($trimmed -match '^\d+\.\s+(.+)$') {
            Flush-MdParagraph
            $indentLen = ([regex]::Match($line, '^\s*')).Value.Length
            $level = [Math]::Max(1, [int][Math]::Floor($indentLen / 2) + 1)
            $elements.Add((New-Element -Type 'numbered' -Data @{
                        Text        = (Normalize-MarkdownInlineText -Text $matches[1] -Attributes $Attributes)
                        Level       = $level
                        MarkerCount = $level
                    }))
            $lineIndex++
            continue
        }

        # draw.io: !drawio[caption](path.drawio)
        if ($trimmed -match '^!drawio\[([^\]]*)\]\(([^\)]+)\)$') {
            Flush-MdParagraph
            $caption = $matches[1].Trim()
            $drawioRef = $matches[2].Trim()

            $drawioOptions = @{ 'file' = $drawioRef }
            $render = Invoke-DrawIoRender `
                -SourceFilePath $absolutePath `
                -Options $drawioOptions `
                -Config $Config

            if ($render.Success) {
                $elements.Add((New-Element -Type 'image' -Data @{
                            Path        = $render.ImagePath
                            Caption     = if (-not [string]::IsNullOrWhiteSpace($caption)) { $caption } else { $pendingCaption }
                            GeneratedBy = 'drawio'
                        }))
            }
            else {
                $elements.Add((New-Element -Type 'admonition' -Data @{
                            Kind = 'WARNING'
                            Text = "[draw.io 画像生成失敗] $($render.ErrorMessage)"
                        }))
            }

            $pendingCaption = $null
            $lineIndex++
            continue
        }

        # !include-json-table[path]
        if ($trimmed -match '^!include-json-table\[([^\]]+)\]$') {
            Flush-MdParagraph
            $jsonRef = $matches[1].Trim()
            $jsonResult = Invoke-JsonTableDirective -FilePath $jsonRef -BaseDirectory $fileDir
            if ($jsonResult.Success) {
                if ($null -ne $jsonResult.Metadata) {
                    $metaDesc = $jsonResult.Metadata.Description
                    if (-not [string]::IsNullOrWhiteSpace($metaDesc)) {
                        $elements.Add((New-Element -Type 'paragraph' -Data @{
                                    Text = [string]$metaDesc
                                }))
                    }
                    $metaRules = $jsonResult.Metadata.Rules
                    if ($null -ne $metaRules) {
                        foreach ($rule in $metaRules) {
                            $rt = [string]$rule
                            if (-not [string]::IsNullOrWhiteSpace($rt)) {
                                $elements.Add((New-Element -Type 'bullet' -Data @{
                                            Text        = $rt
                                            Level       = 1
                                            MarkerCount = 1
                                        }))
                            }
                        }
                    }
                }
                $elements.Add((New-Element -Type 'table' -Data @{
                            tableInfo  = $jsonResult.TableInfo
                            Caption    = $pendingCaption
                            Attributes = @{ 'options' = 'header' }
                        }))
            }
            else {
                $elements.Add((New-Element -Type 'code' -Data @{
                            Text    = (Get-JsonIncludeErrorText -ErrorCode $jsonResult.ErrorCode -FilePath $jsonRef)
                            Caption = $null
                        }))
            }
            $pendingCaption = $null
            $lineIndex++
            continue
        }

        # !include-json-tree[path]
        if ($trimmed -match '^!include-json-tree\[([^\]]+)\]$') {
            Flush-MdParagraph
            $jsonRef = $matches[1].Trim()
            $jsonResult = Invoke-JsonTreeDirective -FilePath $jsonRef -BaseDirectory $fileDir
            if ($jsonResult.Success) {
                $elements.Add((New-Element -Type 'code' -Data @{
                            Text    = $jsonResult.TreeText
                            Caption = $pendingCaption
                        }))
            }
            else {
                $elements.Add((New-Element -Type 'code' -Data @{
                            Text    = (Get-JsonIncludeErrorText -ErrorCode $jsonResult.ErrorCode -FilePath $jsonRef)
                            Caption = $null
                        }))
            }
            $pendingCaption = $null
            $lineIndex++
            continue
        }

        # !include-json-config[path]
        if ($trimmed -match '^!include-json-config\[([^\]]+)\]$') {
            Flush-MdParagraph
            $jsonRef = $matches[1].Trim()
            $jsonResult = Invoke-JsonConfigDirective -FilePath $jsonRef -BaseDirectory $fileDir
            if ($jsonResult.Success) {
                if (-not [string]::IsNullOrWhiteSpace($jsonResult.Description)) {
                    $elements.Add((New-Element -Type 'paragraph' -Data @{
                                Text = [string]$jsonResult.Description
                            }))
                }
                $configRules = $jsonResult.Rules
                if ($null -ne $configRules) {
                    $elements.Add((New-Element -Type 'heading' -Data @{
                                Level = 2
                                Text  = '管理ルール'
                            }))
                    foreach ($rule in $configRules) {
                        $rt = [string]$rule
                        if (-not [string]::IsNullOrWhiteSpace($rt)) {
                            $elements.Add((New-Element -Type 'bullet' -Data @{
                                        Text        = $rt
                                        Level       = 1
                                        MarkerCount = 1
                                    }))
                        }
                    }
                }
                foreach ($cat in $jsonResult.Categories) {
                    $catHeading = if (-not [string]::IsNullOrWhiteSpace([string]$cat.Description)) {
                        [string]$cat.Description
                    }
                    else {
                        [string]$cat.Key
                    }
                    $elements.Add((New-Element -Type 'heading' -Data @{
                                Level = 2
                                Text  = $catHeading
                            }))
                    $catConstants = $cat.Constants
                    if ($null -ne $catConstants -and $catConstants.Count -gt 0) {
                        $rows = New-Object System.Collections.Generic.List[object]
                        $headerCells = @(
                            @{ Text = '定数名'; RowSpan = 1; ColSpan = 1; IsHeader = $true },
                            @{ Text = '値'; RowSpan = 1; ColSpan = 1; IsHeader = $true },
                            @{ Text = '説明'; RowSpan = 1; ColSpan = 1; IsHeader = $true }
                        )
                        $rows.Add($headerCells)
                        foreach ($const in $catConstants) {
                            $dataCells = @(
                                @{ Text = [string]$const.Name; RowSpan = 1; ColSpan = 1; IsHeader = $false },
                                @{ Text = [string]$const.Value; RowSpan = 1; ColSpan = 1; IsHeader = $false },
                                @{ Text = [string]$const.Description; RowSpan = 1; ColSpan = 1; IsHeader = $false }
                            )
                            $rows.Add($dataCells)
                        }
                        $configTableInfo = [pscustomobject]@{
                            Rows       = $rows.ToArray()
                            MaxColumns = 3
                        }
                        $elements.Add((New-Element -Type 'table' -Data @{
                                    tableInfo  = $configTableInfo
                                    Caption    = $null
                                    Attributes = @{ 'options' = 'header' }
                                }))
                    }
                }
            }
            else {
                $elements.Add((New-Element -Type 'code' -Data @{
                            Text    = (Get-JsonIncludeErrorText -ErrorCode $jsonResult.ErrorCode -FilePath $jsonRef)
                            Caption = $null
                        }))
            }
            $pendingCaption = $null
            $lineIndex++
            continue
        }

        # !include-json-enum[path]
        if ($trimmed -match '^!include-json-enum\[([^\]]+)\]$') {
            Flush-MdParagraph
            $jsonRef = $matches[1].Trim()
            $jsonResult = Invoke-JsonEnumDirective -FilePath $jsonRef -BaseDirectory $fileDir
            if ($jsonResult.Success) {
                if (-not [string]::IsNullOrWhiteSpace($jsonResult.Description)) {
                    $elements.Add((New-Element -Type 'paragraph' -Data @{
                                Text = [string]$jsonResult.Description
                            }))
                }
                $enumRules = $jsonResult.Rules
                if ($null -ne $enumRules) {
                    $elements.Add((New-Element -Type 'heading' -Data @{
                                Level = 2
                                Text  = '管理ルール'
                            }))
                    foreach ($rule in $enumRules) {
                        $rt = [string]$rule
                        if (-not [string]::IsNullOrWhiteSpace($rt)) {
                            $elements.Add((New-Element -Type 'bullet' -Data @{
                                        Text        = $rt
                                        Level       = 1
                                        MarkerCount = 1
                                    }))
                        }
                    }
                }
                foreach ($enumItem in $jsonResult.Enums) {
                    $elements.Add((New-Element -Type 'heading' -Data @{
                                Level = 2
                                Text  = [string]$enumItem.EnumName
                            }))
                    if (-not [string]::IsNullOrWhiteSpace([string]$enumItem.Description)) {
                        $elements.Add((New-Element -Type 'paragraph' -Data @{
                                    Text = [string]$enumItem.Description
                                }))
                    }
                    if ($enumItem.HasValues) {
                        $elements.Add((New-Element -Type 'table' -Data @{
                                    tableInfo  = $enumItem.TableInfo
                                    Caption    = $null
                                    Attributes = @{ 'options' = 'header' }
                                }))
                    }
                    else {
                        $elements.Add((New-Element -Type 'paragraph' -Data @{
                                    Text = '定義なし'
                                }))
                    }
                }
            }
            else {
                $elements.Add((New-Element -Type 'code' -Data @{
                            Text    = (Get-JsonIncludeErrorText -ErrorCode $jsonResult.ErrorCode -FilePath $jsonRef)
                            Caption = $null
                        }))
            }
            $pendingCaption = $null
            $lineIndex++
            continue
        }

        # !include-json-workflow[path]
        if ($trimmed -match '^!include-json-workflow\[([^\]]+)\]$') {
            Flush-MdParagraph
            $jsonRef = $matches[1].Trim()
            $jsonResult = Invoke-JsonWorkflowDirective -FilePath $jsonRef -BaseDirectory $fileDir
            if ($jsonResult.Success) {
                $wfActors = $jsonResult.Actors

                if ($wfActors.Count -gt 0) {
                    $elements.Add((New-Element -Type 'heading' -Data @{ Level = 1; Text = 'アクター一覧' }))
                    $rows = New-Object System.Collections.Generic.List[object]
                    $rows.Add(@(
                            @{ Text = 'ロールID'; RowSpan = 1; ColSpan = 1; IsHeader = $true },
                            @{ Text = '名称'; RowSpan = 1; ColSpan = 1; IsHeader = $true }
                        ))
                    foreach ($rid in $wfActors.Keys) {
                        $rows.Add(@(
                                @{ Text = [string]$rid; RowSpan = 1; ColSpan = 1; IsHeader = $false },
                                @{ Text = [string]$wfActors[$rid]; RowSpan = 1; ColSpan = 1; IsHeader = $false }
                            ))
                    }
                    $elements.Add((New-Element -Type 'table' -Data @{
                                tableInfo  = [pscustomobject]@{ Rows = $rows.ToArray(); MaxColumns = 2 }
                                Caption    = $null
                                Attributes = @{ 'options' = 'header' }
                            }))
                }

                foreach ($wf in $jsonResult.Workflows) {
                    $elements.Add((New-Element -Type 'heading' -Data @{ Level = 1; Text = [string]$wf.Key }))
                    if (-not [string]::IsNullOrWhiteSpace([string]$wf.Description)) {
                        $elements.Add((New-Element -Type 'paragraph' -Data @{ Text = [string]$wf.Description }))
                    }

                    # States table
                    $elements.Add((New-Element -Type 'heading' -Data @{ Level = 2; Text = 'ステータス一覧' }))
                    $stRows = New-Object System.Collections.Generic.List[object]
                    $stRows.Add(@(
                            @{ Text = 'ステータスID'; RowSpan = 1; ColSpan = 1; IsHeader = $true },
                            @{ Text = '名称'; RowSpan = 1; ColSpan = 1; IsHeader = $true },
                            @{ Text = '説明'; RowSpan = 1; ColSpan = 1; IsHeader = $true }
                        ))
                    foreach ($sid in $wf.States.Keys) {
                        $s = $wf.States[$sid]
                        $stRows.Add(@(
                                @{ Text = [string]$sid; RowSpan = 1; ColSpan = 1; IsHeader = $false },
                                @{ Text = [string]$s.Label; RowSpan = 1; ColSpan = 1; IsHeader = $false },
                                @{ Text = [string]$s.Description; RowSpan = 1; ColSpan = 1; IsHeader = $false }
                            ))
                    }
                    $elements.Add((New-Element -Type 'table' -Data @{
                                tableInfo  = [pscustomobject]@{ Rows = $stRows.ToArray(); MaxColumns = 3 }
                                Caption    = $null
                                Attributes = @{ 'options' = 'header' }
                            }))

                    # Transitions table
                    $elements.Add((New-Element -Type 'heading' -Data @{ Level = 2; Text = '状態遷移一覧' }))
                    $tColKeys = @()
                    $tColHdrs = @()
                    foreach ($ck in $wf.ColDefs.Keys) { $tColKeys += [string]$ck; $tColHdrs += [string]$wf.ColDefs[$ck] }
                    if ($wf.HasConditions) { $tColKeys += 'condition'; $tColHdrs += '条件' }

                    $trRows = New-Object System.Collections.Generic.List[object]
                    $hdr = @()
                    for ($ci = 0; $ci -lt $tColKeys.Count; $ci++) {
                        $hdr += @{ Text = $tColHdrs[$ci]; RowSpan = 1; ColSpan = 1; IsHeader = $true }
                    }
                    $trRows.Add($hdr)

                    foreach ($tr in $wf.Transitions) {
                        $fromParts = @()
                        foreach ($fv in $tr.FromList) {
                            if ($fv -eq '[*]') { $fromParts += '[初期]' }
                            elseif ($wf.States.Keys -contains $fv) { $fromParts += [string]$wf.States[$fv].Label }
                            else { $fromParts += $fv }
                        }
                        $fromCell = $fromParts -join '\n'

                        $toId = $tr.To
                        $toCell = if ($toId -eq '[*]') { '[完了]' }
                        elseif ($wf.States.Keys -contains $toId) { [string]$wf.States[$toId].Label }
                        else { $toId }

                        $aId = $tr.Actor
                        $actorCell = if (-not [string]::IsNullOrWhiteSpace($aId) -and $wfActors.ContainsKey($aId)) {
                            $wfActors[$aId]
                        }
                        else { $aId }

                        $dCells = @()
                        foreach ($ck in $tColKeys) {
                            $cv = ''
                            if ($ck -eq 'from') { $cv = $fromCell }
                            elseif ($ck -eq 'to') { $cv = $toCell }
                            elseif ($ck -eq 'actor') { $cv = $actorCell }
                            elseif ($ck -eq 'trigger') { $cv = $tr.Trigger }
                            elseif ($ck -eq 'screen') { $cv = $tr.Screen }
                            elseif ($ck -eq 'notes') { $cv = $tr.Notes }
                            elseif ($ck -eq 'condition') { $cv = if ($null -ne $tr.Condition) { $tr.Condition } else { '' } }
                            else {
                                if ($tr.Extras.Keys -contains $ck) { $cv = [string]$tr.Extras[$ck] }
                            }
                            $dCells += @{ Text = [string]$cv; RowSpan = 1; ColSpan = 1; IsHeader = $false }
                        }
                        $trRows.Add($dCells)
                    }

                    $elements.Add((New-Element -Type 'table' -Data @{
                                tableInfo  = [pscustomobject]@{ Rows = $trRows.ToArray(); MaxColumns = $tColKeys.Count }
                                Caption    = $null
                                Attributes = @{ 'options' = 'header' }
                            }))

                    # PlantUML state diagrams (render if PlantUML is configured; fall back to code block)
                    $elements.Add((New-Element -Type 'heading' -Data @{ Level = 2; Text = '状態遷移図' }))
                    $pumlFull = Build-WorkflowPlantUml -Workflow $wf -Actors $wfActors -Filter $null
                    $pumlOpts = @{ target = "wf-$($wf.Key)-all" }
                    $pumlRender = Invoke-PlantUmlRender -PlantUmlSource $pumlFull -SourceFilePath $absolutePath `
                        -Attributes $Attributes -Options $pumlOpts -Config $Config -Sequence 0
                    if ($pumlRender.Success) {
                        $elements.Add((New-Element -Type 'image' -Data @{
                                    Path        = $pumlRender.ImagePath
                                    Caption     = $null
                                    GeneratedBy = 'plantuml'
                                    Source      = $pumlFull
                                }))
                    }
                    else {
                        $elements.Add((New-Element -Type 'code' -Data @{ Text = $pumlFull; Caption = $null }))
                    }

                    if ($wf.HasConditions) {
                        $elements.Add((New-Element -Type 'heading' -Data @{ Level = 2; Text = '状態遷移図（前払い）' }))
                        $pumlPre = Build-WorkflowPlantUml -Workflow $wf -Actors $wfActors -Filter 'PREPAID'
                        $preOpts = @{ target = "wf-$($wf.Key)-prepaid" }
                        $preRender = Invoke-PlantUmlRender -PlantUmlSource $pumlPre -SourceFilePath $absolutePath `
                            -Attributes $Attributes -Options $preOpts -Config $Config -Sequence 0
                        if ($preRender.Success) {
                            $elements.Add((New-Element -Type 'image' -Data @{
                                        Path = $preRender.ImagePath; Caption = $null; GeneratedBy = 'plantuml'; Source = $pumlPre
                                    }))
                        }
                        else {
                            $elements.Add((New-Element -Type 'code' -Data @{ Text = $pumlPre; Caption = $null }))
                        }

                        $elements.Add((New-Element -Type 'heading' -Data @{ Level = 2; Text = '状態遷移図（後払い）' }))
                        $pumlPost = Build-WorkflowPlantUml -Workflow $wf -Actors $wfActors -Filter 'POSTPAID'
                        $postOpts = @{ target = "wf-$($wf.Key)-postpaid" }
                        $postRender = Invoke-PlantUmlRender -PlantUmlSource $pumlPost -SourceFilePath $absolutePath `
                            -Attributes $Attributes -Options $postOpts -Config $Config -Sequence 0
                        if ($postRender.Success) {
                            $elements.Add((New-Element -Type 'image' -Data @{
                                        Path = $postRender.ImagePath; Caption = $null; GeneratedBy = 'plantuml'; Source = $pumlPost
                                    }))
                        }
                        else {
                            $elements.Add((New-Element -Type 'code' -Data @{ Text = $pumlPost; Caption = $null }))
                        }
                    }
                }
            }
            else {
                $elements.Add((New-Element -Type 'code' -Data @{
                            Text    = (Get-JsonIncludeErrorText -ErrorCode $jsonResult.ErrorCode -FilePath $jsonRef)
                            Caption = $null
                        }))
            }
            $pendingCaption = $null
            $lineIndex++
            continue
        }

        # 単独画像行: ![alt](path)
        if ($trimmed -match '^!\[([^\]]*)\]\(([^\)]+)\)$') {
            Flush-MdParagraph
            $altText = $matches[1]
            $imgRef = $matches[2].Trim()
            $imagePath = Get-ImageFullPath -ImageReference $imgRef -CurrentFileDirectory $fileDir -Attributes $Attributes
            $elements.Add((New-Element -Type 'image' -Data @{
                        Path    = $imagePath
                        Caption = if (-not [string]::IsNullOrWhiteSpace($altText)) { $altText } else { $pendingCaption }
                    }))
            $pendingCaption = $null
            $lineIndex++
            continue
        }

        # 通常段落
        $paragraphBuffer.Add($trimmed)
        $lineIndex++
    }

    Flush-MdParagraph

    # タイトルがあるルート文書はセクション番号を自動有効化（AsciiDocの:sectnums:相当）
    if ($LevelOffset -eq 0 -and $metadata.Title -and -not $Attributes.ContainsKey('sectnums')) {
        $Attributes['sectnums'] = $true
    }

    return [pscustomobject]@{
        Metadata   = $metadata
        Elements   = $elements
        Attributes = $Attributes
    }
}

if ($TestMode) { return }

try {
    $configFullPath = Get-AbsolutePath -Path $ConfigFullPath
    $config = Load-JsonConfig -Path $configFullPath

    foreach ($adocPath in $AdocFullPath) {
        $inputFullPath = Get-AbsolutePath -Path $adocPath

        if (-not (Test-Path -LiteralPath $inputFullPath -PathType Leaf)) {
            Write-Output "警告: ファイルが見つかりません: $inputFullPath (スキップ)"
            continue
        }

        $sourceDir = Split-Path -Parent $inputFullPath
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputFullPath)
        $targetDir = if ([string]::IsNullOrWhiteSpace($OutputDir)) { $sourceDir } else { $OutputDir }

        if (-not (Test-Path -LiteralPath $targetDir)) {
            [void](New-Item -ItemType Directory -Path $targetDir -Force)
        }

        $baseOutput = Join-Path $targetDir ($baseName + '.docx')
        if ($Overwrite -or -not (Test-Path -LiteralPath $baseOutput)) {
            $outputFullPath = $baseOutput
        }
        else {
            $counter = 1
            do {
                $outputFullPath = Join-Path $targetDir ($baseName + '_' + $counter + '.docx')
                $counter++
            } while (Test-Path -LiteralPath $outputFullPath)
        }

        try {
            $visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $attributes = @{}
            $extension = [System.IO.Path]::GetExtension($inputFullPath).ToLowerInvariant()
            if ($extension -eq '.md' -or $extension -eq '.markdown') {
                $parsed = Parse-MarkdownFile -Path $inputFullPath -Attributes $attributes -Visited $visited -Config $config
            }
            else {
                $parsed = Parse-AsciiDocFile -Path $inputFullPath -Attributes $attributes -Visited $visited -Config $config
            }
            $TIMESTAMP = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
            Write-Output "$TIMESTAMP パース完了: $inputFullPath"

            Build-WordDocument -Parsed $parsed -Config $config -OutputFullPath $outputFullPath
            $TIMESTAMP = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
            Write-Output "$TIMESTAMP 変換完了: $outputFullPath"
        }
        catch {
            Write-Output "エラー [$inputFullPath]: $($_.Exception.Message)"
            Write-Output "発生箇所: $($_.InvocationInfo.PositionMessage)"
        }
    }
}
catch {
    $message = if ($_.Exception) { $_.Exception.Message } else { [string]$_ }
    Write-Output "致命的エラー: $($_.Exception.Message)"
    Write-Output "発生箇所: $($_.InvocationInfo.PositionMessage)"    
    exit 1
}
