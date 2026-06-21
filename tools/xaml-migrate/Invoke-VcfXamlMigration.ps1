#Requires -Version 5.1
<#
.SYNOPSIS
    Mechanical XAML transforms for Demac.VCF WPF alignment (Phases 1-6).

.DESCRIPTION
    Safe, idempotent-ish transforms for POS bulk migration. Does NOT handle res: includes,
    Scene BackColor, or Button+TextBlock child simplification - see docs/XAML_MIGRATION_PROMPTS.md.

.PARAMETER Path
    File or directory. Directories require -Recurse.

.PARAMETER Recurse
    Process *.xml under directories recursively.

.PARAMETER WhatIf
    Show changes without writing files.

.PARAMETER ReportOnly
    Scan and emit manual-review items only (no transforms).

.PARAMETER Transform
    Subset of transforms. Default: All.

.PARAMETER SelfTest
    Run built-in fixture test and exit.

.EXAMPLE
    .\Invoke-VcfXamlMigration.ps1 -Path ..\..\.Tests\Test0\Resources\XAML -Recurse -WhatIf

.EXAMPLE
    .\Invoke-VcfXamlMigration.ps1 -Path SalesOrder.xml -Transform Layout,ListView
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(ParameterSetName = 'Run', Mandatory = $true)]
    [string[]]$Path,

    [Parameter(ParameterSetName = 'Run')]
    [switch]$Recurse,

    [Parameter(ParameterSetName = 'Run')]
    [switch]$ReportOnly,

    [Parameter(ParameterSetName = 'Run')]
    [ValidateSet('All', 'Layout', 'ListView', 'ThemeResource', 'ButtonText')]
    [string[]]$Transform = @('All'),

    [Parameter(ParameterSetName = 'SelfTest')]
    [switch]$SelfTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-TransformEnabled {
    param([string]$Name)
    return ($Transform -contains 'All') -or ($Transform -contains $Name)
}

function Get-XamlAttributeMap {
    param([string]$AttributeText)

    $map = [ordered]@{}
    if ([string]::IsNullOrWhiteSpace($AttributeText)) { return $map }

    $pattern = '(?<name>[A-Za-z_][\w:\.-]*)\s*=\s*(?<q>''[^'']*''|"[^"]*")'
    foreach ($m in [regex]::Matches($AttributeText, $pattern)) {
        $raw = $m.Groups['q'].Value
        $map[$m.Groups['name'].Value] = $raw.Substring(1, $raw.Length - 2)
    }
    return $map
}

function Format-XamlAttributes {
    param([System.Collections.IDictionary]$Attributes)

    $parts = @()
    foreach ($key in $Attributes.Keys) {
        $val = $Attributes[$key]
        if ($val -match '[\s"<>]') {
            $parts += ('{0}="{1}"' -f $key, ($val -replace '"', '&quot;'))
        }
        else {
            $parts += ('{0}="{1}"' -f $key, $val)
        }
    }
    return ($parts -join ' ')
}

function Convert-XamlOpeningTag {
    param(
        [string]$Line,
        [ref]$ChangeCount,
        [System.Collections.Generic.List[string]]$ManualReview
    )

    $tagMatch = [regex]::Match($Line, '^(?<indent>\s*)<(?<name>[\w:\.-]+)(?<attrs>[^>/]*)(?<self>\s*/)?(?<close>\s*>)')
    if (-not $tagMatch.Success) { return $Line }

    $tagName = $tagMatch.Groups['name'].Value
    $attrs = Get-XamlAttributeMap $tagMatch.Groups['attrs'].Value
    $changed = $false

    if (Test-TransformEnabled 'ListView') {
        if ($tagName -eq 'UnboundListView') {
            $tagName = 'ListView'
            $changed = $true
        }
        if ($attrs.Contains('TargetType') -and $attrs['TargetType'] -eq 'UnboundListView') {
            $attrs['TargetType'] = 'ListView'
            $changed = $true
        }
    }

    if (Test-TransformEnabled 'ButtonText') {
        if ($tagName -eq 'Button' -and $attrs.Contains('Text')) {
            $attrs['Content'] = $attrs['Text']
            $attrs.Remove('Text')
            $changed = $true
        }
    }

    if (Test-TransformEnabled 'Layout') {
        if ($attrs.Contains('DesignWidth')) {
            $attrs['Width'] = $attrs['DesignWidth']
            $attrs.Remove('DesignWidth')
            $changed = $true
        }
        if ($attrs.Contains('DesignHeight')) {
            $attrs['Height'] = $attrs['DesignHeight']
            $attrs.Remove('DesignHeight')
            $changed = $true
        }

        $hasLeft = $attrs.Contains('DesignLeft')
        $hasTop = $attrs.Contains('DesignTop')
        if ($hasLeft -or $hasTop) {
            if ($attrs.Contains('Margin')) {
                $ManualReview.Add("Margin conflict: <$tagName> has Margin and DesignLeft/DesignTop - review manually")
            }
            else {
                $left = if ($hasLeft) { $attrs['DesignLeft'] } else { '0' }
                $top = if ($hasTop) { $attrs['DesignTop'] } else { '0' }
                $attrs['Margin'] = "$left,$top,0,0"
                if ($hasLeft) { $attrs.Remove('DesignLeft') }
                if ($hasTop) { $attrs.Remove('DesignTop') }
                $changed = $true
            }
        }
    }

    if (-not $changed) { return $Line }

    $ChangeCount.Value++
    $attrText = Format-XamlAttributes $attrs
    $selfClose = $tagMatch.Groups['self'].Value
    $close = $tagMatch.Groups['close'].Value
    if ($attrText) {
        return ('{0}<{1} {2}{3}{4}' -f $tagMatch.Groups['indent'].Value, $tagName, $attrText, $selfClose, $close)
    }
    return ('{0}<{1}{2}{3}' -f $tagMatch.Groups['indent'].Value, $tagName, $selfClose, $close)
}

function Convert-ThemeResourceMarkup {
    param(
        [string]$Text,
        [ref]$ChangeCount
    )

    if (-not (Test-TransformEnabled 'ThemeResource')) { return $Text }

    $pattern = '\{ThemeResource\s+(?:Key\s*=\s*)?(?<key>[^}]+)\}'
    $result = [regex]::Replace($Text, $pattern, {
        param($m)
        $script:ThemeResourceHits++
        $key = $m.Groups['key'].Value.Trim()
        return "{DynamicResource $key}"
    })

    if ($script:ThemeResourceHits -gt 0) {
        $ChangeCount.Value += $script:ThemeResourceHits
        $script:ThemeResourceHits = 0
    }
    return $result
}

function Get-ManualReviewNotes {
    param(
        [string[]]$Lines,
        [string]$FilePath
    )

    $notes = [System.Collections.Generic.List[string]]::new()
    $inButton = $false
    $buttonLine = 0

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        $lineNo = $i + 1

        if ($line -match '<Scene\b[^>]*\bBackColor=') {
            $notes.Add("$FilePath`:$lineNo Scene BackColor= - not a Scene DP under StrictXamlLoad; use Style or code")
        }
        if ($line -match '<res:') {
            $notes.Add("$FilePath`:$lineNo res: include - migrate to ResourceDictionary / MergedDictionaries (manual)")
        }
        if ($line -match '\bres:[^\s"<>]+') {
            $notes.Add("$FilePath`:$lineNo res: path reference - review ResourceDictionary migration")
        }
        if ($line -match '@\w+') {
            $notes.Add("$FilePath`:$lineNo @ fragment binding - Phase 7c DataTemplate migration")
        }
        if ($line -match '<Button\b') {
            $inButton = $true
            $buttonLine = $lineNo
        }
        if ($inButton -and $line -match '<TextBlock\b[^>]*\bText=') {
            $notes.Add("$FilePath`:$buttonLine Button with TextBlock child - consider Content= (see XAML_MIGRATION_PROMPTS.md)")
            $inButton = $false
        }
        if ($line -match '</Button>') {
            $inButton = $false
        }
    }

    return $notes
}

function Convert-XamlFileContent {
    param(
        [string]$Content,
        [string]$FilePath,
        [switch]$ReportOnlyMode
    )

    $manual = [System.Collections.Generic.List[string]]::new()
    $manual.AddRange([string[]](Get-ManualReviewNotes -Lines ($Content -split "`r?`n") -FilePath $FilePath))

    if ($ReportOnlyMode) {
        return [pscustomobject]@{
            Content      = $Content
            ChangeCount  = 0
            ManualReview = $manual
        }
    }

    $changeCount = 0
    $lines = $Content -split "`r?`n", -1
    $out = New-Object System.Collections.Generic.List[string]

    foreach ($line in $lines) {
        $script:ThemeResourceHits = 0
        $newLine = Convert-XamlOpeningTag -Line $line -ChangeCount ([ref]$changeCount) -ManualReview $manual
        $newLine = Convert-ThemeResourceMarkup -Text $newLine -ChangeCount ([ref]$changeCount)
        $out.Add($newLine)
    }

    $normalized = ($out -join "`r`n")
    if ($Content.EndsWith("`n") -and -not $normalized.EndsWith("`n")) {
        $normalized += "`r`n"
    }

    return [pscustomobject]@{
        Content      = $normalized
        ChangeCount  = $changeCount
        ManualReview = $manual
    }
}

function Get-TargetFiles {
    param([string[]]$InputPath, [switch]$RecurseFiles)

    $files = [System.Collections.Generic.List[string]]::new()
    foreach ($p in $InputPath) {
        if (-not (Test-Path -LiteralPath $p)) {
            throw "Path not found: $p"
        }
        $item = Get-Item -LiteralPath $p
        if ($item.PSIsContainer) {
            $glob = if ($RecurseFiles) { '*.xml' } else { '*.xml' }
            Get-ChildItem -LiteralPath $p -Filter $glob -File -Recurse:$RecurseFiles |
                ForEach-Object { $files.Add($_.FullName) }
        }
        else {
            $files.Add($item.FullName)
        }
    }
    return @($files)
}

function Normalize-XamlText {
    param([string]$Text)
    return ($Text -replace "`r`n", "`n").TrimEnd()
}

function Invoke-SelfTest {
    $root = $PSScriptRoot
    $before = Join-Path $root 'fixtures\before-sample.xml'
    $expected = Join-Path $root 'fixtures\expected-sample.xml'

    $beforeText = [IO.File]::ReadAllText($before)
    $expectedText = Normalize-XamlText ([IO.File]::ReadAllText($expected))

    $result = Convert-XamlFileContent -Content $beforeText -FilePath $before
    $actual = Normalize-XamlText $result.Content

    if ($actual -ne $expectedText) {
        Write-Error "Self-test failed: transformed output does not match expected-sample.xml"
    }

    if ($result.ManualReview.Count -lt 1) {
        Write-Error "Self-test failed: expected manual review notes (Scene BackColor)"
    }

    Write-Host "Self-test passed ($($result.ChangeCount) transform hits, $($result.ManualReview.Count) manual-review notes)."
}

function Invoke-MigrationRun {
    $files = @(Get-TargetFiles -InputPath $Path -RecurseFiles:$Recurse)
    if ($files.Count -eq 0) {
        Write-Warning 'No XML files found.'
        return
    }

    $totalChanges = 0
    $allReview = [System.Collections.Generic.List[string]]::new()

    foreach ($file in $files) {
        $original = [IO.File]::ReadAllText($file)
        $result = Convert-XamlFileContent -Content $original -FilePath $file -ReportOnlyMode:$ReportOnly
        $allReview.AddRange($result.ManualReview)

        if ($ReportOnly) {
            Write-Host "Report: $file ($($result.ManualReview.Count) note(s))"
            continue
        }

        if ($result.ChangeCount -eq 0 -and $original -eq $result.Content) {
            Write-Verbose "No changes: $file"
            continue
        }

        $totalChanges += $result.ChangeCount
        if ($WhatIfPreference) {
            Write-Host "WhatIf: would update $file ($($result.ChangeCount) transform hit(s))"
            continue
        }

        if ($PSCmdlet.ShouldProcess($file, 'Write migrated XAML')) {
            [IO.File]::WriteAllText($file, $result.Content)
            Write-Host "Updated: $file ($($result.ChangeCount) transform hit(s))"
        }
    }

    if ($allReview.Count -gt 0) {
        Write-Host ''
        Write-Host 'Manual review:'
        $allReview | Select-Object -Unique | ForEach-Object { Write-Host "  $_" }
    }

    if (-not $ReportOnly) {
        Write-Host ''
        Write-Host "Done. Files: $($files.Count), transform hits: $totalChanges"
    }
}

if ($MyInvocation.InvocationName -eq '.') {
    return
}

if ($PSCmdlet.ParameterSetName -eq 'SelfTest') {
    Invoke-SelfTest
    exit 0
}

Invoke-MigrationRun
