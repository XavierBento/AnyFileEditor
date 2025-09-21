# ============================================================================
# Project   : AnyFile Editor (TxtOrganizer)
# File      : Fix-AliasUsings.ps1
# Author    : Xavier Bento
# Version   : v1.0
# Created   : 2025-09-21
# Description: PowerShell helper script for development/build tasks.
# ============================================================================
# File: AnyFileEditor_fixed_all/Fix-AliasUsings.ps1
# Purpose: Build/CI/Script helper. Edit with care; environment-sensitive.
# Notes: Windows-only unless specified. Behavior unchanged; comments only.

<#  Fix-OpenXml-Aliases.ps1
    Fix CS0234 by mapping OpenXML Wordprocessing types from W.* -> WP.*
    - Ensures GlobalUsings.Aliases.cs contains:
        W  = System.Windows
        WP = DocumentFormat.OpenXml.Wordprocessing
        SWD = System.Windows.Documents (for FlowDocument, just in case)
    - Rewrites only Wordprocessing identifiers (Document, Run, Text, Break, Paragraph, Table, borders, etc.)
    - Leaves real WPF usages (W.MessageBox, etc.) untouched.

    Usage (recommended: map to OpenXML WP.*):
      .\Fix-OpenXml-Aliases.ps1 -CsprojPath "C:\Users\xavie\Downloads\Compressed\AnyFileEditor_fixed_all\TxtOrganizer.csproj" -RunBuild

    If a few places actually need FlowDocument types (rare here), you can manually change those lines to SWD.* afterward.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string] $CsprojPath,

  [switch] $Backup,   # keep .bak files when modifying sources
  [switch] $RunBuild  # run dotnet clean/build at the end
)

function Info($m){Write-Host "[INFO] $m" -ForegroundColor Cyan}
function Ok($m){Write-Host   "[OK]  $m" -ForegroundColor Green}
function Err($m){Write-Host  "[ERR] $m" -ForegroundColor Red}

# --- Resolve paths
try { $CsprojPath = (Resolve-Path $CsprojPath).Path } catch { Err "Project not found"; exit 2 }
if ([IO.Path]::GetExtension($CsprojPath).ToLower() -ne ".csproj"){ Err "Not a .csproj: $CsprojPath"; exit 2 }
$ProjectDir       = Split-Path $CsprojPath -Parent
$GlobalUsingsFile = Join-Path $ProjectDir "GlobalUsings.Aliases.cs"

Info "Project : $CsprojPath"
Info "ProjDir : $ProjectDir"

# --- 1) Ensure global aliases exist
$required = @(
  'global using W      = System.Windows;',
  'global using WP     = DocumentFormat.OpenXml.Wordprocessing;',
  'global using SWD    = System.Windows.Documents;',
  'global using SWC    = System.Windows.Controls;',
  'global using SWI    = System.Windows.Input;',
  'global using SMB    = System.Windows.Media;',
  'global using WF     = System.Windows.Forms;',
  'global using MWin32 = Microsoft.Win32;'
)

if (-not (Test-Path $GlobalUsingsFile)) {
  '﻿// Centralized global alias directives (net8.0-windows)' | Set-Content -Path $GlobalUsingsFile -Encoding UTF8
}

$txt = Get-Content -Path $GlobalUsingsFile -Raw -Encoding UTF8
$new = $txt
foreach($line in $required){
  if ($new -notmatch [Regex]::Escape($line)) { $new = ($new.TrimEnd() + "`r`n" + $line) }
}
if ($new -ne $txt) {
  $new | Set-Content -Path $GlobalUsingsFile -Encoding UTF8
  Ok "Updated $GlobalUsingsFile (W, WP, SWD, …)"
} else {
  Ok "GlobalUsings already contains required aliases"
}

# --- 2) Rewrite only Wordprocessing symbols from W.* -> WP.*
# Core Wordprocessing items (from your errors)
$wpExact = @(
  'Document','Body','Paragraph','ParagraphProperties','Justification','JustificationValues',
  'Run','RunProperties','RunFonts','FontSize','Bold','Italic','Underline','UnderlineValues','Color','Text','Break',
  'SpacingBetweenLines','LineSpacingRuleValues',
  'Table','TableRow','TableCell','TableProperties','TableBorders',
  'TopBorder','LeftBorder','BottomBorder','RightBorder','InsideHorizontalBorder','InsideVerticalBorder',
  'TableCellProperties','TableCellBorders','BorderValues'
)

# Get all .cs files except global usings and build outputs
$csFiles = Get-ChildItem -Path $ProjectDir -Recurse -Filter *.cs -File |
  Where-Object {
    $_.FullName -ne $GlobalUsingsFile -and
    $_.DirectoryName -notmatch '\\(bin|obj|\.vs)(\\|$)'
  }

[int]$changed = 0
foreach($f in $csFiles){
  $text = Get-Content -Path $f.FullName -Raw -Encoding UTF8
  $newText = $text

  foreach($t in $wpExact){
    # Replace whole-word: W.<Type>  -> WP.<Type>
    $pattern = "\bW\.$t\b"
    $newText = [Regex]::Replace($newText, $pattern, "WP.$t")
  }

  if ($newText -ne $text){
    if ($Backup) { Copy-Item $f.FullName "$($f.FullName).bak" -Force }
    $newText | Set-Content -Path $f.FullName -Encoding UTF8
    $changed++
    Write-Host "  fixed WP aliases in $($f.FullName)"
  }
}

Ok "Files modified: $changed"

# --- 3) Optional: build to verify
if ($RunBuild) {
  Info "dotnet clean"
  dotnet clean "$CsprojPath" | Out-Host
  Info "dotnet build"
  dotnet build "$CsprojPath" | Out-Host
}

Ok "Done."
