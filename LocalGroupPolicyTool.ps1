#Requires -Version 5.1
<#
.SYNOPSIS
  Applies the JBH Services Local Group Policy baseline and optional policy selections.

.DESCRIPTION
  This tool applies the JBH Services Local Group Policy baseline to the PC.
  Required policy settings are always applied.
  Optional policy settings can be reviewed, changed, previewed, and either included or skipped before changes are written.
  After the changes are applied, Local Group Policy is refreshed and the user can choose whether to restart immediately or restart later.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
  $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = [Security.Principal.WindowsPrincipal]::new($identity)
  return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Restart-Elevated {
  $psExe = Join-Path $PSHOME 'powershell.exe'
  if (-not (Test-Path -LiteralPath $psExe)) {
    $psExe = 'powershell.exe'
  }

  $arguments = @(
    '-NoProfile'
    '-ExecutionPolicy', 'Bypass'
    '-File', ('"{0}"' -f $PSCommandPath)
  )

  Start-Process -FilePath $psExe -Verb RunAs -ArgumentList ($arguments -join ' ')
  exit
}

if (-not (Test-IsAdministrator)) {
  Write-Host 'Restarting script as Administrator...' -ForegroundColor Yellow
  Restart-Elevated
}

# ---------------------------
# Console window size / lock
# ---------------------------

function Set-LockedConsoleWindow {
  param(
    [int]$Width  = 130,
    [int]$Height = 43
  )

  $rawUI = $Host.UI.RawUI

  # Make sure buffer is large enough before setting window size
  $bufferSize = $rawUI.BufferSize
  if ($bufferSize.Width -lt $Width)  { $bufferSize.Width  = $Width }
  if ($bufferSize.Height -lt $Height) { $bufferSize.Height = 3000 }
  $rawUI.BufferSize = $bufferSize

  # Set visible window size
  $windowSize = $rawUI.WindowSize
  $windowSize.Width  = $Width
  $windowSize.Height = $Height
  $rawUI.WindowSize = $windowSize
}

Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class ConsoleWindowLock {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DrawMenuBar(IntPtr hWnd);

    public const int GWL_STYLE = -16;
    public const int WS_SIZEBOX = 0x00040000;
    public const int WS_MAXIMIZEBOX = 0x00010000;

    public static void LockSize() {
        IntPtr hWnd = GetConsoleWindow();
        if (hWnd == IntPtr.Zero) return;

        int style = GetWindowLong(hWnd, GWL_STYLE);
        style &= ~WS_SIZEBOX;
        style &= ~WS_MAXIMIZEBOX;
        SetWindowLong(hWnd, GWL_STYLE, style);
        DrawMenuBar(hWnd);
    }
}
"@

Set-LockedConsoleWindow -Width 130 -Height 43
[ConsoleWindowLock]::LockSize()

$rawUI = $Host.UI.RawUI
$rawUI.BackgroundColor = 'Black'
$rawUI.ForegroundColor = 'Gray'
Clear-Host

# ---------------------------
# Registry.pol helpers
# ---------------------------

$script:RegType = @{
  REG_NONE      = 0
  REG_SZ        = 1
  REG_EXPAND_SZ = 2
  REG_BINARY    = 3
  REG_DWORD     = 4
  REG_MULTI_SZ  = 7
  REG_QWORD     = 11
}

function New-PolicyEntry {
  param(
    [Parameter(Mandatory)][string]$Key,
    [Parameter(Mandatory)][string]$ValueName,
    [Parameter(Mandatory)][int]$ValueType,
    [Parameter(Mandatory)][byte[]]$DataBytes,
    [Parameter(Mandatory)][string]$DisplayName,
    [string]$OptionalKey = ''
  )

  [PSCustomObject]@{
    Key         = $Key
    ValueName   = $ValueName
    ValueType   = $ValueType
    DataBytes   = $DataBytes
    Display     = $DisplayName
    OptionalKey = $OptionalKey
  }
}

function New-DwordPolicyEntry {
  param(
    [Parameter(Mandatory)][string]$Key,
    [Parameter(Mandatory)][string]$ValueName,
    [Parameter(Mandatory)][int]$Value,
    [Parameter(Mandatory)][string]$DisplayName,
    [string]$OptionalKey = ''
  )

  $bytes = [System.BitConverter]::GetBytes([int]$Value)

  New-PolicyEntry -Key $Key -ValueName $ValueName -ValueType $script:RegType.REG_DWORD -DataBytes $bytes -DisplayName $DisplayName -OptionalKey $OptionalKey
}

function New-StringPolicyEntry {
  param(
    [Parameter(Mandatory)][string]$Key,
    [Parameter(Mandatory)][string]$ValueName,
    [Parameter(Mandatory)][string]$Value,
    [Parameter(Mandatory)][string]$DisplayName,
    [string]$OptionalKey = ''
  )

  $bytes = [System.Text.Encoding]::Unicode.GetBytes($Value + "`0")

  New-PolicyEntry -Key $Key -ValueName $ValueName -ValueType 1 -DataBytes $bytes -DisplayName $DisplayName -OptionalKey $OptionalKey
}

function Convert-PolicyEntryToBytes {
  param(
    [Parameter(Mandatory)]$Entry
  )

  $unicode = [System.Text.Encoding]::Unicode
  $bytes = New-Object 'System.Collections.Generic.List[byte]'

  $bytes.AddRange($unicode.GetBytes('['))
  $bytes.AddRange($unicode.GetBytes($Entry.Key + "`0"))
  $bytes.AddRange($unicode.GetBytes(';'))
  $bytes.AddRange($unicode.GetBytes($Entry.ValueName + "`0"))
  $bytes.AddRange($unicode.GetBytes(';'))
  $bytes.AddRange([System.BitConverter]::GetBytes([int]$Entry.ValueType))
  $bytes.AddRange($unicode.GetBytes(';'))
  $bytes.AddRange([System.BitConverter]::GetBytes([int]$Entry.DataBytes.Length))
  $bytes.AddRange($unicode.GetBytes(';'))
  $bytes.AddRange([byte[]]$Entry.DataBytes)
  $bytes.AddRange($unicode.GetBytes(']'))

  return $bytes.ToArray()
}

function Get-UnicodeStringUntilDelimiter {
  param(
    [Parameter(Mandatory)][byte[]]$Bytes,
    [Parameter(Mandatory)][ref]$Index,
    [Parameter(Mandatory)][char]$Delimiter
  )

  $unicode = [System.Text.Encoding]::Unicode
  $start = $Index.Value

  while ($Index.Value -lt ($Bytes.Length - 1)) {
    $char = [System.BitConverter]::ToChar($Bytes, $Index.Value)
    if ($char -eq $Delimiter) {
      break
    }
    $Index.Value += 2
  }

  if ($Index.Value -ge ($Bytes.Length - 1)) {
    throw "Invalid registry.pol format. Delimiter '$Delimiter' was not found."
  }

  $segmentLength = $Index.Value - $start
  [byte[]]$stringBytes = @()
  if ($segmentLength -gt 0) {
    $stringBytes = $Bytes[$start..($Index.Value - 1)]
  }

  $text = if ($stringBytes.Length -gt 0) {
    $unicode.GetString($stringBytes).TrimEnd([char]0)
  }
  else {
    ''
  }

  $Index.Value += 2
  return $text
}

function Read-PolicyFile {
  param(
    [Parameter(Mandatory)][string]$Path
  )

  $entries = New-Object System.Collections.Generic.List[object]

  if (-not (Test-Path -LiteralPath $Path)) {
    return $entries
  }

  [byte[]]$bytes = [System.IO.File]::ReadAllBytes($Path)
  if ($bytes.Length -lt 8) {
    return $entries
  }

  $signature = [System.Text.Encoding]::ASCII.GetString($bytes, 0, 4)
  if ($signature -ne 'PReg') {
    throw "Invalid registry.pol header in '$Path'."
  }

  $version = [System.BitConverter]::ToInt32($bytes, 4)
  if ($version -ne 1) {
    throw "Unsupported registry.pol version '$version' in '$Path'."
  }

  $index = 8

  while ($index -lt ($bytes.Length - 1)) {
    if (($bytes.Length - $index) -lt 2) {
      break
    }

    $opening = [System.BitConverter]::ToChar($bytes, $index)
    if ($opening -ne '[') {
      break
    }
    $index += 2

    $idxRef = [ref]$index
    $key = Get-UnicodeStringUntilDelimiter -Bytes $bytes -Index $idxRef -Delimiter ';'
    $index = $idxRef.Value

    $idxRef = [ref]$index
    $valueName = Get-UnicodeStringUntilDelimiter -Bytes $bytes -Index $idxRef -Delimiter ';'
    $index = $idxRef.Value

    $valueType = [System.BitConverter]::ToInt32($bytes, $index)
    $index += 4

    if ([System.BitConverter]::ToChar($bytes, $index) -ne ';') {
      throw "Invalid registry.pol format in '$Path' after type field."
    }
    $index += 2

    $dataLength = [System.BitConverter]::ToInt32($bytes, $index)
    $index += 4

    if ([System.BitConverter]::ToChar($bytes, $index) -ne ';') {
      throw "Invalid registry.pol format in '$Path' after size field."
    }
    $index += 2

    [byte[]]$dataBytes = @()
    if ($dataLength -gt 0) {
      $dataBytes = $bytes[$index..($index + $dataLength - 1)]
      $index += $dataLength
    }

    if ([System.BitConverter]::ToChar($bytes, $index) -ne ']') {
      throw "Invalid registry.pol format in '$Path' before record close."
    }
    $index += 2

    $entries.Add([PSCustomObject]@{
      Key         = $key
      ValueName   = $valueName
      ValueType   = $valueType
      DataBytes   = $dataBytes
      Display     = ''
      OptionalKey = ''
    })
  }

  return $entries
}

function Write-PolicyFile {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][System.Collections.Generic.List[object]]$Entries
  )

  $directory = Split-Path -Path $Path -Parent
  if (-not (Test-Path -LiteralPath $directory)) {
    New-Item -Path $directory -ItemType Directory -Force | Out-Null
  }

  $output = New-Object System.Collections.Generic.List[byte]
  $output.AddRange([System.BitConverter]::GetBytes(0x67655250))
  $output.AddRange([System.BitConverter]::GetBytes(1))

  foreach ($entry in $Entries) {
    [byte[]]$entryBytes = Convert-PolicyEntryToBytes -Entry $entry
    $output.AddRange($entryBytes)
  }

  [System.IO.File]::WriteAllBytes($Path, $output.ToArray())
}

function Sync-PolicyEntries {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][object[]]$ManagedEntries,
    [Parameter(Mandatory)][object[]]$SelectedEntries
  )

  $current = Read-PolicyFile -Path $Path
  $managedKeys = New-Object 'System.Collections.Generic.HashSet[string]'

  foreach ($entry in $ManagedEntries) {
    $null = $managedKeys.Add(('{0}|{1}' -f $entry.Key, $entry.ValueName))
  }

  $merged = New-Object System.Collections.Generic.List[object]

  foreach ($entry in $current) {
    $identity = '{0}|{1}' -f $entry.Key, $entry.ValueName
    if (-not $managedKeys.Contains($identity)) {
      $merged.Add($entry)
    }
  }

  foreach ($entry in $SelectedEntries) {
    $merged.Add($entry)
  }

  Write-PolicyFile -Path $Path -Entries $merged
}

# ---------------------------
# Menu helpers
# ---------------------------

function Read-Choice {
  param(
    [Parameter(Mandatory)][string]$Prompt,
    [Parameter(Mandatory)][string[]]$Allowed
  )

  while ($true) {
    $inputValue = (Read-Host $Prompt).Trim().ToUpperInvariant()
    if ($Allowed -contains $inputValue) {
      return $inputValue
    }

    Write-Host ("Invalid selection. Allowed values: {0}" -f ($Allowed -join ', ')) -ForegroundColor Yellow
  }
}

function Read-ApplyConfirmation {
  param(
    [string]$Prompt = ' What would you like to do with these optional policy selections?'
  )

  while ($true) {
    Write-Host ''
    Write-Host $Prompt -ForegroundColor Cyan
    Write-Host '   A = Apply these selections now'
    Write-Host '   C = Continue changing optional selections'
    Write-Host '   M = Return to the Main Menu'
    Write-Host '   E = Exit without making changes'
    Write-Host ''

    $answer = (Read-Host ' Choose A, C, M, or E').Trim().ToUpperInvariant()

    switch ($answer) {
      'A' { return 'Apply' }
      'C' { return 'Change' }
      'M' { return 'Main' }
      'E' { return 'Exit' }
      default { Write-Host ' Please enter A, C, M, or E.' -ForegroundColor Yellow }
    }
  }
}

function Read-RestartChoice {
  while ($true) {
    Write-Host ''
    Write-Host ' Restart Required?' -ForegroundColor Cyan
    Write-Host ' Some policy changes may not fully take effect until Windows is restarted.'
    Write-Host ''
    Write-Host '   R = Restart now'
    Write-Host '   N = Not now'
    Write-Host ''

    $answer = (Read-Host ' Choose R or N').Trim().ToUpperInvariant()

    switch ($answer) {
      'R' { return $true }
      'N' { return $false }
      default { Write-Host ' Please enter R or N.' -ForegroundColor Yellow }
    }
  }
}

# ---------------------------
# Optional selections
# ---------------------------

# The keys in this hashtable correspond to the OptionalKey property of the policy entries. The values represent the default selections for this tool, but can be changed by the user during configuration.
$optionalSelections = [ordered]@{
  ApplyOneDriveDefaultSaveDisabled     = $true
  ApplySearchDisableWebSearch          = $true
  ApplySearchNoWebResults              = $true
  ApplyWidgetsDisabled                 = $true
  ApplyHideWindowsInsiderPages         = $true
  ApplyHideFastUserSwitching           = $true
  ApplyDisableLockWorkstation          = $true
  ApplyNoLockScreen                    = $true
  ApplyLetAppsRunInBackground          = $true
  ApplyDisableWindowsErrorReporting    = $true
  ApplyHideSleepOption                 = $true
  ApplyHideTaskViewButton              = $true
  ApplyDisableChatIcon                 = $true
  ApplyStartRemovePersonalizedWebsites = $true
  ApplyStartRemoveRecommendedSection   = $true
  ApplyStartHideMostUsedList           = $true
  ApplyAccountNotificationsOff         = $true
  ApplySpotlightAllOff                 = $true
  ApplySpotlightOnSettingsOff          = $true
  ApplyWelcomeExperienceOff            = $true
  ApplyDisableSearchBoxSuggestions     = $true
}

# Display labels for the optional policies, used during configuration and summary display. Keys must match those in $optionalSelections.
$optionalLabels = [ordered]@{
  ApplyOneDriveDefaultSaveDisabled     = 'Computer > OneDrive > Save documents to OneDrive by default'
  ApplySearchDisableWebSearch          = 'Computer > Search > Do not allow web search'
  ApplySearchNoWebResults              = 'Computer > Search > Don''t search the web or display web results in Search'
  ApplyWidgetsDisabled                 = 'Computer > Widgets > Allow widgets'
  ApplyHideWindowsInsiderPages         = 'Computer > Control Panel > Settings Page Visibility > Hide Windows Insider pages'
  ApplyHideFastUserSwitching           = 'Computer > System > Logon > Hide Fast User Switching'
  ApplyDisableLockWorkstation          = 'Computer > System > Ctrl+Alt+Del Options > Remove Lock Computer'
  ApplyNoLockScreen                    = 'Computer > Personalization > Do not display the lock screen'
  ApplyLetAppsRunInBackground          = 'Computer > App Privacy > Let Windows apps run in the background'
  ApplyDisableWindowsErrorReporting    = 'Computer > Windows Components > Windows Error Reporting > Disable Windows Error Reporting'
  ApplyHideSleepOption                 = 'Computer > Windows Components > File Explorer > Show sleep in the power options menu'
  ApplyHideTaskViewButton              = 'Computer > Start Menu and Taskbar > Hide the TaskView button'
  ApplyDisableChatIcon                 = 'Computer > Windows Components > Chat > Configure the Chat icon on the taskbar'
  ApplyStartRemovePersonalizedWebsites = 'User > Start Menu > Remove Personalized Website Recommendations'
  ApplyStartRemoveRecommendedSection   = 'User > Start Menu > Remove Recommended section'
  ApplyStartHideMostUsedList           = 'User > Start Menu and Taskbar > Show or hide "Most used" list from Start menu'
  ApplyAccountNotificationsOff         = 'User > Account Notifications > Turn off account notifications in Start'
  ApplySpotlightAllOff                 = 'User > Cloud Content > Turn off all Windows spotlight features'
  ApplySpotlightOnSettingsOff          = 'User > Cloud Content > Turn off Windows Spotlight on Settings'
  ApplyWelcomeExperienceOff            = 'User > Cloud Content > Turn off the Windows Welcome Experience'
  ApplyDisableSearchBoxSuggestions     = 'User > Windows Explorer > Turn off display of recent search entries in the File Explorer search box'
}

# Default states for the optional policies, used for display during configuration. These do not affect the actual default selections, which are all set to $true in $optionalSelections.
$optionalDefaultStates = [ordered]@{
  ApplyOneDriveDefaultSaveDisabled     = 'Disabled'
  ApplySearchDisableWebSearch          = 'Enabled'
  ApplySearchNoWebResults              = 'Enabled'
  ApplyWidgetsDisabled                 = 'Disabled'
  ApplyHideWindowsInsiderPages         = 'Enabled'
  ApplyHideFastUserSwitching           = 'Enabled'
  ApplyDisableLockWorkstation          = 'Enabled'
  ApplyNoLockScreen                    = 'Enabled'
  ApplyLetAppsRunInBackground          = 'Force Deny'
  ApplyDisableWindowsErrorReporting    = 'Enabled'
  ApplyHideSleepOption                 = 'Disabled'
  ApplyHideTaskViewButton              = 'Enabled'
  ApplyDisableChatIcon                 = 'Disabled'
  ApplyStartRemovePersonalizedWebsites = 'Enabled'
  ApplyStartRemoveRecommendedSection   = 'Enabled'
  ApplyStartHideMostUsedList           = 'Hide'
  ApplyAccountNotificationsOff         = 'Enabled'
  ApplySpotlightAllOff                 = 'Enabled'
  ApplySpotlightOnSettingsOff          = 'Enabled'
  ApplyWelcomeExperienceOff            = 'Enabled'
  ApplyDisableSearchBoxSuggestions     = 'Enabled'
}

function Show-CurrentSelections {
  param(
    [Parameter(Mandatory)][System.Collections.IDictionary]$Selections,
    [Parameter(Mandatory)][System.Collections.IDictionary]$Labels,
    [switch]$HideTitle
  )

  $appliedCount = 0
  foreach ($key in $Labels.Keys) {
    if ($Selections[$key]) {
      $appliedCount++
    }
  }

  Write-Host ''

  if (-not $HideTitle) {
    Write-Host ' Current Optional Policy Selections:' -ForegroundColor Cyan
    Write-Host ''
  }

  if ($appliedCount -eq 0) {
    Write-Host ' All optional policy selections are currently set to Skip.' -ForegroundColor Yellow
    Write-Host ''
    return
  }

  foreach ($key in $Labels.Keys) {
    if ($Selections[$key]) {
      Write-Host ("  - {0}" -f $Labels[$key])
    }
    else {
      Write-Host ("  - {0} [Skip]" -f $Labels[$key])
    }
  }

  Write-Host ''
}

function Show-SelectionSummary {
  param(
    [Parameter(Mandatory)][System.Collections.IDictionary]$Selections,
    [Parameter(Mandatory)][System.Collections.IDictionary]$Labels
  )

  $appliedCount = 0

  foreach ($key in $Labels.Keys) {
    if ($Selections[$key]) {
      $appliedCount++
    }
  }

  Write-Host ''

  if ($appliedCount -eq 0) {
    Write-Host ' All optional policy selections are currently set to Skip.' -ForegroundColor Yellow
    Write-Host ''
    return
  }

  Write-Host ' Updated Optional Policy Selections:' -ForegroundColor Cyan
  Write-Host ''

  foreach ($key in $Labels.Keys) {
    $state = if ($Selections[$key]) { 'Apply' } else { 'Skip' }
    Write-Host ("  - {0} [{1}]" -f $Labels[$key], $state)
  }

  Write-Host ''
}

function Show-PolicyPreview {
  param(
    [Parameter(Mandatory)][System.Collections.IDictionary]$Selections
  )

  $computerCatalog = Get-ComputerPolicyCatalog
  $userCatalog     = Get-UserPolicyCatalog

  $computerSelected = Get-SelectedPolicyEntries -Catalog $computerCatalog -Selections $Selections
  $userSelected     = Get-SelectedPolicyEntries -Catalog $userCatalog -Selections $Selections

  Clear-Host
  Write-Host ''
  Write-Host ' POLICY PREVIEW:' -ForegroundColor Cyan
  Write-Host ' This preview shows the policies that will be included if you run the tool now.'
  Write-Host ''

  Write-Host ' REQUIRED + SELECTED COMPUTER POLICIES:' -ForegroundColor Green
  if ($computerSelected.Count -gt 0) {
    foreach ($entry in $computerSelected) {
      Write-Host ("  - {0}" -f $entry.Display)
    }
  }
  else {
    Write-Host '  - None'
  }

  Write-Host ''
  Write-Host ' REQUIRED + SELECTED USER POLICIES:' -ForegroundColor Green
  if ($userSelected.Count -gt 0) {
    foreach ($entry in $userSelected) {
      Write-Host ("  - {0}" -f $entry.Display)
    }
  }
  else {
    Write-Host '  - None'
  }

  Write-Host ''
  Write-Host ' OPTIONAL POLICIES CURRENTLY SET TO SKIP:' -ForegroundColor Yellow

  $skippedAny = $false
  foreach ($key in $optionalLabels.Keys) {
    if (-not $Selections[$key]) {
      $skippedAny = $true
      Write-Host ("  - {0}" -f $optionalLabels[$key])
    }
  }

  if (-not $skippedAny) {
    Write-Host '  - None'
  }

  Write-Host ''
  Read-Host ' Press Enter to return to the Main Menu'
}

function Copy-Selections {
  param(
    [Parameter(Mandatory)][System.Collections.IDictionary]$Source
  )

  $copy = [ordered]@{}
  foreach ($key in $Source.Keys) {
    $copy[$key] = $Source[$key]
  }

  return $copy
}

function Read-ConfigChoice {
  param(
    [Parameter(Mandatory)][string]$Question,
    [Parameter(Mandatory)][string]$DefaultState,
    [Parameter(Mandatory)][bool]$CurrentValue
  )

  $currentText = if ($CurrentValue) { 'Apply' } else { 'Skip' }

  while ($true) {
    Write-Host ''
    Write-Host " $Question" -ForegroundColor Cyan
    Write-Host "   JBH default policy state: $DefaultState"
    Write-Host "   Current selection for this run: $currentText"
    Write-Host ''
    $answer = (Read-Host ' Choose [Y=Apply / N=Skip / B=Back / H=Home / E=Exit]').Trim().ToUpperInvariant()

    switch ($answer) {
      'Y' { return 'Y' }
      'N' { return 'N' }
      'B' { return 'B' }
      'H' { return 'H' }
      'E' { return 'E' }
      ''  { return if ($CurrentValue) { 'Y' } else { 'N' } }
      default { Write-Host ' Please enter Y, N, B, H, or E.' -ForegroundColor Yellow }
    }
  }
}

function Configure-OptionalSelections {
  param(
    [Parameter(Mandatory)][System.Collections.IDictionary]$Selections,
    [Parameter(Mandatory)][System.Collections.IDictionary]$Labels,
    [Parameter(Mandatory)][System.Collections.IDictionary]$DefaultStates
  )

  $working = Copy-Selections -Source $Selections

  # Keys are in the order they will be presented to the user during configuration
  $questionKeys = @(
    'ApplyOneDriveDefaultSaveDisabled'
    'ApplySearchDisableWebSearch'
    'ApplySearchNoWebResults'
    'ApplyWidgetsDisabled'
    'ApplyHideWindowsInsiderPages'
    'ApplyHideFastUserSwitching'
    'ApplyDisableLockWorkstation'
    'ApplyNoLockScreen'
    'ApplyLetAppsRunInBackground'
    'ApplyDisableWindowsErrorReporting'
    'ApplyHideSleepOption'
    'ApplyHideTaskViewButton'
    'ApplyDisableChatIcon'
    'ApplyStartRemovePersonalizedWebsites'
    'ApplyStartRemoveRecommendedSection'
    'ApplyStartHideMostUsedList'
    'ApplyAccountNotificationsOff'
    'ApplySpotlightAllOff'
    'ApplySpotlightOnSettingsOff'
    'ApplyWelcomeExperienceOff'
    'ApplyDisableSearchBoxSuggestions'
  )

  $index = 0

  while ($index -lt $questionKeys.Count) {
    $key = $questionKeys[$index]

    $choice = Read-ConfigChoice `
      -Question $Labels[$key] `
      -DefaultState $DefaultStates[$key] `
      -CurrentValue $working[$key]

    switch ($choice) {
      'Y' {
        $working[$key] = $true
        $index++
      }
      'N' {
        $working[$key] = $false
        $index++
      }
      'B' {
        if ($index -gt 0) {
          $index--
        }
      }
      'E' {
        return @{
          Action     = 'Exit'
          Selections = $null
        }
      }
      'H' {
        Write-Host ''
        Write-Host ' Returning to the Main Menu. No changes from this edit session were saved.' -ForegroundColor Yellow
        Write-Host ''
        Read-Host ' Press Enter to continue'
        return @{
          Action     = 'Home'
          Selections = $Selections
        }
      }
    }
  }

  return @{
    Action     = 'Done'
    Selections = $working
  }
}

function Show-MainMenu {
  param(
    [Parameter(Mandatory)][System.Collections.IDictionary]$Selections,
    [Parameter(Mandatory)][System.Collections.IDictionary]$Labels
  )

  while ($true) {
    Clear-Host
    Write-Host ''
    Write-Host ' JBH Services - Local Group Policy Settings Tool' -ForegroundColor DarkGreen
    Write-Host ''
    Write-Host ' DESCRIPTION:' -ForegroundColor Green
    Write-Host '  This tool applies the JBH Services Local Group Policy baseline to this PC.'
    Write-Host '  Some policy settings are required and will always be applied.'
    Write-Host '  Other policy settings are optional and can be reviewed before they are included.'
    Write-Host ''
    Write-Host ''
    Write-Host ' OPTIONAL SETTINGS:' -ForegroundColor Green
    Write-Host '  Apply = include the optional policy in the changes made by this run.'
    Write-Host '  Skip  = leave that optional policy out of this run.'
    Write-Host ''
    Write-Host ''
    Write-Host ' After the changes are applied, Local Group Policy is refreshed and the system restarts after 10 seconds.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host ''
    Write-Host ' MAIN MENU:' -ForegroundColor Cyan
    Write-Host '   D = Use JBH default policy selections'
    Write-Host '   C = Review or change optional selections'
    Write-Host '   P = Preview policies that will be changed'
    Write-Host '   E = Exit without making changes'
    Write-Host ''

    $choice = Read-Choice -Prompt ' Press D to use JBH defaults, C to review/change optional selections, P to preview, or E to exit' -Allowed @('D','C','P','E')

    switch ($choice) {
      'D' { return $true }
      'E' { return $false }
      'P' {
        Show-PolicyPreview -Selections $Selections
        continue
      }

      'C' {
        while ($true) {
          Clear-Host
          Show-CurrentSelections -Selections $Selections -Labels $Labels

          Write-Host ' OPTIONAL POLICY SELECTION MENU:' -ForegroundColor Cyan
          Write-Host '   M = Return to the Main Menu'
          Write-Host '   C = Change optional policy selections'
          Write-Host '   E = Exit without making changes'
          Write-Host ''

          $subChoice = Read-Choice -Prompt ' Press M for the Main Menu, C to change optional selections, or E to exit' -Allowed @('M','C','E')

          if ($subChoice -eq 'M') {
            break
          }

          if ($subChoice -eq 'E') {
            return $false
          }

          if ($subChoice -eq 'C') {
            Clear-Host
            $configResult = Configure-OptionalSelections -Selections $Selections -Labels $Labels -DefaultStates $optionalDefaultStates

            if ($configResult.Action -eq 'Exit') {
              return $false
            }

            if ($configResult.Action -eq 'Home') {
              break
            }

            $selectionKeys = @($Selections.Keys)
            foreach ($key in $selectionKeys) {
              $Selections[$key] = $configResult.Selections[$key]
            }

            Clear-Host
            Write-Host ''
            Write-Host ' Review Optional Policy Settings.' -ForegroundColor Green
            Write-Host ''

            Show-SelectionSummary -Selections $Selections -Labels $Labels

            $nextAction = Read-ApplyConfirmation

            if ($nextAction -eq 'Apply') {
              return $true
            }

            if ($nextAction -eq 'Exit') {
              return $false
            }

            if ($nextAction -eq 'Main') {
              break
            }

            if ($nextAction -eq 'Change') {
              continue
            }
          }
        }
      }
    }
  }
}

# ---------------------------
# Policy catalog
# ---------------------------
function Get-ComputerPolicyCatalog {
  $entries = New-Object System.Collections.Generic.List[object]

  # ============================================================
  # Computer > System / Sign-in
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows NT\CurrentVersion\Winlogon' -ValueName 'SyncForegroundPolicy' -Value 1 -DisplayName 'Computer > System > Logon > Always wait for the network at computer startup and logon = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Microsoft\Windows\CurrentVersion\Policies\System' -ValueName 'HideFastUserSwitching' -Value 1 -DisplayName 'Computer > System > Logon > Hide Fast User Switching = Enabled' -OptionalKey 'ApplyHideFastUserSwitching'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Microsoft\Windows\CurrentVersion\Policies\System' -ValueName 'DisableLockWorkstation' -Value 1 -DisplayName 'Computer > System > Ctrl+Alt+Del Options > Remove Lock Computer = Enabled' -OptionalKey 'ApplyDisableLockWorkstation'))

  # ============================================================
  # Computer > Personalization / Lock Screen
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Personalization' -ValueName 'NoLockScreen' -Value 1 -DisplayName 'Computer > Personalization > Do not display the lock screen = Enabled' -OptionalKey 'ApplyNoLockScreen'))

  # ============================================================
  # Computer > File Explorer / Power Menu
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'ShowSleepOption' -Value 0 -DisplayName 'Computer > Windows Components > File Explorer > Show sleep in the power options menu = Disabled' -OptionalKey 'ApplyHideSleepOption'))

  # ============================================================
  # Computer > Start Menu and Taskbar / Chat
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'HideTaskViewButton' -Value 1 -DisplayName 'Computer > Start Menu and Taskbar > Hide the TaskView button = Enabled' -OptionalKey 'ApplyHideTaskViewButton'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Windows Chat' -ValueName 'ConfigureChatIcon' -Value 3 -DisplayName 'Computer > Windows Components > Chat > Configure the Chat icon on the taskbar = Disabled' -OptionalKey 'ApplyDisableChatIcon'))
  
  # ============================================================
  # Computer > Biometrics
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Biometrics' -ValueName 'Enabled' -Value 0 -DisplayName 'Computer > Windows Components > Biometrics > Allow the use of biometrics = Disabled'))

  # ============================================================
  # Computer > Privacy / Telemetry / CEIP / Error Reporting
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\DataCollection' -ValueName 'DoNotShowFeedbackNotifications' -Value 1 -DisplayName 'Computer > Windows Components > Data Collection and Preview Builds > Do not show feedback notifications = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\DataCollection' -ValueName 'AllowTelemetry' -Value 0 -DisplayName 'Computer > Data Collection > Allow Telemetry = 0'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\DataCollection' -ValueName 'MaxTelemetryAllowed' -Value 0 -DisplayName 'Computer > Data Collection > Max Telemetry Allowed = 0'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\DataCollection' -ValueName 'DisableEnterpriseAuthProxy' -Value 1 -DisplayName 'Computer > Data Collection > Disable Enterprise Auth Proxy = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\SQMClient\Windows' -ValueName 'CEIPEnable' -Value 0 -DisplayName 'Computer > CEIP > CEIP Enable = Disabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\AppPrivacy' -ValueName 'LetAppsRunInBackground' -Value 2 -DisplayName 'Computer > App Privacy > Let Windows apps run in the background = Force Deny' -OptionalKey 'ApplyLetAppsRunInBackground'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Windows Error Reporting' -ValueName 'Disabled' -Value 1 -DisplayName 'Computer > Windows Components > Windows Error Reporting > Disable Windows Error Reporting = Enabled' -OptionalKey 'ApplyDisableWindowsErrorReporting'))

  # ============================================================
  # Computer > Delivery Optimization
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\DeliveryOptimization' -ValueName 'DODownloadMode' -Value 0 -DisplayName 'Computer > Windows Components > Delivery Optimization > Download Mode = HTTP only, no peering'))

  # ============================================================
  # Computer > Cloud Content / Consumer Experience
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableCloudOptimizedContent' -Value 1 -DisplayName 'Computer > Windows Components > Cloud Content > Turn off cloud optimized content = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableConsumerAccountStateContent' -Value 1 -DisplayName 'Computer > Windows Components > Cloud Content > Turn off cloud consumer account state content = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableSoftLanding' -Value 1 -DisplayName 'Computer > Windows Components > Cloud Content > Do not show Windows tips = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableWindowsConsumerFeatures' -Value 1 -DisplayName 'Computer > Windows Components > Cloud Content > Turn off Microsoft consumer experiences = Enabled'))

  # ============================================================
  # Computer > OneDrive
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\OneDrive' -ValueName 'DisableLibrariesDefaultSaveToOneDrive' -Value 1 -DisplayName 'Computer > Windows Components > OneDrive > Save documents to OneDrive by default = Disabled' -OptionalKey 'ApplyOneDriveDefaultSaveDisabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Skydrive' -ValueName 'DisableLibrariesDefaultSaveToSkyDrive' -Value 1 -DisplayName 'Computer > Windows Components > OneDrive > Save documents to OneDrive by default = Disabled (legacy compatibility)' -OptionalKey 'ApplyOneDriveDefaultSaveDisabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\OneDrive' -ValueName 'DisableFileSyncNGSC' -Value 1 -DisplayName 'Computer > OneDrive > Prevent OneDrive sync client = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\OneDrive' -ValueName 'DisableFileSync' -Value 1 -DisplayName 'Computer > OneDrive > Disable legacy file sync = Enabled'))

  # ============================================================
  # Computer > Search / Widgets / Insider
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Windows Search' -ValueName 'DisableWebSearch' -Value 1 -DisplayName 'Computer > Windows Components > Search > Do not allow web search = Enabled' -OptionalKey 'ApplySearchDisableWebSearch'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Windows Search' -ValueName 'ConnectedSearchUseWeb' -Value 0 -DisplayName 'Computer > Windows Components > Search > Don''t search the web or display web results in Search = Enabled' -OptionalKey 'ApplySearchNoWebResults'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Dsh' -ValueName 'AllowNewsAndInterests' -Value 0 -DisplayName 'Computer > Windows Components > Widgets > Allow widgets = Disabled' -OptionalKey 'ApplyWidgetsDisabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate' -ValueName 'ManagePreviewBuilds' -Value 0 -DisplayName 'Computer > Windows Components > Windows Update > Manage preview builds = Disabled'))
  $entries.Add((New-StringPolicyEntry -Key 'Software\Microsoft\Windows\CurrentVersion\Policies\Explorer' -ValueName 'SettingsPageVisibility' -Value 'hide:windowsinsider;windowsinsider-optin' -DisplayName 'Computer > Control Panel > Settings Page Visibility > Hide Windows Insider pages = Enabled' -OptionalKey 'ApplyHideWindowsInsiderPages'))

  # ============================================================
  # Computer > Windows Security
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows Defender Security Center\Account protection' -ValueName 'UILockdown' -Value 1 -DisplayName 'Computer > Windows Components > Windows Security > Hide the Account protection area = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows Defender Security Center\App and Browser protection' -ValueName 'DisallowExploitProtectionOverride' -Value 1 -DisplayName 'Computer > Windows Components > Windows Security > Prevent user from modifying settings = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows Defender Security Center\Family options' -ValueName 'UILockdown' -Value 1 -DisplayName 'Computer > Windows Components > Windows Security > Hide the Family options area = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows Defender Security Center\Notifications' -ValueName 'DisableEnhancedNotifications' -Value 1 -DisplayName 'Computer > Windows Components > Windows Security > Hide non-critical notifications = Enabled'))

  # ============================================================
  # Computer > Windows Update > Core
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate' -ValueName 'ExcludeWUDriversInQualityUpdate' -Value 1 -DisplayName 'Computer > Windows Components > Windows Update > Do not include drivers with Windows Updates = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate' -ValueName 'DeferFeatureUpdates' -Value 1 -DisplayName 'Computer > Windows Components > Windows Update > Select when Preview Builds and Feature Updates are received = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate' -ValueName 'DeferFeatureUpdatesPeriodInDays' -Value 93 -DisplayName 'Computer > Windows Components > Windows Update > Feature Updates deferral = 93 days'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate' -ValueName 'AllowTemporaryEnterpriseFeatureControl' -Value 0 -DisplayName 'Computer > Windows Components > Windows Update > Enable features introduced via servicing that are off by default = Disabled'))

  # ============================================================
  # Computer > Windows Update > Automatic Updates Schedule
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'NoAutoUpdate' -Value 0 -DisplayName 'Computer > Windows Components > Windows Update > Configure Automatic Updates = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'AUOptions' -Value 4 -DisplayName 'Computer > Windows Components > Windows Update > AU option = 4 (Auto download and schedule the install)'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'ScheduledInstallEveryWeek' -Value 1 -DisplayName 'Computer > Windows Components > Windows Update > Scheduled install = Every week'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'ScheduledInstallDay' -Value 7 -DisplayName 'Computer > Windows Components > Windows Update > Scheduled install day = Saturday'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'ScheduledInstallTime' -Value 2 -DisplayName 'Computer > Windows Components > Windows Update > Scheduled install time = 2:00 AM'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'AllowMUUpdateService' -Value 1 -DisplayName 'Computer > Windows Components > Windows Update > Install updates for other Microsoft products = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'NoAutoRebootWithLoggedOnUsers' -Value 1 -DisplayName 'Computer > Windows Components > Windows Update > No auto-restart with logged on user for scheduled automatic updates installations = Enabled'))

  return $entries
}

function Get-UserPolicyCatalog {
  $entries = New-Object System.Collections.Generic.List[object]

  # ============================================================
  # User > System > Ctrl+Alt+Del Options
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Microsoft\Windows\CurrentVersion\Policies\System' -ValueName 'DisableChangePassword' -Value 1 -DisplayName 'User > System > Ctrl+Alt+Del Options > Remove Change Password = Enabled'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Microsoft\Windows\CurrentVersion\Policies\Explorer' -ValueName 'NoLogoff' -Value 1 -DisplayName 'User > System > Ctrl+Alt+Del Options > Remove Logoff = Enabled'))

  # ============================================================
  # User > Start Menu / Explorer
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'HideRecommendedPersonalizedSites' -Value 1 -DisplayName 'User > Start Menu and Taskbar > Remove Personalized Website Recommendations from Recommended = Enabled' -OptionalKey 'ApplyStartRemovePersonalizedWebsites'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'HideRecommendedSection' -Value 1 -DisplayName 'User > Start Menu and Taskbar > Remove Recommended section from Start Menu = Enabled' -OptionalKey 'ApplyStartRemoveRecommendedSection'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'ShowOrHideMostUsedApps' -Value 2 -DisplayName 'User > Start Menu and Taskbar > Show or hide "Most used" list from Start Menu = Hide' -OptionalKey 'ApplyStartHideMostUsedList'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'DisableSearchBoxSuggestions' -Value 1 -DisplayName 'User > Windows Explorer > Turn off display of recent search entries in the File Explorer search box = Enabled' -OptionalKey 'ApplyDisableSearchBoxSuggestions'))

  # ============================================================
  # User > Account Notifications
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CurrentVersion\AccountNotifications' -ValueName 'DisableAccountNotifications' -Value 1 -DisplayName 'User > Windows Components > Account Notifications > Turn off account notifications in Start = Enabled' -OptionalKey 'ApplyAccountNotificationsOff'))

  # ============================================================
  # User > Spotlight / Cloud Content
  # ============================================================
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableWindowsSpotlightFeatures' -Value 1 -DisplayName 'User > Windows Components > Cloud Content > Turn off all Windows spotlight features = Enabled' -OptionalKey 'ApplySpotlightAllOff'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableWindowsSpotlightOnSettings' -Value 1 -DisplayName 'User > Windows Components > Cloud Content > Turn off Windows Spotlight on Settings = Enabled' -OptionalKey 'ApplySpotlightOnSettingsOff'))
  $entries.Add((New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableWindowsSpotlightWindowsWelcomeExperience' -Value 1 -DisplayName 'User > Windows Components > Cloud Content > Turn off the Windows Welcome Experience = Enabled' -OptionalKey 'ApplyWelcomeExperienceOff'))

  return $entries
}

function Get-SelectedPolicyEntries {
  param(
    [Parameter(Mandatory)][object[]]$Catalog,
    [Parameter(Mandatory)][System.Collections.IDictionary]$Selections
  )

  $selected = New-Object System.Collections.Generic.List[object]

  foreach ($entry in $Catalog) {
    if ([string]::IsNullOrWhiteSpace($entry.OptionalKey)) {
      $selected.Add($entry)
      continue
    }

    if ($Selections.Contains($entry.OptionalKey) -and $Selections[$entry.OptionalKey]) {
      $selected.Add($entry)
    }
  }

  return $selected
}

function Apply-PolicyEntries {
  param(
    [Parameter(Mandatory)][object[]]$ManagedEntries,
    [Parameter(Mandatory)][object[]]$SelectedEntries,
    [Parameter(Mandatory)][string]$TargetPath
  )

  $total = $SelectedEntries.Count
  $counter = 0

  foreach ($entry in $SelectedEntries) {
    $counter++
    Write-Host ("[{0}/{1}] {2}" -f $counter, $total, $entry.Display) -ForegroundColor Green
  }

  Sync-PolicyEntries -Path $TargetPath -ManagedEntries $ManagedEntries -SelectedEntries $SelectedEntries
}

# ---------------------------
# Run
# ---------------------------

try {
  $continue = Show-MainMenu -Selections $optionalSelections -Labels $optionalLabels
  if (-not $continue) {
    Write-Host ' Cancelled. No changes were made.' -ForegroundColor Yellow
    exit
  }

  $machinePolicyPath = Join-Path $env:WINDIR 'System32\GroupPolicy\Machine\Registry.pol'
  $userPolicyPath    = Join-Path $env:WINDIR 'System32\GroupPolicy\User\Registry.pol'

  $computerCatalog = Get-ComputerPolicyCatalog
  $userCatalog     = Get-UserPolicyCatalog

  $computerSelected = Get-SelectedPolicyEntries -Catalog $computerCatalog -Selections $optionalSelections
  $userSelected     = Get-SelectedPolicyEntries -Catalog $userCatalog -Selections $optionalSelections

  Write-Host ''
  Write-Host ' Applying Local Group Policy settings...' -ForegroundColor Cyan
  Write-Host ''

  if ($computerSelected.Count -gt 0) {
    Write-Host ' Computer Configuration:' -ForegroundColor Cyan
    Apply-PolicyEntries -ManagedEntries $computerCatalog -SelectedEntries $computerSelected -TargetPath $machinePolicyPath
    Write-Host ''
  }
  else {
    Sync-PolicyEntries -Path $machinePolicyPath -ManagedEntries $computerCatalog -SelectedEntries @()
  }

  if ($userSelected.Count -gt 0) {
    Write-Host ' User Configuration:' -ForegroundColor Cyan
    Apply-PolicyEntries -ManagedEntries $userCatalog -SelectedEntries $userSelected -TargetPath $userPolicyPath
    Write-Host ''
  }
  else {
    Sync-PolicyEntries -Path $userPolicyPath -ManagedEntries $userCatalog -SelectedEntries @()
  }

  Write-Host ' Refreshing Local Group Policy...' -ForegroundColor Cyan
  & gpupdate.exe /force | Out-Null

  $restartNow = Read-RestartChoice

  if ($restartNow) {
    Write-Host ''
    Write-Host ' Done. The system will restart in 3 seconds.' -ForegroundColor Yellow
    shutdown.exe /r /t 3
  }
  else {
    Write-Host ''
    Write-Host ' Done. Local Group Policy has been refreshed.' -ForegroundColor Green
    Write-Host ' Some changes may not fully take effect until the next restart.' -ForegroundColor Yellow
    Write-Host ''
    Read-Host ' Press Enter to close'
  }
}
catch {
  Write-Host ''
  Write-Host ' The script failed:' -ForegroundColor Red
  Write-Host $_.Exception.Message -ForegroundColor Red
  Write-Host ''
  Read-Host ' Press Enter to close'
}
