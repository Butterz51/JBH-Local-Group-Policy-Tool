#Requires -Version 5.1
<#
  AUTHOR: JBH Services
  TITLE: Local Group Policy Script - PowerShell Backend
#>
[CmdletBinding()]
param(
  [string]$ConfigJson,
  [string]$ConfigJsonBase64,
  [string]$ConfigPath,
  [string]$LogFilePath,
  [string]$ResultFilePath,
  [switch]$FromGui
)

$ErrorActionPreference = 'Stop'

$script:LogDirectory = $null
$script:LogFile = $null

function Get-AppRoot {
  $scriptDirectory = if ($PSScriptRoot) {
    $PSScriptRoot
  }
  elseif ($PSCommandPath) {
    Split-Path -Path $PSCommandPath -Parent
  }
  else {
    (Get-Location).Path
  }

  $resolvedScriptPath = $null
  try {
    if ($PSCommandPath -and (Test-Path -LiteralPath $PSCommandPath)) {
      $resolvedScriptPath = (Resolve-Path -LiteralPath $PSCommandPath).Path
    }
  }
  catch {
    $resolvedScriptPath = $PSCommandPath
  }

  $candidate = $scriptDirectory
  for ($i = 0; $i -lt 6; $i++) {
    if ([string]::IsNullOrWhiteSpace($candidate)) {
      break
    }

    $expectedScript = Join-Path $candidate 'Data\Scripts\PowerShell\LocalGroupPolicyTool.ps1'
    if ($resolvedScriptPath -and (Test-Path -LiteralPath $expectedScript)) {
      try {
        if ((Resolve-Path -LiteralPath $expectedScript).Path -eq $resolvedScriptPath) {
          return $candidate
        }
      }
      catch {
        # Ignore resolution failures and continue.
      }
    }

    if ((Test-Path -LiteralPath (Join-Path $candidate 'README.md')) -or (Test-Path -LiteralPath (Join-Path $candidate 'Data'))) {
      return $candidate
    }

    $parent = Split-Path -Path $candidate -Parent
    if ($parent -eq $candidate) {
      break
    }

    $candidate = $parent
  }

  $fallback = $scriptDirectory
  for ($i = 0; $i -lt 3; $i++) {
    $parent = Split-Path -Path $fallback -Parent
    if ($parent -eq $fallback) {
      break
    }
    $fallback = $parent
  }

  return $fallback
}

function Initialize-Logging {
  $appRoot = Get-AppRoot
  $scriptRoot = if ($PSScriptRoot) {
    $PSScriptRoot
  }
  elseif ($PSCommandPath) {
    Split-Path -Path $PSCommandPath -Parent
  }
  else {
    (Get-Location).Path
  }

  if (-not [string]::IsNullOrWhiteSpace($LogFilePath)) {
    $script:LogFile = $LogFilePath
    $script:LogDirectory = Split-Path -Path $script:LogFile -Parent
  }
  else {
    $script:LogDirectory = Join-Path $appRoot 'Logs'
    $timestamp = Get-Date -Format 'MMddyyyy'
    $script:LogFile = Join-Path $script:LogDirectory ("{0}.log" -f $timestamp)
  }

  if (-not (Test-Path -LiteralPath $script:LogDirectory)) {
    New-Item -Path $script:LogDirectory -ItemType Directory -Force | Out-Null
  }

  $header = @(
    ('=' * 90),
    ("[{0}] [INFO] Logging started." -f (Get-Date -Format 'MM-dd-yyyy | HH:mm:ss')),
    ("[{0}] [INFO] AppRoot: {1}" -f (Get-Date -Format 'MM-dd-yyyy | HH:mm:ss'), $appRoot),
    ("[{0}] [INFO] ScriptRoot: {1}" -f (Get-Date -Format 'MM-dd-yyyy | HH:mm:ss'), $scriptRoot),
    ("[{0}] [INFO] ScriptPath: {1}" -f (Get-Date -Format 'MM-dd-yyyy | HH:mm:ss'), $PSCommandPath),
    ("[{0}] [INFO] User: {1}" -f (Get-Date -Format 'MM-dd-yyyy | HH:mm:ss'), $env:USERNAME),
    ("[{0}] [INFO] Computer: {1}" -f (Get-Date -Format 'MM-dd-yyyy | HH:mm:ss'), $env:COMPUTERNAME),
    ("[{0}] [INFO] PID: {1}" -f (Get-Date -Format 'MM-dd-yyyy | HH:mm:ss'), $PID),
    ('=' * 90)
  )

  if (Test-Path -LiteralPath $script:LogFile) {
    Add-Content -LiteralPath $script:LogFile -Value @('', ('=' * 90), ("[{0}] [INFO] Logging resumed in a new PowerShell process." -f (Get-Date -Format 'MM-dd-yyyy | HH:mm:ss'))) -Encoding UTF8
    Add-Content -LiteralPath $script:LogFile -Value $header -Encoding UTF8
  }
  else {
    Set-Content -LiteralPath $script:LogFile -Value $header -Encoding UTF8
  }
}

function Write-Log {
  param(
    [Parameter(Mandatory)][string]$Message,
    [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
  )

  try {
    if (-not $script:LogFile) {
      Initialize-Logging
    }

    $timestamp = Get-Date -Format 'MM-dd-yyyy | HH:mm:ss'
    Add-Content -LiteralPath $script:LogFile -Value ("[{0}] [{1}] {2}" -f $timestamp, $Level, $Message) -Encoding UTF8
  }
  catch {
    # Intentionally swallow logging failures so they do not break policy application.
  }
}

function Write-LogException {
  param([Parameter(Mandatory)]$Record)

  try {
    $message = if ($Record.Exception) { $Record.Exception.Message } else { [string]$Record }
    Write-Log -Level 'ERROR' -Message "Exception: $message"

    if ($Record.InvocationInfo) {
      if ($Record.InvocationInfo.PositionMessage) {
        Write-Log -Level 'ERROR' -Message ("Position: {0}" -f $Record.InvocationInfo.PositionMessage.Trim())
      }

      if ($Record.InvocationInfo.ScriptLineNumber) {
        Write-Log -Level 'ERROR' -Message ("ScriptLineNumber: {0}" -f $Record.InvocationInfo.ScriptLineNumber)
      }
    }

    if ($Record.ScriptStackTrace) {
      Write-Log -Level 'ERROR' -Message ("ScriptStackTrace: {0}" -f $Record.ScriptStackTrace.Trim())
    }
  }
  catch {
    # Intentionally swallow logging failures so they do not hide the original error.
  }
}


function Write-ResultSummary {
  param([Parameter(Mandatory)]$SummaryObject)

  try {
    $json = $SummaryObject | ConvertTo-Json -Compress -Depth 6

    if (-not [string]::IsNullOrWhiteSpace($ResultFilePath)) {
      $resultDirectory = Split-Path -Path $ResultFilePath -Parent
      if (-not [string]::IsNullOrWhiteSpace($resultDirectory) -and -not (Test-Path -LiteralPath $resultDirectory)) {
        New-Item -Path $resultDirectory -ItemType Directory -Force | Out-Null
      }
      $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
      [System.IO.File]::WriteAllText($ResultFilePath, $json, $utf8NoBom)
    }

    Write-Output $json
  }
  catch {
    Write-LogException -Record $_
  }
}

Initialize-Logging
Write-Log -Message 'PowerShell backend launch detected.'

function Test-IsAdministrator {
  $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = [Security.Principal.WindowsPrincipal]::new($identity)
  return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Restart-ElevatedAndWait {
  Write-Log -Message 'Preparing elevated PowerShell relaunch.'

  $psExe = Join-Path $PSHOME 'powershell.exe'
  if (-not (Test-Path -LiteralPath $psExe)) {
    $psExe = 'powershell.exe'
  }

  $argumentList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))

  if (-not [string]::IsNullOrWhiteSpace($ConfigJsonBase64)) {
    $argumentList += @('-ConfigJsonBase64', ('"{0}"' -f $ConfigJsonBase64))
  }
  elseif (-not [string]::IsNullOrWhiteSpace($ConfigJson)) {
    $encoded = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($ConfigJson))
    $argumentList += @('-ConfigJsonBase64', ('"{0}"' -f $encoded))
  }
  elseif (-not [string]::IsNullOrWhiteSpace($ConfigPath)) {
    $argumentList += @('-ConfigPath', ('"{0}"' -f $ConfigPath))
  }

  if (-not [string]::IsNullOrWhiteSpace($LogFilePath)) {
    $argumentList += @('-LogFilePath', ('"{0}"' -f $LogFilePath))
  }

  if (-not [string]::IsNullOrWhiteSpace($ResultFilePath)) {
    $argumentList += @('-ResultFilePath', ('"{0}"' -f $ResultFilePath))
  }

  $process = Start-Process -FilePath $psExe -Verb RunAs -ArgumentList ($argumentList -join ' ') -Wait -PassThru
  Write-Log -Message ("Elevated PowerShell process completed with exit code {0}." -f $process.ExitCode)
  exit $process.ExitCode
}

if (-not (Test-IsAdministrator)) {
  if ($FromGui) {
    $message = 'The GUI must be started as Administrator before applying settings.'
    Write-Log -Level 'ERROR' -Message $message
    throw $message
  }

  Write-Log -Level 'WARN' -Message 'Process is not elevated. Relaunching with administrative rights.'
  Restart-ElevatedAndWait
}

Write-Log -Message 'Administrative context confirmed.'

function Get-ConfigObject {
  $jsonText = $null
  $configSource = $null

  if (-not [string]::IsNullOrWhiteSpace($ConfigJsonBase64)) {
    $configSource = 'ConfigJsonBase64'
    $jsonText = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ConfigJsonBase64))
  }
  elseif (-not [string]::IsNullOrWhiteSpace($ConfigJson)) {
    $configSource = 'ConfigJson'
    $jsonText = $ConfigJson
  }
  elseif (-not [string]::IsNullOrWhiteSpace($ConfigPath)) {
    $configSource = 'ConfigPath'
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
      throw "Config file was not found: $ConfigPath"
    }
    $jsonText = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8
  }
  else {
    throw 'No GUI configuration was supplied. Pass -ConfigJsonBase64, -ConfigJson, or -ConfigPath.'
  }

  Write-Log -Message ("Loading GUI configuration from {0}." -f $configSource)

  $config = $jsonText | ConvertFrom-Json
  if ($null -eq $config) {
    throw 'GUI configuration could not be parsed.'
  }

  return $config
}

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
    [Parameter(Mandatory)][string]$DisplayName
  )

  [PSCustomObject]@{
    Key       = $Key
    ValueName = $ValueName
    ValueType = $ValueType
    DataBytes = $DataBytes
    Display   = $DisplayName
  }
}

function New-DwordPolicyEntry {
  param(
    [Parameter(Mandatory)][string]$Key,
    [Parameter(Mandatory)][string]$ValueName,
    [Parameter(Mandatory)][int]$Value,
    [Parameter(Mandatory)][string]$DisplayName
  )

  $bytes = [System.BitConverter]::GetBytes([int]$Value)
  New-PolicyEntry -Key $Key -ValueName $ValueName -ValueType $script:RegType.REG_DWORD -DataBytes $bytes -DisplayName $DisplayName
}

function New-StringPolicyEntry {
  param(
    [Parameter(Mandatory)][string]$Key,
    [Parameter(Mandatory)][string]$ValueName,
    [Parameter(Mandatory)][string]$Value,
    [Parameter(Mandatory)][string]$DisplayName
  )

  $bytes = [System.Text.Encoding]::Unicode.GetBytes($Value + "`0")
  New-PolicyEntry -Key $Key -ValueName $ValueName -ValueType $script:RegType.REG_SZ -DataBytes $bytes -DisplayName $DisplayName
}

function Convert-PolicyEntryToBytes {
  param([Parameter(Mandatory)]$Entry)

  [byte[]]$record = @()
  $unicode = [System.Text.Encoding]::Unicode

  $record += $unicode.GetBytes('[')
  $record += $unicode.GetBytes($Entry.Key + "`0")
  $record += $unicode.GetBytes(';')
  $record += $unicode.GetBytes($Entry.ValueName + "`0")
  $record += $unicode.GetBytes(';')
  $record += [System.BitConverter]::GetBytes([int]$Entry.ValueType)
  $record += $unicode.GetBytes(';')
  $record += [System.BitConverter]::GetBytes([int]$Entry.DataBytes.Length)
  $record += $unicode.GetBytes(';')
  $record += $Entry.DataBytes
  $record += $unicode.GetBytes(']')

  Write-Output -NoEnumerate ([byte[]]$record)
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
  param([Parameter(Mandatory)][string]$Path)

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
    if (($bytes.Length - $index) -lt 2) { break }

    $opening = [System.BitConverter]::ToChar($bytes, $index)
    if ($opening -ne '[') { break }
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

    [void]$entries.Add([PSCustomObject]@{
      Key       = $key
      ValueName = $valueName
      ValueType = $valueType
      DataBytes = $dataBytes
      Display   = ''
    })
  }

  return $entries
}


function Get-PolicyIdentity {
  param([Parameter(Mandatory)]$Entry)

  return ('{0}|{1}' -f [string]$Entry.Key, [string]$Entry.ValueName).ToLowerInvariant()
}

function Test-ByteArrayEqual {
  param(
    [Parameter(Mandatory)][byte[]]$Left,
    [Parameter(Mandatory)][byte[]]$Right
  )

  if ($Left.Length -ne $Right.Length) { return $false }
  for ($i = 0; $i -lt $Left.Length; $i++) {
    if ($Left[$i] -ne $Right[$i]) { return $false }
  }
  return $true
}

function Test-PolicyEntryEquivalent {
  param(
    [Parameter(Mandatory)]$Left,
    [Parameter(Mandatory)]$Right
  )

  if ((Get-PolicyIdentity -Entry $Left) -ne (Get-PolicyIdentity -Entry $Right)) { return $false }
  if ([int]$Left.ValueType -ne [int]$Right.ValueType) { return $false }

  [byte[]]$leftBytes = if ($null -ne $Left.DataBytes) { [byte[]]$Left.DataBytes } else { @() }
  [byte[]]$rightBytes = if ($null -ne $Right.DataBytes) { [byte[]]$Right.DataBytes } else { @() }

  return (Test-ByteArrayEqual -Left $leftBytes -Right $rightBytes)
}

function Write-PolicyFile {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Entries
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
    [void]$output.AddRange($entryBytes)
  }

  [System.IO.File]::WriteAllBytes($Path, $output.ToArray())
}

function Sync-PolicyEntries {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ManagedEntries,
    [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SelectedEntries
  )

  $current = Read-PolicyFile -Path $Path
  $managedById = @{}
  $selectedById = @{}
  $currentManagedById = @{}
  $merged = New-Object System.Collections.Generic.List[object]

  foreach ($entry in $ManagedEntries) {
    $managedById[(Get-PolicyIdentity -Entry $entry)] = $entry
  }

  foreach ($entry in $SelectedEntries) {
    $selectedById[(Get-PolicyIdentity -Entry $entry)] = $entry
  }

  foreach ($entry in $current) {
    $identity = Get-PolicyIdentity -Entry $entry
    if ($managedById.ContainsKey($identity)) {
      $currentManagedById[$identity] = $entry
    }
    else {
      [void]$merged.Add($entry)
    }
  }

  foreach ($entry in $SelectedEntries) {
    [void]$merged.Add($entry)
  }

  $added = 0
  $updated = 0
  $removed = 0
  $unchanged = 0

  foreach ($identity in $managedById.Keys) {
    $currentEntry = if ($currentManagedById.ContainsKey($identity)) { $currentManagedById[$identity] } else { $null }
    $selectedEntry = if ($selectedById.ContainsKey($identity)) { $selectedById[$identity] } else { $null }

    if ($null -eq $currentEntry -and $null -ne $selectedEntry) {
      $added++
      Write-Log -Message ("{0} ADD: {1}" -f $ScopeName, $displayName)
      continue
    }

    if ($null -ne $currentEntry -and $null -eq $selectedEntry) {
      $removed++
      Write-Log -Message ("{0} REMOVE: {1}" -f $ScopeName, $displayName)
      continue
    }

    if ($null -ne $currentEntry -and $null -ne $selectedEntry) {
      if (Test-PolicyEntryEquivalent -Left $currentEntry -Right $selectedEntry) {
        $unchanged++
        Write-Log -Message ("{0} UNCHANGED: {1}" -f $ScopeName, $displayName)
      }
      else {
        $updated++
        Write-Log -Message ("{0} UPDATE: {1}" -f $ScopeName, $displayName)
      }
    }
  }

  Write-PolicyFile -Path $Path -Entries $merged
  Write-Log -Message ("{0} sync complete. Processed: {1}; Added: {2}; Updated: {3}; Removed: {4}; Unchanged: {5}" -f $ScopeName, $SelectedEntries.Count, $added, $updated, $removed, $unchanged)

  return [pscustomobject]@{
    Processed = $SelectedEntries.Count
    Changed   = ($added + $updated + $removed)
    Added     = $added
    Updated   = $updated
    Removed   = $removed
    Unchanged = $unchanged
  }
}

# ---------------------------
# Config helpers
# ---------------------------

function Get-ConfigBool {
  param(
    [Parameter(Mandatory)]$Config,
    [Parameter(Mandatory)][string]$Name,
    [bool]$Default = $false
  )

  $prop = $Config.PSObject.Properties[$Name]
  if ($null -eq $prop) { return $Default }
  return [bool]$prop.Value
}

function Get-ConfigText {
  param(
    [Parameter(Mandatory)]$Config,
    [Parameter(Mandatory)][string]$Name,
    [string]$Default = ''
  )

  $prop = $Config.PSObject.Properties[$Name]
  if ($null -eq $prop -or $null -eq $prop.Value) { return $Default }
  return [string]$prop.Value
}

function Get-LeadingInt {
  param(
    [string]$Text,
    [int]$Fallback
  )

  if ([string]::IsNullOrWhiteSpace($Text)) { return $Fallback }
  if ($Text -match '^\s*(\d+)') {
    return [int]$Matches[1]
  }
  return $Fallback
}

function Convert-DayNameToNumber {
  param([string]$Day)

  switch ($Day) {
    'Sunday'    { return 1 }
    'Monday'    { return 2 }
    'Tuesday'   { return 3 }
    'Wednesday' { return 4 }
    'Thursday'  { return 5 }
    'Friday'    { return 6 }
    'Saturday'  { return 7 }
    default     { return 7 }
  }
}

function Convert-TimeTextToHour {
  param([string]$TimeText)

  if ([string]::IsNullOrWhiteSpace($TimeText)) { return 2 }

  try {
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $dt = [datetime]::ParseExact($TimeText.Trim(), 'h:mm tt', $culture)
    return [int]$dt.Hour
  }
  catch {
    return 2
  }
}

function Add-EntryIf {
  param(
    [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Entries,
    [Parameter(Mandatory)][bool]$Condition,
    [Parameter(Mandatory)]$Entry
  )

  if ($Condition) {
    [void]$Entries.Add($Entry)
  }
}

function Write-ConfigSnapshot {
  param([Parameter(Mandatory)]$Config)

  Write-Log -Message 'Resolved GUI configuration values:'

  foreach ($property in ($Config.PSObject.Properties | Sort-Object Name)) {
    $value = $property.Value

    if ($value -is [System.Array]) {
      $rendered = ($value -join ', ')
    }
    elseif ($null -eq $value) {
      $rendered = '<null>'
    }
    else {
      $rendered = [string]$value
    }

    Write-Log -Message ("  {0} = {1}" -f $property.Name, $rendered)
  }
}

# ---------------------------
# Catalog builders (GUI driven)
# ---------------------------

function Get-GuiMachinePolicyEntries { 
  param(
    [Parameter(Mandatory)]$Config,
    [switch]$All
  )

  $entries = New-Object System.Collections.Generic.List[object]
  $includeAll = [bool]$All

  # Parsed combo values
  $doMode = Get-LeadingInt -Text (Get-ConfigText $Config 'delivery_optimization_mode' '0 - HTTP only, no peering') -Fallback 0
  $deferDays = Get-LeadingInt -Text (Get-ConfigText $Config 'defer_feature_updates_days' '93 days') -Fallback 93
  $auOption = Get-LeadingInt -Text (Get-ConfigText $Config 'au_option' '4 - Auto download and schedule the install') -Fallback 4
  $scheduleMode = Get-ConfigText $Config 'scheduled_install_mode' 'Every week'
  $scheduleDayText = Get-ConfigText $Config 'scheduled_install_day' 'Saturday'
  $scheduleTimeText = Get-ConfigText $Config 'scheduled_install_time' '2:00 AM'

  $scheduledInstallEveryWeek = if ($scheduleMode -eq 'Every day') { 0 } else { 1 }
  $scheduledInstallDay = if ($scheduleMode -eq 'Every day') { 0 } else { Convert-DayNameToNumber -Day $scheduleDayText }
  $scheduledInstallTime = Convert-TimeTextToHour -TimeText $scheduleTimeText

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'sync_foreground_policy')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows NT\CurrentVersion\Winlogon' -ValueName 'SyncForegroundPolicy' -Value 1 -DisplayName 'Always wait for the network at computer startup and logon')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'hide_fast_user_switching')) (New-DwordPolicyEntry -Key 'Software\Microsoft\Windows\CurrentVersion\Policies\System' -ValueName 'HideFastUserSwitching' -Value 1 -DisplayName 'Hide Fast User Switching')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'no_lock_screen')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Personalization' -ValueName 'NoLockScreen' -Value 1 -DisplayName 'Do not display the lock screen')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'show_sleep_option_disabled')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'ShowSleepOption' -Value 0 -DisplayName 'Do Not Show Sleep in the power options menu')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'hide_taskview_button')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'HideTaskViewButton' -Value 1 -DisplayName 'Hide the TaskView button')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'configure_chat_icon')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Windows Chat' -ValueName 'ConfigureChatIcon' -Value 3 -DisplayName 'Do Not Show the Chat icon on the taskbar')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_feedback_notifications')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\DataCollection' -ValueName 'DoNotShowFeedbackNotifications' -Value 1 -DisplayName 'Do not show feedback notifications')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'allow_telemetry_zero')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\DataCollection' -ValueName 'AllowTelemetry' -Value 0 -DisplayName 'Do Not Allow Telemetry data collection')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'max_telemetry_zero')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\DataCollection' -ValueName 'MaxTelemetryAllowed' -Value 0 -DisplayName 'Do Not Allow Max Telemetry')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_enterprise_auth_proxy')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\DataCollection' -ValueName 'DisableEnterpriseAuthProxy' -Value 1 -DisplayName 'Disable Enterprise Auth Proxy')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'ceip_disabled')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\SQMClient\Windows' -ValueName 'CEIPEnable' -Value 0 -DisplayName 'Disable Customer Experience Improvement Program (CEIP)')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'let_apps_run_in_background_deny')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\AppPrivacy' -ValueName 'LetAppsRunInBackground' -Value 2 -DisplayName 'Do Not Let Windows apps run in the background')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_windows_error_reporting')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Windows Error Reporting' -ValueName 'Disabled' -Value 1 -DisplayName 'Disable Windows Error Reporting')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_cloud_optimized_content')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableCloudOptimizedContent' -Value 1 -DisplayName 'Turn off cloud optimized content')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_consumer_account_state_content')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableConsumerAccountStateContent' -Value 1 -DisplayName 'Turn off cloud consumer account state content')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_soft_landing')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableSoftLanding' -Value 1 -DisplayName 'Do not show Windows tips')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_windows_consumer_features')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableWindowsConsumerFeatures' -Value 1 -DisplayName 'Turn off Microsoft consumer experiences')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_default_save_to_onedrive')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\OneDrive' -ValueName 'DisableLibrariesDefaultSaveToOneDrive' -Value 1 -DisplayName 'Do Not Save documents to OneDrive by default')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_default_save_to_skydrive')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Skydrive' -ValueName 'DisableLibrariesDefaultSaveToSkyDrive' -Value 1 -DisplayName 'Do Not Save documents to OneDrive by default (legacy compatibility)')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_filesync_ngsc')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\OneDrive' -ValueName 'DisableFileSyncNGSC' -Value 1 -DisplayName 'Do Not Allow OneDrive sync client')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_filesync_legacy')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\OneDrive' -ValueName 'DisableFileSync' -Value 1 -DisplayName 'Do Not Allow legacy file sync')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_web_search')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Windows Search' -ValueName 'DisableWebSearch' -Value 1 -DisplayName 'Do not allow web search')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'connected_search_use_web_disabled')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Windows Search' -ValueName 'ConnectedSearchUseWeb' -Value 0 -DisplayName 'Do not search the web or display web results in Search')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'allow_widgets_disabled')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Dsh' -ValueName 'AllowNewsAndInterests' -Value 0 -DisplayName 'Do Not Allow widgets')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'hide_windows_insider_pages')) (New-StringPolicyEntry -Key 'Software\Microsoft\Windows\CurrentVersion\Policies\Explorer' -ValueName 'SettingsPageVisibility' -Value 'hide:windowsinsider;windowsinsider-optin' -DisplayName 'Hide Windows Insider pages')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'hide_account_protection')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows Defender Security Center\Account protection' -ValueName 'UILockdown' -Value 1 -DisplayName 'Hide the Account protection area')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'prevent_user_from_modifying_settings')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows Defender Security Center\App and Browser protection' -ValueName 'DisallowExploitProtectionOverride' -Value 1 -DisplayName 'Prevent user from modifying settings')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'hide_family_options')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows Defender Security Center\Family options' -ValueName 'UILockdown' -Value 1 -DisplayName 'Hide the Family options area')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'hide_non_critical_notifications')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows Defender Security Center\Notifications' -ValueName 'DisableEnhancedNotifications' -Value 1 -DisplayName 'Hide non-critical notifications')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_biometrics')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Biometrics' -ValueName 'Enabled' -Value 0 -DisplayName 'Do Not Allow the use of biometrics')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'configure_delivery_optimization_mode')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\DeliveryOptimization' -ValueName 'DODownloadMode' -Value $doMode -DisplayName 'Configure Delivery Optimization Download Mode')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'exclude_drivers_with_wu')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate' -ValueName 'ExcludeWUDriversInQualityUpdate' -Value 1 -DisplayName 'Do not include drivers with Windows Updates')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'defer_feature_updates_enabled')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate' -ValueName 'DeferFeatureUpdates' -Value 1 -DisplayName 'Select when Preview Builds and Feature Updates are received')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'defer_feature_updates_enabled')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate' -ValueName 'DeferFeatureUpdatesPeriodInDays' -Value $deferDays -DisplayName 'Feature Updates Deferral')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_temp_enterprise_feature_control')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate' -ValueName 'AllowTemporaryEnterpriseFeatureControl' -Value 0 -DisplayName 'Do Not Enable features introduced via servicing that are off by default')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'manage_preview_builds_disabled')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate' -ValueName 'ManagePreviewBuilds' -Value 0 -DisplayName 'Manage preview builds')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'configure_automatic_updates')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'NoAutoUpdate' -Value 0 -DisplayName 'Configure Automatic Updates')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'configure_automatic_updates')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'AUOptions' -Value $auOption -DisplayName 'AU Option')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'configure_automatic_updates')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'ScheduledInstallEveryWeek' -Value $scheduledInstallEveryWeek -DisplayName 'Scheduled install mode')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'configure_automatic_updates')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'ScheduledInstallDay' -Value $scheduledInstallDay -DisplayName 'Scheduled install day')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'configure_automatic_updates')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'ScheduledInstallTime' -Value $scheduledInstallTime -DisplayName 'Scheduled install time')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'allow_mu_update_service')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'AllowMUUpdateService' -Value 1 -DisplayName 'Install updates for other Microsoft products')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'no_auto_reboot_with_logged_on_users')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -ValueName 'NoAutoRebootWithLoggedOnUsers' -Value 1 -DisplayName 'No auto-restart with logged on user for scheduled automatic updates installations')

  return $entries
}

function Get-GuiUserPolicyEntries {
  param(
    [Parameter(Mandatory)]$Config,
    [switch]$All
  )

  $entries = New-Object System.Collections.Generic.List[object]
  $includeAll = [bool]$All

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'remove_lock_computer')) (New-DwordPolicyEntry -Key 'Software\Microsoft\Windows\CurrentVersion\Policies\System' -ValueName 'DisableLockWorkstation' -Value 1 -DisplayName 'Remove Lock Computer')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'remove_change_password')) (New-DwordPolicyEntry -Key 'Software\Microsoft\Windows\CurrentVersion\Policies\System' -ValueName 'DisableChangePassword' -Value 1 -DisplayName 'Remove Change Password')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'remove_logoff')) (New-DwordPolicyEntry -Key 'Software\Microsoft\Windows\CurrentVersion\Policies\Explorer' -ValueName 'NoLogoff' -Value 1 -DisplayName 'Remove Logoff')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'hide_recommended_sites')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'HideRecommendedPersonalizedSites' -Value 1 -DisplayName 'Remove Personalized Website Recommendations from Recommended')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'hide_recommended_section')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'HideRecommendedSection' -Value 1 -DisplayName 'Remove Recommended section from Start Menu')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'hide_most_used_list')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'ShowOrHideMostUsedApps' -Value 2 -DisplayName 'Do Not Show "Most used" list from Start Menu')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_searchbox_suggestions')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\Explorer' -ValueName 'DisableSearchBoxSuggestions' -Value 1 -DisplayName 'Turn off display of recent search entries in the File Explorer search box')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_account_notifications')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CurrentVersion\AccountNotifications' -ValueName 'DisableAccountNotifications' -Value 1 -DisplayName 'Turn off account notifications in Start')

  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_windows_spotlight_features')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableWindowsSpotlightFeatures' -Value 1 -DisplayName 'Turn off all Windows spotlight features')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_windows_spotlight_on_settings')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableWindowsSpotlightOnSettings' -Value 1 -DisplayName 'Turn off Windows Spotlight on Settings')
  Add-EntryIf $entries ($includeAll -or (Get-ConfigBool $Config 'disable_windows_welcome_experience')) (New-DwordPolicyEntry -Key 'Software\Policies\Microsoft\Windows\CloudContent' -ValueName 'DisableWindowsSpotlightWindowsWelcomeExperience' -Value 1 -DisplayName 'Turn off the Windows Welcome Experience')

  return $entries
}

# ---------------------------
# Apply
# ---------------------------

try {
  Write-Log -Message 'Beginning Local Group Policy apply operation.'
  $config = Get-ConfigObject
  Write-ConfigSnapshot -Config $config

  $machinePolicyPath = Join-Path $env:WINDIR 'System32\GroupPolicy\Machine\Registry.pol'
  $userPolicyPath    = Join-Path $env:WINDIR 'System32\GroupPolicy\User\Registry.pol'

  $managedMachineEntries  = Get-GuiMachinePolicyEntries -Config $config -All
  $selectedMachineEntries = Get-GuiMachinePolicyEntries -Config $config
  $managedUserEntries     = Get-GuiUserPolicyEntries -Config $config -All
  $selectedUserEntries    = Get-GuiUserPolicyEntries -Config $config

  Write-Log -Message ("Machine policy path: {0}" -f $machinePolicyPath)
  Write-Log -Message ("User policy path: {0}" -f $userPolicyPath)
  Write-Log -Message ("Managed machine entries: {0}; selected machine entries: {1}" -f $managedMachineEntries.Count, $selectedMachineEntries.Count)
  Write-Log -Message ("Managed user entries: {0}; selected user entries: {1}" -f $managedUserEntries.Count, $selectedUserEntries.Count)

  $machineResult = Sync-PolicyEntries -Path $machinePolicyPath -ManagedEntries $managedMachineEntries -SelectedEntries $selectedMachineEntries -ScopeName 'Machine'
  $userResult = Sync-PolicyEntries -Path $userPolicyPath -ManagedEntries $managedUserEntries -SelectedEntries $selectedUserEntries -ScopeName 'User'

  Write-Log -Message ("Machine changes detected: {0} (Added: {1}, Updated: {2}, Removed: {3}, Unchanged: {4})" -f $machineResult.Changed, $machineResult.Added, $machineResult.Updated, $machineResult.Removed, $machineResult.Unchanged)
  Write-Log -Message ("User changes detected: {0} (Added: {1}, Updated: {2}, Removed: {3}, Unchanged: {4})" -f $userResult.Changed, $userResult.Added, $userResult.Updated, $userResult.Removed, $userResult.Unchanged)

  $totalChanged = [int]$machineResult.Changed + [int]$userResult.Changed
  $restartRecommended = $totalChanged -gt 0
  $gpupdateRan = $false
  $gpupdateExitCode = $null

  if ($totalChanged -gt 0) {
    $gpupdateRan = $true
    Write-Log -Message 'Managed policy changes were detected. Running gpupdate.exe /force.'
    $gpupdateOutput = (& "$env:SystemRoot\System32\gpupdate.exe" /force 2>&1 | Out-String).Trim()
    $gpupdateExitCode = $LASTEXITCODE

    if ($gpupdateOutput) {
        foreach ($line in ($gpupdateOutput -split "`r?`n")) {
            if ($line.Trim()) {
                Write-Log -Message ("gpupdate: {0}" -f $line.Trim())
            }
        }
    }

    Write-Log -Message ("gpupdate.exe /force exit code: {0}" -f $gpupdateExitCode)

    if ($gpupdateExitCode -ne 0) {
        throw "gpupdate.exe /force failed with exit code $gpupdateExitCode."
    }
  }
  else {
    Write-Log -Message 'No managed policy changes detected. Skipping gpupdate.exe /force.'
  }

  $durationSeconds = [Math]::Round(((Get-Date) - $operationStart).TotalSeconds, 2)
  Write-Log -Message ("Policy apply duration: {0} seconds" -f $durationSeconds)

  $summary = [PSCustomObject]@{
    Status                  = 'Success'
    MachineEntriesProcessed = $machineResult.Processed
    UserEntriesProcessed    = $userResult.Processed
    MachineEntriesChanged   = $machineResult.Changed
    UserEntriesChanged      = $userResult.Changed
    MachineEntriesAdded     = $machineResult.Added
    MachineEntriesUnchanged  = $machineResult.Unchanged
    MachineEntriesUpdated   = $machineResult.Updated
    MachineEntriesRemoved   = $machineResult.Removed
    UserEntriesAdded        = $userResult.Added
    UserEntriesUnchanged     = $userResult.Unchanged
    UserEntriesUpdated      = $userResult.Updated
    UserEntriesRemoved      = $userResult.Removed
    TotalEntriesChanged     = $totalChanged
    RestartRecommended       = $restartRecommended
    GpUpdateRan              = $gpupdateRan
    GpUpdateExitCode         = $gpupdateExitCode
    DurationSeconds          = $durationSeconds
    LogFile                  = $script:LogFile
  }

  Write-Log -Message 'Policy application completed successfully.'
  Write-ResultSummary -SummaryObject $summary
  exit 0
}
catch {
  $failedDurationSeconds = $null
  try {
    if ($operationStart) {
      $failedDurationSeconds = [Math]::Round(((Get-Date) - $operationStart).TotalSeconds, 2)
      Write-Log -Message ("Policy apply failed after {0} seconds" -f $failedDurationSeconds) -Level 'ERROR'
    }
  }
  catch {
  }

  Write-LogException -Record $_

  $errorSummary = [PSCustomObject]@{
    Status          = 'Error'
    Message         = $_.Exception.Message
    DurationSeconds = $failedDurationSeconds
    LogFile         = $script:LogFile
  }

  Write-ResultSummary -SummaryObject $errorSummary
  Write-Error $_.Exception.Message
  exit 1
}
