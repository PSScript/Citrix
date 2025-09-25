<#
###############################################################################################
# Migrate-FSLogixOST.ps1 v1.2.2 (deutsch, voll gehärtet)
#
# Zweck: Migration, Analyse & Reporting von Outlook OST-Dateien (FSLogix/Citrix)
# - Von OfficeContainer → UserContainer
# - Analyse verdächtiger/corrupt OSTs (z.B. 64 MB)
# - Reporting (JSON/CSV), Event-Log-Sammlung, Deferred Write, Lazy-Admin-Modus
# - Vollautomatisch, GPO-Ready, produktiv für große Umgebungen
#
# Autoren: OpenAI GPT + Jan Hübener, 2025
###############################################################################################
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # Mapping und Benutzersteuerung
    [Parameter(Mandatory=$true)] [string]$MappingFile,
    [string]$CentralReport,       # Veraltet, durch Lazy-Admin-Parameter ersetzt
    [string]$ProcessedLog,       # Veraltet, durch Lazy-Admin-Parameter ersetzt
    [string[]]$UserList,

    # Lazy-Admin: Report-Ziel automatisch wählen
    [string]$ReportFolder,        # z.B. "C:\Skripte\Reports"
    [string]$ReportFileShare,     # z.B. "\\FS01\OSTMig\Reports"

    # Hauptschalter
    [switch]$ReportOnly,
    [switch]$Migrate,
    [switch]$FixPerms,
    [switch]$PermissiveACL,
    [switch]$StartOutlook,
    [int]$OutlookWarmupSeconds = 20,
    [switch]$CollectLogs,
    [ValidateSet('CurrentSession','HostAll')][string]$EventScope = 'CurrentSession',
    [int]$EventDays = 7,
    [switch]$AnalysisOnly,
    [switch]$AnalyzeOST,
    [switch]$AnalyzeEventLogs,

    # OST/Datei-Prüfung
    [long]$SuspectLowerBytes = 67108864,        # 64 MiB
    [long]$SuspectUpperBytes = 67236864,        # + ca. 128KB
    [long]$SuspectPaddingBytes = 128000,        # Convenience
    [int]$CorruptAgeDays = 2,
    [int]$CorruptAgeHours = 0,
    [switch]$Checksum,
    [switch]$DryRun,
    [switch]$ForceRecheck,
    [ValidateSet('SourceOnly','TargetOnly','Both')][string]$SourceSelect = 'Both',

    # Migration/Backup/Sicherheit
    [int]$MaxFileSizeMB = 16384,
    [int]$FreeSpaceSafetyMB = 512,
    [int]$MaxBackups = 2,
    [int]$ThrottleDelaySec = 0,
    [int]$StopAfterNUsers = 0,
    [switch]$RegkeyBackupAndRemove,
    [switch]$NoRegkeys,
    [switch]$DeferOnPermissionError,

    # Outlook Profil-Heal
    [switch]$AutoHealProfile,
    [string]$PRFFile,

    # Config-Persistenz
    [switch]$SaveConfigJson,
    [switch]$LoadConfigJson,
    [string]$ConfigPath = "$env:ProgramData\OSTMigration\last_config.json",
    [switch]$SaveConfigRegistry,
    [switch]$LoadConfigRegistry,
    [string]$ConfigRegRoot = "HKCU:\\Software\\OSTMigration",
    [switch]$UseLastConfig
)

###############################################################################################
# Hilfsfunktionen: Config-Laden/Speichern, Path-Resolver, Deferred Write, ACL, Logging etc.
###############################################################################################

# --- (1) Config persistieren/laden ---
function Get-ConfigSafeParams {
    param([hashtable]$Bound = $PSBoundParameters)
    $allow = @(
        'MappingFile','ReportFolder','ReportFileShare','CentralReport','ProcessedLog','UserList',
        'ReportOnly','Migrate','FixPerms','StartOutlook','CollectLogs','DryRun','Checksum',
        'AutoHealProfile','PRFFile','RegkeyBackupAndRemove','NoRegkeys',
        'SuspectLowerBytes','SuspectUpperBytes','SuspectPaddingBytes',
        'CorruptAgeDays','CorruptAgeHours','OutlookWarmupSeconds','FreeSpaceSafetyMB',
        'MaxFileSizeMB','MaxBackups','ThrottleDelaySec','StopAfterNUsers',
        'AnalysisOnly','AnalyzeOST','AnalyzeEventLogs','EventScope','EventDays',
        'ForceRecheck','SourceSelect','DeferOnPermissionError','PermissiveACL'
    )
    $cfg = @{}
    foreach ($k in $allow) {
        if ($Bound.ContainsKey($k) -and $null -ne $Bound[$k]) { $cfg[$k] = $Bound[$k] }
    }
    return $cfg
}
function Save-ConfigJson {
    param([hashtable]$Config, [string]$Path)
    $dir = Split-Path $Path -Parent
    if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    $Config | ConvertTo-Json -Depth 6 | Set-Content -Path $Path -Encoding UTF8
    Write-Host "• Konfiguration als JSON gespeichert: $Path"
}
function Load-ConfigJson {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return @{} }
    try { return (Get-Content $Path -Raw | ConvertFrom-Json) | ForEach-Object { $_ } } catch { return @{} }
}
function Save-ConfigRegistry {
    param([hashtable]$Config, [string]$Root = "HKCU:\\Software\\OSTMigration")
    if (-not (Test-Path $Root)) { New-Item -Path $Root -Force | Out-Null }
    $key = Join-Path $Root "Config"
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    foreach ($name in $Config.Keys) {
        $val = $Config[$name]
        if ($val -is [System.Array]) { $val = ($val | ConvertTo-Json -Compress) }
        Set-ItemProperty -Path $key -Name $name -Value $val -Force
    }
    Write-Host "• Konfiguration in Registry gespeichert: $key"
}
function Load-ConfigRegistry {
    param([string]$Root = "HKCU:\\Software\\OSTMigration")
    $key = Join-Path $Root "Config"
    if (-not (Test-Path $key)) { return @{} }
    $props = Get-ItemProperty -Path $key
    $cfg = @{}
    foreach ($p in $props.PSObject.Properties) {
        if ($p.Name -in 'PSPath','PSParentPath','PSChildName','PSDrive','PSProvider') { continue }
        $v = $p.Value
        if ($null -ne $v -and $v -is [string] -and $v.Trim().StartsWith('[')) {
            try { $v = $v | ConvertFrom-Json } catch {}
        }
        $cfg[$p.Name] = $v
    }
    return $cfg
}

# --- (2) Report-Folder/Share-Resolver: Lazy-Admin und Fallback-Logik ---
function _Pick-ReportRoots {
    param(
        [string]$ReportFolder,
        [string]$ReportFileShare,
        [string]$CentralReport,
        [string]$ProcessedLog
    )
    $chosenBase = $null
    if ($ReportFileShare) {
        try {
            if (Test-Path $ReportFileShare) { $chosenBase = $ReportFileShare }
            else {
                $tmp = Join-Path $ReportFileShare ("_probe_" + [IO.Path]::GetRandomFileName())
                New-Item -Path $tmp -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Remove-Item $tmp -Force -ErrorAction SilentlyContinue
                $chosenBase = $ReportFileShare
            }
        } catch { }
    }
    if (-not $chosenBase -and $ReportFolder) { $chosenBase = $ReportFolder }
    if (-not $chosenBase -and $CentralReport) { $chosenBase = $CentralReport }
    if (-not $chosenBase) { $chosenBase = Join-Path $env:ProgramData "OSTMigration\Reports" }
    $finalCentral = $CentralReport
    if (-not $finalCentral) { $finalCentral = $chosenBase }
    $finalProcessed = $ProcessedLog
    if (-not $finalProcessed) { $finalProcessed = Join-Path $chosenBase "ProcessedOSTs.csv" }
    try { if (-not (Test-Path $finalCentral)) { New-Item -Path $finalCentral -ItemType Directory -Force | Out-Null } } catch {}
    try { if (-not (Test-Path (Split-Path $finalProcessed -Parent))) { New-Item -Path (Split-Path $finalProcessed -Parent) -ItemType Directory -Force | Out-Null } } catch {}
    [pscustomobject]@{
        BaseChosen     = $chosenBase
        CentralReport  = $finalCentral
        ProcessedLog   = $finalProcessed
        UsingShare     = ($chosenBase -like '\\*')
    }
}

###############################################################################################
# Hilfsfunktionen: Path-Klassifikation, Deferred Write, ACL, Locks, Regkeys, Logging, etc.
###############################################################################################

# --- (3) Path-Klassifikation (UNC, lokal, gemappt, erreichbar?) ---
function Classify-Path {
    param([Parameter(Mandatory)][string]$Path)
    $out = [ordered]@{
        Input       = $Path
        Type        = 'Unknown'   # LocalFixed | UNC | MappedNetwork | Removable | Unknown
        Root        = $null
        Server      = $null
        Share       = $null
        IsAccessible= $false
    }
    try {
        $resolved = Resolve-Path -Path $Path -ErrorAction Stop | Select-Object -First 1
        $p = $resolved.Path
    } catch { $p = $Path }
    if ($p -match '^(\\\\\\\\)([^\\\\]+)\\\\([^\\\\]+)') {
        $out.Type  = 'UNC'
        $out.Root  = ($matches[0])
        $out.Server= $matches[2]
        $out.Share = $matches[3]
    } elseif ($p -match '^[A-Za-z]:\\\\') {
        $drive = Get-PSDrive -Name $p.Substring(0,1) -ErrorAction SilentlyContinue
        if ($drive) {
            if ($drive.DisplayRoot) { $out.Type = 'MappedNetwork' }
            else { switch ($drive.Provider.Name) { 'FileSystem' { $out.Type = 'LocalFixed' } default { $out.Type = $drive.Provider.Name } } }
            $out.Root = $drive.Root
        }
    }
    try {
        $probe = Test-Path $Path
        if (-not $probe) {
            $parent = Split-Path $Path -Parent
            if ($parent) { $probe = Test-Path $parent }
        }
        $out.IsAccessible = [bool]$probe
    } catch { $out.IsAccessible = $false }
    return [pscustomobject]$out
}

function Test-WriteAccess {
    param([Parameter(Mandatory)][string]$Folder)
    try {
        if (-not (Test-Path $Folder)) {
            New-Item -Path $Folder -ItemType Directory -Force | Out-Null
        }
        $tmp = Join-Path $Folder ('.wtest_'+[IO.Path]::GetRandomFileName())
        Set-Content -Path $tmp -Value 'x' -Encoding ASCII
        Remove-Item $tmp -Force
        return $true
    } catch { return $false }
}

# --- (4) Deferred Write (mit Fallback) ---
function Write-Resilient {
    param(
        [Parameter(Mandatory)][string]$Content,
        [Parameter(Mandatory)][string]$TargetPath,
        [switch]$DeferOnPermissionError,
        [string]$DeferRoot = \"$env:ProgramData\\OSTMigration\\Queue\"
    )
    $targetDir = Split-Path $TargetPath -Parent
    $tmp = Join-Path $env:TEMP ([IO.Path]::GetRandomFileName())
    Set-Content -Path $tmp -Value $Content -Encoding UTF8
    for($i=1;$i -le 3;$i++){
        try {
            if (Test-WriteAccess $targetDir) {
                Move-Item -Path $tmp -Destination $TargetPath -Force
                return @{ Status='OK'; Path=$TargetPath }
            }
        } catch { }
        Start-Sleep -Seconds (2*$i)
    }
    if ($DeferOnPermissionError) {
        try {
            if (-not (Test-Path $DeferRoot)) { New-Item -Path $DeferRoot -ItemType Directory -Force | Out-Null }
            $deferPath = Join-Path $DeferRoot ((Split-Path $TargetPath -Leaf) + \".deferred\")
            Move-Item -Path $tmp -Destination $deferPath -Force
            return @{ Status='DEFERRED'; Path=$deferPath }
        } catch { }
    }
    try { Copy-Item -Path $tmp -Destination $TargetPath -Force } catch { }
    Remove-Item $tmp -ErrorAction SilentlyContinue
    return @{ Status='FAILED'; Path=$TargetPath }
}

# --- (5) ACL-Vergabe, permissiv (Optional) ---
function Set-FilePermissiveACL {
    param([string]$Path, [switch]$PermissiveACL)
    try {
        $acl = Get-Acl $Path
        $acc = New-Object System.Security.AccessControl.FileSystemAccessRule(\"$env:USERDOMAIN\\$env:USERNAME\",\"FullControl\",\"Allow\")
        $acl.SetAccessRule($acc); Set-Acl $Path $acl
        return $true
    } catch {
        if ($PermissiveACL) { Write-Warning \"ACL setzen fehlgeschlagen (permissiv): $Path\"; return $false }
        else { throw \"ACL setzen fehlgeschlagen: $Path — verwende -PermissiveACL für Soft-Fail.\" }
    }
}

# --- (6) User-Lock (pro Benutzer; verhindert parallele Verarbeitung) ---
function New-UserLock {
    param([string]$User, [string]$CentralReport)
    $lockPath = Join-Path $CentralReport \"$($User)_migration.lock\"
    if(Test-Path $lockPath){ Write-Warning \"Lock existiert für $User. Überspringe.\"; return $null }
    try { New-Item -Path $lockPath -ItemType File -Force | Out-Null; return $lockPath } catch { return $null }
}
function Remove-UserLock { param([string]$LockPath) if($LockPath -and (Test-Path $LockPath)){ Remove-Item $LockPath -Force -ErrorAction SilentlyContinue } }

# --- (7) Registry-OOM-Keys Backup, Set und Restore ---
$__RegBackup = $null
function Get-OOMRegPath { \"HKCU:\\Software\\Policies\\Microsoft\\Office\\16.0\\Outlook\\Security\" }
function Backup-OOMKeys {
    $regPath = Get-OOMRegPath
    $keys = \"AdminSecurityMode\",\"ObjectModelGuard\",\"promptoomaddressinformationaccess\"
    $b=@{}; foreach($k in $keys){ $b[$k] = (Get-ItemProperty -Path $regPath -Name $k -ErrorAction SilentlyContinue).$k }
    $__script:__RegBackup = $b
}
function Set-OOMKeys {
    $regPath = Get-OOMRegPath
    New-Item -Path $regPath -Force | Out-Null
    Set-ItemProperty -Path $regPath -Name AdminSecurityMode -Value 3 -Force
    Set-ItemProperty -Path $regPath -Name ObjectModelGuard -Value 2 -Force
    Set-ItemProperty -Path $regPath -Name promptoomaddressinformationaccess -Value 0 -Force
}
function Restore-OOMKeys {
    if($null -eq $__RegBackup){ return }
    $regPath = Get-OOMRegPath
    foreach($k in $__RegBackup.Keys){
        if($null -ne $__RegBackup[$k]){ Set-ItemProperty -Path $regPath -Name $k -Value $__RegBackup[$k] -Force }
        else{ Remove-ItemProperty -Path $regPath -Name $k -ErrorAction SilentlyContinue }
    }
}

# --- (8) Outlook Kill & Preheat ---
function Kill-Outlook {
    Get-Process outlook -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Kill() } catch{} }
    Start-Sleep -Seconds 2
}
function Start-OutlookProfile {
    Start-Process \"outlook.exe\" | Out-Null
    Start-Sleep -Seconds $OutlookWarmupSeconds
    Kill-Outlook
}

###############################################################################################
# OST/Datei-Scan, Eventlog, Migration, Reporting
###############################################################################################

# --- (9) OST-Scan (Verdachtsgrößen, Age, ForceRecheck) ---
function Get-OSTScan {
    param(
        [string]$Folder,
        [datetime]$Cutoff,
        [long]$Lower,
        [long]$Upper,
        [string]$User,
        [string]$ProcessedLog,
        [switch]$ForceRecheck
    )
    $res = @()
    if(-not(Test-Path $Folder)){ return $res }
    $files = Get-ChildItem $Folder -Filter *.ost -ErrorAction SilentlyContinue
    foreach($f in $files){
        $age = (Get-Date) - $f.LastWriteTime
        $suspect = ($f.Length -ge $Lower -and $f.Length -le $Upper -and $f.LastWriteTime -lt $Cutoff)
        $processed = $false
        if(-not $ForceRecheck -and $ProcessedLog -and (Test-Path $ProcessedLog)){
            $csv = Import-Csv $ProcessedLog -Delimiter ';'
            $processed = $csv | Where-Object { $_.Username -eq $User -and $_.OSTFile -eq $f.Name -and $_.LastWriteTime -eq $f.LastWriteTime.ToString() }
        }
        $res += [pscustomobject]@{
            OSTFile = $f.Name
            FullPath = $f.FullName
            SizeMB = [math]::Round($f.Length/1MB,1)
            LastWriteTime = $f.LastWriteTime
            IsSuspect = $suspect
            AgeDays = [int]$age.TotalDays
            AlreadyProcessed = $processed
        }
    }
    return $res
}

# --- (10) Eventlog-Analyse für OST/Outlook-Fehler (nur relevante, nur eigene Userpfade) ---
function Get-HostOutlookOSTErrorsForUser {
    param(
        [string]$Username,
        [string]$OfficeContainerPath,
        [string]$UserContainerPath,
        [int]$Days = 7
    )
    $start = (Get-Date).AddDays(-1*$Days)
    $sources = @('Outlook','Microsoft Office Alerts')
    $errPatterns = @(
        '\\.ost\\b',
        'cannot be found|not found|ERROR_FILE_NOT_FOUND',
        'is not an outlook data file|not an \\.ost file',
        'in use by another process',
        '0x8004010f', '0x8004010d', '0x80040119', '0x80040600'
    )
    $userHints = @()
    if ($OfficeContainerPath) { $userHints += [regex]::Escape($OfficeContainerPath) }
    if ($UserContainerPath)   { $userHints += [regex]::Escape($UserContainerPath) }
    if ($Username)            { $userHints += [regex]::Escape($Username) }
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName    = 'Application'
            StartTime  = $start
            ProviderName = $sources
            Level = 2,3   # Error=2, Warning=3
        } -ErrorAction SilentlyContinue
        $events | Where-Object {
            $msg = $_.Message
            if ([string]::IsNullOrEmpty($msg)) { return $false }
            $matchesErr = $false
            foreach ($pat in $errPatterns) { if ($msg -match $pat) { $matchesErr = $true; break } }
            if (-not $matchesErr) { return $false }
            $matchesUser = $false
            foreach ($hint in $userHints) { if ($msg -match $hint) { $matchesUser = $true; break } }
            return $matchesUser
        } | Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, Message
    } catch { @() }
}

###############################################################################################
# Hauptablauf: Mapping laden, User-Loop, Reporting, Migration etc.
###############################################################################################

# (1) Lazy-Admin Resolver für Pfade (CentralReport/ProcessedLog automatisch)
$resolved = _Pick-ReportRoots -ReportFolder $ReportFolder -ReportFileShare $ReportFileShare `
                              -CentralReport $CentralReport -ProcessedLog $ProcessedLog
$CentralReport = $resolved.CentralReport
$ProcessedLog  = $resolved.ProcessedLog

# (2) Config laden (falls gewünscht)
if ($UseLastConfig) {
    $loaded = @{}
    if (Test-Path $ConfigPath) { $loaded = Load-ConfigJson -Path $ConfigPath }
    if (-not $loaded.Keys.Count) { $loaded = Load-ConfigRegistry -Root $ConfigRegRoot }
    foreach ($k in $loaded.Keys) {
        if (-not $PSBoundParameters.ContainsKey($k)) {
            Set-Variable -Name $k -Value $loaded[$k] -Scope 1
        }
    }
    Write-Host "• Vorherige Konfiguration via -UseLastConfig geladen."
}
if ($LoadConfigJson) {
    $j = Load-ConfigJson -Path $ConfigPath
    foreach ($k in $j.Keys) { if (-not $PSBoundParameters.ContainsKey($k)) { Set-Variable -Name $k -Value $j[$k] -Scope 1 } }
    Write-Host "• Konfiguration aus JSON geladen: $ConfigPath"
}
if ($LoadConfigRegistry) {
    $r = Load-ConfigRegistry -Root $ConfigRegRoot
    foreach ($k in $r.Keys) { if (-not $PSBoundParameters.ContainsKey($k)) { Set-Variable -Name $k -Value $r[$k] -Scope 1 } }
    Write-Host "• Konfiguration aus Registry geladen: $ConfigRegRoot\\Config"
}

# (3) Mapping laden
$map = Import-Csv $MappingFile | Where-Object { -not $UserList -or ($_.Username -in $UserList) }
if($StopAfterNUsers -gt 0){ $map = $map | Select-Object -First $StopAfterNUsers }

$globalSummary = @()

$usrIdx = 0
foreach ($row in $map) {
    $usrIdx++
    $user = $row.Username
    $officeCont = $row.OfficeContainerPath
    $userCont   = $row.UserContainerPath

    # User-Lock setzen
    $lock = New-UserLock -User $user -CentralReport $CentralReport
    if(-not $lock){ continue }

    # Klassifikation für Reporting
    $officeMeta = Classify-Path -Path $officeCont
    $userMeta   = Classify-Path -Path $userCont

    # Eventlog (vorher)
    $eventsBefore = @()
    if($CollectLogs){ $eventsBefore = Get-HostOutlookOSTErrorsForUser -Username $user `
        -OfficeContainerPath $officeCont -UserContainerPath $userCont -Days $EventDays }

    # OST-Analyse
    $cutoff = if($CorruptAgeHours -gt 0){ (Get-Date).AddHours(-1 * $CorruptAgeHours) } else { (Get-Date).AddDays(-1 * $CorruptAgeDays) }
    $scanOffice = $null; $scanUserPre = $null
    if($AnalyzeOST -or $AnalysisOnly){
        if($SourceSelect -in @('SourceOnly','Both')){ $scanOffice = Get-OSTScan -Folder $officeCont -Cutoff $cutoff -Lower $SuspectLowerBytes -Upper $SuspectUpperBytes -User $user -ProcessedLog $ProcessedLog -ForceRecheck:$ForceRecheck }
        if($SourceSelect -in @('TargetOnly','Both')){ $scanUserPre = Get-OSTScan -Folder $userCont   -Cutoff $cutoff -Lower $SuspectLowerBytes -Upper $SuspectUpperBytes -User $user -ProcessedLog $ProcessedLog -ForceRecheck:$ForceRecheck }
    }

    # AnalyseOnly: nur Report, kein Copy, kein Outlook, keine Regkeys
    if($AnalysisOnly){
        $report = [pscustomobject]@{
            Username    = $user
            Hostname    = $env:COMPUTERNAME
            Timestamp   = (Get-Date)
            Mode        = 'AnalysisOnly'
            Params      = [pscustomobject]@{
                EventScope = $EventScope; EventDays = $EventDays
                SuspectRange=\"$SuspectLowerBytes..$SuspectUpperBytes\"
                Cutoff     = $cutoff
                ForceRecheck = [bool]$ForceRecheck
                SourceSelect = $SourceSelect
            }
            PreOfficeScan = $scanOffice
            PreUserScan   = $scanUserPre
            EventsBefore  = $eventsBefore
            Metadata      = [pscustomobject]@{ OfficeContainer = $officeMeta; UserContainer = $userMeta }
        }
        $json = Join-Path $CentralReport \"$user`_AnalysisOnly_$(Get-Date -Format yyyyMMdd_HHmmss).json\"
        $null = Write-Resilient -Content ($report | ConvertTo-Json -Depth 12) -TargetPath $json -DeferOnPermissionError:$DeferOnPermissionError
        Remove-UserLock -LockPath $lock
        continue
    }

    # Optional: OOM-Regkeys sichern/setzen (nur, wenn gewünscht)
    $regSet = $false
    if(-not $NoRegkeys -and $RegkeyBackupAndRemove -and -not $DryRun){
        Backup-OOMKeys
        Set-OOMKeys
        $regSet = $true
    }

    # Outlook schließen (vor Migration, sicherheitshalber)
    if($StartOutlook -or $Migrate){
        Kill-Outlook
    }

    # Migration: OST kopieren, Backup, ACL, ProcessedLog
    $migrateResult = @()
    if($Migrate -and ($SourceSelect -in @('SourceOnly','Both'))){
        $osts = Get-ChildItem $officeCont -Filter *.ost -ErrorAction SilentlyContinue
        foreach($src in $osts){
            $tgt = Join-Path $userCont $src.Name
            $isSuspect = ($src.Length -ge $SuspectLowerBytes -and $src.Length -le $SuspectUpperBytes)
            $sizeOK = ($src.Length -lt ($MaxFileSizeMB*1MB))
            if(-not $sizeOK){ Write-Warning \"Überspringe sehr große OST: $($src.Name)\"; continue }
            # Backup bestehendes Ziel
            if(Test-Path $tgt){
                $backupRoot = Join-Path $userCont (\"Backup_\" + (Get-Date -Format yyyyMMdd))
                if(-not(Test-Path $backupRoot)){ New-Item -Path $backupRoot -ItemType Directory | Out-Null }
                Move-Item -Path $tgt -Destination (Join-Path $backupRoot $src.Name) -Force
            }
            # Migration (robocopy bevorzugt)
            if(-not $DryRun){
                $copied = $false
                $rcLog = Join-Path $env:TEMP (\"rc_\"+[IO.Path]::GetRandomFileName()+\".log\")
                $cmd = \"robocopy\"; $args = @((Split-Path $src.FullName -Parent), (Split-Path $tgt -Parent), $src.Name, \"/FFT\",\"/ZB\",\"/R:2\",\"/W:2\",\"/LOG+:$rcLog\")
                $p = Start-Process $cmd -ArgumentList $args -Wait -PassThru
                $exit = $p.ExitCode
                if($exit -in 0..7 -and (Test-Path $tgt)){ $copied=$true }
                if(-not $copied){
                    # Fallback: Copy-Item
                    try { Copy-Item -Path $src.FullName -Destination $tgt -Force; $copied=(Test-Path $tgt) } catch{}
                }
                # Optional: Checksum
                if($copied -and $Checksum){
                    $srcHash = (Get-FileHash $src.FullName -Algorithm SHA256).Hash
                    $tgtHash = (Get-FileHash $tgt -Algorithm SHA256).Hash
                    if($srcHash -ne $tgtHash){
                        Write-Warning \"Checksum mismatch! $($src.Name)\"
                        $copied = $false
                    }
                }
                # ACL setzen
                if($copied -and $FixPerms){ Set-FilePermissiveACL -Path $tgt -PermissiveACL:$PermissiveACL }
                # ProcessedLog
                if($copied -and $ProcessedLog){
                    $line = \"{0};{1};{2};{3};{4}\" -f $user,$src.Name,$src.FullName,$src.LastWriteTime,$tgt
                    Add-Content -Path $ProcessedLog -Value $line
                }
                $migrateResult += [pscustomobject]@{
                    OSTFile = $src.Name
                    SizeMB = [math]::Round($src.Length/1MB,1)
                    Source = $src.FullName
                    Target = $tgt
                    Copied = $copied
                    Suspect = $isSuspect
                }
            }
        }
    }

    # Outlook (optional) nach Migration starten zum Preheat
    if($StartOutlook -and -not $DryRun){
        Start-OutlookProfile
    }

    # AutoHealProfile (PRF-Import, wenn nach Copy Fehler oder Suspect)
    if($AutoHealProfile -and $PRFFile -and (Test-Path $PRFFile) -and -not $DryRun){
        Start-Process \"outlook.exe\" \"/importprf `\"$PRFFile`\"\" -Wait
        Kill-Outlook
    }

    # Eventlog (nachher)
    $eventsAfter = @()
    if($CollectLogs){ $eventsAfter = Get-HostOutlookOSTErrorsForUser -Username $user `
        -OfficeContainerPath $officeCont -UserContainerPath $userCont -Days $EventDays }

    # Reporting
    $report = [pscustomobject]@{
        Username    = $user
        Hostname    = $env:COMPUTERNAME
        Timestamp   = (Get-Date)
        Parameters  = [pscustomobject]@{
            Migrate=$Migrate; DryRun=$DryRun; FixPerms=$FixPerms; Checksum=$Checksum
            StartOutlook=$StartOutlook; AutoHealProfile=$AutoHealProfile; SourceSelect=$SourceSelect
        }
        PreOfficeScan = $scanOffice
        PreUserScan   = $scanUserPre
        MigrateResult = $migrateResult
        EventsBefore  = $eventsBefore
        EventsAfter   = $eventsAfter
        Metadata      = [pscustomobject]@{ OfficeContainer = $officeMeta; UserContainer = $userMeta }
        Environment   = [pscustomobject]@{
            MachineName = $env:COMPUTERNAME
            User        = \"$env:USERDOMAIN\\$env:USERNAME\"
            SessionId   = (Get-Process -Id $PID).SessionId
        }
        Infrastructure = [pscustomobject]@{
            ReportBase    = $resolved.BaseChosen
            UsingShare    = $resolved.UsingShare
            CentralReport = $CentralReport
            ProcessedLog  = $ProcessedLog
        }
    }
    $jsonOut = Join-Path $CentralReport \"$user`_OST_Migration_$(Get-Date -Format yyyyMMdd_HHmmss).json\"
    $jr = Write-Resilient -Content ($report | ConvertTo-Json -Depth 12) -TargetPath $jsonOut -DeferOnPermissionError:$DeferOnPermissionError
    if ($jr.Status -ne 'OK') { Write-Warning \"Report Write: $($jr.Status) → $($jr.Path)\" }
    $globalSummary += [pscustomobject]@{
        Username  = $user; Host = $env:COMPUTERNAME; Time = (Get-Date); Status = \"Done\"; JSON = $jsonOut
    }

    # OOM-Regkeys zurückspielen (sofern vorher gesetzt)
    if($regSet){ Restore-OOMKeys }

    # Lock entfernen, Delay
    Remove-UserLock -LockPath $lock
    if($ThrottleDelaySec -gt 0){ Start-Sleep -Seconds $ThrottleDelaySec }
}

# Abschluss: Global-CSV schreiben
if($globalSummary.Count -gt 0){
    $csvOut = Join-Path $CentralReport \"MigrationSummary.csv\"
    $csvText = ($globalSummary | ConvertTo-Csv -NoTypeInformation -Delimiter ';') -join [Environment]::NewLine
    $cr = Write-Resilient -Content $csvText -TargetPath $csvOut -DeferOnPermissionError:$DeferOnPermissionError
    if ($cr.Status -ne 'OK') { Write-Warning \"Summary Write: $($cr.Status) → $($cr.Path)\" }
}

# Config-Persistenz (am Ende)
$toSave = Get-ConfigSafeParams -Bound $PSBoundParameters
if ($SaveConfigJson)     { Save-ConfigJson     -Config $toSave -Path $ConfigPath }
if ($SaveConfigRegistry) { Save-ConfigRegistry -Config $toSave -Root $ConfigRegRoot }

Write-Host \"*** Migration abgeschlossen. Alle Berichte unter: $CentralReport ***\"
