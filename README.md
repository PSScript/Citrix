# FSLogix / Outlook OST Migration & Analyse

## Überblick

Dieses PowerShell-Skript unterstützt die Migration und Analyse von Outlook OST-Dateien in **Citrix / FSLogix** Umgebungen.  
Es ermöglicht die Überführung von OST-Dateien aus dem **Office Container** in den **User Container**, erkennt beschädigte Dateien (z. B. Standard-64 MB OST), erstellt zentrale Berichte und bietet optionale Automatisierungen wie Outlook-Preheat oder Profil-Reparatur.

Das Skript ist modular, mandantenfähig und für den Einsatz über **GPO, zentrale Shares oder manuelle Aufrufe** optimiert.

---

## Voraussetzungen

- Windows Server mit Citrix / FSLogix Profilcontainern  
- PowerShell 5.1 oder höher  
- Outlook installiert im Benutzerkontext  
- Schreibrechte auf das definierte Report-Verzeichnis oder alternatives Fallback (lokal/Queue)

---

## Mapping-Datei

CSV-Format mit mindestens folgenden Spalten:

```csv
Username,OfficeContainerPath,UserContainerPath
jan.huebener,\\fs01\OfficeVHD\jan\Outlook,\\fs01\UserVHD\jan\Outlook
anna.schmidt,\\fs01\OfficeVHD\anna\Outlook,\\fs01\UserVHD\anna\Outlook
```

---

## Ausgaben

- **Pro Benutzer JSON-Report**  
  `<ReportRoot>\<username>_OST_Migration_<timestamp>.json`  
- **Zentrale CSV-Zusammenfassung**  
  `<ReportRoot>\MigrationSummary.csv`  
- **Processed-Log** (verarbeitete Benutzer)  
  `<ReportRoot>\ProcessedOSTs.csv`  
- **Optionale Backups**  
  `Backup_YYYYMMDD\<filename>.ost`  
- **Deferred Reports** (bei Berechtigungsfehlern)  
  `%ProgramData%\OSTMigration\Queue\`

---

## Parameterübersicht

| Parameter | Typ | Standard | Beschreibung |
|-----------|-----|----------|--------------|
| `-MappingFile` | String | – | CSV mit Benutzername + Containerpfaden |
| `-CentralReport` | String | – | Basisverzeichnis für Reports (alt, von Lazy-Admin ersetzt) |
| `-ProcessedLog` | String | – | Zentrales Log verarbeiteter Benutzer |
| `-ReportFolder` | String | %ProgramData%\OSTMigration\Reports | Lokales Report-Verzeichnis (Lazy-Admin) |
| `-ReportFileShare` | String | – | UNC-Pfad für zentrale Reports (Lazy-Admin, bevorzugt, Fallback auf lokal) |
| `-UserList` | String[] | – | Einschränkung auf bestimmte Benutzer |
| `-ReportOnly` | Switch | Off | Nur Bericht, keine Migration |
| `-Migrate` | Switch | Off | Migration durchführen |
| `-FixPerms` | Switch | Off | ACLs korrigieren (Benutzer → Vollzugriff) |
| `-PermissiveACL` | Switch | Off | ACL-Fehler werden nur protokolliert, nicht abgebrochen |
| `-StartOutlook` | Switch | Off | Outlook starten, um OST vorzuwärmen |
| `-OutlookWarmupSeconds` | Int | 20 | Dauer des Preheat |
| `-CollectLogs` | Switch | Off | Outlook-Ereignisprotokolle sammeln |
| `-EventScope` | Enum | CurrentSession | Logbereich: `CurrentSession` oder `HostAll` |
| `-EventDays` | Int | 7 | Zeitraum der Log-Abfrage (Tage) |
| `-AnalysisOnly` | Switch | Off | Analyse ohne Migration |
| `-AnalyzeOST` | Switch | Off | OST-Analyse (Größe, Alter) |
| `-AnalyzeEventLogs` | Switch | Off | Nur relevante OST/Outlook-Fehler sammeln |
| `-SuspectLowerBytes` | Long | 67108864 | Untergrenze für 64MB-OST |
| `-SuspectUpperBytes` | Long | 67236864 | Obergrenze für 64MB-OST |
| `-CorruptAgeDays` | Int | 2 | Mindestalter verdächtiger OST in Tagen |
| `-Checksum` | Switch | Off | SHA-256 Verifizierung nach Copy |
| `-DryRun` | Switch | Off | Simulation, keine Änderungen |
| `-RegkeyBackupAndRemove` | Switch | Off | 3 OOM-Registrykeys sichern & entfernen |
| `-NoRegkeys` | Switch | Off | Keine Registrykeys ändern |
| `-DeferOnPermissionError` | Switch | Off | Bei Berechtigungsfehlern lokale Queue statt Abbruch |
| `-MaxFileSizeMB` | Int | 16384 | Max. OST-Dateigröße |
| `-FreeSpaceSafetyMB` | Int | 512 | Zusätzlicher Speicherplatz erforderlich |
| `-MaxBackups` | Int | 2 | Anzahl zu behaltender Backup-Sets |
| `-ThrottleDelaySec` | Int | 0 | Pause zwischen Benutzern |
| `-StopAfterNUsers` | Int | 0 | Limitierte Anzahl pro Lauf |
| `-AutoHealProfile` | Switch | Off | Outlook Profilreparatur mit PRF |
| `-PRFFile` | String | – | Pfad zur PRF-Datei |
| `-ForceRecheck` | Switch | Off | Ignoriert Processed-Log, prüft erneut |
| `-SourceSelect` | Enum | Both | Quelle/Ziel auswählen: `SourceOnly`, `TargetOnly`, `Both` |
| `-SaveConfigJson` | Switch | Off | Letzte Konfiguration in JSON speichern |
| `-LoadConfigJson` | Switch | Off | Konfiguration aus JSON laden |
| `-ConfigPath` | String | %ProgramData%\OSTMigration\last_config.json | Pfad für JSON |
| `-SaveConfigRegistry` | Switch | Off | Letzte Konfiguration in Registry speichern |
| `-LoadConfigRegistry` | Switch | Off | Konfiguration aus Registry laden |
| `-ConfigRegRoot` | String | HKCU:\Software\OSTMigration | Registry-Wurzel |
| `-UseLastConfig` | Switch | Off | Vorherige Konfiguration automatisch laden |

---

## Beispielaufrufe

**1. Analyse (ohne Änderungen):**
```powershell
.\Migrate-FSLogixOST.ps1 `
  -MappingFile \\srv\maps\Mapping.csv `
  -ReportFileShare \\srv\Reports `
  -AnalysisOnly -AnalyzeOST -AnalyzeEventLogs -EventScope HostAll -EventDays 14
```

**2. Migration mit Prüfungen (Pilot):**
```powershell
.\Migrate-FSLogixOST.ps1 `
  -MappingFile \\srv\maps\Mapping.csv `
  -ReportFileShare \\srv\Reports `
  -Migrate -Checksum -StartOutlook -CollectLogs `
  -RegkeyBackupAndRemove -FixPerms
```

**3. Lazy Admin, lokal speichern:**
```powershell
.\Migrate-FSLogixOST.ps1 `
  -MappingFile \\srv\maps\Mapping.csv `
  -ReportFolder "C:\Skripte\Reports" `
  -Migrate -DryRun
```

**4. Config speichern und später erneut nutzen:**
```powershell
# Erstlauf mit vollem Satz:
.\Migrate-FSLogixOST.ps1 -MappingFile \\fs\maps.csv -ReportFileShare \\fs\rep `
  -Migrate -Checksum -CollectLogs -SaveConfigJson -SaveConfigRegistry

# Später nur:
.\Migrate-FSLogixOST.ps1 -UseLastConfig -UserList anna.schmidt
```

---

## Betriebshinweise

- **Immer zuerst mit `-DryRun` oder `-AnalysisOnly` testen**.  
- Bei GPO-Rollout: Skript signieren, ExecutionPolicy prüfen.  
- Deferred-Reports (`%ProgramData%\OSTMigration\Queue`) regelmäßig aufräumen oder nachträglich einsammeln.  
- Processed-Log auf zuverlässigem Share ablegen und sichern.  
