h 

# InvokeZEP

PowerShell-Modul für den Zugriff auf ZEP-Angebote, Angebotspositionen und Mitarbeiter über die ZEP REST API v1.

## Voraussetzungen

- PowerShell 5.1 oder höher
- Modul `BetterCredentials` ≥ 4.5
- ZEP API-Bearer-Token (wird beim ersten Aufruf interaktiv abgefragt und im Windows Credential Manager gespeichert)

## Installation

```powershell
Install-Module -Name InvokeZEP -Repository PSGallery
```

## Schnellstart

```powershell
Import-Module InvokeZEP

# Alle Angebote abrufen (mit lokalem Cache)
$offers = Get-ZEPOffers

# Einzelnes Angebot abrufen
$offer = Get-ZEPOffer -Id 12345

# Positionen eines Angebots abrufen
$items = Get-ZEPOfferItems -OfferId 12345
```

## Konfiguration

```powershell
# Aktuelle Konfiguration anzeigen
Get-ZEPConfiguration

# Konfiguration anpassen
Set-ZEPConfiguration -BaseUri 'https://www.zep-online.de/zepxxxx/next/api/v1' -DefaultPageSize 50
```

## Exportierte Funktionen

| Funktion                            | Beschreibung                                                                                                                      |
| ----------------------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| `Get-ZEPConfiguration`            | Gibt die aktuelle Modulkonfiguration zurück                                                                                      |
| `Set-ZEPConfiguration`            | Ändert Konfigurationsparameter zur Laufzeit                                                                                      |
| `Invoke-ZEPRest`                  | Führt einen beliebigen ZEP REST API-Aufruf aus                                                                                   |
| `Get-ZEPOffer`                    | Ruft ein einzelnes Angebot anhand der ID ab                                                                                       |
| `Get-ZEPOffers`                   | Ruft alle Angebote ab (optional mit Caching und Validierung)                                                                      |
| `Get-ZEPOfferItems`               | Ruft Positionen eines Angebots ab (seitenweise, direkt via API)                                                                   |
| `Test-ZEPOffersCacheNeedsRefresh` | Prüft, ob der lokale Angebots-Cache aktualisiert werden muss                                                                     |
| `Search-ZEPOffers`                | Durchsucht Angebote per Property und Teilstring (`-SkipPropertyValidation` für maximale Geschwindigkeit bei bekannten Feldern) |
| `Get-ZEPRecentOffers`             | Liest nur die neuesten Angebote (rückwärts ab der letzten Seite, `-AfterId` oder `-Last`) – kein Komplett-Download |
| `Get-ZEPOffersById`               | Liest mehrere Angebote per ID in einem Aufruf (`id[]`-Filter, bis 100 IDs je Aufruf) |
| `ConvertTo-ZEPTypedOffer`         | Wandelt ein Angebot in verlässliche Typen um (int-IDs, String-Kundennummern, UTC-Zeiten, Datumswerte) |
| `ConvertFrom-ZEPDateTime`         | Korrigiert ZEP-Zeitstempel (Berliner Ortszeit mit falschem `Z`) nach echtem UTC |
| `Get-ZEPEmployee`                 | Liest einen Mitarbeiter per Benutzername (z. B. für die E-Mail-Adresse), mit Sitzungs-Cache |

## Get-Help Referenz

Schnelle Hilfe mit allen Beispielen:

```powershell
Get-Help Get-ZEPConfiguration -Examples
Get-Help Set-ZEPConfiguration -Examples
Get-Help Invoke-ZEPRest -Examples
Get-Help Get-ZEPOffer -Examples
Get-Help Get-ZEPOffers -Examples
Get-Help Get-ZEPOfferItems -Examples
Get-Help Test-ZEPOffersCacheNeedsRefresh -Examples
Get-Help Search-ZEPOffers -Examples
```

Vollständige Hilfe inklusive Parameterdetails:

```powershell
Get-Help Get-ZEPConfiguration -Full
Get-Help Set-ZEPConfiguration -Full
Get-Help Invoke-ZEPRest -Full
Get-Help Get-ZEPOffer -Full
Get-Help Get-ZEPOffers -Full
Get-Help Get-ZEPOfferItems -Full
Get-Help Test-ZEPOffersCacheNeedsRefresh -Full
Get-Help Search-ZEPOffers -Full
```

## Authentifizierung

Das Modul unterstützt zwei Umgebungen:

**Lokale Sitzung:** Das Bearer-Token wird aus dem Windows Credential Manager unter dem Target `ZEP_hhpberlin_BearerToken` gelesen. Beim ersten Aufruf wird es interaktiv abgefragt und gespeichert.

**Azure Automation Runbook:** Das Token wird über `Get-AutomationPSCredential` gelesen. Der Credential-Name entspricht dem konfigurierten `ServiceUserName`.

## Neue Angebote ohne Komplett-Download

Die ZEP-API ignoriert für `/offers` alle Sortier- und Filterparameter (`orderBy`, `order`, `status`, `modified_after`, …);
die Liste ist immer aufsteigend nach `id` sortiert. `Get-ZEPRecentOffers` liest daher von der letzten Seite rückwärts:

```powershell
# Alle Angebote mit höherer ID als das Wasserzeichen
Get-ZEPRecentOffers -AfterId 18377

# Die neuesten 20 Angebote
Get-ZEPRecentOffers -Last 20

# Bekannte Angebote gezielt nachlesen
Get-ZEPOffersById -Id 18368, 18380
```

## Datum und Zeitzone

ZEP liefert **Berliner Ortszeit, markiert sie aber mit `Z` als UTC** (z. B. `2026-10-07T01:24:35Z` = 23:24:35 UTC).
PowerShell 7 macht daraus ein falsches UTC-Datum, Windows PowerShell 5.1 einen String.
`Get-ZEPRecentOffers` und `Get-ZEPOffersById` liefern deshalb typisierte Angebote mit korrigierten Werten
(`status_since` als UTC-`[datetime]`, `valid_until` u. a. als Datum). Für andere Funktionen:

```powershell
Get-ZEPOffer -Id 18380 -UseCache:$false | ConvertTo-ZEPTypedOffer
ConvertFrom-ZEPDateTime -Value '2026-10-07T01:24:35.000000Z'   # -> 06.10.2026 23:24:35 UTC
```

Die Zeitzone ist einstellbar: `Set-ZEPConfiguration -ServerTimeZone 'Europe/Berlin'` (IANA- und Windows-IDs werden akzeptiert).

## Caching

`Get-ZEPOffers` nutzt einen lokalen JSON-Cache unter `%LOCALAPPDATA%\InvokeZEP`.
Der Cache wird erneuert, wenn er älter als die konfigurierte TTL ist (standardmäßig 60 Minuten) oder wenn sich die Gesamtanzahl der Angebote in ZEP geändert hat.

Für schnelle wiederholte Suchläufe kann `Search-ZEPOffers` mit `-ValidateCache:$false` und optional `-SkipPropertyValidation` auf einem frischen Cache genutzt werden.

## Lizenz

Copyright (c) Klaus Kupferschmid. All rights reserved.
