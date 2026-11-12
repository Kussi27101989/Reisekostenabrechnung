# Reisekosten RB003 – Web-App für iPhone & iPad

Belege fotografieren → automatisch erkennen → am Monatsende die **Original-RB003-Excel** und ein **PDF mit allen Einzelnachweisen** exportieren.

## 1. Veröffentlichen (einmalig, ca. 10 Minuten)

Die App ist eine reine Web-App ohne Server. Sie braucht nur ein HTTPS-Hosting, z. B. **GitHub Pages**:

1. Auf github.com ein neues Repository anlegen, z. B. `reisekosten`.
2. Den **Inhalt** dieses Ordners hochladen (`index.html` muss direkt im Repository-Hauptordner liegen).
3. *Settings → Pages → Branch: main / (root) → Save*.
4. Nach etwa einer Minute ist die App erreichbar unter `https://<benutzername>.github.io/reisekosten/`.

Alternativ geht jeder andere Webspace mit HTTPS (z. B. firmeneigener Webserver, Netlify, Azure Static Web Apps).

> Hinweis: Auf GitHub Pages ist der Ordner öffentlich lesbar. Das betrifft auch die integrierte leere RB003-Vorlage (Firmenformular, ohne Name und Kostenstelle). Belegdaten liegen dagegen **nie** auf dem Server. Wenn das Formular nicht öffentlich sein soll, die App auf einem internen Webserver hosten.

## 2. Auf iPhone / iPad installieren

1. Link in **Safari** öffnen.
2. **Teilen** (Quadrat mit Pfeil) → **„Zum Home-Bildschirm"**.
3. Die App über das neue Symbol starten → **Einstellungen**: Name und Kostenstelle eintragen.

Wichtig: Immer über das Home-Bildschirm-Symbol arbeiten. Nur dort speichert iOS die Daten dauerhaft.

## 3. Texterkennung

| Modus | Funktion |
|---|---|
| **Automatisch** (Standard) | KI-Erkennung, wenn Internet und API-Key vorhanden sind, sonst Offline-Texterkennung |
| Nur KI | Liest Betrag, Datum, MwSt, Händler und Kategorie sehr zuverlässig aus, erkennt bei Hotelrechnungen auch das Frühstück |
| Offline | Tesseract direkt auf dem Gerät, ohne Internet und ohne Kosten. Kategorie und Betrag bitte kontrollieren. |
| **Aus (manuell)** | Keine Texterkennung. Nach dem Foto öffnet sich sofort das Eingabeformular, du trägst Datum (vorbelegt mit heute), Betrag und Kategorie selbst ein. Kein API-Key nötig, nichts verlässt das Gerät. Bei Bedarf liest „Text erkennen" einen einzelnen Beleg trotzdem aus. |

| Nur Google Vision | Sehr genaue Texterkennung von Google (auch bei schlechtem Licht/Handschrift), daraus ermittelt die App Betrag, Datum, MwSt und Kategorie. 1.000 Bilder/Monat kostenlos. |

**Automatisch** nutzt der Reihe nach: Claude (wenn Key) → Google Vision (wenn Key) → Offline-Erkennung. Fällt ein Dienst aus, springt die App auf den nächsten.

### Google Vision einrichten (einmalig, ca. 10 Minuten)

1. console.cloud.google.com öffnen, mit Google-Konto anmelden, **neues Projekt** anlegen (z. B. „Reisekosten").
2. **Abrechnung** aktivieren (Kreditkarte nötig – die ersten 1.000 Bilder pro Monat sind kostenlos, danach 1,50 $ je 1.000).
3. Menü → *APIs & Dienste* → *Bibliothek* → **„Cloud Vision API"** → *Aktivieren*.
4. *APIs & Dienste* → *Anmeldedaten* → *Anmeldedaten erstellen* → **API-Schlüssel**.
5. Schlüssel **einschränken** (wichtig!):
   - *Anwendungseinschränkungen*: **Websites** → `https://kussi27101989.github.io/*` eintragen.
   - *API-Einschränkungen*: **Schlüssel einschränken** → nur **Cloud Vision API**.
6. Empfohlen: *Abrechnung → Budgets & Benachrichtigungen* → Budget z. B. 5 € mit E-Mail-Warnung.
7. Den Schlüssel (beginnt mit `AIza…`) in der App unter **Einstellungen → Google Cloud Vision API-Key** eintragen → **„Google Vision testen"**.

Hinweis: Die 1.000 Gratis-Bilder gelten pro Abrechnungskonto. Nutzen alle Außendienstler denselben Schlüssel, teilen sie sich das Kontingent.

**API-Key für die KI-Erkennung:** auf console.anthropic.com einen Key erstellen und in den Einstellungen eintragen. Kosten: ca. 1–2 Cent pro Beleg. Der Key bleibt auf dem Gerät. Das Belegbild geht nur zur Auswertung an die Anthropic-API.

Tipp: In den Einstellungen einmal **„Offline-Erkennung laden"** tippen (ca. 15 MB, am besten im WLAN). Danach funktioniert die Texterkennung auch im Funkloch.

## 4. Ablauf im Monat

1. **Reise anlegen**: Abreise und Rückkehr eintragen. Die App erzeugt An- und Abreisetag sowie die vollen Zwischentage und berechnet die Verpflegungspauschalen wie die RB003 (Ländertabelle, Kürzung F 20 % / M 40 % / A 40 %).
2. **Belege** mit **＋** fotografieren oder als Foto bzw. PDF importieren (z. B. Hotelrechnung aus der Mail). Die Erkennung ordnet jeden Beleg Datum und Spalte zu.
3. Jeden Beleg kurz **prüfen** und auf **„Geprüft – speichern"** tippen.
4. **Fahrzeug**: km-Stände des Dienstwagens, Mietwagen-km und Privat-PKW-km von Hand eintragen.
5. **Monat exportieren** → **Teilen** (Mail, Teams, OneDrive …):
   - `RB003_Reisekostenabrechnung_JJJJ-MM_Name.xlsx`: die Original-Vorlage, befüllt, mit allen Formeln, Dropdowns und dem Layout
   - `…_Belege.pdf`: Abrechnungsübersicht, Belegliste und je Beleg eine Seite mit der Belegnummer (B01, B02 …). Dieselbe Nummer steht in der Excel in der Spalte „Bezeich.".

## 5. Zuordnung zur RB003

| App-Kategorie | RB003 |
|---|---|
| Verpflegungspauschale (aus den Reisetagen) | Blatt 1, Spalte 1 |
| Verpflegung Einzelnachweis / Übernachtung Pauschale / Hotel | Spalten 2 / 3 / 4 |
| Bewirtung (mit Anlass und Teilnehmern) | Spalte 5 |
| Geschenke 19 % / 7 % | Spalten 6 / 7 |
| Nebenkosten 19 % / 7 % / ohne MwSt | Spalten 8 / 9 / 10 |
| Dienstwagen: Kraftstoff, Öl, Garage, Pflege, Reparatur, km-Stände | Seite 2, Spalte 11 |
| Mietwagen: Kraftstoff, Miete, km | Seite 2, Spalte 12 |
| Privatwagen-km × km-Pauschale | Seite 2, Spalte 13 |
| ÖPNV bis 50 km / Bahn, Flug über 50 km | Seite 2, Spalten 14 / 15 |

Mehrere Belege am selben Tag in derselben Spalte werden addiert, z. B. „B07+B08".

## 6. Datensicherung & neue RB003-Version

- **Einstellungen → Backup erstellen** speichert alle Daten und Belegbilder als ZIP (z. B. in OneDrive). Mit „Backup einspielen" lassen sie sich wiederherstellen, auch auf einem neuen Gerät.
- Gibt die Firma eine neue RB003 heraus: **Einstellungen → Neue RB003-Version laden** und die leere Excel auswählen. Länder und Pauschalen werden aus der Datei übernommen.

## 7. Technik

- Reines HTML/JavaScript, ohne Server und ohne App Store. Die Daten liegen in IndexedDB auf dem Gerät.
- Excel: Das XML der Original-Vorlage wird direkt befüllt (JSZip). Formatierung, Formeln, Länder-Dropdowns und Blattschutz bleiben erhalten. Die Formel-Ergebnisse sind vorberechnet, Excel rechnet beim Öffnen zusätzlich neu.
- PDF: jsPDF, PDF-Import: pdf.js, Offline-OCR: Tesseract.js (Deutsch). Alle Bibliotheken liegen in `vendor/`, die App lädt also nichts von Fremdservern nach.
- Service Worker (`sw.js`) für den Offline-Betrieb. Bei einem Update `VERSION` in `sw.js` hochzählen.
