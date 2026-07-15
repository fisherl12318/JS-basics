# Trinkter — Automatisches Ausfüllen von KV-Gewährleistungs-PDFs für Ultraschallgeräte

## Überblick

Trinkter ist ein Desktop-Tool auf Basis von Python + Tkinter zum automatischen Ausfüllen der PDF-Formulare für die Gewährleistungserklärung von Ultraschallgeräten bei den deutschen KVen (Kassenärztliche Vereinigungen). Anhand des gewählten Gerätemodells und der Schallköpfe kreuzt das Programm automatisch die passenden AK-Checkboxen (Anwendungsklassen) an und füllt die Stammdatenfelder aus.

---

## Dateistruktur

```
Gwe/
├── Trinkter_new.py      # Hauptprogramm (Tkinter-GUI)
├── Pdf_filling.py       # Kernlogik zum Lesen/Schreiben der PDFs
├── Excel.py             # Schallkopf → AK-Zuordnung (liest Excel)
├── data.xlsx            # Zuordnungstabelle Schallkopf → AK (wird von Excel.py gelesen)
├── checkcheck.py        # Eigenständiges Tool: prüft, ob eine PDF-Vorlage alle AK-Formularfelder enthält
├── requirements.txt     # Python-Abhängigkeiten
└── README.md            # Dieses Dokument
```

---

## Verwendung

### 1. Programm starten

```bash
python Trinkter_new.py
```

### 2. Bedienschritte in der Oberfläche

| Schritt | Feld | Beschreibung |
|------|------|------|
| ① | PDF-Datei | Über „浏览" (Durchsuchen) das auszufüllende KV-Gewährleistungs-PDF auswählen |
| ② | Firmenname | Adresse (Name der Firma/Einrichtung) eintragen |
| ③ | Gerätemodell | Gerät aus dem Dropdown wählen — Gerätename und Schallkopfliste werden automatisch aktualisiert |
| ④ | Gerätename | Wird nach Auswahl des Modells automatisch eingetragen (manuell änderbar) |
| ⑤ | Seriennummer (SN) | Seriennummer eingeben, z. B. `360101-M26305870006` |
| ⑥ | Baujahr | Wird **automatisch** aus der SN abgeleitet (nur lesend, keine Eingabe nötig) |
| ⑦ | Datum | Wird **automatisch** mit dem heutigen Datum gefüllt (Format TT.MM.JJJJ, änderbar) |
| ⑧ | Schallkopfauswahl | Schallköpfe links markieren und mit „→" nach rechts verschieben; rechts stehen die tatsächlich verwendeten |
| ⑨ | Ausgabedatei | Pfad/Dateiname des Ausgabe-PDFs angeben |
| ⑩ | Start | Auf „开始" (Start) klicken, um das Ausfüllen zu starten |

---

## Auswertung der Seriennummer

Beispiel für das SN-Format: `360101-M26305870006`

Das Programm nimmt die beiden Ziffern nach dem `M` als letzte zwei Stellen des Baujahrs:

- `M26…` → Baujahr `2026`
- `M23…` → Baujahr `2023`

Entspricht die SN nicht diesem Muster (kein `M` + zwei Ziffern), bleibt das Baujahr-Feld leer.

---

## Neue Geräte und Schallkopf-Zuordnungen hinzufügen

In der `__init__`-Methode von `Trinkter_new.py` das Dictionary `self.device_map` suchen und ein neues Gerät nach folgendem Muster ergänzen:

```python
self.device_map = {
    # vorhandene Geräte ...

    "Neuer Modellname": {
        "name": "Vollständiger Gerätename, wie er im PDF stehen soll",
        "probes": ["Schallkopf1", "Schallkopf2", "Schallkopf3"]
    },
}
```

**Beispiel:**

```python
"Acclarix AX8 Series": {
    "name": "Acclarix AX8 Diagnostic Ultrasound System",
    "probes": ["C5-2Q", "L17-7HQ", "E8-4Q", "P5-1Q"]
},
```

Das neue Gerät erscheint danach automatisch im Dropdown; bei Auswahl wird die Schallkopfliste automatisch aktualisiert.

---

## Frequenz-Zuordnung der Schallköpfe

In `Pdf_filling.py` gibt es ein Dictionary `frequency_map`, das Schallkopfmodelle auf Frequenzangaben abbildet:

```python
frequency_map = {
    "L743-2": "4-3 MHz",
    "C361-2": "6-3 MHz",
    # neuen Schallkopf ergänzen →
    "Neues Modell": "X-Y MHz",
}
```

Folgt das Schallkopfmodell dem Muster `Zahl-Zahl` (z. B. `L17-7HQ`), leitet das Programm die Frequenz automatisch ab (`17-7 MHz`) — ein manueller Eintrag ist nicht nötig. Nur Schallköpfe mit abweichendem Format müssen manuell in `frequency_map` eingetragen werden.

---

## PDF-Felder

Die Funktion `fill_pdf_fields` in `Pdf_filling.py` füllt folgende Felder aus:

| PDF-Feldname | Inhalt | Anmerkung |
|-----------|------|------|
| `Adresse` / `Adresse_1` | Firmenname | auf beiden Seiten |
| `Bezeichnung` / `Bezeichnung_1` | Gerätename | auf beiden Seiten |
| `GeräteNummer` / `GeräteNummer_1` | Seriennummer (SN) | auf beiden Seiten |
| `Baujahr` / `Baujahr_1` | Baujahr | auf beiden Seiten |
| `Auslieferungsdatum` / `Auslieferungsdatum_1` | Lieferdatum | auf beiden Seiten |
| `datum` | aktuelles Datum | bei der Unterschriftszeile |
| `Schallkopf 1`–`Schallkopf 5` | Schallkopfmodelle | maximal 5 |
| `Frequenz 1`–`Frequenz 5` | Schallkopf-Frequenzen | passend zu den Schallköpfen |
| `ak_X_Y` | AK-Checkboxen | ergebnisabhängig aus der Excel-Abfrage |

Weitere Felder lassen sich in `fill_pdf_fields` nach diesem Muster ergänzen:

```python
set_text_field(pdf, "Neuer Feldname", wert)
```

---

## Logik der AK-Checkboxen

1. Nach der Schallkopfauswahl liefert `Excel.find_items()` anhand der Schallkopfliste die anzukreuzenden AK-Nummern aus der Excel-Tabelle (z. B. `["AK 2.1", "AK 3.4", ...]`)
2. Die AK-Nummern werden in das Format `ak_X_Y` normalisiert (Funktion `normalize_name`)
3. Alle Checkbox-Felder des PDFs werden durchlaufen; bei Namensübereinstimmung wird angekreuzt

**Sonderregeln:**
- Beim Gerätenamen `DUS 60` wird automatisch `AK 20.12` ergänzt (Farbkodierte: nein)
- Bei allen anderen Geräten wird automatisch `AK 20.11` ergänzt (Farbkodierte: ja)

---

## checkcheck.py — Prüftool für PDF-Vorlagen (eigenständig)

`checkcheck.py` gehört **nicht zum Hauptprogramm** (`Trinkter_new.py` ruft es nicht auf). Es ist ein eigenständiges Kommandozeilen-Tool, mit dem man vor der Verwendung einer neuen KV-PDF-Vorlage prüft, ob das PDF alle benötigten AK-Checkbox-Formularfelder enthält (Feldnamen wie `ak_1_1`, `ak_2_3`).

**Funktionsweise:**

1. Liest mit PyPDF2 alle Formularfeldnamen aus dem PDF
2. Erzeugt die vollständige AK-Liste (AK 1.1 bis AK 23.1, rund 50 Einträge)
3. Normalisiert jede AK-Nummer in das Format `ak_X_Y` und prüft, ob ein gleichnamiges Feld im PDF existiert
4. Gibt pro Eintrag ✓/✗ sowie den Vollständigkeitsgrad in Prozent aus

**Verwendung:**

In `checkcheck.py` am Dateianfang `PDF_PATH` auf die zu prüfende PDF-Datei setzen:

```python
PDF_PATH = "KV Hessen.pdf"
```

Dann ausführen:

```bash
python checkcheck.py
```

Beispielausgabe:

```
AK 1.1   ✓
AK 2.1   ✓
AK 20.11 ✗
...
完成度: 48/50 (96.0%)
```

Einträge mit ✗ bedeuten: Der PDF-Vorlage fehlt das entsprechende Checkbox-Feld. Das Hauptprogramm kann diese AK dann nicht ankreuzen — das Feld muss zuerst in einem PDF-Editor ergänzt werden (Feldname im Format `ak_X_Y`).

---

## Installation der Abhängigkeiten

```bash
pip install -r requirements.txt
```

Enthält: `pdfrw` (PDF-Verarbeitung im Hauptprogramm), `PyPDF2` (checkcheck.py), `pandas` + `openpyxl` (Excel.py liest data.xlsx).

(Der Rest ist Standardbibliothek: `tkinter`, `re`, `threading`, `datetime`)

---

## Häufige Fragen

**F: Warum werden manche Checkboxen nicht angekreuzt?**

A: Prüfen, ob der tatsächliche Feldname im PDF dem Format `ak_X_Y` entspricht. In der Konsolenausgabe von `Pdf_filling.py` sind die Zielfelder und die Treffer zu sehen. Zur systematischen Prüfung einer Vorlage eignet sich `checkcheck.py`.

**F: Wie wechsle ich die PDF-Vorlage?**

A: Einfach in der Oberfläche eine andere PDF-Datei auswählen. Die Feldnamen der jeweiligen KV-Vorlage müssen zu den Feldnamen in `fill_pdf_fields` passen; bei Abweichungen die Feldnamen in `Pdf_filling.py` entsprechend anpassen.

**F: Lässt sich das Datumsformat ändern?**

A: Ja. In `Trinkter_new.py` folgende Zeile suchen:

```python
self.datum_var = tk.StringVar(value=date.today().strftime("%d.%m.%Y"))
```

und `"%d.%m.%Y"` durch das gewünschte Format ersetzen, z. B. ergibt `"%Y-%m-%d"` das Format `2026-06-29`.
