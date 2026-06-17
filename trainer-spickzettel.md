# Trainer-Spickzettel — Tag 2

Kompakter Begleiter zu [README.md](README.md) und [handout.md](handout.md).
Für den Blick zwischendurch: Was sage ich, wo hängen die Leute, was ist der
Notausgang.

---

## Die EINE Botschaft des Tages

> „Ein Web Part ist eine Klasse mit einer `render()`-Methode. Property Pane,
> Lesen, Schreiben, React, Redux sind Schichten darauf. Heute bauen wir das
> Skelett selbst und ergänzen eine Fähigkeit nach der anderen — am Ende ist
> klar, wie größere SPFx-Apps das strukturieren."

Diesen Satz am Anfang sagen und am Ende wieder aufgreifen.

---

## Einordnung am Morgen (M0, 15 min)

- **Komplexität einordnen**: „Reale Apps bringen viele Schichten gleichzeitig mit.
  Heute drehen wir das auseinander und bauen Stück für Stück selbst."
- **Sachlich bleiben**: kein Hype, keine Jubelsprache. Die Teilnehmer sind erfahren.
- **Tagesziel benennen**: „Bis zum Nachmittag steht ein Web Part, das eine Liste
  liest und Einträge anlegen, ändern und löschen kann — einmal mit `SPHttpClient`,
  einmal mit `@pnp/sp`."

---

## Meilensteine — kurz benennen, nicht zelebrieren

| Nach Block | Stand |
|---|---|
| Block 2 | Eigenes Web Part rendert in der Workbench. |
| Block 3 | Property-Wert kommt in `render()` an. |
| Block 4 | Echte Listendaten als Tabelle. |
| Block 5 | Liste per Dropdown wählbar statt getippt. |
| Block 6 | Anlegen, Bearbeiten, Löschen funktionieren. |
| Block 7 | Beide Datenzugriffe (`SPHttpClient` / `@pnp/sp`) verglichen. |

---

## Timing-Notizen pro Block

| Block | Dauer | Wo es hakt → erste Hilfe |
|---|---|---|
| 1 Scaffold | 45 | `yo` lädt lange → Anatomie parallel am Whiteboard erklären, nicht warten. |
| 2 Serve | 30 | `gulp trust-dev-cert` vergessen → weiße/unsichere Seite. Schritt laut ansagen. |
| 3 Property-Textfeld | 30 | `PropertyPaneTextField` aus `@microsoft/sp-webpart-base` (nicht `sp-property-pane`). |
| 4 Lesen | 60 | `localhost`-Workbench hat keine Liste → SharePoint-gehostete Workbench. **Häufigster Stolperstein.** |
| 5 Listen-Picker | 30 | `npm install` setzt `PropertyControlStrings` in `config.json` automatisch. Picker fehlt trotzdem? Eintrag prüfen, `gulp serve` neu starten. |
| 6 Schreiben | 75 | 403 beim Schreiben = fehlende **Bearbeiten**-Rechte. Metadaten-Fehler = `odata=nometadata` setzen. |
| 7 `@pnp/sp` | 45 | Falsche pnp-Version → Compile-Fehler. **Nur `@pnp/sp@1.3.11`.** |
| 8 Rückblick | 30 | Reiner Vortrag/Diskussion, kein Code. |

> **Wenn der Vormittag rutscht**: Block 7 (`@pnp/sp`) auf „nur zeigen und
> vergleichen" kürzen (Code aus [loesung-pnp.md](loesung-pnp.md) vorführen). Die
> Schreiboperationen in Block 6 sind das Herzstück — nicht kürzen.

---

## Backup-Plan für typische Probleme

| Problem | Erste Hilfe |
|---|---|
| `yo` bricht / Generator fehlt | `npm install -g gulp-cli@2.3.0 yo@3.1.1 @microsoft/generator-sharepoint@1.10.0` (mit `fnm use 8.17.0`). Generator 1.10 + Ziel „SharePoint 2019 onwards" erzeugt das SPFx-1.4.1-Projekt. |
| `gulp serve` bricht mit node-gyp/Python | `node -v` ≠ `v8.17.0` → `fnm use 8.17.0`. |
| Workbench weiß / „insecure" | `gulp trust-dev-cert` lief nicht, oder Chrome-Flag `chrome://flags/#allow-insecure-localhost`. Alternativ Edge. |
| Listen-Picker erscheint nicht | `PropertyControlStrings` wird von `npm install` automatisch in `config.json` gesetzt; fehlt der Eintrag, ergänzen und `gulp serve` neu starten. |
| Lesen: 404/400 | Liste in **dieser** Site vorhanden? Picker liefert die Id — URL `lists(guid'...')`. |
| Schreiben: 403 | Teilnehmer-Account braucht **Bearbeiten**-Rechte, nicht nur Lesen. |
| Anlegen: Metadaten-/„type"-Fehler | `odata=nometadata` in Accept **und** Content-type. |
| Schreiben: „OData-Version … 3.0 or 4.0" | `'odata-version': ''` aus den Headern entfernen — `SPHttpClient.configurations.v1` validiert ihn, der Client setzt selbst einen gültigen Wert. |
| Tabelle aktualisiert sich nach Schreiben nicht | Nach der Schreiboperation `_ladeListe()` / `this._load()` aufrufen. |
| `@pnp/sp`-Compile-Fehler | Falsche Version installiert. Nur `@pnp/sp@1.3.11` (v1) passt zu SPFx 1.4.1 / TS 2.4.2. |
| `?.` / `??` Compile-Fehler | TS 2.4.2 kennt das nicht. `||` / `&&` verwenden. |
| `dataVersion` rot unterringelt | Bekannter Typkonflikt in den SPFx-1.4.1-Typen; der Build läuft. `// @ts-ignore` direkt über `get dataVersion()` setzen. |
| React: „valid element (or null) must be returned" | Ein `render()`-Pfad gibt `undefined` zurück. Jeder Zweig muss returnen (`null` ist erlaubt). |

---

## Vorbereitung am Vorabend / 30 min vorher

- [ ] **Test-Liste `Aufgaben`** auf einer Workshop-Site anlegen — Spalten
      `Prioritaet` (Choice), `Status` (Choice), `FaelligAm` (Datum). **Interne
      Namen ASCII halten** (erst ASCII anlegen, dann Anzeigenamen umbenennen).
      5–8 gemischte Einträge.
- [ ] **Bearbeiten-Rechte** für die Teilnehmer-Accounts auf der Liste setzen —
      sonst scheitern Block 6/7 mit 403.
- [ ] SharePoint-Workbench-URL bereithalten:
      `http://<server>/sites/<site>/_layouts/15/workbench.aspx`
- [ ] Auf der Trainer-Maschine das Web Part einmal komplett durchbauen
      (`SPHttpClient` **und** `@pnp/sp`), Timings notieren.
- [ ] Ein **fertig gescaffoldetes `listenpflege`-Projekt** als Fallback bereit
      (gezippt), falls jemandes `yo` streikt.
- [ ] [loesung-vanilla.md](loesung-vanilla.md), [loesung-pnp.md](loesung-pnp.md),
      [loesung-react.md](loesung-react.md) offen im zweiten Monitor.
- [ ] `gulp trust-dev-cert` auf der Demo-Maschine schon erledigt.

---

## Rückblick: Aufbau größerer SPFx-Apps (Block 8) — Talking Points

Whiteboard in zwei Spalten: links „dieses Web Part", rechts „größere App".

1. **Datenzugriff**: `spHttpClient` / `sp.web...` ⟷ eigener **Service**
   „Der Service ist genau dieser Datenzugriff, in eine eigene Klasse gezogen."
2. **Übersetzung**: `daten.value` direkt ⟷ **Mapper**
   „Der Mapper übersetzt SP-Spaltennamen in saubere Objekt-Felder. Heute brauchten
   wir ihn nicht — bei vielen Feldern lohnt er sich."
3. **Zustand**: Tabelle aus `properties`/`state` ⟷ **State-Container** (z. B. Redux)
   „Der Container ist `this.state`, aus der Komponente herausgezogen, damit viele
   Komponenten denselben Zustand teilen."
4. **Konfiguration**: `properties.listId` aus der Property Pane ⟷ Web-Part-Klasse
   „Die Web-Part-Klasse reicht die Konfiguration als Props in den Komponentenbaum."

**Abschluss-Satz**: „Eine große App ist im Kern dasselbe wie das Web Part von heute
— nur sind Datenzugriff, Übersetzung und Zustand in eigene Dateien ausgelagert,
weil ein echtes Formular viele Felder hat."

---

## Nach dem Workshop

- Kurzes Feedback: Welcher Block hat am meisten gebracht?
- Hat der `SPHttpClient`-vs-`@pnp/sp`-Vergleich die Entscheidung „wann was" klarer
  gemacht?
- Hat der Architektur-Rückblick (Block 8) die Schichten größerer Apps greifbar
  gemacht? → Wenn ja, ist das Tagesziel erreicht.
- Neue Stolperfallen in die Backup-Tabelle oben einpflegen.
