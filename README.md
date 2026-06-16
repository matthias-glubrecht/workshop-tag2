# Workshop Tag 2 — Ein Web Part von Grund auf: „Listenpflege"

Ein Tag zum **selber bauen**. Die Teilnehmer erzeugen ein SPFx-Web-Part von null
und entwickeln es Schritt für Schritt zu einem kleinen, aber vollständigen
Werkzeug: Es zeigt die Einträge einer SharePoint-Liste und erlaubt, **neue
anzulegen, vorhandene zu bearbeiten und zu löschen**.

Unterwegs lernen die Teilnehmer zwei Dinge, die in jedem realen SPFx-Projekt
vorkommen:

- die **Property Pane** — zunächst mit einem einfachen Textfeld für den
  Listennamen, dann mit einem echten **Listen-Selektor** aus den SPFx Property
  Controls,
- den Datenzugriff auf zwei Arten im direkten Vergleich: **`SPHttpClient`** (der
  in SPFx eingebaute Weg) und **`@pnp/sp`** (die Bibliothek, die vieles abkürzt).

> **Funktion des Web Parts in einem Satz**: „Listenpflege" zeigt die Einträge
> einer konfigurierten Liste als Tabelle und bietet Anlegen, Bearbeiten und
> Löschen — also die vier Grundoperationen Lesen, Erstellen, Ändern, Entfernen
> (CRUD) an einem realen Beispiel.

## Warum so aufgebaut

Reale SPFx-Apps bringen React, einen State-Container (z. B. Redux), einen
Service-Layer und Mapper oft gleichzeitig mit. Das ist viel auf einmal. Hier
bauen wir das **nackte Skelett** selbst und ergänzen **eine Fähigkeit nach der
anderen** — am Ende steht eine klare Vorstellung davon, wie größere Apps denselben
Datenzugriff strukturieren.

---

## Lernziele

Am Ende kann jede Teilnehmerin / jeder Teilnehmer:

- ein SPFx-Projekt (1.4.1, on-premises-tauglich) mit dem Yeoman-Generator erzeugen,
- die Anatomie eines Web Parts benennen (Manifest, Web-Part-Klasse, `render()`),
- `gulp serve` ausführen und das Ergebnis in der lokalen Workbench sehen,
- eine Property in der Property Pane anlegen — als Textfeld **und** als Listen-Selektor,
- eine SharePoint-Liste lesen und als Tabelle anzeigen,
- Einträge anlegen, bearbeiten und löschen,
- denselben Datenzugriff mit `SPHttpClient` und mit `@pnp/sp` schreiben und die
  beiden Wege bewerten,
- nachvollziehen, wie größere SPFx-Apps denselben Datenzugriff in Service-,
  Mapper- und State-Schichten aufteilen.

Alles davon ist **übertragbares SPFx-Grundwissen** und gilt für jedes SPFx-Projekt.

---

## Tagesablauf (Start 08:00, Ende 15:45)

| Zeit | Dauer | Block | Inhalt | Meilenstein |
|---|---|---|---|---|
| **08:00 – 08:15** | 15 min | M0 | Begrüßung + Einordnung des Tages | — |
| **08:15 – 09:00** | 45 min | Block 1 | Projekt scaffolden, Anatomie des Web Parts | — |
| **09:00 – 09:30** | 30 min | Block 2 | `gulp serve`, erstes Rendern in der Workbench | Web Part rendert |
| **09:30 – 09:45** | 15 min | Pause | Kaffee | — |
| **09:45 – 10:15** | 30 min | Block 3 | Property Pane: Listenname als Textfeld | Konfiguration greift |
| **10:15 – 11:15** | 60 min | Block 4 | Liste lesen, als Tabelle anzeigen (`SPHttpClient`) | Echte Daten |
| **11:15 – 11:45** | 30 min | Block 5 | Property Pane: echter Listen-Selektor (Property Controls) | Auswahl statt Tippen |
| **11:45 – 12:45** | 60 min | Pause | Mittag | — |
| **12:45 – 14:00** | 75 min | Block 6 | Anlegen, Bearbeiten, Löschen (`SPHttpClient`) | Vollständiges CRUD |
| **14:00 – 14:15** | 15 min | Pause | Kaffee | — |
| **14:15 – 15:00** | 45 min | Block 7 | Gegenüberstellung `@pnp/sp` (lesen und schreiben) | Zwei Wege bewertet |
| **15:00 – 15:30** | 30 min | Block 8 | Rückblick: Aufbau größerer SPFx-Apps | Architektur eingeordnet |
| **15:30 – 15:45** | 15 min | Wrap-up | Q&A + Feedback | — |

> Die **React-Variante** (dasselbe in einer Komponente mit State) liegt als
> vollständige Referenz in [loesung-react.md](loesung-react.md) bereit — als
> optionaler Pfad für schnelle Teilnehmer oder als Vorlage für eine Folgeeinheit.

---

## Inhalt des Ordners

```
workshop-tag2/
├── README.md                — diese Datei (Überblick + Ablauf)
├── handout.md               — Schritt-für-Schritt für die Teilnehmer
├── trainer-spickzettel.md   — Ablauf, Sprech-Anker, Backup-Plan
├── loesung-vanilla.md       — fertiges Web Part mit SPHttpClient (ohne Framework)
├── loesung-pnp.md           — dasselbe Web Part mit @pnp/sp (direkter Vergleich)
└── loesung-react.md         — dasselbe Web Part in React
```

> Die Lösungen liegen bewusst als **Markdown-Code-Listings** vor, nicht als echte
> `.ts`-Dateien. So werden sie nicht versehentlich von einem Build mitkompiliert.
> Zum Einsatz werden sie in ein **frisch gescaffoldetes Projekt** kopiert.

---

## Voraussetzungen

| Werkzeug | Wofür | Prüfen mit |
|---|---|---|
| **fnm + Node 8.17.0** | SPFx-1.4.1-Zwang | `node -v` → `v8.17.0` |
| **Yeoman + SP-Generator 1.4.1** | Projekt scaffolden | `yo --version` |
| **gulp 3** | Build/Serve | `gulp -v` |
| **Chrome/Edge** | Workbench | — |

Globale Pakete (falls auf einem Laptop noch nicht vorhanden):

```pwsh
fnm use 8.17.0
npm install -g yo gulp@3 @microsoft/generator-sharepoint@1.4.1
```

> **Wichtig**: Das neue Projekt wird in einem **eigenen, kurzen Pfad** angelegt
> (z. B. `C:\repos\listenpflege`), NICHT in OneDrive.

---

## Der rote Faden

1. **Scaffold** → ein Web Part ist eine Klasse mit einer `render()`-Methode.
2. **Serve** → der eigene Code rendert live in der Workbench.
3. **Property Pane (Textfeld)** → Konfiguration kommt als `this.properties.xxx` rein.
4. **Lesen** → echte Listendaten über die SharePoint-REST-API.
5. **Property Pane (Picker)** → Liste auswählen statt Namen tippen.
6. **Schreiben** → anlegen, bearbeiten, löschen.
7. **`@pnp/sp`** → derselbe Datenzugriff, kürzer geschrieben — und die Frage, wann sich das lohnt.
8. **Größere Apps** → dieselbe Idee, in mehr Schichten (Service, Mapper, State).

Vollständige Schritte: **[handout.md](handout.md)**.
Trainer-Anker und Notfallplan: **[trainer-spickzettel.md](trainer-spickzettel.md)**.
