# Handout — Das Web Part „Listenpflege"

> Schritt für Schritt vom leeren Ordner zu einem Web Part, das die Einträge einer
> SharePoint-Liste anzeigt und das **Anlegen, Bearbeiten und Löschen** erlaubt.
> Wenn etwas klemmt: die fertigen Dateien liegen in
> [loesung-vanilla.md](loesung-vanilla.md) (mit `SPHttpClient`),
> [loesung-pnp.md](loesung-pnp.md) (mit `@pnp/sp`) und
> [loesung-react.md](loesung-react.md) (React).

**Faustregel für den ganzen Tag**: Nach jeder Änderung im Browser **hart neu laden**
(`Strg+Shift+R`). SPFx cached aggressiv.

Wir verwenden durchgehend die SharePoint-Begriffe so, wie sie im Code stehen
(`render`, `properties`, `SPHttpClient`), damit nichts mental übersetzt werden muss.

---

## Teil 0 — Vorbereitung (5 min)

Neues, kurzes Arbeitsverzeichnis — **nicht** in OneDrive:

```pwsh
fnm use 8.17.0
node -v        # muss v8.17.0 zeigen
mkdir C:\repos\listenpflege
cd C:\repos\listenpflege
```

Falls `yo` / Generator fehlen:

```pwsh
npm install -g yo gulp@3 @microsoft/generator-sharepoint@1.4.1
```

---

## Teil 1 — Projekt scaffolden (Block 1, ~25 min)

```pwsh
yo @microsoft/sharepoint
```

Antworten im Assistenten:

| Frage | Antwort |
|---|---|
| What is your solution name? | `listenpflege` |
| Which baseline packages? | **SharePoint 2016 onwards including 2019 and SharePoint Online** *(die on-premises-taugliche Wahl — NICHT „SharePoint Online only")* |
| Where to place the files? | **Use the current folder** |
| Deploy to all sites? | `N` |
| Permissions / CDN-Frage | Enter (Default) |
| Type of client-side component? | **WebPart** |
| Web part name? | `ListEditor` |
| Description | `Einträge einer Liste anzeigen und pflegen` |
| Framework? | **No JavaScript framework** |

> Der Generator lädt jetzt Abhängigkeiten. Das dauert ein paar Minuten — Zeit für
> den Anatomie-Teil unten.

### Anatomie des erzeugten Projekts

```
listenpflege/
├── config/
│   └── config.json                   → welche Bundles gebaut werden
├── src/webparts/listEditor/
│   ├── ListEditorWebPart.ts           ← hier passiert alles
│   ├── ListEditorWebPart.manifest.json← Id, Titel, Default-Properties
│   └── loc/                           → Texte (de-de, en-us)
└── gulpfile.js
```

Die zentrale Datei ist `ListEditorWebPart.ts`. Öffne sie. Das Skelett:

```typescript
export default class ListEditorWebPart extends BaseClientSideWebPart<IListEditorWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `... HTML als Text ...`;
  }
}
```

Der Kern ist überschaubar: Ein Web Part ist eine Klasse. `render()` schreibt HTML
in `this.domElement` (ein `<div>`, das SharePoint bereitstellt). Alles Weitere
heute baut darauf auf.

---

## Teil 2 — Zum Laufen bringen (Block 2, ~30 min) → Meilenstein: Web Part rendert

Die lokale Workbench braucht ein Zertifikat (einmalig pro Maschine):

```pwsh
gulp trust-dev-cert
```

Dann starten:

```pwsh
gulp serve --nobrowser
```

Browser öffnen auf die **lokale** Workbench:

```
https://localhost:4321/temp/workbench.html
```

> Die Workbench bleibt weiß / „insecure"? → Chrome-Flag
> `chrome://flags/#allow-insecure-localhost` aktivieren oder Edge nutzen.

In der Workbench: **+** klicken → das Web Part **ListEditor** auswählen. Es erscheint.

Ändere jetzt die Überschrift in `render()`, damit der Name die Funktion trägt:

```typescript
public render(): void {
  this.domElement.innerHTML = `
    <div style="padding:16px; font-family:Segoe UI, sans-serif;">
      <h2>Listenpflege</h2>
      <p>Hier erscheinen gleich die Einträge einer SharePoint-Liste.</p>
    </div>`;
}
```

Speichern → `gulp serve` rebuildet automatisch → Browser hart neu laden
(`Strg+Shift+R`). Die Änderung ist sichtbar.

> Das ist die Arbeitsschleife für den ganzen Tag (und für jedes echte Projekt):
> **ändern → speichern → hart neu laden → prüfen.**

---

## Teil 3 — Property Pane: Listenname als Textfeld (Block 3, ~30 min) → Meilenstein: Konfiguration greift

Ziel: Im Bearbeiten-Bereich rechts soll ein **Listenname** eingetragen werden
können. Den brauchen wir in Teil 4 zum Lesen.

### 3a) Property deklarieren

Oben in `ListEditorWebPart.ts` das Props-Interface festlegen:

```typescript
export interface IListEditorWebPartProps {
  listName: string;
}
```

### 3b) Eingabefeld in die Property Pane

`PropertyPaneTextField` kommt in SPFx 1.4.1 noch aus `@microsoft/sp-webpart-base`
(in neueren Versionen aus `@microsoft/sp-property-pane` — bei uns **nicht**):

```typescript
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
```

```typescript
groupFields: [
  PropertyPaneTextField('listName', { label: 'Name der SharePoint-Liste' })
]
```

### 3c) Wert anzeigen

In `render()` den Wert ausgeben — **immer escapen** gegen XSS:

```typescript
import { escape } from '@microsoft/sp-lodash-subset';
```

```typescript
public render(): void {
  this.domElement.innerHTML = `
    <div style="padding:16px;">
      <h2>Listenpflege: ${escape(this.properties.listName)}</h2>
    </div>`;
}
```

Hart neu laden → rechts im Property-Bereich tippen → der Wert erscheint **live**.
Konfiguration kommt also als `this.properties.xxx` in `render()` an.

> **Sicherheitshinweis**: Benutzereingaben (wie `listName`) nie roh in `innerHTML`
> schreiben — immer `escape()`. Sonst XSS. React nimmt einem das später ab.

---

## Teil 4 — Liste lesen und anzeigen (Block 4, ~60 min) → Meilenstein: echte Daten

### 4a) Test-Liste anlegen (einmalig)

In der echten Site (nicht der lokalen Workbench, z. B.
`https://<server>/sites/dev`): Liste **Aufgaben** anlegen (Typ „Liste").

Damit die Tabelle später beim Bonus etwas hergibt, lohnen sich Spalten mit
verschiedenen Typen — alle kommen direkt per REST zurück:

| Anzeigename | Anlegen als (= interner Name) | Typ | Werte |
|---|---|---|---|
| Titel | `Title` (schon da) | Text | — |
| Priorität | `Prioritaet` | Auswahl | Hoch · Mittel · Niedrig |
| Status | `Status` | Auswahl | Offen · In Arbeit · Erledigt |
| Fällig am | `FaelligAm` | Datum | — |

> **Interner Name ≠ Anzeigename**: SharePoint friert beim **Anlegen** den internen
> Namen ein. „Fällig am" würde intern zu `F_x00e4_llig_x0020_am` — genau den
> bräuchte man dann in der REST-URL. Deshalb Spalten zuerst mit **ASCII-Namen ohne
> Leerzeichen** anlegen (`Prioritaet`, `FaelligAm`), danach optional den
> Anzeigenamen umbenennen. Das ist eine SPFx-Grundregel.

5–8 Einträge mit gemischten Werten anlegen.

> Zum echten Datenlesen brauchst Du die **SharePoint-gehostete** Workbench:
> `https://<server>/sites/dev/_layouts/15/workbench.aspx`.
> Die `localhost`-Workbench hat keine echte Liste im Hintergrund.

### 4b) Lesen mit `SPHttpClient`

SPFx liefert `this.context.spHttpClient` — den vorkonfigurierten HTTP-Client mit
Authentifizierung. Importe und eine Typ-Form für ein Listen-Element:

```typescript
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IListItem {
  Id: number;
  Title: string;
}
```

Den ersten Lese-Versuch halten wir bewusst bei `Id,Title` — schlicht und robust:

```typescript
private _ladeListe(): void {
  const status: Element = this.domElement.querySelector('#status');
  if (!this.properties.listName) {
    status.innerHTML = 'Bitte rechts einen Listennamen eintragen.';
    return;
  }

  const url: string = this.context.pageContext.web.absoluteUrl
    + "/_api/web/lists/getbytitle('" + this.properties.listName
    + "')/items?$select=Id,Title&$top=50";

  this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then((antwort: SPHttpClientResponse): Promise<{ value: IListItem[] }> => {
      if (!antwort.ok) { throw new Error('HTTP ' + antwort.status); }
      return antwort.json();
    })
    .then((daten: { value: IListItem[] }): void => {
      let zeilen: string = '';
      for (let i: number = 0; i < daten.value.length; i++) {
        zeilen = zeilen + '<tr><td>' + daten.value[i].Id + '</td><td>'
          + escape(daten.value[i].Title) + '</td></tr>';
      }
      status.innerHTML =
        '<table border="1" cellpadding="6" style="border-collapse:collapse;">'
        + '<thead><tr><th>Id</th><th>Titel</th></tr></thead>'
        + '<tbody>' + zeilen + '</tbody></table>';
    })
    .catch((fehler: Error): void => {
      status.innerHTML = 'Fehler: ' + escape(fehler.message);
    });
}
```

`render()` legt den Container an und ruft `_ladeListe()` auf. Vollständige Datei:
[loesung-vanilla.md](loesung-vanilla.md).

Hart neu laden in der **SharePoint-Workbench** → die Tabelle füllt sich mit echten
Einträgen.

### Stolperfallen

- **TS 3.6 / target es5**: kein `??`, kein `?.`. Nur `||`, `&&`, klassische `if`.
- **`escape()`** bei jedem Wert, der in `innerHTML` geht.
- **403 / CORS** in der `localhost`-Workbench ist normal — echte Daten nur in der
  SharePoint-gehosteten Workbench.

---

## Teil 5 — Property Pane: echter Listen-Selektor (Block 5, ~30 min) → Meilenstein: Auswahl statt Tippen

Ein Textfeld ist fehleranfällig (Tippfehler, Groß-/Kleinschreibung). Besser: ein
Dropdown aller Listen der Site. Das liefert `PropertyFieldListPicker` aus den
**SPFx Property Controls** — fertig nutzbar, kein Eigenbau.

### 5a) Paket installieren

```pwsh
npm install @pnp/spfx-property-controls@1.20.0
```

In `config/config.json` die mitgelieferten Texte registrieren (unter
`localizedResources`):

```json
"PropertyControlStrings": "node_modules/@pnp/spfx-property-controls/lib/loc/{locale}.js"
```

> `gulp serve` nach dem Installieren **einmal neu starten** — neue Pakete werden
> sonst nicht gebündelt.

### 5b) Property umstellen: vom Namen zur Id

Der Picker liefert die **Listen-Id (GUID)**, nicht den Namen. Das ist robuster
(der Name kann sich ändern, die Id nicht). Property umbenennen:

```typescript
export interface IListEditorWebPartProps {
  listId: string;
}
```

### 5c) Picker einbinden

```typescript
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
```

```typescript
groupFields: [
  PropertyFieldListPicker('listId', {
    label: 'Liste auswählen',
    selectedList: this.properties.listId,
    includeHidden: false,
    orderBy: PropertyFieldListPickerOrderBy.Title,
    disabled: false,
    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
    properties: this.properties,
    context: this.context,
    deferredValidationTime: 0,
    key: 'listPickerFieldId'
  })
]
```

### 5d) Lese-URL auf die Id umstellen

Statt `getbytitle('<name>')` jetzt die Id:

```typescript
const url: string = this.context.pageContext.web.absoluteUrl
  + "/_api/web/lists(guid'" + this.properties.listId
  + "')/items?$select=Id,Title&$top=50";
```

Hart neu laden → rechts erscheint ein Dropdown mit allen Listen. Auswahl treffen →
die Tabelle lädt. Kein Listenname mehr zu tippen.

---

## Teil 6 — Anlegen, Bearbeiten, Löschen (Block 6, ~75 min) → Meilenstein: vollständiges CRUD

Bisher nur Lesen (`GET`). Jetzt die drei schreibenden Operationen. Alle gehen über
`spHttpClient.post(...)`; die Unterscheidung steckt in den **Headern**.

> **Digest**: Schreibende SharePoint-Aufrufe brauchen ein Sicherheits-Token
> (`X-RequestDigest`). Gute Nachricht: `SPHttpClient` setzt es **automatisch** —
> bei rohem `fetch` müsste man es selbst von `_api/contextinfo` holen.

### 6a) Anlegen (POST)

```typescript
private _anlegen(titel: string): Promise<void> {
  const url: string = this.context.pageContext.web.absoluteUrl
    + "/_api/web/lists(guid'" + this.properties.listId + "')/items";

  return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': ''
    },
    body: JSON.stringify({ Title: titel })
  }).then((antwort: SPHttpClientResponse): void => {
    if (!antwort.ok) { throw new Error('Anlegen fehlgeschlagen: ' + antwort.status); }
  });
}
```

> **`odata=nometadata`** erspart das Mitschicken des Entity-Typs
> (`__metadata: { type: 'SP.Data.AufgabenListItem' }`). Mit Verbose-Metadaten
> müsste man diesen exakten Namen kennen — eine klassische Stolperfalle.

### 6b) Bearbeiten (MERGE)

Gleiche URL plus die Item-Id; der Header `X-HTTP-Method: MERGE` macht aus dem POST
ein Update, `IF-MATCH: *` heißt „aktuelle Version, egal welche":

```typescript
private _bearbeiten(id: number, titel: string): Promise<void> {
  const url: string = this.context.pageContext.web.absoluteUrl
    + "/_api/web/lists(guid'" + this.properties.listId + "')/items(" + id + ")";

  return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': '',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
    },
    body: JSON.stringify({ Title: titel })
  }).then((antwort: SPHttpClientResponse): void => {
    if (!antwort.ok) { throw new Error('Bearbeiten fehlgeschlagen: ' + antwort.status); }
  });
}
```

### 6c) Löschen (DELETE)

Kein Body, Header `X-HTTP-Method: DELETE`:

```typescript
private _loeschen(id: number): Promise<void> {
  const url: string = this.context.pageContext.web.absoluteUrl
    + "/_api/web/lists(guid'" + this.properties.listId + "')/items(" + id + ")";

  return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', 'odata-version': '' }
  }).then((antwort: SPHttpClientResponse): void => {
    if (!antwort.ok) { throw new Error('Löschen fehlgeschlagen: ' + antwort.status); }
  });
}
```

### 6d) Bedienelemente

Ein Eingabefeld + „Hinzufügen"-Button über der Tabelle, pro Zeile ein
„Bearbeiten"- und ein „Löschen"-Button. Nach jeder Schreiboperation `_ladeListe()`
erneut aufrufen, damit die Tabelle den neuen Stand zeigt. Die vollständige
Verdrahtung (Event-Handler an die per `innerHTML` erzeugten Buttons) steht in
[loesung-vanilla.md](loesung-vanilla.md).

> **Reihenfolge merken**: Schreiben → auf das Promise warten → neu laden. Sonst
> zeigt die Tabelle noch den alten Stand.

---

## Teil 7 — Gegenüberstellung `@pnp/sp` (Block 7, ~45 min) → Meilenstein: zwei Wege bewertet

Dieselben vier Operationen mit `@pnp/sp`. Wichtig für die eingefrorene Plattform:
es muss die **Version 1** sein (`@pnp/sp@1.3.11`), passend zu SPFx 1.4.1 / TS 3.6.
Modernes pnp v3/v4 braucht neueres TypeScript und Node und läuft hier **nicht**.

```pwsh
npm install @pnp/sp@1.3.11
```

Einmalig einrichten (z. B. in `onInit`), damit pnp den SPFx-Kontext kennt:

```typescript
import { sp } from '@pnp/sp';

protected onInit(): Promise<void> {
  return super.onInit().then((): void => {
    sp.setup({ spfxContext: this.context });
  });
}
```

Die vier Operationen:

```typescript
import { sp } from '@pnp/sp';

// Lesen
const items = await sp.web.lists.getById(this.properties.listId)
  .items.select('Id', 'Title').top(50).get();

// Anlegen
await sp.web.lists.getById(this.properties.listId).items.add({ Title: titel });

// Bearbeiten
await sp.web.lists.getById(this.properties.listId)
  .items.getById(id).update({ Title: titel });

// Löschen
await sp.web.lists.getById(this.properties.listId).items.getById(id).delete();
```

> `async/await` ist mit TS 3.6 / target es5 erlaubt — der Compiler erzeugt den
> passenden Promise-Code. Wer bei `.then()` bleiben will, kann das auch.

### Die direkte Gegenüberstellung

| Operation | `SPHttpClient` (eingebaut) | `@pnp/sp` (Bibliothek) |
|---|---|---|
| Lesen | URL bauen, `get`, `antwort.json()`, `.value` auspacken | `...items.select(...).top(50).get()` |
| Anlegen | `post` + Accept/Content-type-Header + Body | `...items.add({ Title })` |
| Bearbeiten | `post` + `X-HTTP-Method: MERGE` + `IF-MATCH` + Body | `...items.getById(id).update({ Title })` |
| Löschen | `post` + `X-HTTP-Method: DELETE` + `IF-MATCH` | `...items.getById(id).delete()` |
| Digest | automatisch durch `SPHttpClient` | automatisch durch pnp |
| Entity-Typ | mit `odata=nometadata` umgangen | von pnp intern gelöst |

**Bewertung — was lernen wir daraus?**

- `SPHttpClient` ist **ohne zusätzliche Abhängigkeit** da und macht jeden Schritt
  sichtbar. Genau das ist für das Verständnis wertvoll: man sieht, dass ein Update
  „nur" ein POST mit einem speziellen Header ist.
- `@pnp/sp` **verbirgt** Header, Metadaten und URL-Bau hinter einer lesbaren Kette.
  Bei vielen Operationen spart das spürbar Code und Fehlerquellen.
- Faustregel: **Ein, zwei Aufrufe** → `SPHttpClient` reicht und spart eine
  Dependency. **Viel Datenzugriff** → `@pnp/sp` lohnt sich.
- Auf der eingefrorenen Plattform unbedingt **pnp v1** verwenden. Die Version ist
  hier kein Detail, sondern Voraussetzung.

---

## Teil 8 — Rückblick: Wie größere SPFx-Apps das strukturieren (Block 8, ~30 min)

Das fertige Web Part hält Datenzugriff, Anzeige und Konfiguration in einer Klasse.
Für ein kleines Web Part ist das genau richtig. Wächst eine Anwendung, werden diese
Aufgaben üblicherweise in eigene Schichten getrennt:

| Hier (klein) | In größeren Apps | Aufgabe |
|---|---|---|
| `spHttpClient` / `sp.web...` im Web Part | eigener **Service** | Daten von SharePoint holen und schreiben |
| `daten.value` direkt benutzt | **Mapper** | SP-Felder ↔ saubere Domänen-Objekte übersetzen |
| Tabelle aus `this.properties`/`state` heraus | **State-Container** (z. B. Redux) | Zustand zentral halten |
| `this.properties.listId` aus der Property Pane | Web-Part-Klasse reicht Props hinein | Konfiguration in den Komponentenbaum reichen |

**Kernaussage**: Eine große SPFx-App macht im Kern dasselbe wie dieses Web Part —
nur sind Datenzugriff (Service), Übersetzung (Mapper) und Zustand (State-Container)
in eigene Dateien ausgelagert, damit ein großes Formular wartbar bleibt. Wer dieses
Web Part versteht, versteht das Skelett der großen Variante.

---

## Wenn die Zeit reicht — Bonus

Die Test-Liste mit mehreren Spalten zahlt sich jetzt aus (immer **interne Namen**):

- **Mehr Spalten anzeigen**: `$select=Id,Title,Prioritaet,Status,FaelligAm` und je
  eine Tabellenspalte mehr.
- **Sortieren**: `&$orderby=FaelligAm asc` oder `&$orderby=Prioritaet desc`.
- **Filtern**: `&$filter=Status eq 'Offen'`.
- **Farbig nach Priorität**: beim Rendern je nach `Prioritaet` eine
  Hintergrundfarbe setzen.
- **Bearbeiten auf mehr Felder ausweiten**: `update({ Title, Status })` —
  derselbe Aufruf, mehr Felder im Objekt.

**Typ-Hinweise beim Rendern:**

- **Datum** (`FaelligAm`) kommt als ISO-String in UTC → mit
  `new Date(wert).toLocaleDateString('de-DE')` lesbar machen.
- **Auswahl** (`Prioritaet`, `Status`) kommt als String — bei **einfacher** Auswahl.
- **Personen-/Nachschlagespalten** kommen nicht direkt mit; sie brauchen
  `$expand=...&$select=.../Title`. Bewusst weggelassen, schönes Thema für „danach".

Und natürlich: **dasselbe in React** — vollständig in [loesung-react.md](loesung-react.md).
