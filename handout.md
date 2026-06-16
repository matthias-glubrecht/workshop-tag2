# Handout — Ein Web Part zur Listenpflege

Diese Anleitung beschreibt Schritt für Schritt, wie ein einfaches Web Part zur
Listenpflege entsteht. Wir bauen es gemeinsam vom leeren Ordner bis zu einem
kleinen, aber vollständigen Werkzeug, das die Einträge einer SharePoint-Liste
anzeigt und das **Anlegen, Bearbeiten und Löschen** erlaubt. Dabei wird jeder
Schritt genau erklärt, und jede Entscheidung, die unterwegs zu treffen ist, wird
ausführlich begründet — so bleibt stets nachvollziehbar, *warum* wir etwas tun,
nicht nur, *dass* wir es tun.

Sollte einmal etwas klemmen, liegen die fertigen Dateien zum Vergleich bereit: in
[loesung-vanilla.md](loesung-vanilla.md) (mit `SPHttpClient`),
[loesung-pnp.md](loesung-pnp.md) (mit `@pnp/sp`) und
[loesung-react.md](loesung-react.md) (in React).

Zwei Dinge begleiten uns den ganzen Tag. Erstens: Nach jeder Änderung laden wir die
Seite im Browser **hart neu** (`Strg+Shift+R`), denn SPFx hält Dateien hartnäckig
im Cache. Zweitens: Wir benennen die Bausteine durchgehend so, wie sie auch im Code
heißen (`render`, `properties`, `SPHttpClient`) — dann muss niemand zwischen dem
deutschen Begriff und dem Code-Bezeichner hin- und herübersetzen.

---

## Infrastruktur — Zugriff auf den Test-SharePoint

Bevor wir mit dem Bauen beginnen, stellen wir sicher, dass unser Rechner den
Test-SharePoint überhaupt erreichen kann. Dieser Server ist nur über seinen
internen Hostnamen ansprechbar und im Netz **nicht per DNS bekannt** — sein Name
ist also nirgends zentral hinterlegt. Damit Browser, `gulp serve` und die
REST-Aufrufe den Namen trotzdem auflösen können, tragen wir ihn auf **jedem
Teilnehmerrechner einmalig** von Hand in die HOSTS-Datei ein:

| Hostname | IP-Adresse |
|---|---|
| `sharepoint.pangaea.local` | `192.168.2.212` |

### HOSTS-Eintrag setzen (Administratorrechte nötig)

PowerShell **als Administrator** öffnen und die Zeile anhängen:

```pwsh
Add-Content -Path "$env:windir\System32\drivers\etc\hosts" -Value "192.168.2.212`tsharepoint.pangaea.local"
```

(Das `` `t `` erzeugt einen Tabulator zwischen IP und Name.)

Alternativ die Datei `C:\Windows\System32\drivers\etc\hosts` von Hand in einem
**als Administrator** gestarteten Editor öffnen und diese Zeile ergänzen:

```
192.168.2.212    sharepoint.pangaea.local
```

### Prüfen

```pwsh
ping sharepoint.pangaea.local
```

Entscheidend ist, dass der Name zu `192.168.2.212` **aufgelöst** wird. Ob die
Ping-Pakete selbst beantwortet werden, ist nebensächlich (ICMP kann geblockt sein).

Danach ist der Server im Browser erreichbar — die Workshop-Site und die
SharePoint-Workbench liegen unter diesem Host:

```
https://sharepoint.pangaea.local/sites/<workshop-site>/_layouts/15/workbench.aspx
```

> Überall im weiteren Verlauf, wo `<server>` steht (z. B. in Teil 4 und Teil 5),
> ist `sharepoint.pangaea.local` gemeint.

---

## Teil 0 — Vorbereitung (5 min)

Zunächst legen wir ein Arbeitsverzeichnis an, in dem alle Dateien unseres Projekts
abgelegt werden. Wir wählen dafür bewusst einen **kurzen Pfad** und meiden OneDrive
— synchronisierte Pfade führen bei SPFx erfahrungsgemäß zu schwer auffindbaren
Problemen.

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

Ein SPFx-Projekt schreibt man nicht von Hand, sondern lässt sich das Grundgerüst
von einem Generator erzeugen — dem Yeoman-Generator für SharePoint. Er stellt uns
eine Reihe Fragen und legt daraus ein vollständiges, sofort lauffähiges Projekt an.
Wir starten ihn mit:

```pwsh
yo @microsoft/sharepoint
```

und beantworten seine Fragen wie folgt:

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

## Teil 2 — Zum Laufen bringen (Block 2, ~30 min)

Bevor wir eigenen Code schreiben, bringen wir das frisch erzeugte Web Part einmal
zum Laufen. So überzeugen wir uns früh davon, dass die gesamte Werkzeugkette
funktioniert — und sehen am Ende dieses Teils unser Web Part zum ersten Mal in der
Workbench.

Die lokale Workbench benötigt zunächst ein Entwicklungs-Zertifikat. Das richten wir
einmalig pro Maschine ein:

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

## Teil 3 — Property Pane: Listenname als Textfeld (Block 3, ~30 min)

Web Parts besitzen eine sogenannte Property Pane — den Bereich, der beim Bearbeiten
einer Seite rechts aufklappt und in dem Benutzer das Web Part konfigurieren. In
diesem Teil geben wir unserem Web Part seine erste Konfigurationseigenschaft (eine
Web-Part-Property): einen **Listennamen**. Diesen Namen verwenden wir anschließend
in Teil 4, um die passende Liste auszulesen.

Wir gehen dazu in drei kleinen Schritten vor: die Eigenschaft im Code deklarieren,
ein Eingabefeld in der Property Pane anbieten und den eingegebenen Wert anzeigen.

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

## Teil 4 — Liste lesen und anzeigen (Block 4, ~60 min)

Jetzt wird es spannend: Wir holen zum ersten Mal echte Daten aus SharePoint. Dafür
brauchen wir zweierlei — eine Liste, die wir auslesen können, und den Code, der das
tut. Also legen wir zunächst eine Test-Liste an und lesen sie anschließend über die
REST-Schnittstelle von SharePoint aus.

### 4a) Test-Liste anlegen (einmalig)

Wir beginnen mit der Datenquelle. In der echten Site (nicht der lokalen Workbench,
z. B. `https://<server>/sites/dev`) legen wir eine Liste **Aufgaben** an (Typ
„Liste").

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

## Teil 5 — Property Pane: echter Listen-Selektor (Block 5, ~30 min)

Den Listennamen von Hand einzutippen ist bequem, aber fehleranfällig — ein
Tippfehler oder die falsche Groß-/Kleinschreibung, und nichts wird gefunden. In
diesem Teil ersetzen wir das Textfeld deshalb durch einen echten Listen-Selektor:
ein Auswahlfeld, das uns alle Listen der Site zur Auswahl anbietet.
Praktischerweise müssen wir dieses Steuerelement nicht selbst bauen — es steckt
fertig in den **SPFx Property Controls** und heißt dort `PropertyFieldListPicker`.

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

## Teil 6 — Anlegen, Bearbeiten, Löschen (Block 6, ~75 min)

Bisher haben wir nur gelesen. In diesem Teil bringen wir unserem Web Part bei, die
Daten auch zu verändern: neue Einträge anzulegen, bestehende zu bearbeiten und
welche zu löschen. Damit ist die Listenpflege vollständig — Fachleute fassen diese
vier Grundoperationen (Lesen, Erstellen, Ändern, Löschen) unter dem Kürzel **CRUD**
zusammen.

Alle drei schreibenden Operationen laufen über denselben Aufruf,
`spHttpClient.post(...)`. Was sie unterscheidet, sind nicht etwa verschiedene
Methoden, sondern die mitgeschickten **HTTP-Header** — ein Detail, das wir uns gleich
genauer ansehen.

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

> **`IF-MATCH` ist mehr als Pflichtprogramm — es ist optimistische
> Nebenläufigkeitskontrolle.** Jedes Listenelement trägt ein **ETag**: eine
> Versionskennung, die SharePoint bei jeder Änderung hochzählt. Der
> `IF-MATCH`-Header legt fest, *unter welcher Bedingung* geschrieben werden darf:
>
> - `IF-MATCH: '*'` → schreiben **ohne** Prüfung. Das ist **last-write-wins**: Hat
>   jemand zwischenzeitlich geändert, wird seine Änderung kommentarlos überschrieben.
> - `IF-MATCH: '<etag>'` → schreiben **nur**, wenn das Element noch exakt die
>   Version hat, die ich gelesen habe. Sonst lehnt SharePoint mit **HTTP 412
>   Precondition Failed** ab — und **nichts** wird überschrieben.
>
> „Optimistisch" heißt: Das Element wird beim Bearbeiten **nicht gesperrt**.
> Geprüft wird erst beim Speichern. Das ist im Web der Normalfall — man kann einen
> Datensatz nicht blockieren, nur weil bei jemandem ein Formular offen steht.

So sieht das optimistische Speichern aus. Voraussetzung: beim Lesen das ETag pro
Element **mitnehmen** — dazu `odata=minimalmetadata` statt `nometadata`, dann steht
in jedem Element ein Feld `odata.etag` (z. B. `"4"`):

```typescript
private _bearbeitenGeprueft(id: number, titel: string, etag: string): Promise<void> {
  const url: string = this.context.pageContext.web.absoluteUrl
    + "/_api/web/lists(guid'" + this.properties.listId + "')/items(" + id + ")";

  return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': '',
      'IF-MATCH': etag,            // das gemerkte ETag statt '*'
      'X-HTTP-Method': 'MERGE'
    },
    body: JSON.stringify({ Title: titel })
  }).then((antwort: SPHttpClientResponse): void => {
    if (antwort.status === 412) {
      throw new Error('Der Eintrag wurde zwischenzeitlich von jemand anderem '
        + 'geändert. Bitte neu laden und die Änderung erneut eingeben.');
    }
    if (!antwort.ok) { throw new Error('Bearbeiten fehlgeschlagen: ' + antwort.status); }
  });
}
```

> **Im Workshop** bleiben wir der Einfachheit halber bei `'*'`. Wer das echte
> optimistische Speichern zeigen will, liest die Liste mit `odata=minimalmetadata`,
> merkt sich pro Zeile das `odata.etag` und reicht es hier herein. Der 412-Zweig
> ist dann das Erfolgserlebnis: zwei Browser-Tabs nebeneinander, im einen ändern,
> im anderen speichern → saubere Fehlermeldung statt stillem Überschreiben.

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

> **Beim Löschen lohnt sich `IF-MATCH` genauso.** `IF-MATCH: '*'` löscht
> bedingungslos. `IF-MATCH: '<etag>'` löscht **nur**, wenn das Element seit dem
> Lesen unverändert ist — sonst **412**. Das schützt davor, fremde frische Arbeit
> wegzuwerfen: Du siehst eine Aufgabe „Offen" und willst sie entfernen; in der
> Zwischenzeit hat ein Kollege sie auf „In Arbeit" gesetzt und Details ergänzt. Mit
> dem ETag-Check scheitert dein Löschen mit 412, statt seine Änderung zu vernichten.

```typescript
private _loeschenGeprueft(id: number, etag: string): Promise<void> {
  const url: string = this.context.pageContext.web.absoluteUrl
    + "/_api/web/lists(guid'" + this.properties.listId + "')/items(" + id + ")";

  return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: { 'IF-MATCH': etag, 'X-HTTP-Method': 'DELETE', 'odata-version': '' }
  }).then((antwort: SPHttpClientResponse): void => {
    if (antwort.status === 412) {
      throw new Error('Der Eintrag wurde zwischenzeitlich geändert. Bitte prüfen, '
        + 'ob das Löschen noch gewünscht ist.');
    }
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

## Teil 7 — Gegenüberstellung `@pnp/sp` (Block 7, ~45 min)

Den Datenzugriff, den wir uns gerade Schritt für Schritt von Hand erarbeitet haben,
gibt es auch in deutlich kürzerer Form. In diesem Teil schreiben wir dieselben vier
Operationen ein zweites Mal — diesmal mit der Bibliothek `@pnp/sp` — und stellen
beide Wege nebeneinander, um abschließend zu beurteilen, wann sich welcher lohnt.

Ein Hinweis vorweg, der auf unserer eingefrorenen Plattform entscheidend ist: Wir
brauchen zwingend die **Version 1** (`@pnp/sp@1.3.11`), passend zu SPFx 1.4.1 und
TypeScript 3.6. Das moderne `@pnp/sp` v3/v4 setzt neueres TypeScript und Node voraus
und läuft hier **nicht**.

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

Zum Abschluss treten wir einen Schritt zurück. Unser Web Part hält alles in einer
einzigen Klasse — Datenzugriff, Anzeige und Konfiguration. Für ein kleines Web Part
ist das genau richtig. Wächst eine Anwendung jedoch, werden diese Aufgaben
üblicherweise in eigene Schichten getrennt:

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
