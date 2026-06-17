# Lösung — „Listenpflege" mit `SPHttpClient` (ohne Framework)

> Vollständiges Web Part: Liste auswählen (Property-Picker), Einträge anzeigen,
> anlegen, bearbeiten, löschen — alles über den eingebauten `SPHttpClient`.
> Die Variante mit `@pnp/sp` steht in [loesung-pnp.md](loesung-pnp.md), die
> React-Variante in [loesung-react.md](loesung-react.md).

## So setzt Du es ein

1. `yo @microsoft/sharepoint` → Solution `listenpflege`, Web-Part `ListEditor`,
   Framework **No JavaScript framework** (Antworten siehe [handout.md](handout.md), Teil 1).
2. Property Controls installieren und in `config/config.json` registrieren:
   ```pwsh
   npm install @pnp/spfx-property-controls@1.20.0
   ```
   ```json
   "localizedResources": {
     "ListEditorWebPartStrings": "lib/webparts/listEditor/loc/{locale}.js",
     "PropertyControlStrings": "node_modules/@pnp/spfx-property-controls/lib/loc/{locale}.js"
   }
   ```
3. Die Dateien unten in `src/webparts/listEditor/` ersetzen bzw. anlegen.
4. Im generierten `ListEditorWebPart.ts` den vom Generator erzeugten Import der
   `.module.scss`-Datei und den `styles`-Bezug **entfernen** — wir rendern ohne SCSS.
5. `gulp serve --nobrowser`, in der **SharePoint-Workbench** öffnen, rechts eine
   Liste auswählen.

> Die `id`/`alias` im Manifest kannst Du so lassen, wie der Generator sie erzeugt
> hat — der `alias` muss zum Klassennamen `ListEditorWebPart` passen.

---

## `src/webparts/listEditor/ListEditorWebPart.ts`

```typescript
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'ListEditorWebPartStrings';

export interface IListEditorWebPartProps {
  listId: string;
}

interface IListItem {
  Id: number;
  Title: string;
}

export default class ListEditorWebPart extends BaseClientSideWebPart<IListEditorWebPartProps> {

  public render(): void {
    if (!this.properties.listId) {
      this.domElement.innerHTML =
        '<div style="padding:16px;">Bitte rechts in den Einstellungen eine Liste auswählen.</div>';
      return;
    }

    this.domElement.innerHTML =
      '<div style="padding:16px; font-family:Segoe UI, sans-serif;">'
      + '<h2>Listenpflege</h2>'
      + '<div style="margin:8px 0;">'
      + '<input id="neuTitel" type="text" placeholder="Neuer Titel" />'
      + ' <button id="hinzufuegen">Hinzufügen</button>'
      + '</div>'
      + '<div id="status">Lade Daten ...</div>'
      + '</div>';

    this._verdrahten();
    this._ladeListe();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('listId', {
                  label: strings.ListFieldLabel,
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
            }
          ]
        }
      ]
    };
  }

  private _baseUrl(): string {
    return this.context.pageContext.web.absoluteUrl
      + `/_api/web/lists(guid'${this.properties.listId}')`;
  }

  private _verdrahten(): void {
    const add: Element = this.domElement.querySelector('#hinzufuegen');
    add.addEventListener('click', (): void => {
      const feld: HTMLInputElement = this.domElement.querySelector('#neuTitel') as HTMLInputElement;
      const titel: string = feld.value.trim();
      if (titel.length === 0) { return; }
      this._anlegen(titel)
        .then((): void => { feld.value = ''; this._ladeListe(); })
        .catch((f: Error): void => this._zeigeFehler(f));
    });
  }

  private _ladeListe(): void {
    const status: Element = this.domElement.querySelector('#status');
    const url: string = this._baseUrl() + '/items?$select=Id,Title&$top=50';

    this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((antwort: SPHttpClientResponse): Promise<{ value: IListItem[] }> => {
        if (!antwort.ok) { throw new Error('HTTP ' + antwort.status); }
        return antwort.json();
      })
      .then((daten: { value: IListItem[] }): void => {
        if (daten.value.length === 0) {
          status.innerHTML = 'Die Liste ist leer.';
          return;
        }
        let zeilen: string = '';
        for (let i: number = 0; i < daten.value.length; i++) {
          const item: IListItem = daten.value[i];
          zeilen = zeilen
            + '<tr><td>' + item.Id + '</td>'
            + '<td>' + escape(item.Title) + '</td>'
            + '<td><button class="bearbeiten" data-id="' + item.Id
            + '" data-titel="' + escape(item.Title) + '">Bearbeiten</button> '
            + '<button class="loeschen" data-id="' + item.Id + '">Löschen</button></td></tr>';
        }
        status.innerHTML =
          '<table border="1" cellpadding="6" style="border-collapse:collapse;">'
          + '<thead><tr><th>Id</th><th>Titel</th><th></th></tr></thead>'
          + '<tbody>' + zeilen + '</tbody></table>';
        this._verdrahteZeilen();
      })
      .catch((f: Error): void => this._zeigeFehler(f));
  }

  private _verdrahteZeilen(): void {
    const bearbeiten: NodeListOf<Element> = this.domElement.querySelectorAll('.bearbeiten');
    for (let i: number = 0; i < bearbeiten.length; i++) {
      bearbeiten[i].addEventListener('click', (e: Event): void => {
        const el: HTMLElement = e.currentTarget as HTMLElement;
        const id: number = parseInt(el.getAttribute('data-id') || '', 10);
        const alt: string = el.getAttribute('data-titel') || '';
        const neu: string = window.prompt('Neuer Titel:', alt);
        if (neu === null || neu.trim().length === 0) { return; }
        this._bearbeiten(id, neu.trim())
          .then((): void => this._ladeListe())
          .catch((f: Error): void => this._zeigeFehler(f));
      });
    }

    const loeschen: NodeListOf<Element> = this.domElement.querySelectorAll('.loeschen');
    for (let i: number = 0; i < loeschen.length; i++) {
      loeschen[i].addEventListener('click', (e: Event): void => {
        const el: HTMLElement = e.currentTarget as HTMLElement;
        const id: number = parseInt(el.getAttribute('data-id') || '', 10);
        if (!window.confirm('Eintrag ' + id + ' wirklich löschen?')) { return; }
        this._loeschen(id)
          .then((): void => this._ladeListe())
          .catch((f: Error): void => this._zeigeFehler(f));
      });
    }
  }

  private _anlegen(titel: string): Promise<void> {
    return this.context.spHttpClient.post(this._baseUrl() + `/items`,
      SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify({ Title: titel })
      })
      .then((antwort: SPHttpClientResponse): void => {
        if (!antwort.ok) { throw new Error('Anlegen fehlgeschlagen: ' + antwort.status); }
      });
  }

  private _bearbeiten(id: number, titel: string): Promise<void> {
    return this.context.spHttpClient.post(this._baseUrl() + `/items(${id})`,
      SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify({ Title: titel })
      })
      .then((antwort: SPHttpClientResponse): void => {
        if (!antwort.ok) { throw new Error('Bearbeiten fehlgeschlagen: ' + antwort.status); }
      });
  }

  private _loeschen(id: number): Promise<void> {
    return this.context.spHttpClient.post(this._baseUrl() + `/items(${id})`,
      SPHttpClient.configurations.v1, {
        headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', 'odata-version': '' }
      })
      .then((antwort: SPHttpClientResponse): void => {
        if (!antwort.ok) { throw new Error('Löschen fehlgeschlagen: ' + antwort.status); }
      });
  }

  private _zeigeFehler(fehler: Error): void {
    const status: Element = this.domElement.querySelector('#status');
    if (status) { status.innerHTML = 'Fehler: ' + escape(fehler.message); }
  }
}
```

> **Sicherheit**: Jeder Wert, der in `innerHTML` geht (`Title`, auch in
> `data-titel`), läuft durch `escape()` — `escape` maskiert auch Anführungszeichen,
> deshalb ist das Attribut sicher. Ohne Escaping wäre ein Titel wie
> `"><img src=x onerror=alert(1)>` eine XSS-Lücke.

> **`window.prompt` / `window.confirm`** halten das Beispiel ohne Framework kurz.
> In React ersetzt man sie durch ein eigenes Eingabefeld bzw. einen Dialog —
> siehe [loesung-react.md](loesung-react.md).

---

## `src/webparts/listEditor/ListEditorWebPart.manifest.json`

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json",
  "id": "b3f7a1c9-2d4e-4a6b-9c8d-1e0f2a3b4c5d",
  "alias": "ListEditorWebPart",
  "componentType": "WebPart",
  "version": "*",
  "manifestVersion": 2,
  "requiresCustomScript": false,
  "preconfiguredEntries": [{
    "groupId": "5c03119e-3074-46fd-976b-c60198311f70",
    "group": { "default": "Other" },
    "title": { "default": "Listenpflege" },
    "description": { "default": "Einträge einer Liste anzeigen und pflegen" },
    "officeFabricIconFontName": "EditTable",
    "properties": {
      "listId": ""
    }
  }]
}
```

---

## `src/webparts/listEditor/loc/mystrings.d.ts`

```typescript
declare interface IListEditorWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListFieldLabel: string;
}

declare module 'ListEditorWebPartStrings' {
  const strings: IListEditorWebPartStrings;
  export = strings;
}
```

---

## `src/webparts/listEditor/loc/de-de.js`

```javascript
define([], function() {
  return {
    "PropertyPaneDescription": "Einstellungen für die Listenpflege",
    "BasicGroupName": "Allgemein",
    "ListFieldLabel": "Liste auswählen"
  };
});
```

---

## `src/webparts/listEditor/loc/en-us.js`

```javascript
define([], function() {
  return {
    "PropertyPaneDescription": "Settings for the list editor",
    "BasicGroupName": "General",
    "ListFieldLabel": "Select a list"
  };
});
```

---

## Häufige Fehler beim Einsetzen

| Symptom | Ursache / Fix |
|---|---|
| Listen-Picker erscheint nicht / Fehler zu `PropertyControlStrings` | `config/config.json` → `localizedResources` um `PropertyControlStrings` ergänzen, dann `gulp serve` neu starten. |
| `Cannot find module 'ListEditorWebPartStrings'` | `loc/mystrings.d.ts` fehlt oder Modulname weicht von `config.json` ab. Generator-Default passt. |
| Schreiben schlägt mit 403 fehl | Teilnehmer-Account braucht **Bearbeiten**-Rechte auf der Liste (Lesen reicht fürs Lesen, nicht fürs Schreiben). |
| Anlegen schlägt mit „type" / Metadaten-Fehler fehl | `odata=nometadata` in **Accept** und **Content-type** gesetzt? Dann ist kein `__metadata.type` nötig. |
| Tabelle aktualisiert sich nach Schreiben nicht | Nach `_anlegen`/`_bearbeiten`/`_loeschen` im `.then(...)` `_ladeListe()` aufrufen. |
| Weiße Fläche, Konsole zeigt 403/CORS | In der `localhost`-Workbench gibt es keine echte Liste. SharePoint-gehostete Workbench nutzen. |
| `tsc`-Fehler bei `?.` / `??` | TS 3.6 kennt das nicht. Nur `||` / `&&` / klassische `if`. |
