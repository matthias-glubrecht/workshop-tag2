# Lösung — „Listenpflege" mit `@pnp/sp` (ohne Framework)

> Dasselbe Web Part wie [loesung-vanilla.md](loesung-vanilla.md), aber der
> Datenzugriff läuft über `@pnp/sp` statt über `SPHttpClient`. Gut geeignet, um
> beide Wege **nebeneinander** zu zeigen: Die UI-Verdrahtung ist identisch, nur
> die vier Datenmethoden unterscheiden sich.

## So setzt Du es ein

1. Basis wie in [loesung-vanilla.md](loesung-vanilla.md) (Scaffold, Property
   Controls, Manifest, `loc/`-Dateien sind **identisch**).
2. `@pnp/sp` in **Version 1** installieren — passend zu SPFx 1.4.1 / TS 3.6:
   ```pwsh
   npm install @pnp/sp@1.3.11
   ```
   > **Wichtig**: Modernes `@pnp/sp` v3/v4 braucht neueres TypeScript und Node und
   > läuft auf dieser eingefrorenen Plattform **nicht**. Die `1.3.11` ist hier
   > Voraussetzung, kein Detail.
3. Nur die Datei `ListEditorWebPart.ts` unten verwenden (Manifest/`loc/` wie vanilla).

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
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';

import * as strings from 'ListEditorWebPartStrings';

export interface IListEditorWebPartProps {
  listId: string;
}

interface IListItem {
  Id: number;
  Title: string;
}

export default class ListEditorWebPart extends BaseClientSideWebPart<IListEditorWebPartProps> {

  // Einmalig: pnp den SPFx-Kontext bekanntgeben (Auth, Site-URL).
  protected onInit(): Promise<void> {
    return super.onInit().then((): void => {
      sp.setup({ spfxContext: this.context });
    });
  }

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

  // ----- Datenzugriff über @pnp/sp -----

  private _ladeListe(): void {
    const status: Element = this.domElement.querySelector('#status');
    sp.web.lists.getById(this.properties.listId)
      .items.select('Id', 'Title').top(50).get()
      .then((items: IListItem[]): void => {
        if (items.length === 0) { status.innerHTML = 'Die Liste ist leer.'; return; }
        let zeilen: string = '';
        for (let i: number = 0; i < items.length; i++) {
          const item: IListItem = items[i];
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

  private _anlegen(titel: string): Promise<void> {
    return sp.web.lists.getById(this.properties.listId)
      .items.add({ Title: titel })
      .then((): void => { return; });
  }

  private _bearbeiten(id: number, titel: string): Promise<void> {
    return sp.web.lists.getById(this.properties.listId)
      .items.getById(id).update({ Title: titel })
      .then((): void => { return; });
  }

  private _loeschen(id: number): Promise<void> {
    return sp.web.lists.getById(this.properties.listId)
      .items.getById(id).delete();
  }

  // ----- UI-Verdrahtung (identisch zur SPHttpClient-Variante) -----

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

  private _zeigeFehler(fehler: Error): void {
    const status: Element = this.domElement.querySelector('#status');
    if (status) { status.innerHTML = 'Fehler: ' + escape(fehler.message); }
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
}
```

---

## Was im direkten Vergleich auffällt

| | `SPHttpClient` | `@pnp/sp` |
|---|---|---|
| Lesen | URL bauen, `get`, `antwort.json()`, `.value` auspacken | `...items.select('Id','Title').top(50).get()` — Ergebnis ist direkt das Array |
| Anlegen | `post` + zwei Header + `JSON.stringify` | `...items.add({ Title })` |
| Bearbeiten | `post` + `X-HTTP-Method: MERGE` + `IF-MATCH` | `...items.getById(id).update({ Title })` |
| Löschen | `post` + `X-HTTP-Method: DELETE` + `IF-MATCH` | `...items.getById(id).delete()` |
| Setup | keiner (`spHttpClient` ist da) | einmal `sp.setup({ spfxContext: this.context })` |
| Abhängigkeit | keine zusätzliche | `@pnp/sp@1.3.11` |

**Einordnung**: `@pnp/sp` nimmt einem Header, Metadaten und das Auspacken der
Antwort ab — bei vielen Operationen weniger Code und weniger Fehlerquellen. Dafür
kommt eine Abhängigkeit dazu, die auf der eingefrorenen Plattform exakt auf v1
festgenagelt sein muss. `SPHttpClient` macht jeden Schritt sichtbar und braucht
nichts extra. Für ein, zwei Aufrufe reicht `SPHttpClient`; bei datenintensiven
Lösungen lohnt sich `@pnp/sp`.
