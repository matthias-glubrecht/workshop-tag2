# Lösung — „Listenpflege" in React

> Dasselbe Web Part als React-Komponente mit State. Der Datenzugriff läuft hier
> über `SPHttpClient` (die `@pnp/sp`-Methoden aus [loesung-pnp.md](loesung-pnp.md)
> ließen sich genauso in die Handler einsetzen). Manifest und `loc/`-Dateien sind
> identisch zu [loesung-vanilla.md](loesung-vanilla.md).

## So setzt Du es ein

1. `yo @microsoft/sharepoint` → Solution `listenpflege`, Web-Part `ListEditor`,
   Framework **React**.
2. Property Controls installieren und in `config/config.json` registrieren
   (siehe [loesung-vanilla.md](loesung-vanilla.md), Schritt 2).
3. Die drei Dateien unten anlegen/ersetzen.
4. `gulp serve --nobrowser`, in der **SharePoint-Workbench** öffnen, Liste auswählen.

---

## `src/webparts/listEditor/ListEditorWebPart.ts` (Wurzel)

```typescript
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'ListEditorWebPartStrings';
import ListEditor from './components/ListEditor';
import { IListEditorProps } from './components/IListEditorProps';

export interface IListEditorWebPartProps {
  listId: string;
}

export default class ListEditorWebPart extends BaseClientSideWebPart<IListEditorWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IListEditorProps> = React.createElement(
      ListEditor,
      {
        listId: this.properties.listId,
        spHttpClient: this.context.spHttpClient,
        webUrl: this.context.pageContext.web.absoluteUrl
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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

## `src/webparts/listEditor/components/IListEditorProps.ts`

```typescript
import { SPHttpClient } from '@microsoft/sp-http';

export interface IListEditorProps {
  listId: string;
  spHttpClient: SPHttpClient;
  webUrl: string;
}
```

---

## `src/webparts/listEditor/components/ListEditor.tsx`

```tsx
import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListEditorProps } from './IListEditorProps';

interface IListItem {
  Id: number;
  Title: string;
}

interface IListEditorState {
  items: IListItem[];
  neuTitel: string;
  loading: boolean;
  error: string;
}

export default class ListEditor extends React.Component<IListEditorProps, IListEditorState> {

  constructor(props: IListEditorProps) {
    super(props);
    this.state = { items: [], neuTitel: '', loading: false, error: '' };
    this._onTitelChange = this._onTitelChange.bind(this);
    this._onAdd = this._onAdd.bind(this);
  }

  public componentDidMount(): void {
    this._load();
  }

  public componentWillReceiveProps(nextProps: IListEditorProps): void {
    if (nextProps.listId !== this.props.listId) {
      this._load(nextProps.listId);
    }
  }

  private _base(listId?: string): string {
    return this.props.webUrl + "/_api/web/lists(guid'"
      + (listId || this.props.listId) + "')";
  }

  private _load(listId?: string): void {
    const liste: string = listId || this.props.listId;
    if (!liste) { this.setState({ items: [], loading: false, error: '' }); return; }

    this.setState({ loading: true, error: '' });
    const url: string = this._base(liste) + "/items?$select=Id,Title&$top=50";

    this.props.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((antwort: SPHttpClientResponse): Promise<{ value: IListItem[] }> => {
        if (!antwort.ok) { throw new Error('HTTP ' + antwort.status); }
        return antwort.json();
      })
      .then((daten: { value: IListItem[] }): void => {
        this.setState({ items: daten.value, loading: false });
      })
      .catch((f: Error): void => {
        this.setState({ items: [], loading: false, error: f.message });
      });
  }

  private _onTitelChange(e: React.FormEvent<HTMLInputElement>): void {
    this.setState({ neuTitel: e.currentTarget.value });
  }

  private _onAdd(): void {
    const titel: string = this.state.neuTitel.trim();
    if (titel.length === 0) { return; }

    this.props.spHttpClient.post(this._base() + "/items",
      SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify({ Title: titel })
      })
      .then((antwort: SPHttpClientResponse): void => {
        if (!antwort.ok) { throw new Error('Anlegen: ' + antwort.status); }
        this.setState({ neuTitel: '' });
        this._load();
      })
      .catch((f: Error): void => this.setState({ error: f.message }));
  }

  private _onEdit(item: IListItem): void {
    const neu: string = window.prompt('Neuer Titel:', item.Title);
    if (neu === null || neu.trim().length === 0) { return; }

    this.props.spHttpClient.post(this._base() + "/items(" + item.Id + ")",
      SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify({ Title: neu.trim() })
      })
      .then((antwort: SPHttpClientResponse): void => {
        if (!antwort.ok) { throw new Error('Bearbeiten: ' + antwort.status); }
        this._load();
      })
      .catch((f: Error): void => this.setState({ error: f.message }));
  }

  private _onDelete(item: IListItem): void {
    if (!window.confirm('Eintrag ' + item.Id + ' wirklich löschen?')) { return; }

    this.props.spHttpClient.post(this._base() + "/items(" + item.Id + ")",
      SPHttpClient.configurations.v1, {
        headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', 'odata-version': '' }
      })
      .then((antwort: SPHttpClientResponse): void => {
        if (!antwort.ok) { throw new Error('Löschen: ' + antwort.status); }
        this._load();
      })
      .catch((f: Error): void => this.setState({ error: f.message }));
  }

  public render(): JSX.Element {
    if (!this.props.listId) {
      return <p>Bitte rechts in den Einstellungen eine Liste auswählen.</p>;
    }

    return (
      <div style={{ padding: 16, fontFamily: 'Segoe UI, sans-serif' }}>
        <h2>Listenpflege</h2>

        <div style={{ margin: '8px 0' }}>
          <input
            type='text'
            placeholder='Neuer Titel'
            value={this.state.neuTitel}
            onChange={this._onTitelChange}
          />
          <button onClick={this._onAdd}>Hinzufügen</button>
        </div>

        {this.state.error
          ? <p style={{ color: 'red' }}>Fehler: {this.state.error}</p>
          : null}

        {this.state.loading
          ? <p>Lade Daten ...</p>
          : <table border={1} cellPadding={6} style={{ borderCollapse: 'collapse' }}>
              <thead>
                <tr><th>Id</th><th>Titel</th><th></th></tr>
              </thead>
              <tbody>
                {this.state.items.map((item: IListItem) =>
                  <tr key={item.Id}>
                    <td>{item.Id}</td>
                    <td>{item.Title}</td>
                    <td>
                      <button onClick={() => this._onEdit(item)}>Bearbeiten</button>{' '}
                      <button onClick={() => this._onDelete(item)}>Löschen</button>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>}
      </div>
    );
  }
}
```

---

## Was die React-Variante zeigt (Sprech-Anker)

- **JSX escapet automatisch** → kein `escape()` mehr nötig.
- **`state` statt direktem DOM-Schreiben**: `setState(...)` löst das Neu-Rendern
  aus. Nach jeder Schreiboperation `this._load()` → die Tabelle aktualisiert sich.
- **`render()` gibt nie `undefined` zurück** — jeder Zweig endet mit einem Element
  oder `null` (z. B. die Fehlerzeile `... ? <p/> : null`). Das ist eine
  klassische React-15.6-Falle.
- **Kein Fragment** (`<>...</>`) — alles in einem Wurzel-`<div>`.
- **`componentWillReceiveProps`** ist in React 15.6 das normale Lifecycle, um auf
  eine geänderte `listId` zu reagieren.
- **Controlled Input**: `value` + `onChange` halten Feldinhalt und State synchron —
  derselbe Mechanismus, den größere Formular-Apps für ihre Felder nutzen.
