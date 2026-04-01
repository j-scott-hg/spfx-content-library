import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/views';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { IListInfo, IViewInfo, IListItem, IFieldDefinition } from '../models/IListItem';

export class SharePointDataService {
  private _sp: SPFI;
  private _context: WebPartContext;

  constructor(context: WebPartContext) {
    this._context = context;
    this._sp = spfi().using(SPFx(context));
  }

  /**
   * Returns a new SPFI instance scoped to the given absolute site URL,
   * falling back to the current site if siteUrl is empty.
   */
  private _getSp(siteUrl?: string): SPFI {
    if (siteUrl && siteUrl.trim() !== '' && siteUrl.trim() !== this._context.pageContext.web.absoluteUrl) {
      return spfi(siteUrl.trim()).using(SPFx(this._context));
    }
    return this._sp;
  }

  // ─── Lists & Libraries ──────────────────────────────────────────────────────

  public async getLists(siteUrl?: string): Promise<IListInfo[]> {
    try {
      const sp = this._getSp(siteUrl);
      const lists = await sp.web.lists
        .filter('Hidden eq false and IsCatalog eq false')
        .select('Id', 'Title', 'BaseTemplate')
        .orderBy('Title')();

      return lists.map((l: { Id: string; Title: string; BaseTemplate: number }) => ({
        id: l.Id,
        title: l.Title,
        baseTemplate: l.BaseTemplate,
        isDocumentLibrary: l.BaseTemplate === 101,
      }));
    } catch (err) {
      console.error('[SharePointDataService] getLists error:', err);
      return [];
    }
  }

  // ─── Views ──────────────────────────────────────────────────────────────────

  public async getViews(listId: string, siteUrl?: string): Promise<IViewInfo[]> {
    if (!listId) return [];
    try {
      const sp = this._getSp(siteUrl);

      // Fetch view metadata without expand — combining $filter and $expand on views
      // is unreliable across list types (especially document libraries).
      const views: Array<{ Id: string; Title: string; RowLimit: number; ServerRelativeUrl: string }> =
        await sp.web.lists.getById(listId).views
          .filter('Hidden eq false and PersonalView eq false')
          .select('Id', 'Title', 'RowLimit', 'ServerRelativeUrl')
          .orderBy('Title')();

      // Fetch ViewFields for each view individually.
      // view.fields is a getter returning an IViewFields (invokable collection);
      // calling it as a function returns { Items: string[] }.
      const results = await Promise.all(
        views.map(async v => {
          try {
            const viewObj = sp.web.lists.getById(listId).views.getById(v.Id);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const vf: { Items: string[] } = await (viewObj as any).fields();
            return {
              id: v.Id,
              title: v.Title,
              viewFields: vf?.Items ?? [],
              rowLimit: v.RowLimit,
              serverRelativeUrl: v.ServerRelativeUrl,
            };
          } catch {
            return {
              id: v.Id,
              title: v.Title,
              viewFields: [] as string[],
              rowLimit: v.RowLimit,
              serverRelativeUrl: v.ServerRelativeUrl,
            };
          }
        })
      );

      return results;
    } catch (err) {
      console.error('[SharePointDataService] getViews error:', err);
      return [];
    }
  }

  // ─── Fields ─────────────────────────────────────────────────────────────────

  public async getFields(listId: string, siteUrl?: string): Promise<IFieldDefinition[]> {
    if (!listId) return [];
    try {
      const sp = this._getSp(siteUrl);
      const fields = await sp.web.lists.getById(listId).fields
        .filter('Hidden eq false')
        .select('InternalName', 'Title', 'TypeAsString')
        .orderBy('Title')();

      return fields.map((f: { InternalName: string; Title: string; TypeAsString: string }) => ({
        internalName: f.InternalName,
        displayName: f.Title,
        fieldType: f.TypeAsString,
      }));
    } catch (err) {
      console.error('[SharePointDataService] getFields error:', err);
      return [];
    }
  }

  // ─── Display Form URL ───────────────────────────────────────────────────────

  /**
   * Returns the server-relative URL of the list's default display form
   * (e.g. /sites/mysite/Lists/MyList/DispForm.aspx).
   * Used to open the native SharePoint item details panel.
   */
  public async getListDispFormUrl(listId: string, siteUrl?: string): Promise<string> {
    if (!listId) return '';
    try {
      const sp = this._getSp(siteUrl);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = await (sp.web.lists.getById(listId) as any)
        .select('DefaultDisplayFormUrl')();
      return result?.DefaultDisplayFormUrl ?? '';
    } catch (err) {
      console.error('[SharePointDataService] getListDispFormUrl error:', err);
      return '';
    }
  }

  // ─── Items ──────────────────────────────────────────────────────────────────

  public async getItems(
    listId: string,
    viewFields: string[],
    itemLimit: number,
    isDocumentLibrary: boolean,
    sortField?: string,
    sortAsc?: boolean,
    siteUrl?: string,
    extraFields?: string[]  // additional fields to fetch (e.g. filter field not in view)
  ): Promise<IListItem[]> {
    if (!listId) return [];

    try {
      const sp = this._getSp(siteUrl);

      // Build the select fields — always include core fields.
      // SMTotalSize is excluded: it doesn't exist on all libraries and causes 400 errors.
      // File_x0020_Type uses the encoded internal name which SharePoint REST handles correctly.
      const coreFields = isDocumentLibrary
        ? ['Id', 'Title', 'FileLeafRef', 'FileRef', 'File_x0020_Type', 'Modified', 'Editor/Title', 'Editor/Id', 'Created', 'Author/Title', 'FSObjType', 'ContentTypeId']
        : ['Id', 'Title', 'Modified', 'Editor/Title', 'Editor/Id', 'Created', 'Author/Title', 'ContentTypeId'];

      // Virtual/UI-only view field names that cannot be used in $select
      const UNSELECTABLES = new Set([
        'LinkFilename', 'LinkTitle', 'LinkFilenameNoMenu', 'DocIcon',
        'Edit', 'SelectTitle', '_UIVersionString', 'SMTotalSize',
      ]);

      // Merge view fields, extra fields, and core fields — avoiding duplicates
      const combined = [
        ...coreFields,
        ...viewFields.filter(f => !UNSELECTABLES.has(f)),
        ...(extraFields ?? []).filter(f => !UNSELECTABLES.has(f)),
      ];
      const seen: Record<string, boolean> = {};
      const selectFields: string[] = [];
      combined.forEach(f => { if (!seen[f]) { seen[f] = true; selectFields.push(f); } });

      let query = sp.web.lists.getById(listId).items
        .select(...selectFields)
        .expand('Editor', 'Author')
        .top(itemLimit || 100);

      if (sortField) {
        query = query.orderBy(sortField, sortAsc !== false);
      } else {
        query = query.orderBy('Modified', false);
      }

      const rawItems = await query();

      return rawItems.map((item: Record<string, unknown>) => this._mapItem(item, isDocumentLibrary));
    } catch (err) {
      console.error('[SharePointDataService] getItems error:', err);
      return [];
    }
  }

  // ─── Private Mapping ────────────────────────────────────────────────────────

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _mapItem(raw: any, isDocumentLibrary: boolean): IListItem {
    const editor = raw.Editor as Record<string, unknown> | undefined;
    const author = raw.Author as Record<string, unknown> | undefined;

    const mapped: IListItem = {
      id: String(raw.Id ?? ''),
      title: String(raw.Title ?? raw.FileLeafRef ?? ''),
      modified: raw.Modified ? String(raw.Modified) : undefined,
      modifiedBy: editor ? String(editor.Title ?? '') : undefined,
      modifiedById: editor ? Number(editor.Id ?? 0) : undefined,
      created: raw.Created ? String(raw.Created) : undefined,
      createdBy: author ? String(author.Title ?? '') : undefined,
      contentTypeId: raw.ContentTypeId ? String(raw.ContentTypeId) : undefined,
    };

    if (isDocumentLibrary) {
      mapped.fileLeafRef = raw.FileLeafRef ? String(raw.FileLeafRef) : undefined;
      mapped.fileRef = raw.FileRef ? String(raw.FileRef) : undefined;
      // File_x0020_Type uses underscore encoding — access via bracket to avoid dot-notation lint
      // eslint-disable-next-line dot-notation
      mapped.fileType = raw['File_x0020_Type'] ? String(raw['File_x0020_Type']).toLowerCase() : undefined;
      mapped.isFolder = raw.FSObjType === 1 || raw.FSObjType === '1';
      mapped.name = mapped.fileLeafRef;
      const sizeRaw = raw.SMTotalSize;
      if (sizeRaw !== undefined && sizeRaw !== null) {
        mapped.size = Number(sizeRaw);
      }
    }

    // Copy all remaining fields for dynamic column rendering
    for (const key of Object.keys(raw)) {
      if (!(key in mapped) && key !== 'Editor' && key !== 'Author' && key !== 'odata.type' && key !== 'odata.id' && key !== 'odata.etag' && key !== 'odata.editLink') {
        mapped[key] = raw[key];
      }
    }

    return mapped;
  }
}
