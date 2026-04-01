export interface IListItem {
  id: string;
  title: string;
  name?: string; // document library file name
  fileRef?: string; // server-relative URL for documents
  fileLeafRef?: string; // file name with extension
  fileType?: string; // extension e.g. 'docx', 'xlsx', 'pdf'
  contentTypeId?: string;
  modified?: string;
  modifiedBy?: string;
  modifiedById?: number;
  created?: string;
  createdBy?: string;
  size?: number;
  isFolder?: boolean;
  folderChildCount?: number;
  description?: string;
  // dynamic additional fields from the selected view
  [key: string]: unknown;
}

export interface IFieldDefinition {
  internalName: string;
  displayName: string;
  fieldType: string; // 'Text', 'Note', 'DateTime', 'User', 'Lookup', 'Choice', 'Boolean', 'Number', 'Currency', 'URL', 'Computed', 'File'
  isHidden?: boolean;
  isReadOnly?: boolean;
}

export interface IListInfo {
  id: string;
  title: string;
  baseTemplate: number; // 100 = list, 101 = document library
  isDocumentLibrary: boolean;
  defaultViewUrl?: string;
  fields?: IFieldDefinition[];
}

export interface IViewInfo {
  id: string;
  title: string;
  viewFields: string[];
  query?: string;
  rowLimit?: number;
  serverRelativeUrl?: string;
}
