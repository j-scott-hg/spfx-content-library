/**
 * Maps file extensions to Fluent UI icon names and color classes.
 * Uses the standard Microsoft file type color conventions.
 */

export interface IFileIconInfo {
  iconName: string;
  colorClass: string;
  label: string;
}

const FILE_ICON_MAP: Record<string, IFileIconInfo> = {
  docx: { iconName: 'WordDocument', colorClass: 'fileIcon--word', label: 'Word Document' },
  doc:  { iconName: 'WordDocument', colorClass: 'fileIcon--word', label: 'Word Document' },
  xlsx: { iconName: 'ExcelDocument', colorClass: 'fileIcon--excel', label: 'Excel Workbook' },
  xls:  { iconName: 'ExcelDocument', colorClass: 'fileIcon--excel', label: 'Excel Workbook' },
  csv:  { iconName: 'ExcelDocument', colorClass: 'fileIcon--excel', label: 'CSV File' },
  pptx: { iconName: 'PowerPointDocument', colorClass: 'fileIcon--powerpoint', label: 'PowerPoint Presentation' },
  ppt:  { iconName: 'PowerPointDocument', colorClass: 'fileIcon--powerpoint', label: 'PowerPoint Presentation' },
  pdf:  { iconName: 'PDF', colorClass: 'fileIcon--pdf', label: 'PDF Document' },
  one:  { iconName: 'OneNoteLogo', colorClass: 'fileIcon--onenote', label: 'OneNote Notebook' },
  onetoc2: { iconName: 'OneNoteLogo', colorClass: 'fileIcon--onenote', label: 'OneNote Notebook' },
  vsdx: { iconName: 'VisioDocument', colorClass: 'fileIcon--visio', label: 'Visio Diagram' },
  vsd:  { iconName: 'VisioDocument', colorClass: 'fileIcon--visio', label: 'Visio Diagram' },
  mpp:  { iconName: 'ProjectDocument', colorClass: 'fileIcon--project', label: 'Project File' },
  zip:  { iconName: 'ZipFolder', colorClass: 'fileIcon--archive', label: 'ZIP Archive' },
  rar:  { iconName: 'ZipFolder', colorClass: 'fileIcon--archive', label: 'Archive' },
  txt:  { iconName: 'TextDocument', colorClass: 'fileIcon--text', label: 'Text File' },
  msg:  { iconName: 'Mail', colorClass: 'fileIcon--email', label: 'Email Message' },
  eml:  { iconName: 'Mail', colorClass: 'fileIcon--email', label: 'Email Message' },
  png:  { iconName: 'FileImage', colorClass: 'fileIcon--image', label: 'Image' },
  jpg:  { iconName: 'FileImage', colorClass: 'fileIcon--image', label: 'Image' },
  jpeg: { iconName: 'FileImage', colorClass: 'fileIcon--image', label: 'Image' },
  gif:  { iconName: 'FileImage', colorClass: 'fileIcon--image', label: 'Image' },
  svg:  { iconName: 'FileImage', colorClass: 'fileIcon--image', label: 'SVG Image' },
  mp4:  { iconName: 'Video', colorClass: 'fileIcon--video', label: 'Video' },
  mov:  { iconName: 'Video', colorClass: 'fileIcon--video', label: 'Video' },
  avi:  { iconName: 'Video', colorClass: 'fileIcon--video', label: 'Video' },
  mp3:  { iconName: 'MusicNote', colorClass: 'fileIcon--audio', label: 'Audio' },
  wav:  { iconName: 'MusicNote', colorClass: 'fileIcon--audio', label: 'Audio' },
  html: { iconName: 'Globe', colorClass: 'fileIcon--web', label: 'Web Page' },
  htm:  { iconName: 'Globe', colorClass: 'fileIcon--web', label: 'Web Page' },
  aspx: { iconName: 'Globe', colorClass: 'fileIcon--web', label: 'Web Page' },
};

const FOLDER_ICON: IFileIconInfo = {
  iconName: 'FabricFolder',
  colorClass: 'fileIcon--folder',
  label: 'Folder',
};

const DEFAULT_ICON: IFileIconInfo = {
  iconName: 'Document',
  colorClass: 'fileIcon--generic',
  label: 'File',
};

export function getFileIconInfo(fileType?: string, isFolder?: boolean): IFileIconInfo {
  if (isFolder) return FOLDER_ICON;
  if (!fileType) return DEFAULT_ICON;
  return FILE_ICON_MAP[fileType.toLowerCase()] ?? DEFAULT_ICON;
}

export function getFileIconColorHex(fileType?: string, isFolder?: boolean): string {
  if (isFolder) return '#FFB900';
  const ext = fileType?.toLowerCase();
  const colorMap: Record<string, string> = {
    docx: '#185ABD', doc: '#185ABD',
    xlsx: '#107C41', xls: '#107C41', csv: '#107C41',
    pptx: '#C43E1C', ppt: '#C43E1C',
    pdf: '#D93025',
    one: '#7719AA', onetoc2: '#7719AA',
    vsdx: '#3955A3', vsd: '#3955A3',
    mpp: '#217346',
    png: '#0078D4', jpg: '#0078D4', jpeg: '#0078D4', gif: '#0078D4', svg: '#0078D4',
    mp4: '#881798', mov: '#881798', avi: '#881798',
    mp3: '#038387', wav: '#038387',
    zip: '#8A8886', rar: '#8A8886',
    txt: '#8A8886',
    msg: '#0078D4', eml: '#0078D4',
    html: '#0078D4', htm: '#0078D4', aspx: '#0078D4',
  };
  return ext ? (colorMap[ext] ?? '#8A8886') : '#8A8886';
}
