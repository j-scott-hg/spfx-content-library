import { IFieldDefinition } from '../models/IListItem';

/**
 * Maps view field internal names to display-friendly column definitions
 * using the full field list from the list/library.
 */
export interface IColumnDef {
  internalName: string;
  displayName: string;
  fieldType: string;
  width?: number;
}

export class ViewMapper {
  /**
   * Given a list of view field internal names and the full field definitions,
   * returns ordered column definitions for rendering.
   */
  public static mapViewFields(
    viewFieldNames: string[],
    allFields: IFieldDefinition[]
  ): IColumnDef[] {
    const fieldMap = new Map<string, IFieldDefinition>(
      allFields.map(f => [f.internalName, f])
    );

    // Normalize common computed field names to their actual internal names
    const normalize = (name: string): string => {
      const aliases: Record<string, string> = {
        'LinkFilename': 'FileLeafRef',
        'LinkTitle': 'Title',
        'LinkFilenameNoMenu': 'FileLeafRef',
        'DocIcon': 'File_x0020_Type',
        '_UIVersionString': '_UIVersionString',
      };
      return aliases[name] ?? name;
    };

    const columns: IColumnDef[] = [];

    for (const rawName of viewFieldNames) {
      const internalName = normalize(rawName);

      // Skip purely UI fields
      if (internalName === 'DocIcon' || internalName === 'Edit' || internalName === 'SelectTitle') {
        continue;
      }

      const field = fieldMap.get(internalName);
      if (field) {
        columns.push({
          internalName: field.internalName,
          displayName: field.displayName,
          fieldType: field.fieldType,
        });
      } else {
        // Field not in the field list (e.g. computed) — include with best-guess display name
        columns.push({
          internalName: internalName,
          displayName: ViewMapper._humanize(internalName),
          fieldType: 'Text',
        });
      }
    }

    return columns;
  }

  private static _humanize(internalName: string): string {
    return internalName
      .replace(/_x0020_/g, ' ')
      .replace(/([A-Z])/g, ' $1')
      .replace(/^_/, '')
      .trim();
  }
}
