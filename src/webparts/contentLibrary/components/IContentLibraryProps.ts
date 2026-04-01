import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWebPartConfig, IItemIconOverride } from '../models/IWebPartConfig';

export interface IContentLibraryProps {
  config: IWebPartConfig;
  context: WebPartContext;
  isEditMode: boolean;
  /** Called when the user saves an icon override so the web part can persist it.
   *  Pass undefined to remove the override (reset to default). */
  onIconOverrideSave: (itemId: string, override: IItemIconOverride | undefined) => void;
  /** Called whenever the set of live category keys changes, so the property pane
   *  colour picker can stay in sync with the actual data. */
  onCategoryKeysChange?: (keys: string[]) => void;
}
