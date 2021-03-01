import { UiActionsActionDefinition as ActionDefinition } from '../../../../src/plugins/ui_actions/public';
import { ISearchEmbeddable } from '../../../../src/plugins/discover/public';
import { ViewMode } from '../../../../src/plugins/embeddable/public';

interface ActionContext {
  embeddable: ISearchEmbeddable;
}

export class GetExcelPanelAction implements ActionDefinition<ActionContext> {
  readonly id = 'downloadExcelReport';

  public isCompatible = async (context: ActionContext) => {
    const { embeddable } = context;
    return embeddable.getInput().viewMode !== ViewMode.EDIT && embeddable.type === 'search';
  };

  public getIconType(): string {
    return 'document';
  }

  public getDisplayName(): string {
    return 'Export as Excel';
  }

  public execute = async (context: ActionContext) => {};
}
