import moment from 'moment-timezone';
import { CoreSetup } from 'src/core/public';
import { writeFile, read } from 'xlsx';
import { IncompatibleActionError } from '../../../../src/plugins/ui_actions/public';
import type { UiActionsActionDefinition as ActionDefinition } from '../../../../src/plugins/ui_actions/public';
import type { ISearchEmbeddable, SavedSearch } from '../../../../src/plugins/discover/public';
import {
  loadSharingDataHelpers,
  SEARCH_EMBEDDABLE_TYPE,
} from '../../../../src/plugins/discover/public';
import { IEmbeddable, ViewMode } from '../../../../src/plugins/embeddable/public';
import { API_GENERATE_IMMEDIATE } from '../../common/constants';
import type { JobParamsDownloadCSV } from '../../../../x-pack/plugins/reporting/server/export_types/csv_searchsource_immediate/types';

interface ActionContext {
  embeddable: ISearchEmbeddable;
}

function isSavedSearchEmbeddable(
  embeddable: IEmbeddable | ISearchEmbeddable
): embeddable is ISearchEmbeddable {
  return embeddable.type === SEARCH_EMBEDDABLE_TYPE;
}

export class GetExcelPanelAction implements ActionDefinition<ActionContext> {
  readonly id = 'downloadExcelReport';

  private isDownloading: boolean;
  private core: any;

  constructor(core: CoreSetup) {
    this.isDownloading = false;
    this.core = core;
  }

  public isCompatible = async (context: ActionContext) => {
    const { embeddable } = context;
    return embeddable.getInput().viewMode !== ViewMode.EDIT && isSavedSearchEmbeddable(embeddable);
  };

  public getIconType(): string {
    return 'document';
  }

  public getDisplayName(): string {
    return 'Export as Excel';
  }

  public async getSearchSource(savedSearch: SavedSearch, embeddable: ISearchEmbeddable) {
    const { getSharingData } = await loadSharingDataHelpers();
    return await getSharingData(
      savedSearch.searchSource,
      savedSearch, // TODO: get unsaved state (using embeddable.searchScope): https://github.com/elastic/kibana/issues/43977
      embeddable.services
    );
  }

  public execute = async (context: ActionContext) => {
    const { embeddable } = context;

    if (!isSavedSearchEmbeddable(embeddable)) {
      throw new IncompatibleActionError();
    }

    if (this.isDownloading) {
      return;
    }

    const savedSearch = embeddable.getSavedSearch();
    const { columns, getSearchSource } = await this.getSearchSource(savedSearch, embeddable);

    const kibanaTimezone = this.core.uiSettings.get('dateFormat:tz');
    const browserTimezone = kibanaTimezone === 'Browser' ? moment.tz.guess() : kibanaTimezone;

    const immediateJobParams: JobParamsDownloadCSV = {
      searchSource: getSearchSource(),
      columns,
      browserTimezone,
      title: savedSearch.title,
    };

    const body = JSON.stringify(immediateJobParams);

    this.isDownloading = true;

    this.core.notifications.toasts.addSuccess({
      title: `Excel Download Started`,
      text: `Your Excel will download momentarily.`,
      'data-test-subj': 'csvDownloadStarted',
    });

    await this.core.http
      .post(`${API_GENERATE_IMMEDIATE}`, { body })
      .then((rawResponse: string) => {
        this.isDownloading = false;

        const workbook = read(rawResponse, { type: 'string', raw: true });
        writeFile(workbook, `${embeddable.getSavedSearch().title}.xlsx`, { type: 'binary' });
      })
      .catch(this.onGenerationFail.bind(this));
  };

  private onGenerationFail(error: Error) {
    this.isDownloading = false;
    this.core.notifications.toasts.addDanger({
      title: `Excel download failed`,
      text: `We couldn't generate your Excel at this time.`,
      'data-test-subj': 'downloadExcelFail',
    });
  }
}
