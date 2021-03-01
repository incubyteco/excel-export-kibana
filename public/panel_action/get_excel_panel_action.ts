import _ from 'lodash';
import moment from 'moment-timezone';
import dateMath from '@elastic/datemath';
import { CoreSetup } from 'src/core/public';
import { writeFile, read } from 'xlsx';
import {
  IncompatibleActionError,
  UiActionsActionDefinition as ActionDefinition,
} from '../../../../src/plugins/ui_actions/public';
import { ISearchEmbeddable, SEARCH_EMBEDDABLE_TYPE } from '../../../../src/plugins/discover/public';
import { IEmbeddable, ViewMode } from '../../../../src/plugins/embeddable/public';
import { API_GENERATE_IMMEDIATE } from '../../common/constants';

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

  public getSearchRequestBody({ searchEmbeddable }: { searchEmbeddable: any }) {
    const adapters = searchEmbeddable.getInspectorAdapters();
    if (!adapters) {
      return {};
    }

    if (adapters.requests.requests.length === 0) {
      return {};
    }

    return searchEmbeddable.getSavedSearch().searchSource.getSearchRequestBody();
  }

  public execute = async (context: ActionContext) => {
    const { embeddable } = context;

    if (!isSavedSearchEmbeddable(embeddable)) {
      throw new IncompatibleActionError();
    }

    if (this.isDownloading) {
      return;
    }

    const {
      timeRange: { to, from },
    } = embeddable.getInput();

    const searchEmbeddable = embeddable;
    const searchRequestBody = await this.getSearchRequestBody({ searchEmbeddable });
    const state = _.pick(searchRequestBody, ['sort', 'docvalue_fields', 'query']);
    const kibanaTimezone = this.core.uiSettings.get('dateFormat:tz');

    const id = `search:${embeddable.getSavedSearch().id}`;
    const timezone = kibanaTimezone === 'Browser' ? moment.tz.guess() : kibanaTimezone;
    const fromTime = dateMath.parse(from);
    const toTime = dateMath.parse(to, { roundUp: true });

    if (!fromTime || !toTime) {
      return this.onGenerationFail(
        new Error(`Invalid time range: From: ${fromTime}, To: ${toTime}`)
      );
    }

    const body = JSON.stringify({
      timerange: {
        min: fromTime.format(),
        max: toTime.format(),
        timezone,
      },
      state,
    });

    this.isDownloading = true;

    this.core.notifications.toasts.addSuccess({
      title: `Excel Download Started`,
      text: `Your Excel will download momentarily.`,
      'data-test-subj': 'csvDownloadStarted',
    });

    await this.core.http
      .post(`${API_GENERATE_IMMEDIATE}/${id}`, { body })
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
