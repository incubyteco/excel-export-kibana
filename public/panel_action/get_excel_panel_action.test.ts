import { GetExcelPanelAction } from './get_excel_panel_action';
import { IncompatibleActionError } from '../../../../src/plugins/ui_actions/public/actions';
import { read, writeFile } from 'xlsx';
jest.mock('xlsx');

describe('GetExcelReportPanelAction', () => {
  let core: any;
  let context: any;

  beforeEach(() => {
    core = {
      http: {
        post: jest.fn().mockImplementation(() => Promise.resolve(true)),
      },
      notifications: {
        toasts: {
          addSuccess: jest.fn(),
          addDanger: jest.fn(),
        },
      },
      uiSettings: {
        get: () => 'Browser',
      },
    } as any;

    context = {
      embeddable: {
        type: 'search',
        getSavedSearch: () => ({ id: 'saved_search' }),
        getTitle: () => `Test Saved Search `,
        getInspectorAdapters: () => null,
        getInput: () => ({
          viewMode: 'list',
          timeRange: {
            to: 'now',
            from: 'now-7d',
          },
        }),
      },
    } as any;
  });

  it('should return action display name', () => {
    const action = new GetExcelPanelAction(core);
    expect(action.getDisplayName()).toBe('Export as Excel');
  });

  it('should return icon type', () => {
    const action = new GetExcelPanelAction(core);
    expect(action.getIconType()).toBe('document');
  });

  it('should be compatible with search', () => {
    const action = new GetExcelPanelAction(core);

    return action.isCompatible(context).then((data) => {
      expect(data).toBe(true);
    });
  });

  it('should download for valid context', async () => {
    const action = new GetExcelPanelAction(core);
    await action.execute(context);

    expect(core.http.post).toHaveBeenCalled();
    expect(read).toHaveBeenCalled();
    expect(writeFile).toHaveBeenCalled();
  });

  it('should throw Error on incompatible embeddable', () => {
    const badContext = {
      ...context,
      embeddable: {
        type: 'visualization',
        getInput: () => ({
          viewMode: 'list',
        }),
      },
    } as any;

    const action = new GetExcelPanelAction(core);

    return expect(action.execute(badContext)).rejects.toThrow(IncompatibleActionError);
  });

  it('shows a notification when it successfully starts', async () => {
    const panel = new GetExcelPanelAction(core);
    await panel.execute(context);

    expect(core.notifications.toasts.addSuccess).toHaveBeenCalled();
    expect(core.notifications.toasts.addDanger).not.toHaveBeenCalled();
  });

  it('shows a notification when it fails', async () => {
    const coreFails: any = {
      ...core,
      http: {
        post: jest.fn().mockImplementation(() => Promise.reject('No more ram!')),
      },
    };

    const panel = new GetExcelPanelAction(coreFails);
    await panel.execute(context);

    expect(core.notifications.toasts.addDanger).toHaveBeenCalled();
  });
});
