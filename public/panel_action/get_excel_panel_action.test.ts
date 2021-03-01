import { GetExcelPanelAction } from './get_excel_panel_action';

describe('GetExcelReportPanelAction', () => {
  it('should return action display name', () => {
    const action = new GetExcelPanelAction();
    expect(action.getDisplayName()).toBe('Export as Excel');
  });

  it('should return icon type', () => {
    const action = new GetExcelPanelAction();
    expect(action.getIconType()).toBe('document');
  });

  it('should be compatible with search', () => {
    const context = {
      embeddable: {
        type: 'search',
        getInput: () => ({
          viewMode: 'list',
        }),
      },
    } as any;

    const action = new GetExcelPanelAction();

    return action.isCompatible(context).then((data) => {
      expect(data).toBe(true);
    });
  });
});
