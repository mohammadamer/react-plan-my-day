import * as React from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { IPlanMyDayData } from '../models';
import { getMockData } from '../mockData';
import { askCopilot, CopilotSearchResponse, loadPlanMyDayData, markTaskComplete } from '../graphService';
import { isMockMode } from '../utils';
import { LauncherBar } from './LauncherBar';
import { PlanMyDayPanel } from './PlanMyDayPanel';
import * as strings from 'PlanMyDayApplicationCustomizerStrings';

export interface IAppProps {
  context: ApplicationCustomizerContext;
}

export const App: React.FC<IAppProps> = ({ context }) => {
  const mockMode: boolean = React.useMemo(() => isMockMode(), []);

  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [data, setData] = React.useState<IPlanMyDayData | undefined>(undefined);
  const [busyTaskId, setBusyTaskId] = React.useState<string | undefined>(undefined);
  const [isCopilotLoading, setIsCopilotLoading] = React.useState<boolean>(false);
  const [copilotResult, setCopilotResult] = React.useState<CopilotSearchResponse | undefined>(undefined);
  const [copilotError, setCopilotError] = React.useState<string | undefined>(undefined);

  const loadData = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(undefined);
    try {
      const result: IPlanMyDayData = mockMode ? getMockData() : await loadPlanMyDayData(context);
      setData(result);
    } catch {
      setError(strings.ErrorLabel);
    } finally {
      setIsLoading(false);
    }
  }, [context, mockMode]);

  const handleOpen = (): void => {
    setIsOpen(true);
    if (!data) {
      loadData().catch(() => {
        /* errors are surfaced via the error state */
      });
    }
  };

  const handleMarkTaskComplete = async (taskId: string): Promise<void> => {
    setBusyTaskId(taskId);
    try {
      if (!mockMode) {
        await markTaskComplete(context, taskId);
      }
      setData(previous => previous && { ...previous, tasks: previous.tasks.filter(task => task.id !== taskId) });
    } catch {
      setError(strings.ErrorLabel);
    } finally {
      setBusyTaskId(undefined);
    }
  };

  const handleAskCopilot = async (prompt: string): Promise<void> => {
    setIsCopilotLoading(true);
    setCopilotError(undefined);
    setCopilotResult(undefined);
    try {
      const result: CopilotSearchResponse = mockMode
        ? {
            totalCount: 1,
            searchHits: [
              {
                webUrl: 'https://contoso.sharepoint.com/sites/demo/Shared%20Documents/sample.docx',
                preview: `Sample Copilot result for: "${prompt}"`,
                resourceType: 'driveItem',
                resourceMetadata: { title: 'Sample document.docx', author: 'Adele Vance' }
              }
            ]
          }
        : await askCopilot(context, prompt);
      setCopilotResult(result);
    } catch (copilotException) {
      console.error('[PlanMyDay] Copilot search failed', copilotException);
      setCopilotError(strings.CopilotLicenseError);
    } finally {
      setIsCopilotLoading(false);
    }
  };

  return (
    <FluentProvider theme={webLightTheme}>
      <LauncherBar onOpen={handleOpen} />
      <PlanMyDayPanel
        isOpen={isOpen}
        onClose={() => setIsOpen(false)}
        onRefresh={() => loadData().catch(() => { /* errors are surfaced via the error state */ })}
        isLoading={isLoading}
        error={error}
        data={data}
        isMockMode={mockMode}
        onMarkTaskComplete={(taskId) => handleMarkTaskComplete(taskId).catch(() => { /* errors are surfaced via the error state */ })}
        busyTaskId={busyTaskId}
        onAskCopilot={(prompt) => handleAskCopilot(prompt).catch(() => { /* errors are surfaced via the copilotError state */ })}
        isCopilotLoading={isCopilotLoading}
        copilotResult={copilotResult}
        copilotError={copilotError}
      />
    </FluentProvider>
  );
};
