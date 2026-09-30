import * as React from 'react';
import {
  Button,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  MessageBar,
  MessageBarBody,
  OverlayDrawer,
  Spinner,
  makeStyles,
  tokens
} from '@fluentui/react-components';
import { Dismiss24Regular, ArrowClockwise24Regular } from '@fluentui/react-icons';
import * as strings from 'PlanMyDayApplicationCustomizerStrings';
import { IPlanMyDayData } from '../models';
import { CopilotSearchResponse } from '../graphService';
import { GreetingSection } from './sections/GreetingSection';
import { AgendaSection } from './sections/AgendaSection';
import { TasksSection } from './sections/TasksSection';
import { EmailsSection } from './sections/EmailsSection';
import { NewsSection } from './sections/NewsSection';
import { QuickActionsSection } from './sections/QuickActionsSection';
import { CopilotSection } from './sections/CopilotSection';

const useStyles = makeStyles({
  drawer: {
    width: 'min(960px, 100vw)',
    maxWidth: '100vw',
    backgroundColor: tokens.colorNeutralBackground3
  },
  header: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1
  },
  body: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalXL,
    paddingTop: tokens.spacingVerticalXL,
    paddingBottom: tokens.spacingVerticalXXXL
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(340px, 1fr))',
    gap: tokens.spacingHorizontalL,
    alignItems: 'start'
  },
  column: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalL,
    minWidth: 0
  },
  centered: {
    display: 'flex',
    justifyContent: 'center',
    paddingTop: tokens.spacingVerticalXXXL
  }
});

export interface IPlanMyDayPanelProps {
  isOpen: boolean;
  onClose: () => void;
  onRefresh: () => void;
  isLoading: boolean;
  error?: string;
  data?: IPlanMyDayData;
  isMockMode: boolean;
  onMarkTaskComplete: (taskId: string) => void;
  busyTaskId?: string;
  onAskCopilot: (prompt: string) => void;
  isCopilotLoading: boolean;
  copilotResult?: CopilotSearchResponse;
  copilotError?: string;
}

export const PlanMyDayPanel: React.FC<IPlanMyDayPanelProps> = (props) => {
  const {
    isOpen, onClose, onRefresh, isLoading, error, data, isMockMode,
    onMarkTaskComplete, busyTaskId, onAskCopilot, isCopilotLoading, copilotResult, copilotError
  } = props;
  const styles = useStyles();

  return (
    <OverlayDrawer
      open={isOpen}
      onOpenChange={(_ev, state) => !state.open && onClose()}
      position="end"
      size="large"
      className={styles.drawer}
    >
      <DrawerHeader className={styles.header}>
        <DrawerHeaderTitle
          action={
            <>
              <Button appearance="subtle" icon={<ArrowClockwise24Regular />} onClick={onRefresh} disabled={isLoading} title={strings.RefreshButtonLabel} />
              <Button appearance="subtle" icon={<Dismiss24Regular />} onClick={onClose} title={strings.CloseButtonLabel} />
            </>
          }
        >
          {strings.PanelTitle}
        </DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody className={styles.body}>
        {isMockMode && (
          <MessageBar intent="info">
            <MessageBarBody>{strings.MockDataBanner}</MessageBarBody>
          </MessageBar>
        )}
        {isLoading && (
          <div className={styles.centered}>
            <Spinner label={strings.LoadingLabel} />
          </div>
        )}
        {!isLoading && error && (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}
        {!isLoading && !error && data && (
          <>
            <GreetingSection
              displayName={data.displayName}
              actions={
                <QuickActionsSection
                  nextMeeting={data.meetings[0]}
                  nextTask={data.tasks[0]}
                  onMarkTaskComplete={onMarkTaskComplete}
                  busyTaskId={busyTaskId}
                />
              }
            />
            <div className={styles.grid}>
              <div className={styles.column}>
                <NewsSection news={data.news} />
                <CopilotSection
                  onAsk={onAskCopilot}
                  isLoading={isCopilotLoading}
                  result={copilotResult}
                  error={copilotError}
                />
              </div>
              <div className={styles.column}>
                <EmailsSection emails={data.emails} />
                <AgendaSection meetings={data.meetings} />
                <TasksSection tasks={data.tasks} onMarkComplete={onMarkTaskComplete} busyTaskId={busyTaskId} />
              </div>
            </div>
          </>
        )}
      </DrawerBody>
    </OverlayDrawer>
  );
};
