import * as React from 'react';
import { Button, makeStyles, tokens } from '@fluentui/react-components';
import { CheckmarkCircle16Regular, Video16Regular } from '@fluentui/react-icons';
import * as strings from 'PlanMyDayApplicationCustomizerStrings';
import { IMeeting, IPlannerTask } from '../../models';

const useStyles = makeStyles({
  actions: {
    display: 'flex',
    flexWrap: 'wrap',
    columnGap: tokens.spacingHorizontalS,
    rowGap: tokens.spacingVerticalS
  }
});

export interface IQuickActionsSectionProps {
  nextMeeting?: IMeeting;
  nextTask?: IPlannerTask;
  onMarkTaskComplete: (taskId: string) => void;
  busyTaskId?: string;
}

export const QuickActionsSection: React.FC<IQuickActionsSectionProps> = ({ nextMeeting, nextTask, onMarkTaskComplete, busyTaskId }) => {
  const styles = useStyles();
  const canJoinMeeting = nextMeeting?.joinUrl && nextMeeting.joinUrl.trim().length > 0;
  return (
    <div className={styles.actions}>
      <Button
        appearance="primary"
        icon={<Video16Regular />}
        disabled={!canJoinMeeting}
        as={canJoinMeeting ? 'a' : 'button'}
        href={canJoinMeeting ? nextMeeting?.joinUrl : undefined}
        target={canJoinMeeting ? '_blank' : undefined}
        rel={canJoinMeeting ? 'noopener noreferrer' : undefined}
      >
        {nextMeeting ? strings.QuickActionJoinNextMeeting : strings.QuickActionNoMeeting}
      </Button>
      <Button
        appearance="secondary"
        icon={<CheckmarkCircle16Regular />}
        disabled={!nextTask || busyTaskId === nextTask.id}
        onClick={() => nextTask && onMarkTaskComplete(nextTask.id)}
      >
        {nextTask ? strings.QuickActionMarkTaskComplete : strings.QuickActionNoTask}
      </Button>
    </div>
  );
};
