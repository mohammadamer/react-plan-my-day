import * as React from 'react';
import { Button, Caption1, Text } from '@fluentui/react-components';
import { CheckmarkCircle16Regular, TaskListSquareLtr20Regular } from '@fluentui/react-icons';
import * as strings from 'PlanMyDayApplicationCustomizerStrings';
import { IPlannerTask } from '../../models';
import { SectionCard } from './SectionCard';
import { useListStyles } from './sectionStyles';

export interface ITasksSectionProps {
  tasks: IPlannerTask[];
  onMarkComplete: (taskId: string) => void;
  busyTaskId?: string;
}

export const TasksSection: React.FC<ITasksSectionProps> = ({ tasks, onMarkComplete, busyTaskId }) => {
  const styles = useListStyles();
  return (
    <SectionCard title={strings.TasksSectionTitle} icon={<TaskListSquareLtr20Regular />} count={tasks.length}>
      {tasks.length === 0 && <Caption1 className={styles.empty}>{strings.TasksEmpty}</Caption1>}
      <div className={styles.list}>
        {tasks.map(task => (
          <div key={task.id} className={styles.row}>
            <div className={styles.rowMain}>
              <Text weight="semibold" className={styles.truncate} title={task.title}>{task.title}</Text>
              {task.planTitle && <Caption1 className={styles.meta}>{task.planTitle}</Caption1>}
            </div>
            <Button
              size="small"
              appearance="secondary"
              icon={<CheckmarkCircle16Regular />}
              disabled={busyTaskId === task.id}
              onClick={() => onMarkComplete(task.id)}
            >
              {strings.TasksMarkComplete}
            </Button>
          </div>
        ))}
      </div>
    </SectionCard>
  );
};
