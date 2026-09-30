import * as React from 'react';
import { Button, Caption1, Text } from '@fluentui/react-components';
import { CalendarLtr20Regular, Clock16Regular, Video16Regular } from '@fluentui/react-icons';
import * as strings from 'PlanMyDayApplicationCustomizerStrings';
import { IMeeting } from '../../models';
import { SectionCard } from './SectionCard';
import { useListStyles } from './sectionStyles';
import { formatTime } from '../../utils';

export interface IAgendaSectionProps {
  meetings: IMeeting[];
}

export const AgendaSection: React.FC<IAgendaSectionProps> = ({ meetings }) => {
  const styles = useListStyles();
  return (
    <SectionCard title={strings.AgendaSectionTitle} icon={<CalendarLtr20Regular />} count={meetings.length}>
      {meetings.length === 0 && <Caption1 className={styles.empty}>{strings.AgendaEmpty}</Caption1>}
      <div className={styles.list}>
        {meetings.map(meeting => (
          <div key={meeting.id} className={styles.row}>
            <div className={styles.rowMain}>
              <Text weight="semibold" className={styles.truncate} title={meeting.subject}>{meeting.subject}</Text>
              <Caption1 className={styles.meta}>
                <Clock16Regular />
                {formatTime(meeting.start)} &ndash; {formatTime(meeting.end)}
              </Caption1>
            </div>
            {meeting.joinUrl && (
              <Button
                size="small"
                appearance="secondary"
                icon={<Video16Regular />}
                as="a"
                href={meeting.joinUrl}
                target="_blank"
                rel="noopener noreferrer"
              >
                {strings.AgendaJoinButton}
              </Button>
            )}
          </div>
        ))}
      </div>
    </SectionCard>
  );
};
