import * as React from 'react';
import { Caption1, Link } from '@fluentui/react-components';
import { Mail20Regular } from '@fluentui/react-icons';
import * as strings from 'PlanMyDayApplicationCustomizerStrings';
import { IImportantEmail } from '../../models';
import { SectionCard } from './SectionCard';
import { useListStyles } from './sectionStyles';
import { formatTime } from '../../utils';

export interface IEmailsSectionProps {
  emails: IImportantEmail[];
}

export const EmailsSection: React.FC<IEmailsSectionProps> = ({ emails }) => {
  const styles = useListStyles();
  return (
    <SectionCard title={strings.EmailsSectionTitle} icon={<Mail20Regular />} count={emails.length}>
      {emails.length === 0 && <Caption1 className={styles.empty}>{strings.EmailsEmpty}</Caption1>}
      <div className={styles.list}>
        {emails.map(email => (
          <div key={email.id} className={styles.row}>
            <div className={styles.rowMain}>
              <Link
                href={email.webLink}
                target="_blank"
                rel="noopener noreferrer"
                className={styles.truncate}
                title={email.subject}
              >
                {email.subject}
              </Link>
              <Caption1 className={styles.meta}>{email.from}</Caption1>
            </div>
            <Caption1 className={styles.meta}>{formatTime(email.receivedDateTime)}</Caption1>
          </div>
        ))}
      </div>
    </SectionCard>
  );
};
