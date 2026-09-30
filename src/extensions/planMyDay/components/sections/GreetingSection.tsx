import * as React from 'react';
import { Avatar, Caption1, makeStyles, Title2, tokens } from '@fluentui/react-components';
import * as strings from 'PlanMyDayApplicationCustomizerStrings';
import { getGreetingKey, formatTemplate } from '../../utils';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalL,
    padding: `${tokens.spacingVerticalXL} ${tokens.spacingHorizontalXL}`,
    borderRadius: tokens.borderRadiusXLarge,
    backgroundImage: `linear-gradient(120deg, ${tokens.colorBrandBackground2} 0%, ${tokens.colorNeutralBackground1} 70%)`,
    border: `1px solid ${tokens.colorNeutralStroke2}`
  },
  identity: {
    display: 'flex',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalL,
    minWidth: 0
  },
  text: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalXXS,
    minWidth: 0
  },
  greeting: {
    margin: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis'
  },
  date: {
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em'
  }
});

export interface IGreetingSectionProps {
  displayName: string;
  actions?: React.ReactNode;
}

export const GreetingSection: React.FC<IGreetingSectionProps> = ({ displayName, actions }) => {
  const styles = useStyles();
  const template: string = strings[getGreetingKey()] as string;
  const today: string = new Date().toLocaleDateString([], {
    weekday: 'long',
    month: 'long',
    day: 'numeric'
  });

  return (
    <div className={styles.root}>
      <div className={styles.identity}>
        <Avatar name={displayName} size={56} color="colorful" />
        <div className={styles.text}>
          <Caption1 className={styles.date}>{today}</Caption1>
          <Title2 as="h2" className={styles.greeting}>{formatTemplate(template, displayName)}</Title2>
        </div>
      </div>
      {actions}
    </div>
  );
};
