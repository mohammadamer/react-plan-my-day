import * as React from 'react';
import { Badge, Caption1, makeStyles, tokens } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    boxShadow: tokens.shadow2,
    overflow: 'hidden'
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2
  },
  icon: {
    display: 'flex',
    color: tokens.colorBrandForeground1,
    fontSize: '20px'
  },
  title: {
    flexGrow: 1,
    margin: 0,
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1
  },
  body: {
    padding: tokens.spacingHorizontalS
  }
});

export interface ISectionCardProps {
  title: string;
  icon?: React.ReactNode;
  count?: number;
  children: React.ReactNode;
}

export const SectionCard: React.FC<ISectionCardProps> = ({ title, icon, count, children }) => {
  const styles = useStyles();
  return (
    <section className={styles.root}>
      <div className={styles.header}>
        {icon && <span className={styles.icon}>{icon}</span>}
        <Caption1 as="h3" className={styles.title}>{title}</Caption1>
        {count !== undefined && count > 0 && <Badge appearance="tint" color="brand" size="small">{count}</Badge>}
      </div>
      <div className={styles.body}>{children}</div>
    </section>
  );
};
