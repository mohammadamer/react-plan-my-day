import * as React from 'react';
import { Button, makeStyles, tokens } from '@fluentui/react-components';
import * as strings from 'PlanMyDayApplicationCustomizerStrings';

const useStyles = makeStyles({
  bar: {
    display: 'flex',
    justifyContent: 'flex-end',
    alignItems: 'center',
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`
  }
});

export interface ILauncherBarProps {
  onOpen: () => void;
}

export const LauncherBar: React.FC<ILauncherBarProps> = ({ onOpen }) => {
  const styles = useStyles();
  return (
    <div className={styles.bar}>
      <Button appearance="primary" onClick={onOpen}>{strings.LauncherButtonLabel}</Button>
    </div>
  );
};
