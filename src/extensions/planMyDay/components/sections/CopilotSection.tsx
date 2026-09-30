import * as React from 'react';
import { Button, Caption1, Link, MessageBar, MessageBarBody, Spinner, Text, Textarea, makeStyles, tokens } from '@fluentui/react-components';
import { Sparkle20Regular } from '@fluentui/react-icons';
import * as strings from 'PlanMyDayApplicationCustomizerStrings';
import { CopilotSearchHit, CopilotSearchResponse } from '../../graphService';
import { formatTemplate } from '../../utils';
import { SectionCard } from './SectionCard';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalS,
    padding: tokens.spacingHorizontalXS
  },
  textarea: {
    width: '100%'
  },
  footer: {
    display: 'flex',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalS
  },
  count: {
    color: tokens.colorNeutralForeground3
  },
  results: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalXS,
    maxHeight: '420px',
    overflowY: 'auto'
  },
  hit: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalXXS,
    padding: tokens.spacingHorizontalS,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    borderLeft: `3px solid ${tokens.colorBrandStroke1}`
  },
  title: {
    display: 'block',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap'
  },
  meta: {
    color: tokens.colorNeutralForeground3
  },
  preview: {
    color: tokens.colorNeutralForeground2,
    display: '-webkit-box',
    WebkitLineClamp: 3,
    WebkitBoxOrient: 'vertical',
    overflow: 'hidden'
  }
});

// Microsoft Search wraps matched terms in <c0> markers; render as plain text rather than HTML.
function toPlainText(value: unknown): string {
  if (typeof value !== 'string') {
    return '';
  }
  let text: string = '';
  let insideTag: boolean = false;
  for (const character of value) {
    if (character === '<') {
      insideTag = true;
    } else if (character === '>') {
      insideTag = false;
    } else if (!insideTag) {
      text += character;
    }
  }
  return text.trim();
}

interface ICopilotResultItem {
  title: string;
  url?: string;
  subtitle: string;
  preview: string;
}

function toResultItem(hit: CopilotSearchHit): ICopilotResultItem {
  const extractText: string = (hit.extracts || []).map(extract => toPlainText(extract?.text)).join(' ');
  return {
    title: toPlainText(hit.resourceMetadata?.title) || hit.webUrl || '',
    url: typeof hit.webUrl === 'string' ? hit.webUrl : undefined,
    subtitle: [toPlainText(hit.resourceMetadata?.author), hit.resourceType]
      .filter(part => !!part)
      .join(' \u00b7 '),
    preview: toPlainText(hit.preview) || extractText
  };
}

export interface ICopilotSectionProps {
  onAsk: (prompt: string) => void;
  isLoading: boolean;
  result?: CopilotSearchResponse;
  error?: string;
}

export const CopilotSection: React.FC<ICopilotSectionProps> = ({ onAsk, isLoading, result, error }) => {
  const styles = useStyles();
  const [prompt, setPrompt] = React.useState<string>('');

  const submit = (): void => {
    if (prompt.trim().length > 0) {
      onAsk(prompt.trim());
    }
  };

  const items: ICopilotResultItem[] = React.useMemo(
    () => (Array.isArray(result?.searchHits) ? result!.searchHits.map(toResultItem) : []),
    [result]
  );
  const totalCount: number = typeof result?.totalCount === 'number' ? result.totalCount : items.length;

  return (
    <SectionCard title={strings.CopilotSectionTitle} icon={<Sparkle20Regular />} count={result ? totalCount : undefined}>
      <div className={styles.root}>
        <Textarea
          className={styles.textarea}
          value={prompt}
          placeholder={strings.CopilotPlaceholder}
          onChange={(_ev, data) => setPrompt(data.value)}
          resize="vertical"
        />
        <div className={styles.footer}>
          <Button appearance="primary" onClick={submit} disabled={isLoading || prompt.trim().length === 0}>
            {isLoading ? strings.CopilotLoading : strings.CopilotSubmit}
          </Button>
          {isLoading && <Spinner size="tiny" />}
        </div>
        {!isLoading && error && (
          <MessageBar intent="warning">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}
        {!isLoading && !error && result && items.length === 0 && (
          <Caption1 className={styles.count}>{strings.CopilotNoResults}</Caption1>
        )}
        {!isLoading && !error && items.length > 0 && (
          <>
            <Caption1 className={styles.count}>
              {formatTemplate(strings.CopilotResultsCount || '{0} / {1}', String(items.length), String(totalCount))}
            </Caption1>
            <div className={styles.results}>
              {items.map((item, index) => (
                <div key={`${item.url || item.title}-${index}`} className={styles.hit}>
                  {item.url ? (
                    <Link
                      href={item.url}
                      target="_blank"
                      rel="noopener noreferrer"
                      className={styles.title}
                      title={item.title}
                    >
                      {item.title}
                    </Link>
                  ) : (
                    <Text weight="semibold" className={styles.title} title={item.title}>{item.title}</Text>
                  )}
                  {item.subtitle && <Caption1 className={styles.meta}>{item.subtitle}</Caption1>}
                  {item.preview && <Text size={200} className={styles.preview}>{item.preview}</Text>}
                </div>
              ))}
            </div>
          </>
        )}
      </div>
    </SectionCard>
  );
};
