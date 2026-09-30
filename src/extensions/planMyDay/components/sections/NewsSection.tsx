import * as React from 'react';
import { Caption1, Link, makeStyles, tokens } from '@fluentui/react-components';
import { News20Regular } from '@fluentui/react-icons';
import * as strings from 'PlanMyDayApplicationCustomizerStrings';
import { INewsPost } from '../../models';
import { SectionCard } from './SectionCard';
import { useListStyles } from './sectionStyles';

const useStyles = makeStyles({
  thumbnail: {
    width: '56px',
    height: '40px',
    objectFit: 'cover',
    borderRadius: tokens.borderRadiusMedium,
    flexShrink: 0
  },
  item: {
    display: 'flex',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalS,
    minWidth: 0,
    flexGrow: 1
  }
});

export interface INewsSectionProps {
  news: INewsPost[];
}

export const NewsSection: React.FC<INewsSectionProps> = ({ news }) => {
  const styles = useListStyles();
  const localStyles = useStyles();
  return (
    <SectionCard title={strings.NewsSectionTitle} icon={<News20Regular />} count={news.length}>
      {news.length === 0 && <Caption1 className={styles.empty}>{strings.NewsEmpty}</Caption1>}
      <div className={styles.list}>
        {news.map(post => (
          <div key={post.id} className={styles.row}>
            <div className={localStyles.item}>
              {post.imageUrl && <img src={post.imageUrl} alt="" className={localStyles.thumbnail} />}
              <div className={styles.rowMain}>
                <Link
                  href={post.webUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className={styles.truncate}
                  title={post.title}
                >
                  {post.title}
                </Link>
                <Caption1 className={styles.meta}>{post.publishedDateTime.toLocaleDateString()}</Caption1>
              </div>
            </div>
          </div>
        ))}
      </div>
    </SectionCard>
  );
};
