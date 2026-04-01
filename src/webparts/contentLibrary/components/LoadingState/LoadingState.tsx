import * as React from 'react';
import { ItemDisplayStyle } from '../../models/IWebPartConfig';
import styles from '../../styles/ContentLibrary.module.scss';

export interface ILoadingStateProps {
  displayStyle: ItemDisplayStyle;
  itemCount?: number;
}

const LoadingState: React.FC<ILoadingStateProps> = ({
  displayStyle,
  itemCount = 6,
}) => {
  if (displayStyle === 'card-grid' || displayStyle === 'icon-grid') {
    return (
      <div
        className={styles.cardGrid}
        style={{ '--grid-cols': 3 } as React.CSSProperties}
        aria-label="Loading content"
        aria-busy="true"
      >
        {new Array(itemCount).fill(0).map((_: number, i: number) => (
          <div key={i} className={styles.shimmerCard} />
        ))}
      </div>
    );
  }

  if (displayStyle === 'tile-grid') {
    return (
      <div
        className={styles.tileGrid}
        style={{ '--grid-cols': 4 } as React.CSSProperties}
        aria-label="Loading content"
        aria-busy="true"
      >
        {new Array(itemCount).fill(0).map((_: number, i: number) => (
          <div key={i} className={styles.shimmerCard} style={{ height: 100 }} />
        ))}
      </div>
    );
  }

  // Table / list shimmer rows
  return (
    <div className={styles.loadingContainer} aria-label="Loading content" aria-busy="true">
      {new Array(itemCount).fill(0).map((_: number, i: number) => (
        <div key={i} className={styles.shimmerRow}>
          <div className={styles.shimmerIcon} />
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 6 }}>
            <div className={styles.shimmerTitle} />
            <div className={styles.shimmerMeta} />
          </div>
        </div>
      ))}
    </div>
  );
};

export default LoadingState;
