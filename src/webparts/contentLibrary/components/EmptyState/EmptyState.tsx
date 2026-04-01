import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from '../../styles/ContentLibrary.module.scss';

export interface IEmptyStateProps {
  message?: string;
  hasSearch?: boolean;
  hasFilter?: boolean;
}

const EmptyState: React.FC<IEmptyStateProps> = ({
  message,
  hasSearch,
  hasFilter,
}) => {
  const defaultMessage = hasSearch || hasFilter
    ? 'No items match your current search or filter. Try adjusting your criteria.'
    : (message || 'No items found.');

  return (
    <div className={styles.emptyState} role="status" aria-live="polite">
      <div className={styles.emptyStateIcon} aria-hidden="true">
        <Icon iconName="SearchIssue" />
      </div>
      <div className={styles.emptyStateTitle}>No results</div>
      <div className={styles.emptyStateMessage}>{defaultMessage}</div>
    </div>
  );
};

export default EmptyState;
