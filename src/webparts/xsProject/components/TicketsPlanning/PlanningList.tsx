import * as React from 'react';

import styles from './TicketsPlanning.module.scss';
import { getErrorMessage } from '../../../../models/ApiError';
import { IPlanningGroup, IPlanningItem, formatWorkDate } from '../../../../models/IntegrationWork';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';

export interface IPlanningListProps {
  integrationService: IProjectIntegrationService;
}

interface IPlanningListState {
  isLoading: boolean;
  error?: string;
  groups: IPlanningGroup[];
}

/**
 * The KYP planning items the user is responsible for, per project - the reference's Planning
 * tab (`_PlanningOverview.cshtml`): start, end, days elapsed of the total, and the name linking
 * to the project's planning in KYP.
 */
export default class PlanningList extends React.Component<IPlanningListProps, IPlanningListState> {
  private _isActive: boolean = false;

  public constructor(props: IPlanningListProps) {
    super(props);

    this.state = { isLoading: true, groups: [] };
  }

  public componentDidMount(): void {
    this._isActive = true;
    this._load();
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<IPlanningListProps> {
    const { isLoading, error, groups } = this.state;

    if (isLoading) {
      return <div className={styles.status} role="status" aria-live="polite">Loading planning&hellip;</div>;
    }
    if (error) {
      return (
        <div className={styles.error} role="alert">
          {error}
          <button type="button" className={styles.button} onClick={this._onRetry}>Try again</button>
        </div>
      );
    }
    if (groups.length === 0) {
      return <div className={styles.status}>No planning found.</div>;
    }

    return (
      <>
        {groups.map((group: IPlanningGroup) => (
          <div key={group.projectId}>
            <div className={styles.projectName}>{group.projectName}</div>
            {group.message && <div className={styles.error} role="alert">{group.message}</div>}
            <ul className={styles.list}>
              {group.items.map((item: IPlanningItem, index: number) => (
                <li key={`${group.projectId}-${index}`} className={styles.item}>
                  <div className={styles.itemHeader}>
                    {group.url ? (
                      <a href={group.url} target="_blank" rel="noopener noreferrer">{item.name}</a>
                    ) : (
                      <span>{item.name}</span>
                    )}
                    <span className={styles.meta}>
                      Start: {formatWorkDate(item.startDate)} · End: {formatWorkDate(item.endDate)} · Number of days:{' '}
                      {item.elapsedDays} / {item.totalDays}
                    </span>
                  </div>
                </li>
              ))}
            </ul>
          </div>
        ))}
      </>
    );
  }

  private _load(): void {
    this.setState({ isLoading: true, error: undefined });
    this.props.integrationService.getMyPlannings().then(
      (groups: IPlanningGroup[]): void => {
        if (this._isActive) {
          this.setState({ isLoading: false, groups });
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({ isLoading: false, error: getErrorMessage(error, 'Your KYP planning could not be loaded.') });
        }
      }
    );
  }

  private _onRetry = (): void => {
    this._load();
  };
}
