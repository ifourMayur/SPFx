import * as React from 'react';

import styles from './TicketsPlanning.module.scss';
import { getErrorMessage } from '../../../../models/ApiError';
import { ITicket, ITicketGroup, formatWorkDate } from '../../../../models/IntegrationWork';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';

export interface ITicketListProps {
  integrationService: IProjectIntegrationService;
}

interface ITicketListState {
  isLoading: boolean;
  error?: string;
  groups: ITicketGroup[];
}

/**
 * The user's open ED Controls tickets, per project - the reference's Tickets tab
 * (`_TicketListOverView.cshtml`): number, created date, state, title (linking to the ticket in
 * ED Controls) and text. Photos are counted rather than shown; see `models/IntegrationWork.ts`.
 */
export default class TicketList extends React.Component<ITicketListProps, ITicketListState> {
  private _isActive: boolean = false;

  public constructor(props: ITicketListProps) {
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

  public render(): React.ReactElement<ITicketListProps> {
    const { isLoading, error, groups } = this.state;

    if (isLoading) {
      return <div className={styles.status} role="status" aria-live="polite">Loading tickets&hellip;</div>;
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
      return <div className={styles.status}>No tickets found.</div>;
    }

    return (
      <>
        {groups.map((group: ITicketGroup) => (
          <div key={group.projectId}>
            <div className={styles.projectName}>{group.projectName}</div>
            {group.message && <div className={styles.error} role="alert">{group.message}</div>}
            <ul className={styles.list}>
              {group.tickets.map((ticket: ITicket) => (
                <li key={ticket.id} className={styles.item}>
                  <div className={styles.itemHeader}>
                    {ticket.url ? (
                      <a href={ticket.url} target="_blank" rel="noopener noreferrer">{ticket.title || ticket.ticketNumber}</a>
                    ) : (
                      <span>{ticket.title || ticket.ticketNumber}</span>
                    )}
                    <span className={styles.meta}>
                      #{ticket.ticketNumber} · {ticket.state === 'created' ? 'Created' : 'Started'}
                      {ticket.createdOn && <> · {formatWorkDate(ticket.createdOn)}</>}
                      {ticket.dueDate && <> · due {formatWorkDate(ticket.dueDate)}</>}
                      {ticket.attachmentCount > 0 && (
                        <> · {ticket.attachmentCount} photo{ticket.attachmentCount === 1 ? '' : 's'} in ED Controls</>
                      )}
                    </span>
                  </div>
                  {ticket.body && <p className={styles.body}>{ticket.body}</p>}
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
    this.props.integrationService.getMyTickets().then(
      (groups: ITicketGroup[]): void => {
        if (this._isActive) {
          this.setState({ isLoading: false, groups });
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({ isLoading: false, error: getErrorMessage(error, 'Your ED Controls tickets could not be loaded.') });
        }
      }
    );
  }

  private _onRetry = (): void => {
    this._load();
  };
}
