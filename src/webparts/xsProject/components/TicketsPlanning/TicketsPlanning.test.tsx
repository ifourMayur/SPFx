import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';

import PlanningList from './PlanningList';
import TicketList from './TicketList';
import { ApiError } from '../../../../models/ApiError';
import { IPlanningGroup, ITicketGroup } from '../../../../models/IntegrationWork';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';

function service(tickets: ITicketGroup[] | Error, plannings: IPlanningGroup[] | Error = []): IProjectIntegrationService {
  return {
    getMyTickets: async (): Promise<ITicketGroup[]> => {
      if (tickets instanceof Error) {
        throw tickets;
      }
      return tickets;
    },
    getMyPlannings: async (): Promise<IPlanningGroup[]> => {
      if (plannings instanceof Error) {
        throw plannings;
      }
      return plannings;
    }
  } as unknown as IProjectIntegrationService;
}

describe('Tickets & Planning', () => {
  let host: HTMLDivElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  });

  async function render(element: React.ReactElement): Promise<void> {
    await act(async () => {
      ReactDOM.render(element, host);
    });
  }

  it('lists the tickets per project, linking to ED Controls and keeping the text as text', async () => {
    await render(<TicketList integrationService={service([
      {
        projectId: 3,
        projectName: 'Damrak',
        message: '',
        tickets: [{
          id: 'c1',
          ticketNumber: 'A1B2C3',
          title: 'Leak',
          body: '<img src=x onerror=alert(1)>',
          state: 'created',
          createdOn: '2026-11-01T10:00:00',
          dueDate: '',
          url: 'https://web.edcontrols.com/#/projects/db/tickets/c1',
          attachmentCount: 2
        }]
      },
      { projectId: 4, projectName: 'Rokin', message: 'ED Controls tickets could not be read.', tickets: [] }
    ])} />);

    const link: HTMLAnchorElement = host.querySelector('a') as HTMLAnchorElement;
    expect(link.textContent).toBe('Leak');
    expect(link.href).toBe('https://web.edcontrols.com/#/projects/db/tickets/c1');
    expect(link.target).toBe('_blank');
    expect(host.textContent).toContain('#A1B2C3');
    expect(host.textContent).toContain('01/11/2026');
    expect(host.textContent).toContain('2 photos in ED Controls');
    expect(host.querySelector('img')).toBeNull();
    expect(host.textContent).toContain('<img src=x onerror=alert(1)>');
    expect(host.textContent).toContain('ED Controls tickets could not be read.');
  });

  it('says when there are no tickets, and offers a retry when they could not be loaded', async () => {
    await render(<TicketList integrationService={service([])} />);
    expect(host.textContent).toContain('No tickets found.');

    // A new key mounts a fresh list, so it loads again with the failing service.
    await render(<TicketList key="failing" integrationService={service(new ApiError('ED Controls is not set up for your organisation.', { status: 200 }))} />);
    expect(host.textContent).toContain('ED Controls is not set up for your organisation.');
    expect(host.textContent).toContain('Try again');
  });

  it('lists the planning items per project with their dates and days', async () => {
    await render(<PlanningList integrationService={service([], [{
      projectId: 3,
      projectName: 'Damrak',
      url: 'https://kyp.nl/project/k/planning?lang=nl',
      message: '',
      items: [{ name: 'Demolition', startDate: '2026-11-01T00:00:00', endDate: '2026-11-11T00:00:00', totalDays: 10, elapsedDays: 4 }]
    }])} />);

    expect(host.textContent).toContain('Damrak');
    expect((host.querySelector('a') as HTMLAnchorElement).href).toBe('https://kyp.nl/project/k/planning?lang=nl');
    expect(host.textContent).toContain('Start: 01/11/2026');
    expect(host.textContent).toContain('End: 11/11/2026');
    expect(host.textContent).toContain('Number of days: 4 / 10');
  });
});
