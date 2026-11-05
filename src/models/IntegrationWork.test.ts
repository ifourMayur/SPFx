import { formatWorkDate, toPlanningGroups, toTicketGroups } from './IntegrationWork';

describe('IntegrationWork', () => {
  it('reads ticket groups, keeping only https links', () => {
    const groups = toTicketGroups([
      {
        projectId: 3,
        projectName: 'Damrak',
        message: null,
        tickets: [
          { id: 'c1', ticketNumber: 'A1', title: 'Leak', body: '<b>wet</b>', state: 'Created', createdOn: '2026-11-01T10:00:00', dueDate: null, url: 'https://web.edcontrols.com/#/x', attachmentCount: 2 },
          { id: 'c2', title: 'Insecure link', url: 'http://web.edcontrols.com/#/x' }
        ]
      },
      'not a group'
    ]);

    expect(groups).toHaveLength(1);
    expect(groups[0].tickets[0]).toMatchObject({ state: 'created', body: '<b>wet</b>', attachmentCount: 2, dueDate: '' });
    expect(groups[0].tickets[1].url).toBe('');
  });

  it('reads planning groups and their items', () => {
    expect(toPlanningGroups([
      {
        projectId: 3,
        projectName: 'Damrak',
        url: 'https://kyp.nl/project/k/planning?lang=nl',
        message: 'KYP did not accept the token.',
        items: [{ name: 'Demolition', startDate: '2026-11-01T00:00:00', endDate: '2026-11-11T00:00:00', totalDays: 10, elapsedDays: 4 }]
      }
    ])).toEqual([
      {
        projectId: 3,
        projectName: 'Damrak',
        url: 'https://kyp.nl/project/k/planning?lang=nl',
        message: 'KYP did not accept the token.',
        items: [{ name: 'Demolition', startDate: '2026-11-01T00:00:00', endDate: '2026-11-11T00:00:00', totalDays: 10, elapsedDays: 4 }]
      }
    ]);
    expect(toPlanningGroups(null)).toEqual([]);
  });

  it('shows dates as dd/mm/yyyy without a time-zone shift', () => {
    expect(formatWorkDate('2026-11-01T23:30:00Z')).toBe('01/11/2026');
    expect(formatWorkDate('')).toBe('');
  });
});
