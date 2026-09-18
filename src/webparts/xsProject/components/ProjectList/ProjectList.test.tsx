/**
 * Render tests for the project overview table.
 *
 * The site a session works in is chosen before the app renders, by `SharePointSites`, and
 * from then on only its title is visible - in the shell's site bar. These cover that the
 * listing also names the site's URL, which is what tells the user where a project's
 * folders are actually being created.
 *
 * `react-dom/test-utils` and the jsdom environment are already available, so this needs no
 * new dependency. Every `ReactDOM.render` here is unmounted in `afterEach`.
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';

import ProjectList from './ProjectList';
import { IProjectListProps } from './IProjectListProps';
import { SAMPLE_PROJECT_ROWS } from '../../../../models/ProjectListRow';

const SITE_URL: string = 'https://ifourtechnolab.sharepoint.com/sites/Projects';

function props(overrides?: Partial<IProjectListProps>): IProjectListProps {
  return {
    rows: SAMPLE_PROJECT_ROWS,
    canAddEdit: true,
    siteUrl: SITE_URL,
    onAddProject: (): void => undefined,
    onEditProject: (): void => undefined,
    onToggleFavorite: (): void => undefined,
    onDismissNotification: (): void => undefined,
    ...overrides
  };
}

describe('ProjectList', () => {
  let host: HTMLDivElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  });

  function render(overrides?: Partial<IProjectListProps>): void {
    act(() => {
      ReactDOM.render(<ProjectList {...props(overrides)} />, host);
    });
  }

  it('names the URL of the site the session is working in', () => {
    render();

    expect(host.textContent).toContain(SITE_URL);
  });

  /**
   * The URL is shown for the user to read and to copy into a browser, so it is rendered as
   * a link to the site rather than as text that has to be retyped.
   */
  it('links to the site', () => {
    render();

    const link: HTMLAnchorElement | null = host.querySelector(`a[href="${SITE_URL}"]`);

    expect(link).not.toBeNull();
    expect(link!.textContent).toBe(SITE_URL);
  });

  /**
   * The listing is rendered by `AppShell` only once a site has been chosen, so an empty URL
   * should not happen - but an empty label reading "Site URL:" with nothing after it would
   * be worse than no label at all.
   */
  it('omits the line entirely when no site URL is known', () => {
    render({ siteUrl: '' });

    expect(host.textContent).not.toContain('Site URL');
  });
});
