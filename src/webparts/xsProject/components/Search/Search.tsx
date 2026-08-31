import * as React from 'react';
import styles from './Search.module.scss';

/**
 * Placeholder for the "Search" item in the shared `Menu`. No backing service exists yet -
 * wire one in following the `Project`/`ProjectService` pattern (see `CLAUDE.md`) once the
 * Web API exposes a search endpoint.
 */
export default class Search extends React.Component {
  public render(): React.ReactElement {
    return (
      <section className={styles.search}>
        <h3 className={styles.title}>Search</h3>
        <div className={styles.status}>Coming soon.</div>
      </section>
    );
  }
}
