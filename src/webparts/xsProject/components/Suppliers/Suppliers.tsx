import * as React from 'react';
import styles from './Suppliers.module.scss';

/**
 * Placeholder for the "Suppliers" item in the shared `Menu`. No backing service exists yet -
 * wire one in following the `Project`/`ProjectService` pattern (see `CLAUDE.md`) once the
 * Web API exposes a suppliers endpoint.
 */
export default class Suppliers extends React.Component {
  public render(): React.ReactElement {
    return (
      <section className={styles.suppliers}>
        <h3 className={styles.title}>Suppliers</h3>
        <div className={styles.status}>Coming soon.</div>
      </section>
    );
  }
}
