import * as React from 'react';
import styles from './Document.module.scss';

/**
 * Placeholder for the "Document" item under the Project submenu, opened from the shared
 * `Menu`. No backing service exists yet - wire one in following the `Project`/`ProjectService`
 * pattern (see `CLAUDE.md`) once the Web API exposes a documents endpoint.
 */
export default class Document extends React.Component {
  public render(): React.ReactElement {
    return (
      <section className={styles.document}>
        <h3 className={styles.title}>Document</h3>
        <div className={styles.status}>Coming soon.</div>
      </section>
    );
  }
}
