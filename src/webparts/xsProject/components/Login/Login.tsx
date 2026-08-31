import * as React from 'react';
import { Log } from '@microsoft/sp-core-library';

import styles from './Login.module.scss';
import type { ILoginProps } from './ILoginProps';
import { ILoginFailure, ILoginResponse, ILoginState, resolveLoginFailure } from '../../../../models/Auth';

/** Source name used for SPFx log entries emitted by this component. */
const LOG_SOURCE: string = 'Login';

/**
 * Signs the current Microsoft 365 user in to the XSProject Web API.
 *
 * The component owns nothing but view state: `ILoginService` acquires the SPFx access
 * token and posts it to `Authenticate/SPFxLogin`, so the token never reaches this class,
 * the rendered output or browser storage. The service writes the complete API response to
 * the console, and it is rendered under the success message, for development
 * verification - see `showResponse` on `ILoginProps`.
 *
 * The handshake runs on mount unless `autoLogin` is `false`, in which case the user
 * triggers it from the button.
 */
export default class Login extends React.Component<ILoginProps, ILoginState> {
  /** Guards against completing a request after the component has been unmounted. */
  private _isActive: boolean = false;

  public constructor(props: ILoginProps) {
    super(props);

    this.state = { isLoading: false, isAuthenticated: false };
  }

  public componentDidMount(): void {
    this._isActive = true;

    if (this.props.autoLogin !== false) {
      this._runLogin();
    }
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<ILoginProps> {
    const { isLoading, isAuthenticated, response, failure } = this.state;

    return (
      <section className={styles.login}>
        <h3 className={styles.title}>XSProject sign-in</h3>

        {isLoading && (
          <div className={`${styles.status} ${styles.pending}`} role="status" aria-live="polite">
            <span className={styles.spinner} aria-hidden="true" />
            <span>Signing you in&hellip;</span>
          </div>
        )}

        {!isLoading && isAuthenticated && (
          <>
            <div className={`${styles.status} ${styles.success}`} role="status" aria-live="polite">
              <span>{Login._getSuccessMessage(response)}</span>
            </div>

            {response && this.props.showResponse !== false && (
              <details className={styles.response} open>
                <summary className={styles.responseSummary}>API response</summary>
                <pre className={styles.responseBody}>{JSON.stringify(response, undefined, 2)}</pre>
              </details>
            )}
          </>
        )}

        {!isLoading && failure && (
          <div className={`${styles.status} ${styles.failure}`} role="alert">
            <div>{failure.message}</div>
            {/* A 403 will not resolve itself: the account has to be granted access first. */}
            {failure.kind !== 'authorization' && (
              <button type="button" className={styles.action} onClick={this._onSignInClick}>
                Try again
              </button>
            )}
          </div>
        )}

        {!isLoading && !isAuthenticated && !failure && (
          <button type="button" className={styles.action} onClick={this._onSignInClick}>
            Sign in
          </button>
        )}
      </section>
    );
  }

  /**
   * Runs the handshake and reduces the outcome to view state.
   *
   * Every failure - token acquisition, 401, 403 or any other API error - is classified by
   * `resolveLoginFailure`, so this method never inspects status codes itself.
   */
  private async _login(): Promise<void> {
    this.setState({ isLoading: true, isAuthenticated: false, response: undefined, failure: undefined });

    let response: ILoginResponse | undefined;
    let failure: ILoginFailure | undefined;

    try {
      response = await this.props.loginService.loginAsCurrentUser(this.props.domainUrl);
    } catch (error) {
      failure = resolveLoginFailure(error);
    }

    if (this._isActive) {
      this.setState({ isLoading: false, isAuthenticated: !failure, response, failure });
    }

    // Notified outside the try block, so a handler that throws is not reported back to the
    // user as a failed sign-in.
    if (failure) {
      if (this.props.onLoginFailed) {
        this.props.onLoginFailed(failure);
      }
    } else if (response && this.props.onLoginSucceeded) {
      this.props.onLoginSucceeded(response);
    }
  }

  /**
   * Starts the handshake from a synchronous caller.
   *
   * `_login` reports every sign-in failure through the component state; this catch only
   * covers an unexpected render-time error, so it can never surface as an unhandled
   * rejection.
   */
  private _runLogin(): void {
    this._login().catch((error: unknown): void => {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
    });
  }

  // Assigned as a property so the `this` pointer is bound without a per-render closure.
  private _onSignInClick = (): void => {
    this._runLogin();
  };

  /**
   * Names the signed-in user from whichever identity field the API populated.
   *
   * `contactpersoon` is the display name on the XSProject user record, with the account
   * name and then the mailbox as fallbacks.
   */
  private static _getSuccessMessage(response: ILoginResponse | undefined): string {
    const name: string = Login._firstNonEmpty([
      response?.contactpersoon,
      response?.userName,
      response?.email
    ]);
    const company: string = Login._firstNonEmpty([
      response?.companyName,
      response?.bedrijf,
      response?.masterCompanyName
    ]);

    if (!name) {
      return 'Signed in to XSProject.';
    }

    return company ? `Signed in to XSProject as ${name} (${company}).` : `Signed in to XSProject as ${name}.`;
  }

  /** First value that holds text once trimmed, or an empty string. */
  private static _firstNonEmpty(values: (string | undefined)[]): string {
    for (let index: number = 0; index < values.length; index++) {
      const value: string = (values[index] || '').trim();
      if (value) {
        return value;
      }
    }

    return '';
  }
}
