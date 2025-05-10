/**
 * Configuration interface for Microsoft Outlook OAuth settings
 */
export interface MicrosoftOutlookConfig {
  /**
   * The client id for the Microsoft Outlook OAuth settings
   */
  clientId: string;
  /**
   * The client secret for the Microsoft Outlook OAuth settings
   */
  clientSecret: string;
  /**
   * The path of the redirect uri. e.g. auth/microsoft/callback
   */
  redirectPath: string;
  /**
   * The base url of the backend. e.g. https://dev.dashboard.checkfirstapp.com
   */
  backendBaseUrl: string;
  /**
   * The base path of the backend. e.g. api/v1
   */
  basePath?: string;
}
