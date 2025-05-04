export interface MicrosoftTokenApiResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  token_type: string;
  scope: string;
  ext_expires_in?: number;
  id_token?: string;
}
