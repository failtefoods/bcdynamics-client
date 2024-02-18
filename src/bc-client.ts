import { Customer } from "./types/customers";

type BCClientConfig = {
  client_id: string;
  client_secret: string;
  scope: string;
  access_token_url: string;

  tenant_id: string;
  sandbox?: boolean;
};

/**
 * Business Central Client
 *
 * This class is used to interact with the Business Central API.
 */
export class BCClient {
  _accessToken: string | undefined;
  _accessTokenExpires: number | undefined;
  _refreshToken: string | undefined;
  _refreshTokenExpires: number | undefined;

  _clientId: string;
  _clientSecret: string;
  _scope: string;
  _grantType: string;
  _accessTokenUrl: string;

  // ID of BC app
  _tenantId: string;

  _sandbox: boolean = false;

  /**
   *
   * @param config
   */
  constructor(config: BCClientConfig) {
    this._clientId = config.client_id;
    this._clientSecret = config.client_secret;
    this._scope = config.scope;
    this._accessTokenUrl = config.access_token_url;

    this._tenantId = config.tenant_id;
    this._sandbox = config.sandbox || false;
  }

  connect() {
    console.log("BCClient connect");
  }

  /**
   *
   * @param path Path of the API endpoint
   * @returns
   */
  _buildURL(path: string) {
    if (path.startsWith("/")) {
      path = path.slice(1);
    }
    return `https://api.businesscentral.dynamics.com/v2.0/${this._tenantId}${this._sandbox == true ? "/Sandbox" : ""}/ODataV4/${path}`;
  }

  /**
   * Get a new access token
   */
  async _getNewToken() {
    console.log("BCClient getNewToken");

    const body = {
      grant_type: "client_credentials",
      client_id: this._clientId,
      client_secret: this._clientSecret,
      scope: this._scope,
    };
    try {
      const response = await fetch(this._accessTokenUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: `grant_type=client_credentials&client_id=${this._clientId}&client_secret=${this._clientSecret}&scope=${this._scope}`,
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const tokenData = await response.json();

      this._accessToken = tokenData.access_token;
      this._accessTokenExpires =
        new Date().getTime() + tokenData.expires_in * 1000;
      this._refreshToken = tokenData.refresh_token;
      this._refreshTokenExpires = tokenData.refresh_token_expires_in;
    } catch (err) {
      throw new Error(`Error getting token: ${err}`);
    }
  }

  async _validateToken() {
    console.log("BCClient _validateToken");

    if (!this._accessToken) {
      await this._getNewToken();
    }

    const now = new Date().getTime();
    if (now >= this._accessTokenExpires) {
      await this._getNewToken();
    }

    return this._accessToken;
  }

  /**
   * Fetch customers from Business Central
   * @returns List of customers
   */
  async getCustomers() {
    const path = this._buildURL("/Company('My%20Company')/Customers");

    const token = await this._validateToken();
    const response = await fetch(path, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (!response.ok) {
      throw new Error(
        `HTTP error! status: ${response.status} with message: ${response.statusText}`
      );
    }

    const json_response = await response.json();
    return json_response.value as Customer[];
  }
}
