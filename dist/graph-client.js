import logger from "./logger.js";
import { refreshAccessToken } from "./lib/microsoft-auth.js";
import { encode as toonEncode } from "@toon-format/toon";
import { getCloudEndpoints } from "./cloud-config.js";
import { getRequestTokens } from "./request-context.js";
class GraphClient {
  constructor(authManager, secrets, outputFormat = "json") {
    this.outputFormat = "json";
    this.authManager = authManager;
    this.secrets = secrets;
    this.outputFormat = outputFormat;
  }
  async makeRequest(endpoint, options = {}) {
    const contextTokens = getRequestTokens();
    let accessToken = options.accessToken ?? contextTokens?.accessToken ?? await this.authManager.getToken();
    const refreshToken = options.refreshToken ?? contextTokens?.refreshToken;
    if (!accessToken) {
      throw new Error("No access token available");
    }
    try {
      let response = await this.performRequest(endpoint, accessToken, options);
      if (response.status === 401 && refreshToken) {
        const newTokens = await this.refreshAccessToken(refreshToken);
        accessToken = newTokens.accessToken;
        response = await this.performRequest(endpoint, accessToken, options);
      }
      if (response.status === 403) {
        const errorText = await response.text();
        if (errorText.includes("scope") || errorText.includes("permission")) {
          throw new Error(
            `Microsoft Graph API scope error: ${response.status} ${response.statusText} - ${errorText}. This tool requires organization mode. Please restart with --org-mode flag.`
          );
        }
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${errorText}`
        );
      }
      if (!response.ok) {
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${await response.text()}`
        );
      }
      const text = await response.text();
      let result;
      if (text === "") {
        result = { message: "OK!" };
      } else {
        try {
          result = JSON.parse(text);
        } catch {
          result = { message: "OK!", rawResponse: text };
        }
      }
      if (options.includeHeaders) {
        const etag = response.headers.get("ETag") || response.headers.get("etag");
        if (result && typeof result === "object" && !Array.isArray(result)) {
          return {
            ...result,
            _etag: etag || "no-etag-found"
          };
        }
      }
      return result;
    } catch (error) {
      logger.error("Microsoft Graph API request failed:", error);
      throw error;
    }
  }
  async refreshAccessToken(refreshToken) {
    const tenantId = this.secrets.tenantId || "common";
    const clientId = this.secrets.clientId;
    const clientSecret = this.secrets.clientSecret;
    if (clientSecret) {
      logger.info("GraphClient: Refreshing token with confidential client");
    } else {
      logger.info("GraphClient: Refreshing token with public client");
    }
    const response = await refreshAccessToken(
      refreshToken,
      clientId,
      clientSecret,
      tenantId,
      this.secrets.cloudType
    );
    return {
      accessToken: response.access_token,
      refreshToken: response.refresh_token
    };
  }
  async performRequest(endpoint, accessToken, options) {
    const cloudEndpoints = getCloudEndpoints(this.secrets.cloudType);
    const url = `${cloudEndpoints.graphApi}/v1.0${endpoint}`;
    logger.info(`[GRAPH CLIENT] Final URL being sent to Microsoft: ${url}`);
    const headers = {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...options.headers
    };
    return fetch(url, {
      method: options.method || "GET",
      headers,
      body: options.body
    });
  }
  serializeData(data, outputFormat, pretty = false) {
    if (outputFormat === "toon") {
      try {
        return toonEncode(data);
      } catch (error) {
        logger.warn(`Failed to encode as TOON, falling back to JSON: ${error}`);
        return JSON.stringify(data, null, pretty ? 2 : void 0);
      }
    }
    return JSON.stringify(data, null, pretty ? 2 : void 0);
  }
  async graphRequest(endpoint, options = {}) {
    try {
      logger.info(`Calling ${endpoint} with options: ${JSON.stringify(options)}`);
      const result = await this.makeRequest(endpoint, options);
      return this.formatJsonResponse(result, options.rawResponse, options.excludeResponse);
    } catch (error) {
      logger.error(`Error in Graph API request: ${error}`);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: error.message }) }],
        isError: true
      };
    }
  }
  formatJsonResponse(data, rawResponse = false, excludeResponse = false) {
    if (excludeResponse) {
      return {
        content: [{ type: "text", text: this.serializeData({ success: true }, this.outputFormat) }]
      };
    }
    if (data && typeof data === "object" && "_headers" in data) {
      const responseData = data;
      const meta = {};
      if (responseData._etag) {
        meta.etag = responseData._etag;
      }
      if (responseData._headers) {
        meta.headers = responseData._headers;
      }
      if (rawResponse) {
        return {
          content: [
            { type: "text", text: this.serializeData(responseData.data, this.outputFormat) }
          ],
          _meta: meta
        };
      }
      if (responseData.data === null || responseData.data === void 0) {
        return {
          content: [
            { type: "text", text: this.serializeData({ success: true }, this.outputFormat) }
          ],
          _meta: meta
        };
      }
      const removeODataProps2 = (obj) => {
        if (typeof obj === "object" && obj !== null) {
          Object.keys(obj).forEach((key) => {
            if (key.startsWith("@odata.")) {
              delete obj[key];
            } else if (typeof obj[key] === "object") {
              removeODataProps2(obj[key]);
            }
          });
        }
      };
      removeODataProps2(responseData.data);
      return {
        content: [
          { type: "text", text: this.serializeData(responseData.data, this.outputFormat, true) }
        ],
        _meta: meta
      };
    }
    if (rawResponse) {
      return {
        content: [{ type: "text", text: this.serializeData(data, this.outputFormat) }]
      };
    }
    if (data === null || data === void 0) {
      return {
        content: [{ type: "text", text: this.serializeData({ success: true }, this.outputFormat) }]
      };
    }
    const removeODataProps = (obj) => {
      if (typeof obj === "object" && obj !== null) {
        Object.keys(obj).forEach((key) => {
          if (key.startsWith("@odata.")) {
            delete obj[key];
          } else if (typeof obj[key] === "object") {
            removeODataProps(obj[key]);
          }
        });
      }
    };
    removeODataProps(data);
    return {
      content: [{ type: "text", text: this.serializeData(data, this.outputFormat, true) }]
    };
  }
}
var graph_client_default = GraphClient;
export {
  graph_client_default as default
};
