import logger from "../logger.js";
import { getCloudEndpoints } from "../cloud-config.js";
const microsoftBearerTokenAuthMiddleware = (req, res, next) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    res.status(401).json({ error: "Missing or invalid access token" });
    return;
  }
  const accessToken = authHeader.substring(7);
  const refreshToken = req.headers["x-microsoft-refresh-token"] || "";
  req.microsoftAuth = {
    accessToken,
    refreshToken
  };
  next();
};
async function exchangeCodeForToken(code, redirectUri, clientId, clientSecret, tenantId = "common", codeVerifier, cloudType = "global") {
  const cloudEndpoints = getCloudEndpoints(cloudType);
  const params = new URLSearchParams({
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    client_id: clientId
  });
  if (clientSecret) {
    params.append("client_secret", clientSecret);
  }
  if (codeVerifier) {
    params.append("code_verifier", codeVerifier);
  }
  const response = await fetch(`${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: params
  });
  if (!response.ok) {
    const error = await response.text();
    logger.error(`Failed to exchange code for token: ${error}`);
    throw new Error(`Failed to exchange code for token: ${error}`);
  }
  return response.json();
}
async function refreshAccessToken(refreshToken, clientId, clientSecret, tenantId = "common", cloudType = "global") {
  const cloudEndpoints = getCloudEndpoints(cloudType);
  const params = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    client_id: clientId
  });
  if (clientSecret) {
    params.append("client_secret", clientSecret);
  }
  const response = await fetch(`${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: params
  });
  if (!response.ok) {
    const error = await response.text();
    logger.error(`Failed to refresh token: ${error}`);
    throw new Error(`Failed to refresh token: ${error}`);
  }
  return response.json();
}
export {
  exchangeCodeForToken,
  microsoftBearerTokenAuthMiddleware,
  refreshAccessToken
};
