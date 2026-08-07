/**
 * Device Code Flow for Microsoft OAuth2
 *
 * Enables authentication without browser redirect — ideal for
 * headless/remote environments (SSH, VPS, containers).
 *
 * The user gets a short code, visits https://microsoft.com/devicelogin
 * on any device, and enters it. No auth server or port forwarding needed.
 */
const https = require('https');
const querystring = require('querystring');
const config = require('../config');

// Fail fast if the OAuth endpoint is unreachable (e.g. blocked outbound
// egress in a sandboxed connector) instead of hanging indefinitely. (#213)
const REQUEST_TIMEOUT_MS = 15000;

/**
 * POST helper for OAuth2 endpoints
 * @param {string} url - Full URL to POST to
 * @param {string} postData - URL-encoded form data
 * @returns {Promise<{statusCode: number, body: object}>}
 */
function postRequest(url, postData) {
  return new Promise((resolve, reject) => {
    const req = https.request(
      url,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Content-Length': Buffer.byteLength(postData),
        },
      },
      (res) => {
        let data = '';
        res.on('data', (chunk) => (data += chunk));
        res.on('end', () => {
          try {
            resolve({ statusCode: res.statusCode, body: JSON.parse(data) });
          } catch (_e) {
            reject(new Error(`Failed to parse response: ${data}`));
          }
        });
      }
    );
    req.on('error', reject);
    req.setTimeout(REQUEST_TIMEOUT_MS, () => {
      req.destroy(
        new Error(
          'Device code request timed out — network egress to login.microsoftonline.com may be blocked.'
        )
      );
    });
    req.write(postData);
    req.end();
  });
}

/**
 * Initiates the device code flow by requesting a device code from Azure.
 * @param {string} clientId - Azure app client ID
 * @param {string[]} scopes - OAuth2 scopes to request
 * @returns {Promise<{userCode: string, verificationUri: string, deviceCode: string, expiresIn: number, interval: number, message: string}>}
 */
async function initiateDeviceCodeFlow(clientId, scopes) {
  const postData = querystring.stringify({
    client_id: clientId,
    scope: scopes.join(' '),
  });

  const endpoint = config.AUTH_CONFIG.deviceCodeEndpoint;
  const { statusCode, body } = await postRequest(endpoint, postData);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(
      body.error_description ||
        `Device code request failed with status ${statusCode}`
    );
  }

  return {
    userCode: body.user_code,
    verificationUri: body.verification_uri,
    deviceCode: body.device_code,
    expiresIn: body.expires_in,
    interval: body.interval || 5,
    message: body.message,
  };
}

/**
 * Polls the token endpoint until the user completes authentication.
 * @param {string} clientId - Azure app client ID
 * @param {string} deviceCode - Device code from initiateDeviceCodeFlow
 * @param {number} interval - Polling interval in seconds
 * @param {number} expiresIn - Seconds until the device code expires
 * @returns {Promise<{access_token: string, refresh_token: string, expires_in: number, scope: string, token_type: string}>}
 */
async function pollForToken(clientId, deviceCode, interval, expiresIn) {
  const endpoint = config.AUTH_CONFIG.tokenEndpoint;
  const deadline = Date.now() + expiresIn * 1000;
  let pollInterval = interval;

  while (Date.now() < deadline) {
    await new Promise((resolve) => {
      setTimeout(resolve, pollInterval * 1000);
    });

    const postData = querystring.stringify({
      client_id: clientId,
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      device_code: deviceCode,
    });

    const { statusCode, body } = await postRequest(endpoint, postData);

    if (statusCode >= 200 && statusCode < 300) {
      return body;
    }

    switch (body.error) {
      case 'authorization_pending':
        // User hasn't completed auth yet — keep polling
        break;
      case 'slow_down':
        // Server asked us to slow down — increase interval by 5s
        pollInterval += 5;
        break;
      case 'authorization_declined':
        throw new Error('Authentication was declined by the user.');
      case 'expired_token':
        throw new Error(
          'Device code expired. Please restart the authentication process.'
        );
      default: {
        // Attach the raw OAuth payload so callers (e.g. handleDeviceCodeComplete)
        // can classify the failure — notably scope-consent rejections that should
        // trigger a base-scopes fallback. Keep the existing message text.
        const e = new Error(
          body.error_description ||
            `Token polling failed: ${body.error || `status ${statusCode}`}`
        );
        e.oauth = {
          error: body.error,
          error_codes: body.error_codes,
          suberror: body.suberror,
          error_description: body.error_description,
        };
        throw e;
      }
    }
  }

  throw new Error(
    'Device code expired. Please restart the authentication process.'
  );
}

// AADSTS codes that indicate the signed-in account cannot consent to one or
// more requested scopes (the `.Shared` scopes for a personal Microsoft account
// is the case we care about). We match the FAMILY defensively because we cannot
// test against a live personal account here:
//   650053 — "The application '<app>' asked for scope '<x>' that doesn't exist
//             on the resource" (commonly surfaced for unsupported scopes)
//   65001  — user/admin has not consented to the application
//   70011  — invalid scope value
//   28000  — invalid request / unsupported scope (seen on some tenants)
// CAVEAT: the EXACT code returned for the personal-account `.Shared` rejection
// should be confirmed empirically against a real personal Microsoft account.
// Until then we match the whole family + the OAuth `error` values defensively.
const SCOPE_CONSENT_AADSTS_CODES = ['650053', '65001', '70011', '28000'];

/**
 * Predicate: does this error look like a scope-consent rejection that warrants
 * falling back to base scopes? Matches defensively against the OAuth `error`
 * value AND the AADSTS code family found in `err.oauth.error_codes` (array) or
 * anywhere in `err.oauth.error_description` / `err.message` (string).
 * @param {Error & {oauth?: object}} err
 * @returns {boolean}
 */
function isScopeConsentError(err) {
  if (!err) {
    return false;
  }
  const oauth = err.oauth || {};

  // OAuth-level error values that map to a scope/consent problem.
  if (oauth.error === 'invalid_scope' || oauth.error === 'invalid_grant') {
    return true;
  }

  // AADSTS codes can arrive as a numeric array (`error_codes`) ...
  if (Array.isArray(oauth.error_codes)) {
    const codes = oauth.error_codes.map(String);
    if (codes.some((c) => SCOPE_CONSENT_AADSTS_CODES.includes(c))) {
      return true;
    }
  }

  // ... or embedded in free-text (error_description / message).
  const haystack = `${oauth.error_description || ''} ${err.message || ''}`;
  return SCOPE_CONSENT_AADSTS_CODES.some((code) =>
    haystack.includes(`AADSTS${code}`)
  );
}

module.exports = {
  initiateDeviceCodeFlow,
  pollForToken,
  isScopeConsentError,
};
