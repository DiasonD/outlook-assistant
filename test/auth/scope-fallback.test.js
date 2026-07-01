/**
 * Tests for the dynamic scope fallback (work/school = full scopes; personal =
 * automatic fallback to base scopes; refresh re-requests only granted scopes).
 *
 * Covers:
 *   - isScopeConsentError predicate (auth/device-code.js)
 *   - BASE_SCOPES / SHARED_SCOPES / AUTH_CONFIG scope exports (config.js)
 *   - resolveRefreshScopes scope selection (auth/token-storage.js)
 */
const { isScopeConsentError } = require('../../auth/device-code');
const config = require('../../config');
const { resolveRefreshScopes } = require('../../auth/token-storage');

describe('isScopeConsentError', () => {
  it('returns true for OAuth error=invalid_scope', () => {
    const err = new Error('AADSTS70011: invalid scope');
    err.oauth = { error: 'invalid_scope' };
    expect(isScopeConsentError(err)).toBe(true);
  });

  it('returns true for OAuth error=invalid_grant', () => {
    const err = new Error('consent required');
    err.oauth = { error: 'invalid_grant' };
    expect(isScopeConsentError(err)).toBe(true);
  });

  it('returns true for AADSTS650053 in error_codes array', () => {
    const err = new Error('scope rejected');
    err.oauth = { error: 'some_error', error_codes: [650053] };
    expect(isScopeConsentError(err)).toBe(true);
  });

  it('returns true for AADSTS65001 in error_codes array (as strings)', () => {
    const err = new Error('not consented');
    err.oauth = { error: 'some_error', error_codes: ['65001'] };
    expect(isScopeConsentError(err)).toBe(true);
  });

  it('returns true for AADSTS650053 embedded in error_description string', () => {
    const err = new Error('Token polling failed');
    err.oauth = {
      error: 'invalid_request',
      error_description:
        "AADSTS650053: The application asked for scope 'Mail.Read.Shared' that doesn't exist.",
    };
    expect(isScopeConsentError(err)).toBe(true);
  });

  it('returns true for AADSTS70011 embedded in the message string', () => {
    const err = new Error(
      'AADSTS70011: The provided value for the input parameter scope is not valid.'
    );
    expect(isScopeConsentError(err)).toBe(true);
  });

  it('returns false for authorization_pending', () => {
    const err = new Error('authorization pending');
    err.oauth = { error: 'authorization_pending' };
    expect(isScopeConsentError(err)).toBe(false);
  });

  it('returns false for expired_token', () => {
    const err = new Error('Device code expired.');
    err.oauth = { error: 'expired_token' };
    expect(isScopeConsentError(err)).toBe(false);
  });

  it('returns false for a network error (no oauth payload)', () => {
    const err = new Error('ECONNRESET');
    expect(isScopeConsentError(err)).toBe(false);
  });

  it('returns false for a generic message with no scope markers', () => {
    const err = new Error('Something unrelated went wrong');
    err.oauth = { error: 'server_error', error_codes: [50000] };
    expect(isScopeConsentError(err)).toBe(false);
  });

  it('returns false for null / undefined', () => {
    expect(isScopeConsentError(null)).toBe(false);
    expect(isScopeConsentError(undefined)).toBe(false);
  });
});

describe('config scope exports', () => {
  it('BASE_SCOPES excludes the .Shared scopes', () => {
    expect(config.BASE_SCOPES).not.toContain('Mail.Read.Shared');
    expect(config.BASE_SCOPES).not.toContain('Mail.ReadWrite.Shared');
    expect(config.BASE_SCOPES).toContain('Mail.Read');
    expect(config.BASE_SCOPES).toContain('offline_access');
  });

  it('SHARED_SCOPES contains exactly the two .Shared scopes', () => {
    expect(config.SHARED_SCOPES).toEqual([
      'Mail.Read.Shared',
      'Mail.ReadWrite.Shared',
    ]);
  });

  it('AUTH_CONFIG.scopes is base + shared (includes both .Shared)', () => {
    expect(config.AUTH_CONFIG.scopes).toContain('Mail.Read.Shared');
    expect(config.AUTH_CONFIG.scopes).toContain('Mail.ReadWrite.Shared');
    expect(config.AUTH_CONFIG.scopes).toEqual([
      ...config.BASE_SCOPES,
      ...config.SHARED_SCOPES,
    ]);
  });

  it('AUTH_CONFIG.fallbackScopes matches BASE_SCOPES content', () => {
    expect(config.AUTH_CONFIG.fallbackScopes).toEqual(config.BASE_SCOPES);
  });
});

describe('resolveRefreshScopes — refresh uses granted, not configured, scopes', () => {
  const FULL = [...config.BASE_SCOPES, ...config.SHARED_SCOPES];

  it('prefers granted_scopes (array) when present', () => {
    const tokens = { granted_scopes: config.BASE_SCOPES };
    const result = resolveRefreshScopes(tokens, FULL);
    expect(result).toEqual(config.BASE_SCOPES);
    expect(result).not.toContain('Mail.Read.Shared');
  });

  it('parses the scope string when granted_scopes is absent', () => {
    const tokens = { scope: 'offline_access User.Read Mail.Read' };
    const result = resolveRefreshScopes(tokens, FULL);
    expect(result).toEqual(['offline_access', 'User.Read', 'Mail.Read']);
  });

  it('falls back to configured scopes when neither granted_scopes nor scope exist', () => {
    const tokens = { access_token: 'x' };
    const result = resolveRefreshScopes(tokens, FULL);
    expect(result).toEqual(FULL);
  });

  it('falls back to configured scopes when tokens is null', () => {
    expect(resolveRefreshScopes(null, FULL)).toEqual(FULL);
  });

  it('ignores an empty granted_scopes array and falls through to scope string', () => {
    const tokens = { granted_scopes: [], scope: 'offline_access Mail.Read' };
    expect(resolveRefreshScopes(tokens, FULL)).toEqual([
      'offline_access',
      'Mail.Read',
    ]);
  });
});
