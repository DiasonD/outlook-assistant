/**
 * Tests for OUTLOOK_AUTH_AUDIENCE — the configurable OAuth audience that
 * picks which segment is used in the Microsoft login URLs:
 *   common (default), consumers, organizations, <tenant-guid>
 */

describe('OUTLOOK_AUTH_AUDIENCE', () => {
  let originalAudience;

  beforeEach(() => {
    originalAudience = process.env.OUTLOOK_AUTH_AUDIENCE;
    // Force a fresh load of config.js — Node caches required modules,
    // so without this the env var change wouldn't take effect.
    jest.resetModules();
  });

  afterEach(() => {
    if (originalAudience === undefined) {
      delete process.env.OUTLOOK_AUTH_AUDIENCE;
    } else {
      process.env.OUTLOOK_AUTH_AUDIENCE = originalAudience;
    }
    jest.resetModules();
  });

  test('defaults to "common" when env var is unset', () => {
    delete process.env.OUTLOOK_AUTH_AUDIENCE;
    const { AUTH_CONFIG } = require('../../config');
    expect(AUTH_CONFIG.audience).toBe('common');
    expect(AUTH_CONFIG.tokenEndpoint).toBe(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    );
    expect(AUTH_CONFIG.deviceCodeEndpoint).toBe(
      'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode'
    );
    expect(AUTH_CONFIG.authorizeEndpoint).toBe(
      'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
    );
  });

  test('routes to /consumers/ for personal-only apps', () => {
    process.env.OUTLOOK_AUTH_AUDIENCE = 'consumers';
    const { AUTH_CONFIG } = require('../../config');
    expect(AUTH_CONFIG.audience).toBe('consumers');
    expect(AUTH_CONFIG.tokenEndpoint).toBe(
      'https://login.microsoftonline.com/consumers/oauth2/v2.0/token'
    );
    expect(AUTH_CONFIG.deviceCodeEndpoint).toBe(
      'https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode'
    );
    expect(AUTH_CONFIG.authorizeEndpoint).toBe(
      'https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize'
    );
  });

  test('routes to /organizations/ for work/school-only apps', () => {
    process.env.OUTLOOK_AUTH_AUDIENCE = 'organizations';
    const { AUTH_CONFIG } = require('../../config');
    expect(AUTH_CONFIG.audience).toBe('organizations');
    expect(AUTH_CONFIG.tokenEndpoint).toBe(
      'https://login.microsoftonline.com/organizations/oauth2/v2.0/token'
    );
  });

  test('accepts a tenant GUID for single-tenant apps', () => {
    const tenantId = '11111111-2222-3333-4444-555555555555';
    process.env.OUTLOOK_AUTH_AUDIENCE = tenantId;
    const { AUTH_CONFIG } = require('../../config');
    expect(AUTH_CONFIG.audience).toBe(tenantId);
    expect(AUTH_CONFIG.tokenEndpoint).toBe(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`
    );
    expect(AUTH_CONFIG.authorizeEndpoint).toBe(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`
    );
  });
});
