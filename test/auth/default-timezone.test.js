/**
 * Tests for OUTLOOK_DEFAULT_TIMEZONE — the env-var override for the
 * default IANA timezone applied to calendar events when the caller
 * doesn't explicitly specify one.
 */

describe('OUTLOOK_DEFAULT_TIMEZONE', () => {
  let originalTz;

  beforeEach(() => {
    originalTz = process.env.OUTLOOK_DEFAULT_TIMEZONE;
    jest.resetModules();
  });

  afterEach(() => {
    if (originalTz === undefined) {
      delete process.env.OUTLOOK_DEFAULT_TIMEZONE;
    } else {
      process.env.OUTLOOK_DEFAULT_TIMEZONE = originalTz;
    }
    jest.resetModules();
  });

  test('defaults to Australia/Melbourne when env var is unset', () => {
    delete process.env.OUTLOOK_DEFAULT_TIMEZONE;
    const { DEFAULT_TIMEZONE } = require('../../config');
    expect(DEFAULT_TIMEZONE).toBe('Australia/Melbourne');
  });

  test('honours OUTLOOK_DEFAULT_TIMEZONE when set', () => {
    process.env.OUTLOOK_DEFAULT_TIMEZONE = 'Europe/London';
    const { DEFAULT_TIMEZONE } = require('../../config');
    expect(DEFAULT_TIMEZONE).toBe('Europe/London');
  });

  test('passes the configured timezone through to create-event', async () => {
    process.env.OUTLOOK_DEFAULT_TIMEZONE = 'America/New_York';
    jest.resetModules();
    jest.doMock('../../utils/graph-api');
    jest.doMock('../../auth');
    const { callGraphAPI } = require('../../utils/graph-api');
    const { ensureAuthenticated } = require('../../auth');
    const handleCreateEvent = require('../../calendar/create');

    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1' });

    await handleCreateEvent({
      subject: 'NY meeting',
      start: '2026-06-01T09:00:00',
      end: '2026-06-01T10:00:00',
    });

    const body = callGraphAPI.mock.calls[0][3];
    expect(body.start.timeZone).toBe('America/New_York');
    expect(body.end.timeZone).toBe('America/New_York');
  });
});
