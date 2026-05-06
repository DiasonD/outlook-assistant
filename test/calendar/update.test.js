const handleUpdateEvent = require('../../calendar/update');
const { DEFAULT_TIMEZONE } = require('../../config');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleUpdateEvent', () => {
  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
  });

  test('returns error when eventId is missing', async () => {
    const result = await handleUpdateEvent({ subject: 'foo' });
    expect(result.content[0].text).toBe(
      'Event ID is required to update an event.'
    );
    expect(ensureAuthenticated).not.toHaveBeenCalled();
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('returns error when no updatable fields are provided', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    const result = await handleUpdateEvent({ eventId: 'evt_1' });
    expect(result.content[0].text).toMatch(/No fields to update/);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('PATCHes only the subject when subject-only update is requested', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1', subject: 'New' });

    await handleUpdateEvent({ eventId: 'evt_1', subject: 'New' });

    expect(callGraphAPI).toHaveBeenCalledTimes(1);
    const [, method, endpoint, body] = callGraphAPI.mock.calls[0];
    expect(method).toBe('PATCH');
    expect(endpoint).toBe('me/events/evt_1');
    expect(body).toEqual({ subject: 'New' });
  });

  test('uses default timezone when start is provided without timezone', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1' });

    await handleUpdateEvent({
      eventId: 'evt_1',
      start: '2026-05-12T10:00:00',
    });

    const body = callGraphAPI.mock.calls[0][3];
    expect(body.start.dateTime).toBe('2026-05-12T10:00:00');
    expect(body.start.timeZone).toBe(DEFAULT_TIMEZONE);
    expect(body.end).toBeUndefined();
  });

  test('honours explicit timezone on start when provided as object', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1' });

    await handleUpdateEvent({
      eventId: 'evt_1',
      start: { dateTime: '2026-05-12T10:00:00', timeZone: 'Europe/London' },
    });

    const body = callGraphAPI.mock.calls[0][3];
    expect(body.start.timeZone).toBe('Europe/London');
  });

  test('updates end with default timezone when string is provided', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1' });

    await handleUpdateEvent({
      eventId: 'evt_1',
      end: '2026-05-12T17:00:00',
    });

    const body = callGraphAPI.mock.calls[0][3];
    expect(body.end.dateTime).toBe('2026-05-12T17:00:00');
    expect(body.end.timeZone).toBe(DEFAULT_TIMEZONE);
    expect(body.start).toBeUndefined();
  });

  test('replaces the attendee list when attendees is provided', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1' });

    await handleUpdateEvent({
      eventId: 'evt_1',
      attendees: ['alice@example.com', 'bob@example.com'],
    });

    const body = callGraphAPI.mock.calls[0][3];
    expect(body.attendees).toEqual([
      { emailAddress: { address: 'alice@example.com' }, type: 'required' },
      { emailAddress: { address: 'bob@example.com' }, type: 'required' },
    ]);
  });

  test('clears attendees when an empty array is passed', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1' });

    await handleUpdateEvent({ eventId: 'evt_1', attendees: [] });

    const body = callGraphAPI.mock.calls[0][3];
    expect(body.attendees).toEqual([]);
  });

  test('updates body as HTML content', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1' });

    await handleUpdateEvent({
      eventId: 'evt_1',
      body: 'New notes',
    });

    const body = callGraphAPI.mock.calls[0][3];
    expect(body.body).toEqual({ contentType: 'HTML', content: 'New notes' });
  });

  test('updates location as displayName', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1' });

    await handleUpdateEvent({
      eventId: 'evt_1',
      location: '16 Apex Drive',
    });

    const body = callGraphAPI.mock.calls[0][3];
    expect(body.location).toEqual({ displayName: '16 Apex Drive' });
  });

  test('combines multiple fields in a single PATCH', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1', subject: 'Reschedule' });

    await handleUpdateEvent({
      eventId: 'evt_1',
      subject: 'Reschedule',
      start: '2026-05-12T10:00:00',
      end: '2026-05-12T17:00:00',
    });

    expect(callGraphAPI).toHaveBeenCalledTimes(1);
    const body = callGraphAPI.mock.calls[0][3];
    expect(body.subject).toBe('Reschedule');
    expect(body.start.dateTime).toBe('2026-05-12T10:00:00');
    expect(body.end.dateTime).toBe('2026-05-12T17:00:00');
  });

  test('handles authentication errors with the standard message', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleUpdateEvent({
      eventId: 'evt_1',
      subject: 'New',
    });
    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('surfaces Graph API errors to the caller', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockRejectedValue(new Error('Graph API exploded'));

    const result = await handleUpdateEvent({
      eventId: 'evt_1',
      subject: 'New',
    });
    expect(result.content[0].text).toBe(
      'Error updating event: Graph API exploded'
    );
  });

  test('reports which fields were changed in the success output', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'evt_1', subject: 'Reschedule' });

    const result = await handleUpdateEvent({
      eventId: 'evt_1',
      subject: 'Reschedule',
      location: '16 Apex Drive',
    });

    const text = result.content[0].text;
    expect(text).toMatch(/updated successfully/);
    expect(text).toMatch(/Fields changed:.*subject/);
    expect(text).toMatch(/Fields changed:.*location/);
    expect(result._meta.fieldsChanged).toEqual(
      expect.arrayContaining(['subject', 'location'])
    );
  });
});
