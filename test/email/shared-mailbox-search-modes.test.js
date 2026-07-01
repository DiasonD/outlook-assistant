/**
 * Regression tests (issues #3 / #4): the `search-emails` delta and
 * Message-ID lookup modes must route to the owning shared/delegated mailbox
 * when `sharedMailbox`/`email` is supplied, rather than silently querying the
 * signed-in account.
 *
 * folder-utils is partially mocked: the real `buildMailboxPrefix` runs so the
 * `me` vs `users/{mailbox}` construction is exercised, while `resolveFolderRef`
 * is stubbed to a deterministic value so delta-mode endpoint assertions don't
 * depend on live folder enumeration.
 */
const handleListEmailsDelta = require('../../email/delta');
const { handleSearchByMessageId } = require('../../email/search');
const { handler: searchEmailsHandler } =
  require('../../email/index').emailTools.find(
    (t) => t.name === 'search-emails'
  );
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const folderUtils = require('../../email/folder-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../email/folder-utils', () => {
  const actual = jest.requireActual('../../email/folder-utils');
  return {
    ...actual,
    resolveFolderRef: jest.fn(),
  };
});

const TOKEN = 'test_token';
const MAILBOX = 'office@werdropo.com';
const RESOLVED_REF = 'AAMkResolvedFolderRef=';

/** Return the endpoint (3rd positional arg) of the Nth callGraphAPI call. */
function endpointOfCall(n = 0) {
  return callGraphAPI.mock.calls[n][2];
}

beforeEach(() => {
  jest.clearAllMocks();
  jest.spyOn(console, 'error').mockImplementation(() => {});
  ensureAuthenticated.mockResolvedValue(TOKEN);
  folderUtils.resolveFolderRef.mockResolvedValue(RESOLVED_REF);
});

afterEach(() => {
  console.error.mockRestore();
});

describe('delta initial sync — shared-mailbox routing', () => {
  test('routes to /users/{mailbox} when sharedMailbox is set', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });
    await handleListEmailsDelta({ folder: 'Vendors', sharedMailbox: MAILBOX });
    expect(folderUtils.resolveFolderRef).toHaveBeenCalledWith(
      TOKEN,
      'Vendors',
      MAILBOX
    );
    expect(endpointOfCall()).toBe(
      `users/${MAILBOX}/mailFolders/${RESOLVED_REF}/messages/delta`
    );
  });

  test('accepts `email` as an alias', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });
    await handleListEmailsDelta({ email: MAILBOX });
    expect(endpointOfCall()).toBe(
      `users/${MAILBOX}/mailFolders/${RESOLVED_REF}/messages/delta`
    );
  });

  test('defaults to /me when no mailbox supplied', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });
    await handleListEmailsDelta({ folder: 'inbox' });
    expect(folderUtils.resolveFolderRef).toHaveBeenCalledWith(
      TOKEN,
      'inbox',
      null
    );
    expect(endpointOfCall()).toBe(
      `me/mailFolders/${RESOLVED_REF}/messages/delta`
    );
  });

  test('continuation (deltaToken) branch is left untouched', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });
    const deltaLink = 'me/mailFolders/inbox/messages/delta?$deltatoken=abc';
    await handleListEmailsDelta({
      deltaToken: deltaLink,
      sharedMailbox: MAILBOX,
    });
    // The deltaLink URL is used verbatim; resolveFolderRef must NOT run.
    expect(folderUtils.resolveFolderRef).not.toHaveBeenCalled();
    expect(endpointOfCall()).toBe(deltaLink);
  });

  test('returns a clean folder-not-found response when ref is null', async () => {
    folderUtils.resolveFolderRef.mockResolvedValue(null);
    const result = await handleListEmailsDelta({
      folder: 'Nope',
      sharedMailbox: MAILBOX,
    });
    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(result.content[0].text).toMatch(/not found/i);
  });
});

describe('message-id lookup — shared-mailbox routing', () => {
  test('routes GET to /users/{mailbox}/messages when sharedMailbox is set', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });
    await handleSearchByMessageId({
      messageId: '<abc123@example.com>',
      sharedMailbox: MAILBOX,
    });
    expect(endpointOfCall()).toBe(`users/${MAILBOX}/messages`);
  });

  test('accepts `email` as an alias', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });
    await handleSearchByMessageId({
      messageId: '<abc123@example.com>',
      email: MAILBOX,
    });
    expect(endpointOfCall()).toBe(`users/${MAILBOX}/messages`);
  });

  test('defaults to /me/messages when no mailbox supplied', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });
    await handleSearchByMessageId({ messageId: '<abc123@example.com>' });
    expect(endpointOfCall()).toBe('me/messages');
  });
});

describe('search-emails router — forwards sharedMailbox to message-id lookup', () => {
  test('internetMessageId + sharedMailbox reaches /users/{mailbox}/messages', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });
    await searchEmailsHandler({
      internetMessageId: '<abc123@example.com>',
      sharedMailbox: MAILBOX,
    });
    // Proves the router no longer drops sharedMailbox before the handler.
    expect(endpointOfCall()).toBe(`users/${MAILBOX}/messages`);
  });

  test('internetMessageId without mailbox stays on /me/messages', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });
    await searchEmailsHandler({ internetMessageId: '<abc123@example.com>' });
    expect(endpointOfCall()).toBe('me/messages');
  });
});
