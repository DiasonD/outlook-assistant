const {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  resolveFolderRef,
  getFolderIdByName,
  getAllFolders,
  buildMailboxPrefix,
} = require('../../email/folder-utils');
const { callGraphAPI } = require('../../utils/graph-api');

jest.mock('../../utils/graph-api');

describe('resolveFolderPath', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    // Mock console.error to avoid cluttering test output
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('well-known folders', () => {
    test('should return inbox endpoint when no folder name is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, null);
      expect(result).toBe(WELL_KNOWN_FOLDERS.inbox);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return inbox endpoint when undefined folder name is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, undefined);
      expect(result).toBe(WELL_KNOWN_FOLDERS.inbox);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return inbox endpoint when empty string is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, '');
      expect(result).toBe(WELL_KNOWN_FOLDERS.inbox);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return correct endpoint for well-known folders', async () => {
      const result = await resolveFolderPath(mockAccessToken, 'drafts');
      expect(result).toBe(WELL_KNOWN_FOLDERS.drafts);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should handle case-insensitive well-known folder names', async () => {
      const result1 = await resolveFolderPath(mockAccessToken, 'INBOX');
      const result2 = await resolveFolderPath(mockAccessToken, 'Drafts');
      const result3 = await resolveFolderPath(mockAccessToken, 'SENT');

      expect(result1).toBe(WELL_KNOWN_FOLDERS.inbox);
      expect(result2).toBe(WELL_KNOWN_FOLDERS.drafts);
      expect(result3).toBe(WELL_KNOWN_FOLDERS.sent);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should resolve Graph API well-known folder names (sentitems, deleteditems, junkemail)', async () => {
      const result1 = await resolveFolderPath(mockAccessToken, 'sentitems');
      const result2 = await resolveFolderPath(mockAccessToken, 'deleteditems');
      const result3 = await resolveFolderPath(mockAccessToken, 'junkemail');
      const result4 = await resolveFolderPath(mockAccessToken, 'outbox');

      expect(result1).toBe(WELL_KNOWN_FOLDERS.sentitems);
      expect(result2).toBe(WELL_KNOWN_FOLDERS.deleteditems);
      expect(result3).toBe(WELL_KNOWN_FOLDERS.junkemail);
      expect(result4).toBe(WELL_KNOWN_FOLDERS.outbox);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should resolve display name aliases (Sent Items, Deleted Items, Junk Email)', async () => {
      const result1 = await resolveFolderPath(mockAccessToken, 'Sent Items');
      const result2 = await resolveFolderPath(mockAccessToken, 'Deleted Items');
      const result3 = await resolveFolderPath(mockAccessToken, 'Junk Email');
      const result4 = await resolveFolderPath(mockAccessToken, 'spam');

      expect(result1).toBe(WELL_KNOWN_FOLDERS.sent);
      expect(result2).toBe(WELL_KNOWN_FOLDERS.deleted);
      expect(result3).toBe(WELL_KNOWN_FOLDERS.junk);
      expect(result4).toBe(WELL_KNOWN_FOLDERS.junk);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });
  });

  describe('custom folders', () => {
    test('should resolve custom folder by ID when found', async () => {
      const customFolderId = 'custom-folder-id-123';
      const customFolderName = 'MyCustomFolder';

      callGraphAPI.mockResolvedValueOnce({
        value: [{ id: customFolderId, displayName: customFolderName }],
      });

      const result = await resolveFolderPath(mockAccessToken, customFolderName);

      expect(result).toBe(`me/mailFolders/${customFolderId}/messages`);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/mailFolders',
        null,
        { $filter: `displayName eq '${customFolderName}'` }
      );
    });

    test('should try case-insensitive search when exact match fails', async () => {
      const customFolderId = 'custom-folder-id-456';
      const customFolderName = 'ProjectAlpha';

      // First call returns empty (exact match fails)
      callGraphAPI.mockResolvedValueOnce({ value: [] });

      // Second call returns all folders for case-insensitive match
      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: 'other-id', displayName: 'OtherFolder' },
          { id: customFolderId, displayName: 'projectalpha' },
        ],
      });

      const result = await resolveFolderPath(mockAccessToken, customFolderName);

      expect(result).toBe(`me/mailFolders/${customFolderId}/messages`);
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
    });

    test('should throw error when custom folder is not found', async () => {
      const nonExistentFolder = 'NonExistentFolder';

      // First call returns empty (exact match fails)
      callGraphAPI.mockResolvedValueOnce({ value: [] });

      // Second call returns folders without a match
      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: 'id1', displayName: 'Folder1' },
          { id: 'id2', displayName: 'Folder2' },
        ],
      });

      await expect(
        resolveFolderPath(mockAccessToken, nonExistentFolder)
      ).rejects.toThrow('not found');
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
    });

    test('should throw error when API call fails', async () => {
      const customFolderName = 'CustomFolder';

      callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

      await expect(
        resolveFolderPath(mockAccessToken, customFolderName)
      ).rejects.toThrow('not found');
      expect(callGraphAPI).toHaveBeenCalledTimes(1);
    });
  });
});

describe('getFolderIdByName', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should return folder ID when exact match is found', async () => {
    const folderId = 'folder-id-123';
    const folderName = 'TestFolder';

    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: folderId, displayName: folderName }],
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBe(folderId);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/mailFolders',
      null,
      { $filter: `displayName eq '${folderName}'` }
    );
  });

  test('should return folder ID when case-insensitive match is found', async () => {
    const folderId = 'folder-id-456';
    const folderName = 'TestFolder';

    // First call returns empty (exact match fails)
    callGraphAPI.mockResolvedValueOnce({ value: [] });

    // Second call returns folders with case-insensitive match
    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: folderId, displayName: 'testfolder' }],
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBe(folderId);
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
  });

  test('should return null when folder is not found', async () => {
    const folderName = 'NonExistentFolder';

    // First call returns empty
    callGraphAPI.mockResolvedValueOnce({ value: [] });

    // Second call returns folders without a match
    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: 'id1', displayName: 'OtherFolder' }],
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBeNull();
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
  });

  test('should return null when API call fails', async () => {
    const folderName = 'TestFolder';

    callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBeNull();
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  test('should find a nested subfolder by name via recursion', async () => {
    // Top-level exact filter miss, then top-level list, then childFolders
    callGraphAPI
      .mockResolvedValueOnce({ value: [] }) // exact filter
      .mockResolvedValueOnce({
        value: [
          { id: 'inbox-id', displayName: 'Inbox', childFolderCount: 2 },
          { id: 'sent-id', displayName: 'Sent Items', childFolderCount: 0 },
        ],
      })
      .mockResolvedValueOnce({
        value: [
          { id: 'vendor-id', displayName: 'Vendors', childFolderCount: 0 },
          { id: 'acme-id', displayName: 'Acme', childFolderCount: 0 },
        ],
      });

    const result = await getFolderIdByName(mockAccessToken, 'Acme');

    expect(result).toBe('acme-id');
    expect(callGraphAPI).toHaveBeenCalledTimes(3);
  });

  test('should target a shared mailbox when mailbox is provided', async () => {
    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: 'shared-folder', displayName: 'Projekte' }],
    });

    const result = await getFolderIdByName(
      mockAccessToken,
      'Projekte',
      'shared@company.com'
    );

    expect(result).toBe('shared-folder');
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'users/shared@company.com/mailFolders',
      null,
      { $filter: `displayName eq 'Projekte'` }
    );
  });
});

describe('buildMailboxPrefix', () => {
  test('returns "me" for null/empty', () => {
    expect(buildMailboxPrefix(null)).toBe('me');
    expect(buildMailboxPrefix('')).toBe('me');
    expect(buildMailboxPrefix(undefined)).toBe('me');
  });

  test('builds users/{email} prefix for a shared mailbox', () => {
    expect(buildMailboxPrefix('shared@company.com')).toBe(
      'users/shared@company.com'
    );
  });

  test('passes through an already-qualified prefix', () => {
    expect(buildMailboxPrefix('users/x@y.com')).toBe('users/x@y.com');
    expect(buildMailboxPrefix('me')).toBe('me');
  });
});

describe('resolveFolderRef', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('resolves a well-known folder to its segment without an API call', async () => {
    const result = await resolveFolderRef(mockAccessToken, 'archive');
    expect(result).toBe('archive');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('resolves a nested folder path by walking the tree', async () => {
    // first segment "Inbox" is well-known → segment "inbox"
    // then childFolders of inbox → Vendors, then childFolders of Vendors → Acme
    callGraphAPI
      .mockResolvedValueOnce({
        value: [{ id: 'vendors-id', displayName: 'Vendors' }],
      })
      .mockResolvedValueOnce({
        value: [{ id: 'acme-id', displayName: 'Acme' }],
      });

    const result = await resolveFolderRef(
      mockAccessToken,
      'Inbox/Vendors/Acme',
      'shared@company.com'
    );

    expect(result).toBe('acme-id');
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      1,
      mockAccessToken,
      'GET',
      'users/shared@company.com/mailFolders/inbox/childFolders',
      null,
      { $top: 100 }
    );
  });

  test('passes through a raw folder ID when nothing resolves', async () => {
    // looks-like-ID heuristic: long, no spaces, base64url chars
    callGraphAPI
      .mockResolvedValueOnce({ value: [] }) // exact filter
      .mockResolvedValueOnce({ value: [] }); // top-level list

    const rawId =
      'AAMkADRmMDExLT1234567890abcdefABCDEF_ghijklmnopqrstuvwxyz0987654321=';
    const result = await resolveFolderRef(mockAccessToken, rawId);

    expect(result).toBe(rawId);
  });

  test('returns null for an unresolved short name', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [] })
      .mockResolvedValueOnce({ value: [{ id: 'x', displayName: 'Other' }] });

    const result = await resolveFolderRef(mockAccessToken, 'Nope');
    expect(result).toBeNull();
  });
});

describe('resolveFolderPath (shared mailbox)', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('builds a shared-mailbox well-known path without an API call', async () => {
    const result = await resolveFolderPath(
      mockAccessToken,
      'inbox',
      'shared@company.com'
    );
    expect(result).toBe('users/shared@company.com/mailFolders/inbox/messages');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('resolves a shared-mailbox custom folder to its message endpoint', async () => {
    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: 'custom-id', displayName: 'Archiv' }],
    });

    const result = await resolveFolderPath(
      mockAccessToken,
      'Archiv',
      'shared@company.com'
    );

    expect(result).toBe(
      'users/shared@company.com/mailFolders/custom-id/messages'
    );
  });
});

describe('getAllFolders', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('enumerates a shared mailbox tree recursively with paths', async () => {
    callGraphAPI
      .mockResolvedValueOnce({
        value: [
          {
            id: 'inbox-id',
            displayName: 'Inbox',
            parentFolderId: 'root',
            childFolderCount: 1,
          },
          {
            id: 'sent-id',
            displayName: 'Sent Items',
            parentFolderId: 'root',
            childFolderCount: 0,
          },
        ],
      })
      .mockResolvedValueOnce({
        value: [
          {
            id: 'vendors-id',
            displayName: 'Vendors',
            parentFolderId: 'inbox-id',
            childFolderCount: 0,
          },
        ],
      });

    const folders = await getAllFolders(mockAccessToken, {
      mailbox: 'shared@company.com',
    });

    expect(folders).toHaveLength(3);
    const vendors = folders.find((f) => f.displayName === 'Vendors');
    expect(vendors.folderPath).toBe('Inbox/Vendors');
    expect(vendors.parentFolder).toBe('Inbox');
    expect(vendors.isTopLevel).toBe(false);
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      1,
      mockAccessToken,
      'GET',
      'users/shared@company.com/mailFolders',
      null,
      expect.objectContaining({ $top: 100 })
    );
  });
});
