const {
  handleListFolders,
  handleCreateFolder,
  handleMoveEmails,
  handleGetFolderStats,
  handleDeleteFolder,
} = require('../../folder');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const {
  getFolderIdByName,
  buildMailboxPrefix,
  resolveFolderRef,
} = require('../../email/folder-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../email/folder-utils');

const mockAccessToken = 'test_token';

const mockFolders = [
  {
    id: 'folder-1',
    displayName: 'Inbox',
    parentFolderId: 'root',
    childFolderCount: 0,
    totalItemCount: 42,
    unreadItemCount: 5,
  },
  {
    id: 'folder-2',
    displayName: 'Sent Items',
    parentFolderId: 'root',
    childFolderCount: 0,
    totalItemCount: 100,
    unreadItemCount: 0,
  },
];

beforeEach(() => {
  jest.resetAllMocks();
  jest.spyOn(console, 'error').mockImplementation();
  ensureAuthenticated.mockResolvedValue(mockAccessToken);
  // folder-utils is fully mocked; restore the real prefix logic so the move
  // handler builds correct `me` / `users/{email}` paths.
  buildMailboxPrefix.mockImplementation((mailbox) => {
    if (!mailbox) return 'me';
    if (mailbox === 'me' || mailbox.startsWith('users/')) return mailbox;
    return `users/${mailbox}`;
  });
});

afterEach(() => {
  console.error.mockRestore();
});

describe('handleListFolders', () => {
  it('should list folders as flat list', async () => {
    callGraphAPI.mockResolvedValue({ value: mockFolders });

    const result = await handleListFolders({});

    expect(result.content[0].text).toContain('Found 2 folders');
    expect(result.content[0].text).toContain('Inbox');
    expect(result.content[0].text).toContain('Sent Items');
  });

  it('should include item counts when requested', async () => {
    callGraphAPI.mockResolvedValue({ value: mockFolders });

    const result = await handleListFolders({ includeItemCounts: true });

    expect(result.content[0].text).toContain('42 items');
    expect(result.content[0].text).toContain('5 unread');
  });

  it('should format as hierarchy when requested', async () => {
    callGraphAPI.mockResolvedValue({ value: mockFolders });

    const result = await handleListFolders({ includeChildren: true });

    expect(result.content[0].text).toContain('Folder Hierarchy');
  });

  it('should handle empty folders', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });

    const result = await handleListFolders({});

    expect(result.content[0].text).toContain('No folders found');
  });

  it('should enumerate a shared mailbox when sharedMailbox is provided', async () => {
    callGraphAPI.mockResolvedValue({ value: mockFolders });

    const result = await handleListFolders({
      sharedMailbox: 'shared@company.com',
    });

    expect(result.content[0].text).toContain('Mailbox: shared@company.com');
    // Folder enumeration should target the shared mailbox endpoint
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'users/shared@company.com/mailFolders',
      null,
      expect.objectContaining({ $top: 100 })
    );
  });

  it('should accept email as an alias for sharedMailbox', async () => {
    callGraphAPI.mockResolvedValue({ value: mockFolders });

    await handleListFolders({ email: 'shared@company.com' });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'users/shared@company.com/mailFolders',
      null,
      expect.objectContaining({ $top: 100 })
    );
  });

  it('should handle auth error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleListFolders({});

    expect(result.content[0].text).toContain('Authentication required');
  });

  it('should handle API error', async () => {
    callGraphAPI.mockRejectedValue(new Error('API Error'));

    const result = await handleListFolders({});

    expect(result.content[0].text).toBe('Error listing folders: API Error');
  });
});

describe('handleCreateFolder', () => {
  it('should create a folder at root level', async () => {
    getFolderIdByName.mockResolvedValueOnce(null); // No existing folder
    callGraphAPI.mockResolvedValue({ id: 'new-folder-id' });

    const result = await handleCreateFolder({ name: 'My Folder' });

    expect(result.content[0].text).toContain('Successfully created folder');
    expect(result.content[0].text).toContain('My Folder');
    expect(result.content[0].text).toContain('root level');
  });

  it('should include the new folder ID in response (F-31)', async () => {
    getFolderIdByName.mockResolvedValueOnce(null);
    callGraphAPI.mockResolvedValue({ id: 'new-folder-id-12345' });

    const result = await handleCreateFolder({ name: 'WithId' });

    expect(result.content[0].text).toMatch(/\*\*ID\*\*: new-folder-id-12345/);
    expect(result._meta.folderId).toBe('new-folder-id-12345');
  });

  it('should create a folder inside a parent', async () => {
    getFolderIdByName
      .mockResolvedValueOnce(null) // No existing folder with same name
      .mockResolvedValueOnce('parent-id'); // Parent folder found
    callGraphAPI.mockResolvedValue({ id: 'new-folder-id' });

    const result = await handleCreateFolder({
      name: 'Subfolder',
      parentFolder: 'Inbox',
    });

    expect(result.content[0].text).toContain('Successfully created folder');
    expect(result.content[0].text).toContain('inside "Inbox"');
  });

  it('should not create if folder already exists', async () => {
    getFolderIdByName.mockResolvedValueOnce('existing-id');

    const result = await handleCreateFolder({ name: 'Existing' });

    expect(result.content[0].text).toContain('already exists');
  });

  it('should handle missing parent folder', async () => {
    getFolderIdByName
      .mockResolvedValueOnce(null) // No existing folder
      .mockResolvedValueOnce(null); // Parent not found

    const result = await handleCreateFolder({
      name: 'Subfolder',
      parentFolder: 'NonExistent',
    });

    expect(result.content[0].text).toContain('not found');
  });

  it('should create a folder inside a shared mailbox at root', async () => {
    getFolderIdByName.mockResolvedValueOnce(null); // no existing folder
    callGraphAPI.mockResolvedValue({ id: 'shared-folder-id' });

    const result = await handleCreateFolder({
      name: 'Lieferanten',
      sharedMailbox: 'shared@company.com',
    });

    expect(result.content[0].text).toContain('Successfully created folder');
    // Duplicate check must be scoped to the shared mailbox
    expect(getFolderIdByName).toHaveBeenCalledWith(
      mockAccessToken,
      'Lieferanten',
      'shared@company.com'
    );
    // Create POST must target the shared mailbox
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'POST',
      'users/shared@company.com/mailFolders',
      { displayName: 'Lieferanten' }
    );
  });

  it('should create a folder inside a shared mailbox parent (email alias)', async () => {
    getFolderIdByName
      .mockResolvedValueOnce(null) // no existing folder
      .mockResolvedValueOnce('parent-id'); // parent resolved in shared mailbox
    callGraphAPI.mockResolvedValue({ id: 'child-id' });

    await handleCreateFolder({
      name: 'Acme',
      parentFolder: 'Lieferanten',
      email: 'shared@company.com',
    });

    expect(getFolderIdByName).toHaveBeenNthCalledWith(
      2,
      mockAccessToken,
      'Lieferanten',
      'shared@company.com'
    );
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'POST',
      'users/shared@company.com/mailFolders/parent-id/childFolders',
      { displayName: 'Acme' }
    );
  });

  it('should default folder creation to the signed-in mailbox (me)', async () => {
    getFolderIdByName.mockResolvedValueOnce(null);
    callGraphAPI.mockResolvedValue({ id: 'new-id' });

    await handleCreateFolder({ name: 'Personal' });

    expect(getFolderIdByName).toHaveBeenCalledWith(
      mockAccessToken,
      'Personal',
      null
    );
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'POST',
      'me/mailFolders',
      { displayName: 'Personal' }
    );
  });

  it('should require folder name', async () => {
    const result = await handleCreateFolder({});

    expect(result.content[0].text).toBe('Folder name is required.');
  });

  it('should handle auth error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleCreateFolder({ name: 'Test' });

    expect(result.content[0].text).toContain('Authentication required');
  });

  it('should handle API error', async () => {
    getFolderIdByName.mockResolvedValueOnce(null);
    callGraphAPI.mockRejectedValue(new Error('Create failed'));

    const result = await handleCreateFolder({ name: 'Test' });

    expect(result.content[0].text).toBe('Error creating folder: Create failed');
  });
});

describe('handleMoveEmails', () => {
  it('should move emails to target folder', async () => {
    getFolderIdByName.mockResolvedValue('target-folder-id');
    callGraphAPI.mockResolvedValue({});

    const result = await handleMoveEmails({
      emailIds: 'msg-1,msg-2',
      targetFolder: 'Archive',
    });

    expect(result.content[0].text).toContain('Successfully moved 2 email(s)');
    expect(result.content[0].text).toContain('Archive');
  });

  it('should resolve the folder and move within a shared mailbox', async () => {
    getFolderIdByName.mockResolvedValue('shared-folder-id');
    callGraphAPI.mockResolvedValue({});

    const result = await handleMoveEmails({
      emailIds: 'msg-1',
      targetFolder: 'Vendors/Acme',
      sharedMailbox: 'shared@company.com',
    });

    expect(result.content[0].text).toContain('Successfully moved 1 email(s)');
    // Folder lookup must be scoped to the shared mailbox
    expect(getFolderIdByName).toHaveBeenCalledWith(
      mockAccessToken,
      'Vendors/Acme',
      'shared@company.com'
    );
    // Move POST must target the shared mailbox
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'POST',
      'users/shared@company.com/messages/msg-1/move',
      { destinationId: 'shared-folder-id' }
    );
  });

  it('should default the move to the signed-in mailbox (me)', async () => {
    getFolderIdByName.mockResolvedValue('target-folder-id');
    callGraphAPI.mockResolvedValue({});

    await handleMoveEmails({ emailIds: 'msg-1', targetFolder: 'Archive' });

    expect(getFolderIdByName).toHaveBeenCalledWith(
      mockAccessToken,
      'Archive',
      null
    );
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'POST',
      'me/messages/msg-1/move',
      { destinationId: 'target-folder-id' }
    );
  });

  it('should handle target folder not found', async () => {
    getFolderIdByName.mockResolvedValue(null);

    const result = await handleMoveEmails({
      emailIds: 'msg-1',
      targetFolder: 'NonExistent',
    });

    expect(result.content[0].text).toContain('not found');
  });

  it('should handle partial failures', async () => {
    getFolderIdByName.mockResolvedValue('target-id');
    callGraphAPI
      .mockResolvedValueOnce({}) // First email succeeds
      .mockRejectedValueOnce(new Error('Move failed')); // Second fails

    const result = await handleMoveEmails({
      emailIds: 'msg-1,msg-2',
      targetFolder: 'Archive',
    });

    expect(result.content[0].text).toContain('Successfully moved 1');
    expect(result.content[0].text).toContain('Failed to move 1');
  });

  it('should require email IDs', async () => {
    const result = await handleMoveEmails({ targetFolder: 'Archive' });

    expect(result.content[0].text).toContain('Email IDs are required');
  });

  it('should require target folder', async () => {
    const result = await handleMoveEmails({ emailIds: 'msg-1' });

    expect(result.content[0].text).toContain('Target folder name is required');
  });

  it('should handle auth error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleMoveEmails({
      emailIds: 'msg-1',
      targetFolder: 'Archive',
    });

    expect(result.content[0].text).toContain('Authentication required');
  });
});

describe('handleGetFolderStats', () => {
  it('should return folder statistics', async () => {
    // resolveFolderName now delegates to the shared resolveFolderRef
    resolveFolderRef.mockResolvedValue('inbox-id');
    callGraphAPI
      .mockResolvedValueOnce({
        // folder details
        id: 'inbox-id',
        displayName: 'Inbox',
        totalItemCount: 42,
        unreadItemCount: 5,
        childFolderCount: 0,
      })
      .mockResolvedValueOnce({
        // newest email
        value: [{ receivedDateTime: '2024-01-15T10:00:00Z' }],
      })
      .mockResolvedValueOnce({
        // oldest email
        value: [{ receivedDateTime: '2024-01-01T08:00:00Z' }],
      });

    const result = await handleGetFolderStats({});

    expect(result.content[0].text).toContain('Inbox');
    expect(result.content[0].text).toContain('42');
    expect(result._meta.totalItems).toBe(42);
    // Default mailbox: folder details fetched from `me`
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/mailFolders/inbox-id',
      null,
      expect.any(Object)
    );
  });

  it('should return stats for a shared mailbox', async () => {
    resolveFolderRef.mockResolvedValue('shared-folder-id');
    callGraphAPI
      .mockResolvedValueOnce({
        id: 'shared-folder-id',
        displayName: 'Lieferanten',
        totalItemCount: 7,
        unreadItemCount: 1,
        childFolderCount: 0,
      })
      .mockResolvedValueOnce({
        value: [{ receivedDateTime: '2024-02-15T10:00:00Z' }],
      })
      .mockResolvedValueOnce({
        value: [{ receivedDateTime: '2024-02-01T08:00:00Z' }],
      });

    const result = await handleGetFolderStats({
      folder: 'Lieferanten',
      sharedMailbox: 'shared@company.com',
    });

    expect(result.content[0].text).toContain('Lieferanten');
    // Folder resolution must be scoped to the shared mailbox
    expect(resolveFolderRef).toHaveBeenCalledWith(
      mockAccessToken,
      'Lieferanten',
      'shared@company.com'
    );
    // Folder details GET must target the shared mailbox
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'users/shared@company.com/mailFolders/shared-folder-id',
      null,
      expect.any(Object)
    );
    // Date-range fetch must also target the shared mailbox
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'users/shared@company.com/mailFolders/shared-folder-id/messages',
      null,
      expect.any(Object)
    );
  });

  it('should accept email as an alias for sharedMailbox', async () => {
    resolveFolderRef.mockResolvedValue('shared-folder-id');
    callGraphAPI.mockResolvedValueOnce({
      id: 'shared-folder-id',
      displayName: 'Lieferanten',
      totalItemCount: 0,
      unreadItemCount: 0,
    });

    await handleGetFolderStats({
      folder: 'Lieferanten',
      email: 'shared@company.com',
      outputVerbosity: 'minimal',
    });

    expect(resolveFolderRef).toHaveBeenCalledWith(
      mockAccessToken,
      'Lieferanten',
      'shared@company.com'
    );
  });

  it('should handle folder not found', async () => {
    // Resolver returns null for an unknown folder
    resolveFolderRef.mockResolvedValue(null);

    const result = await handleGetFolderStats({ folder: 'NonExistent' });

    expect(result.content[0].text).toContain('not found');
  });

  it('should handle minimal verbosity', async () => {
    resolveFolderRef.mockResolvedValue('inbox-id');
    callGraphAPI.mockResolvedValueOnce({
      // folder details
      id: 'inbox-id',
      displayName: 'Inbox',
      totalItemCount: 10,
      unreadItemCount: 2,
    });
    // minimal verbosity skips date range fetch

    const result = await handleGetFolderStats({
      folder: 'inbox',
      outputVerbosity: 'minimal',
    });

    expect(result.content[0].text).toMatch(/Inbox.*10 items.*2 unread/);
  });

  it('should handle auth error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleGetFolderStats({});

    expect(result.content[0].text).toContain('Authentication required');
  });

  it('should handle API error', async () => {
    // resolveFolderName succeeds, but folder details call fails
    resolveFolderRef.mockResolvedValue('inbox-id');
    callGraphAPI.mockRejectedValueOnce(new Error('Stats failed')); // folder details

    const result = await handleGetFolderStats({});

    expect(result.content[0].text).toBe(
      'Error getting folder stats: Stats failed'
    );
  });
});

describe('handleDeleteFolder', () => {
  it('should delete a folder by name in the signed-in mailbox (me)', async () => {
    getFolderIdByName.mockResolvedValue('custom-folder-id');
    callGraphAPI.mockResolvedValue({});

    const result = await handleDeleteFolder({ folderName: 'Old Project' });

    expect(result.content[0].text).toContain('deleted successfully');
    // Name resolution scoped to `me`
    expect(getFolderIdByName).toHaveBeenCalledWith(
      mockAccessToken,
      'Old Project',
      null
    );
    // DELETE issued against `me`
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'DELETE',
      'me/mailFolders/custom-folder-id'
    );
  });

  it('should delete a folder inside a shared mailbox', async () => {
    getFolderIdByName.mockResolvedValue('shared-folder-id');
    callGraphAPI.mockResolvedValue({});

    const result = await handleDeleteFolder({
      folderName: 'Old Project',
      sharedMailbox: 'shared@company.com',
    });

    expect(result.content[0].text).toContain('deleted successfully');
    // Name resolution must be scoped to the shared mailbox
    expect(getFolderIdByName).toHaveBeenCalledWith(
      mockAccessToken,
      'Old Project',
      'shared@company.com'
    );
    // DELETE must target the shared mailbox, NOT `me`
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'DELETE',
      'users/shared@company.com/mailFolders/shared-folder-id'
    );
  });

  it('should accept email as an alias for sharedMailbox', async () => {
    getFolderIdByName.mockResolvedValue('shared-folder-id');
    callGraphAPI.mockResolvedValue({});

    await handleDeleteFolder({
      folderName: 'Old Project',
      email: 'shared@company.com',
    });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'DELETE',
      'users/shared@company.com/mailFolders/shared-folder-id'
    );
  });

  it('should delete by raw folderId against the shared mailbox without resolving', async () => {
    callGraphAPI.mockResolvedValue({});

    await handleDeleteFolder({
      folderId: 'raw-id-123',
      sharedMailbox: 'shared@company.com',
    });

    expect(getFolderIdByName).not.toHaveBeenCalled();
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'DELETE',
      'users/shared@company.com/mailFolders/raw-id-123'
    );
  });

  it('should block protected folders regardless of sharedMailbox', async () => {
    const result = await handleDeleteFolder({
      folderName: 'Inbox',
      sharedMailbox: 'shared@company.com',
    });

    expect(result.content[0].text).toContain('Cannot delete protected folder');
    expect(getFolderIdByName).not.toHaveBeenCalled();
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  it('should report when the folder name cannot be resolved', async () => {
    getFolderIdByName.mockResolvedValue(null);

    const result = await handleDeleteFolder({
      folderName: 'Ghost',
      sharedMailbox: 'shared@company.com',
    });

    expect(result.content[0].text).toContain('not found');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  it('should require folderId or folderName', async () => {
    const result = await handleDeleteFolder({});

    expect(result.content[0].text).toContain(
      'Either folderId or folderName is required'
    );
  });

  it('should handle auth error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleDeleteFolder({ folderName: 'Old Project' });

    expect(result.content[0].text).toContain('Authentication required');
  });
});
