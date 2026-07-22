/**
 * Delete folder functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolder } = require('./resolve');

/**
 * Protected folders that cannot be deleted. `PROTECTED_FOLDERS` is a fast
 * literal-name short-circuit; `PROTECTED_WELL_KNOWN` is the authoritative
 * check applied to the RESOLVED folder's Graph `wellKnownName`, so a path or
 * ID that resolves to a system folder can't bypass the guard. (#216 review)
 */
const PROTECTED_FOLDERS = [
  'inbox',
  'drafts',
  'sentitems',
  'deleteditems',
  'junkemail',
  'archive',
  'outbox',
];
const PROTECTED_WELL_KNOWN = new Set(PROTECTED_FOLDERS);

/**
 * Delete folder handler
 * @param {object} args - Tool arguments
 * @param {string} [args.folderId] - Folder ID to delete
 * @param {string} [args.folderName] - Folder name to delete (resolved to ID)
 * @returns {object} - MCP response
 */
async function handleDeleteFolder(args) {
  const { folderId, folderName } = args;

  if (!folderId && !folderName) {
    return {
      content: [
        {
          type: 'text',
          text: 'Either folderId or folderName is required.',
        },
      ],
    };
  }

  // Fast literal-name guard (clear message for the obvious case).
  if (
    folderName &&
    PROTECTED_FOLDERS.includes(folderName.toLowerCase().trim())
  ) {
    return {
      content: [
        {
          type: 'text',
          text: `Cannot delete protected folder "${folderName}". Protected folders: Inbox, Drafts, Sent Items, Deleted Items, Junk Email, Archive, Outbox.`,
        },
      ],
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Always resolve (by name/path OR explicit ID) so protection can be
    // enforced on the CANONICAL folder identity — a path/ID that lands on a
    // system folder must not slip past the literal-name guard. (#216 review)
    let resolved;
    try {
      resolved = await resolveFolder(accessToken, {
        id: folderId,
        name: folderName,
      });
    } catch (resolveError) {
      return {
        content: [{ type: 'text', text: resolveError.message }],
      };
    }

    if (
      resolved.wellKnownName &&
      PROTECTED_WELL_KNOWN.has(resolved.wellKnownName.toLowerCase())
    ) {
      return {
        content: [
          {
            type: 'text',
            text: `Cannot delete protected folder "${resolved.displayName}" (system folder: ${resolved.wellKnownName}).`,
          },
        ],
      };
    }

    const resolvedId = resolved.id;
    const displayName = resolved.path;

    // Delete the folder
    await callGraphAPI(accessToken, 'DELETE', `me/mailFolders/${resolvedId}`);
    return {
      content: [
        {
          type: 'text',
          text: `Folder "${displayName}" deleted successfully.`,
        },
      ],
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [
          {
            type: 'text',
            text: "Authentication required. Please use the 'auth' tool with action=authenticate first.",
          },
        ],
      };
    }
    return {
      content: [
        {
          type: 'text',
          text: `Error deleting folder: ${error.message}`,
        },
      ],
    };
  }
}

module.exports = handleDeleteFolder;
