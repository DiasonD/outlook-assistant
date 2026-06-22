/**
 * Email folder utilities
 *
 * Folder name/path → ID resolution. All resolvers accept an optional `mailbox`
 * argument so the same logic works against the signed-in user (`me`) or a
 * shared/delegated mailbox (`users/{email}`). This is what lets the shared
 * mailbox tools reach custom subfolders and localized folder names rather than
 * only the handful of Microsoft well-known folders.
 */
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Cache of folder information to reduce API calls
 * Format: { userId: { folderName: { id, path } } }
 */
const _folderCache = {};

/**
 * Well-known folder name → Graph folder segment.
 *
 * Includes Graph API well-known names (sentitems, deleteditems, junkemail),
 * common display names (Sent Items, Deleted Items, Junk Email), and short
 * aliases (sent, deleted, junk) for consistent resolution across tools. These
 * segments are valid in any mailbox, so `mailFolders/inbox` works for both
 * `me` and `users/{shared}`.
 */
const WELL_KNOWN_SEGMENTS = {
  inbox: 'inbox',
  drafts: 'drafts',
  sent: 'sentItems',
  sentitems: 'sentItems',
  'sent items': 'sentItems',
  deleted: 'deletedItems',
  deleteditems: 'deletedItems',
  'deleted items': 'deletedItems',
  junk: 'junkemail',
  junkemail: 'junkemail',
  'junk email': 'junkemail',
  spam: 'junkemail',
  archive: 'archive',
  outbox: 'outbox',
};

/**
 * Well-known folder names mapped to their `me` message endpoints.
 * Preserved for backwards compatibility with existing callers/tests.
 */
const WELL_KNOWN_FOLDERS = Object.fromEntries(
  Object.entries(WELL_KNOWN_SEGMENTS).map(([name, segment]) => [
    name,
    `me/mailFolders/${segment}/messages`,
  ])
);

/**
 * Build the Graph resource prefix for a mailbox.
 * @param {string|null} [mailbox] - Shared mailbox email address, or null/empty for the signed-in user
 * @returns {string} - `me` or `users/{mailbox}`
 */
function buildMailboxPrefix(mailbox) {
  if (!mailbox) {
    return 'me';
  }
  // Already a fully-qualified prefix (e.g. caller passed `users/x`)
  if (mailbox === 'me' || mailbox.startsWith('users/')) {
    return mailbox;
  }
  // Email addresses are valid in a Graph path segment unencoded; this matches
  // the long-standing `users/{mailbox}` construction used elsewhere.
  return `users/${mailbox}`;
}

/**
 * Escape a value for use inside an OData single-quoted string literal.
 * @param {string} value
 * @returns {string}
 */
function escapeODataString(value) {
  return String(value).replace(/'/g, "''");
}

/**
 * Heuristic: does this string look like a raw Graph mail-folder ID rather than
 * a display name? Graph folder IDs are long, contain no spaces, and are made up
 * of base64url/base64 characters. We only fall back to treating an unresolved
 * reference as an ID when it plausibly is one — otherwise a typo would be
 * forwarded to Graph and produce the opaque "Id is malformed" error this module
 * exists to avoid.
 * @param {string} value
 * @returns {boolean}
 */
function looksLikeFolderId(value) {
  if (!value || /\s/.test(value)) {
    return false;
  }
  // Well-known segment names are short and explicitly handled elsewhere.
  return value.length >= 40 && /^[A-Za-z0-9_\-=+/.]+$/.test(value);
}

/**
 * Search a folder's descendant tree for a child folder matching `lowerName`.
 * Recurses only into folders that report child folders, so a shallow tree
 * costs a single request per child-bearing branch.
 * @param {string} accessToken
 * @param {string} parentId - Folder ID (or well-known segment) to search under
 * @param {string} lowerName - Lower-cased display name to match
 * @param {string} prefix - Mailbox prefix (`me` or `users/{email}`)
 * @returns {Promise<string|null>}
 */
async function findChildFolderId(accessToken, parentId, lowerName, prefix) {
  let response;
  try {
    response = await callGraphAPI(
      accessToken,
      'GET',
      `${prefix}/mailFolders/${parentId}/childFolders`,
      null,
      { $top: 100 }
    );
  } catch (error) {
    console.error(
      `Error listing child folders of "${parentId}": ${error.message}`
    );
    return null;
  }

  const children = response.value || [];

  // Direct children first (breadth-first keeps the closest match)
  for (const child of children) {
    if (child.displayName && child.displayName.toLowerCase() === lowerName) {
      return child.id;
    }
  }

  // Then recurse into branches that have their own children
  for (const child of children.filter((c) => c.childFolderCount > 0)) {
    const deepMatch = await findChildFolderId(
      accessToken,
      child.id,
      lowerName,
      prefix
    );
    if (deepMatch) {
      return deepMatch;
    }
  }

  return null;
}

/**
 * Get the ID of a mail folder by its display name.
 *
 * Resolution order: exact-match $filter on top-level folders, case-insensitive
 * scan of top-level folders, then a recursive scan of every subfolder. The
 * recursive step is what makes custom subfolders (e.g. an Inbox subfolder)
 * reachable by name.
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name of the folder to find
 * @param {string|null} [mailbox] - Shared mailbox to search, or null for the signed-in user
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName, mailbox = null) {
  const prefix = buildMailboxPrefix(mailbox);
  try {
    // First try with exact match filter
    console.error(`Looking for folder with name "${folderName}"`);
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `${prefix}/mailFolders`,
      null,
      { $filter: `displayName eq '${escapeODataString(folderName)}'` }
    );

    if (response.value && response.value.length > 0) {
      console.error(
        `Found folder "${folderName}" with ID: ${response.value[0].id}`
      );
      return response.value[0].id;
    }

    // If exact match fails, get all top-level folders for a case-insensitive
    // comparison, then recurse into any subfolders.
    console.error(
      `No exact match found for "${folderName}", trying case-insensitive search`
    );
    const allFoldersResponse = await callGraphAPI(
      accessToken,
      'GET',
      `${prefix}/mailFolders`,
      null,
      { $top: 100 }
    );

    const topLevel = allFoldersResponse.value || [];
    const lowerFolderName = folderName.toLowerCase();

    const matchingFolder = topLevel.find(
      (folder) =>
        folder.displayName &&
        folder.displayName.toLowerCase() === lowerFolderName
    );
    if (matchingFolder) {
      console.error(
        `Found case-insensitive match for "${folderName}" with ID: ${matchingFolder.id}`
      );
      return matchingFolder.id;
    }

    // Recurse into subfolders of any folder that reports children.
    for (const parent of topLevel.filter((f) => f.childFolderCount > 0)) {
      const childMatch = await findChildFolderId(
        accessToken,
        parent.id,
        lowerFolderName,
        prefix
      );
      if (childMatch) {
        console.error(
          `Found nested match for "${folderName}" with ID: ${childMatch}`
        );
        return childMatch;
      }
    }

    console.error(`No folder found matching "${folderName}"`);
    return null;
  } catch (error) {
    console.error(`Error finding folder "${folderName}": ${error.message}`);
    return null;
  }
}

/**
 * Resolve a slash- or backslash-separated folder path (e.g. `Inbox/Vendors/Acme`)
 * to a folder ID by walking the tree one segment at a time. The first segment
 * may be a well-known folder name; subsequent segments are matched against
 * child-folder display names (case-insensitive).
 * @param {string} accessToken
 * @param {string} folderPath
 * @param {string} prefix - Mailbox prefix
 * @returns {Promise<string|null>}
 */
async function resolveFolderPathSegments(accessToken, folderPath, prefix) {
  const parts = folderPath
    .split(/[\\/]+/)
    .map((p) => p.trim())
    .filter(Boolean);

  if (parts.length === 0) {
    return null;
  }

  // Resolve the first segment: well-known name or a top-level display name.
  let currentId;
  const firstLower = parts[0].toLowerCase();
  if (WELL_KNOWN_SEGMENTS[firstLower]) {
    currentId = WELL_KNOWN_SEGMENTS[firstLower];
  } else {
    currentId = await getTopLevelFolderId(accessToken, parts[0], prefix);
    if (!currentId) {
      return null;
    }
  }

  // Walk the remaining segments through child folders.
  for (let i = 1; i < parts.length; i++) {
    const nextId = await getDirectChildFolderId(
      accessToken,
      currentId,
      parts[i],
      prefix
    );
    if (!nextId) {
      return null;
    }
    currentId = nextId;
  }

  return currentId;
}

/**
 * Find a top-level folder ID by display name (exact then case-insensitive).
 */
async function getTopLevelFolderId(accessToken, name, prefix) {
  try {
    const exact = await callGraphAPI(
      accessToken,
      'GET',
      `${prefix}/mailFolders`,
      null,
      { $filter: `displayName eq '${escapeODataString(name)}'` }
    );
    if (exact.value && exact.value.length > 0) {
      return exact.value[0].id;
    }
    const all = await callGraphAPI(
      accessToken,
      'GET',
      `${prefix}/mailFolders`,
      null,
      { $top: 100 }
    );
    const lower = name.toLowerCase();
    const match = (all.value || []).find(
      (f) => f.displayName && f.displayName.toLowerCase() === lower
    );
    return match ? match.id : null;
  } catch (error) {
    console.error(`Error finding top-level folder "${name}": ${error.message}`);
    return null;
  }
}

/**
 * Find a direct child folder ID by display name under a given parent.
 */
async function getDirectChildFolderId(accessToken, parentId, name, prefix) {
  try {
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `${prefix}/mailFolders/${parentId}/childFolders`,
      null,
      { $top: 100 }
    );
    const lower = name.toLowerCase();
    const match = (response.value || []).find(
      (f) => f.displayName && f.displayName.toLowerCase() === lower
    );
    return match ? match.id : null;
  } catch (error) {
    console.error(
      `Error finding child folder "${name}" under "${parentId}": ${error.message}`
    );
    return null;
  }
}

/**
 * Resolve a folder reference to an identifier usable in
 * `${prefix}/mailFolders/{ref}`. Accepts well-known names, folder paths
 * (`A/B/C`), single display names (searched recursively), and raw folder IDs.
 * @param {string} accessToken
 * @param {string} folderRef - Folder name, path, or ID
 * @param {string|null} [mailbox] - Shared mailbox, or null for the signed-in user
 * @returns {Promise<string|null>} - A well-known segment or folder ID, or null if unresolved
 */
async function resolveFolderRef(accessToken, folderRef, mailbox = null) {
  if (!folderRef) {
    return 'inbox';
  }

  const prefix = buildMailboxPrefix(mailbox);
  const lower = folderRef.toLowerCase();

  if (WELL_KNOWN_SEGMENTS[lower]) {
    return WELL_KNOWN_SEGMENTS[lower];
  }

  // Path reference (contains a separator)
  if (/[\\/]/.test(folderRef)) {
    const pathId = await resolveFolderPathSegments(
      accessToken,
      folderRef,
      prefix
    );
    if (pathId) {
      return pathId;
    }
  } else {
    // Single display name — search the whole tree
    const id = await getFolderIdByName(accessToken, folderRef, mailbox);
    if (id) {
      return id;
    }
  }

  // Last resort: the caller may have passed a raw folder ID.
  if (looksLikeFolderId(folderRef)) {
    return folderRef;
  }

  return null;
}

/**
 * Resolve a folder name/path to its message-collection endpoint path.
 * @param {string} accessToken - Access token
 * @param {string} folderName - Folder name, path, or ID to resolve
 * @param {string|null} [mailbox] - Shared mailbox email, or null for the signed-in user
 * @returns {Promise<string>} - Resolved endpoint path (e.g. `me/mailFolders/{id}/messages`)
 */
async function resolveFolderPath(accessToken, folderName, mailbox = null) {
  const prefix = buildMailboxPrefix(mailbox);

  // Default to inbox if no folder specified
  if (!folderName) {
    return `${prefix}/mailFolders/inbox/messages`;
  }

  // Well-known folders resolve without an API call
  const lowerFolderName = folderName.toLowerCase();
  if (WELL_KNOWN_SEGMENTS[lowerFolderName]) {
    console.error(`Using well-known folder path for "${folderName}"`);
    return `${prefix}/mailFolders/${WELL_KNOWN_SEGMENTS[lowerFolderName]}/messages`;
  }

  try {
    const folderId = await resolveFolderRef(accessToken, folderName, mailbox);
    if (folderId) {
      const path = `${prefix}/mailFolders/${folderId}/messages`;
      console.error(`Resolved folder "${folderName}" to path: ${path}`);
      return path;
    }

    // If not found, throw error instead of silently falling back
    throw new Error(
      `Folder "${folderName}" not found. Use the folders tool (action=list) to see available folders.`
    );
  } catch (error) {
    if (error.message.includes('not found')) {
      throw error;
    }
    throw new Error(
      `Error resolving folder "${folderName}": ${error.message}. Use the folders tool (action=list) to see available folders.`,
      { cause: error }
    );
  }
}

/**
 * Enumerate a mailbox's full folder tree (recursively) as a flat array.
 *
 * Each returned folder carries `isTopLevel`, `parentFolderId`, and (for
 * non-top-level folders) a `parentFolder` display-name and `folderPath` like
 * `Inbox/Vendors/Acme` so callers can render a hierarchy and resolve nested
 * folders.
 * @param {string} accessToken
 * @param {object} [options]
 * @param {string|null} [options.mailbox] - Shared mailbox email, or null for the signed-in user
 * @param {boolean} [options.includeItemCounts] - Include total/unread counts
 * @returns {Promise<Array>} - Flat array of folder objects
 */
async function getAllFolders(accessToken, options = {}) {
  const { mailbox = null, includeItemCounts = true } = options;
  const prefix = buildMailboxPrefix(mailbox);
  const selectFields = includeItemCounts
    ? 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
    : 'id,displayName,parentFolderId,childFolderCount';

  const collected = [];

  async function walk(endpoint, parentDisplayName, parentPath, isTopLevel) {
    let response;
    try {
      response = await callGraphAPI(accessToken, 'GET', endpoint, null, {
        $top: 100,
        $select: selectFields,
      });
    } catch (error) {
      console.error(`Error listing folders at "${endpoint}": ${error.message}`);
      return;
    }

    const folders = response.value || [];
    for (const folder of folders) {
      const folderPath = parentPath
        ? `${parentPath}/${folder.displayName}`
        : folder.displayName;

      collected.push({
        ...folder,
        isTopLevel,
        parentFolder: parentDisplayName || undefined,
        folderPath,
      });

      if (folder.childFolderCount > 0) {
        await walk(
          `${prefix}/mailFolders/${folder.id}/childFolders`,
          folder.displayName,
          folderPath,
          false
        );
      }
    }
  }

  await walk(`${prefix}/mailFolders`, null, '', true);
  return collected;
}

module.exports = {
  WELL_KNOWN_FOLDERS,
  WELL_KNOWN_SEGMENTS,
  buildMailboxPrefix,
  resolveFolderPath,
  resolveFolderRef,
  getFolderIdByName,
  getAllFolders,
};
