/**
 * List folders functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List folders handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFolders(args) {
  const includeItemCounts = args.includeItemCounts === true;
  const includeChildren = args.includeChildren === true;
  // Target a shared/delegated mailbox instead of the signed-in account.
  const sharedMailbox = args.sharedMailbox || args.email || null;

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Get all mail folders (signed-in account or a shared mailbox)
    const folders = await getAllFoldersHierarchy(
      accessToken,
      includeItemCounts,
      sharedMailbox
    );

    const heading = sharedMailbox ? `\n\nMailbox: ${sharedMailbox}` : '';

    // If including children, format as hierarchy
    if (includeChildren) {
      return {
        content: [
          {
            type: 'text',
            text: formatFolderHierarchy(folders, includeItemCounts) + heading,
          },
        ],
      };
    } else {
      // Otherwise, format as flat list
      return {
        content: [
          {
            type: 'text',
            text: formatFolderList(folders, includeItemCounts) + heading,
          },
        ],
      };
    }
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [
          {
            type: 'text',
            text: "Authentication required. Please use the 'authenticate' tool first.",
          },
        ],
      };
    }

    return {
      content: [
        {
          type: 'text',
          text: `Error listing folders: ${error.message}`,
        },
      ],
    };
  }
}

/**
 * Get all mail folders with hierarchy information
 * @param {string} accessToken - Access token
 * @param {boolean} includeItemCounts - Include item counts in response
 * @param {string|null} [sharedMailbox] - Shared mailbox email, or null for the signed-in account
 * @returns {Promise<Array>} - Array of folder objects with hierarchy
 */
async function getAllFoldersHierarchy(
  accessToken,
  includeItemCounts,
  sharedMailbox = null
) {
  const prefix = sharedMailbox ? `users/${sharedMailbox}` : 'me';

  try {
    // Determine select fields based on whether to include counts
    const selectFields = includeItemCounts
      ? 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
      : 'id,displayName,parentFolderId,childFolderCount';

    // Get all mail folders
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `${prefix}/mailFolders`,
      null,
      {
        $top: 100,
        $select: selectFields,
      }
    );

    if (!response.value) {
      return [];
    }

    const topLevelFolders = response.value.map((folder) => ({
      ...folder,
      isTopLevel: true,
    }));

    // Recursively gather descendants for any folder reporting children.
    async function gatherChildren(folder) {
      let childResponse;
      try {
        childResponse = await callGraphAPI(
          accessToken,
          'GET',
          `${prefix}/mailFolders/${folder.id}/childFolders`,
          null,
          { $top: 100, $select: selectFields }
        );
      } catch (error) {
        console.error(
          `Error getting child folders for "${folder.displayName}": ${error.message}`
        );
        return [];
      }

      const children = childResponse.value || [];
      const result = [];
      for (const child of children) {
        child.parentFolder = folder.displayName;
        result.push(child);
        if (child.childFolderCount > 0) {
          const grandchildren = await gatherChildren(child);
          result.push(...grandchildren);
        }
      }
      return result;
    }

    const allChildFolders = [];
    for (const folder of topLevelFolders.filter(
      (f) => f.childFolderCount > 0
    )) {
      const descendants = await gatherChildren(folder);
      allChildFolders.push(...descendants);
    }

    // Combine all folders
    return [...topLevelFolders, ...allChildFolders];
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    throw error;
  }
}

/**
 * Format folders as a flat list
 * @param {Array} folders - Array of folder objects
 * @param {boolean} includeItemCounts - Whether to include item counts
 * @returns {string} - Formatted list
 */
function formatFolderList(folders, includeItemCounts) {
  if (!folders || folders.length === 0) {
    return 'No folders found.';
  }

  // Sort folders alphabetically, with well-known folders first
  const wellKnownFolderNames = [
    'Inbox',
    'Drafts',
    'Sent Items',
    'Deleted Items',
    'Junk Email',
    'Archive',
  ];

  const sortedFolders = [...folders].sort((a, b) => {
    // Well-known folders come first
    const aIsWellKnown = wellKnownFolderNames.includes(a.displayName);
    const bIsWellKnown = wellKnownFolderNames.includes(b.displayName);

    if (aIsWellKnown && !bIsWellKnown) return -1;
    if (!aIsWellKnown && bIsWellKnown) return 1;

    if (aIsWellKnown && bIsWellKnown) {
      // Sort well-known folders by their index in the array
      return (
        wellKnownFolderNames.indexOf(a.displayName) -
        wellKnownFolderNames.indexOf(b.displayName)
      );
    }

    // Sort other folders alphabetically
    return a.displayName.localeCompare(b.displayName);
  });

  // Format each folder
  const folderLines = sortedFolders.map((folder) => {
    let folderInfo = folder.displayName;

    // Add parent folder info if available
    if (folder.parentFolder) {
      folderInfo += ` (in ${folder.parentFolder})`;
    }

    // Add item counts if requested
    if (includeItemCounts) {
      const unreadCount = folder.unreadItemCount || 0;
      const totalCount = folder.totalItemCount || 0;
      folderInfo += ` - ${totalCount} items`;

      if (unreadCount > 0) {
        folderInfo += ` (${unreadCount} unread)`;
      }
    }

    return folderInfo;
  });

  return `Found ${folders.length} folders:\n\n${folderLines.join('\n')}`;
}

/**
 * Format folders as a hierarchical tree
 * @param {Array} folders - Array of folder objects
 * @param {boolean} includeItemCounts - Whether to include item counts
 * @returns {string} - Formatted hierarchy
 */
function formatFolderHierarchy(folders, includeItemCounts) {
  if (!folders || folders.length === 0) {
    return 'No folders found.';
  }

  // Build folder hierarchy
  const folderMap = new Map();
  const rootFolders = [];

  // First pass: create map of all folders
  folders.forEach((folder) => {
    folderMap.set(folder.id, {
      ...folder,
      children: [],
    });

    if (folder.isTopLevel) {
      rootFolders.push(folder.id);
    }
  });

  // Second pass: build hierarchy
  folders.forEach((folder) => {
    if (!folder.isTopLevel && folder.parentFolderId) {
      const parent = folderMap.get(folder.parentFolderId);
      if (parent) {
        parent.children.push(folder.id);
      } else {
        // Fallback for orphaned folders
        rootFolders.push(folder.id);
      }
    }
  });

  // Format hierarchy recursively
  function formatSubtree(folderId, level = 0) {
    const folder = folderMap.get(folderId);
    if (!folder) return '';

    const indent = '  '.repeat(level);
    let line = `${indent}${folder.displayName}`;

    // Add item counts if requested
    if (includeItemCounts) {
      const unreadCount = folder.unreadItemCount || 0;
      const totalCount = folder.totalItemCount || 0;
      line += ` - ${totalCount} items`;

      if (unreadCount > 0) {
        line += ` (${unreadCount} unread)`;
      }
    }

    // Add children
    const childLines = folder.children
      .map((childId) => formatSubtree(childId, level + 1))
      .filter((childLine) => childLine.length > 0)
      .join('\n');

    return childLines.length > 0 ? `${line}\n${childLines}` : line;
  }

  // Format all root folders
  const formattedHierarchy = rootFolders
    .map((folderId) => formatSubtree(folderId))
    .join('\n');

  return `Folder Hierarchy:\n\n${formattedHierarchy}`;
}

module.exports = handleListFolders;
