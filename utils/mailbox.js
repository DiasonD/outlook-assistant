/**
 * Mailbox scoping helper.
 *
 * Every Graph path in this server is built as `${prefix}/...`. The prefix is
 * `me` for the signed-in account, or `users/{email}` for a shared/delegated
 * mailbox. Keeping the construction in one place is what lets the shared-mailbox
 * parameter be threaded through readers, writers, and folder resolution without
 * each call site re-deciding the shape.
 */

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

module.exports = { buildMailboxPrefix };
