/**
 * Mailbox scoping helper.
 *
 * Every Graph path in this server is built as `${prefix}/...`. The prefix is
 * `me` for the signed-in account, or `users/{email}` for a shared/delegated
 * mailbox. Keeping the construction in one place is what lets the shared-mailbox
 * parameter be threaded through readers, writers, and folder resolution without
 * each call site re-deciding the shape.
 */

// Pragmatic SMTP address / UPN shape — deliberately not RFC 5322. The point is
// to keep caller input inside a single Graph path segment: something@something
// .something with no whitespace and none of `/ ? # %` (path, query, fragment
// and percent-escape delimiters). The tool schemas advertise an email address
// only, so bare user GUIDs are not accepted.
const MAILBOX_PATTERN = /^[^\s/?#%@]+@[^\s/?#%@]+\.[^\s/?#%@]+$/;

/**
 * Build the Graph resource prefix for a mailbox.
 * @param {string|null} [mailbox] - Shared mailbox email address, or null/empty for the signed-in user
 * @returns {string} - `me` or `users/{mailbox}`
 * @throws {Error} If `mailbox` is non-empty but not a plausible email address
 */
function buildMailboxPrefix(mailbox) {
  const trimmed = typeof mailbox === 'string' ? mailbox.trim() : mailbox;
  if (!trimmed) {
    return 'me';
  }
  if (trimmed === 'me') {
    return 'me';
  }
  if (!MAILBOX_PATTERN.test(trimmed)) {
    throw new Error(
      `Invalid mailbox "${mailbox}" — expected a shared mailbox email address (e.g. "team@contoso.com").`
    );
  }
  // Encode as a single path segment. `@` is a legal path character (RFC 3986
  // pchar) and Graph/the existing endpoints expect it literal, so restore it;
  // a literal `%40` in the input is impossible because `%` is rejected above.
  return `users/${encodeURIComponent(trimmed).replace(/%40/g, '@')}`;
}

module.exports = { buildMailboxPrefix };
