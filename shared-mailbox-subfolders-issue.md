# Cannot list or access custom subfolders of a shared/delegated mailbox

## Summary

The Outlook MCP tools can read a shared mailbox's *well-known* folders but provide no
way to discover or open *custom* subfolders within it. There is no tool that enumerates
a shared mailbox's folder tree, and the shared-mailbox reader only accepts well-known
folder names — so custom folders (and localized folder names) are unreachable.

## Environment

- Outlook MCP connector (tools: `access-shared-mailbox`, `search-emails`, `folders`, `read-email`, `auth`)
- Shared/delegated mailbox accessed by a signed-in user with delegate rights

## What works

- `access-shared-mailbox` resolves Microsoft well-known folder names: `inbox`,
  `sentitems`, `archive`, `junkemail`, `deleteditems`, `drafts`.

## What doesn't work

1. **No folder enumeration for shared mailboxes.** `folders` (action `list`, even with
   `includeChildren`) only returns the *signed-in account's* folder tree. It has no
   parameter to target a shared mailbox, so the shared mailbox's folders/subfolders
   can't be listed.

2. **Custom subfolders can't be opened by name.** `access-shared-mailbox` with a custom
   folder name (e.g. an inbox subfolder) or a localized name (e.g. `Archiv`, `sent`)
   fails with:

   ```
   API call failed with status 400: {"error":{"code":"ErrorInvalidIdMalformed","message":"Id is malformed."}}
   ```

   It appears the `folder` value is passed through as a folder ID when it isn't a
   recognized well-known name, producing a malformed-ID error.

3. **No way to obtain folder IDs.** Because (1) can't enumerate the shared mailbox,
   there's no path to discover the folder IDs that (2) would need.

## Impact

Users who organize a shared mailbox with custom subfolders (e.g. per-vendor,
per-project) cannot have those folders read, searched, or triaged. Only top-level
well-known folders are usable. This is a functional gap, not a permissions issue — the
existing delegation already grants access; the tooling just can't address the folders.

## Steps to reproduce

1. Use a shared mailbox that has custom subfolders under Inbox.
2. `folders` action `list`, `includeChildren: true` → returns only the signed-in
   account's tree, not the shared mailbox's.
3. `access-shared-mailbox` with `email: <shared mailbox>`, `folder: <custom subfolder name>`
   → `ErrorInvalidIdMalformed`.

## Expected behavior

- A way to enumerate a shared mailbox's full folder hierarchy (e.g. a
  `sharedMailbox`/`email` parameter on the `folders` tool), and/or
- `access-shared-mailbox` and `search-emails` should accept a custom folder path or
  folder ID scoped to the shared mailbox and resolve it correctly.

## Suggested fix

Add shared-mailbox targeting to folder enumeration (Graph: `/users/{mailbox}/mailFolders`
+ `/childFolders`) and accept resolved folder IDs / folder paths in the shared-mailbox
read and search tools.
