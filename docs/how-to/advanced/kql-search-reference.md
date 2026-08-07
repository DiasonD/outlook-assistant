---
title: "KQL Search Reference for Outlook Assistant"
description: "Use Microsoft Graph $search expressions (KQL-style syntax) to write precise search queries for emails by date, sender, subject, and message properties."
tags: [outlook-assistant, advanced, how-to, reference]
---

# KQL Search Reference

The `searchExpression` parameter lets you pass a raw Microsoft Graph `$search` expression — KQL-style syntax, but **not** full KQL — for precise searches when the standard filter parameters aren't enough.

## Using `searchExpression` in Outlook Assistant

Pass a raw Graph `$search` expression through the `searchExpression` parameter (formerly `kqlQuery`, still accepted as a deprecated alias):

```
tool: search-emails
params:
  searchExpression: "from:sarah@company.com AND subject:quarterly"
```

When you use `searchExpression`, it bypasses other search parameters (`from`, `subject`, etc.) and sends the raw expression directly to the Graph API.

## Basic Keyword Search

Search across all fields:

```
searchExpression: "budget approval"
```

## Search by Field

| Field | Example |
|-------|---------|
| Sender | `from:sarah@company.com` |
| Recipient | `to:team@company.com` |
| Subject | `subject:quarterly report` |
| Body | `body:action required` |
| CC | `cc:manager@company.com` |

> **Personal accounts (issue [#217](https://github.com/littlebearapps/outlook-assistant/issues/217))**: Field-scoped `$search` expressions like `subject:"invoice"` or `from:github.com` are **best-effort** on personal Outlook.com accounts — they often return zero results even when matching emails exist, because the raw `searchExpression` branch does not fall back to other strategies. For reliable personal-account search, prefer the `query` parameter (progressive OData plus client-side fallback) or the structured filters (`from`, `subject`, `to`, `receivedAfter`).

## Combine Conditions

### AND — both must match

```
searchExpression: "from:sarah AND subject:quarterly"
```

### OR — either must match

```
searchExpression: "from:sarah OR from:james"
```

### NOT — exclude results

```
searchExpression: "subject:report NOT from:noreply@github.com"
```

### Parentheses — group conditions

```
searchExpression: "(from:sarah OR from:james) AND subject:review"
```

## Date Ranges

```
searchExpression: "received>=2026-01-01 AND received<=2026-01-31"
```

```
searchExpression: "sent>=2026-02-01"
```

## Attachment and Flag Filters

```
searchExpression: "hasAttachment:true"
```

```
searchExpression: "isRead:false"
```

## Common Patterns

| Goal | Search Expression |
|------|-------------------|
| Emails from a specific sender this year | `from:boss@company.com AND received>=2026-01-01` |
| Unread emails with attachments | `isRead:false AND hasAttachment:true` |
| Emails about a project from anyone | `subject:\"Project Alpha\" OR body:\"Project Alpha\"` |
| Emails from a domain | `from:@company.com` |
| Recent emails excluding newsletters | `received>=2026-02-01 NOT from:newsletter@` |

## `searchExpression` vs Filter Parameters

| Approach | Personal Account | Work/School Account | When to use |
|----------|-----------------|---------------------|-------------|
| Filter params (`from`, `subject`, etc.) | Full support | Full support | **Recommended default** — reliable on all accounts |
| `query` param | Limited (auto-fallback) | Full support | Free-text search; falls back to filters on personal accounts |
| `searchExpression` param | Best-effort (no fallback) | Full support | Complex queries: AND/OR/NOT, date ranges, multi-field (work accounts only) |

> **Important**: On personal Outlook.com accounts, `query` and `searchExpression` use Microsoft's `$search` API, which is not fully supported and may silently return no results. Unlike `query` — which falls back to OData filters when `$search` comes up empty — the raw `searchExpression` branch does **not** fall back (see issue [#217](https://github.com/littlebearapps/outlook-assistant/issues/217)). Always prefer structured filter parameters (`from`, `subject`, `to`, `receivedAfter`, `hasAttachments`, `unreadOnly`) — these use OData `$filter` and work reliably on all account types.
>
> If you must use `searchExpression`, test with a simple expression first to confirm it works with your account type.

## Tips

- Enclose multi-word phrases in escaped quotes: `subject:\"Project Alpha\"`
- Graph `$search` matching is case-insensitive
- Date format is `YYYY-MM-DD`
- If a `searchExpression` returns no results, try the simpler `query` parameter first — it's more forgiving

## Related

- [Find Emails](../email/find-emails.md) — standard search with filter parameters
- [Find Emails — Search Across All Folders](../email/find-emails.md#search-across-all-folders)
- [Tools Reference — search-emails](../../quickrefs/tools-reference.md#email-6-tools)
