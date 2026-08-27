# mcp-outlook-mail

Outlook Mail (Microsoft 365) MCP Pack

Part of [Pipeworx](https://pipeworx.io) — an MCP gateway connecting AI agents to 1481+ live data sources.

## Tools

| Tool | Description |
|------|-------------|
| `outlook_list_messages` | List email messages from an Outlook / Microsoft 365 mailbox folder (default: inbox). Returns compact message summaries (subject, sender, received time, preview, read status, attachment flag). Supports keyword search and unread-only filtering. Use to browse a user's Outlook email. |
| `outlook_get_message` | Get the full content of a single Outlook / Microsoft 365 email message by its ID, including subject, sender, recipients (to/cc), received time, and the full body (HTML or text). Use after outlook_list_messages or outlook_search_messages to read an email in full. |
| `outlook_list_folders` | List the mail folders in an Outlook / Microsoft 365 mailbox, with display names, unread counts, and total item counts. Use to discover folder names/IDs before listing messages. |
| `outlook_search_messages` | Search across an entire Outlook / Microsoft 365 mailbox by keyword. Returns matching email summaries (subject, sender, received time, preview). Use to find emails by topic, sender, or content regardless of folder. |
| `outlook_get_profile` | Get the signed-in user's Microsoft 365 / Outlook profile: display name, email address, user principal name, job title, and mobile phone. Use to identify whose mailbox is connected. |

## Quick Start

Add to your MCP client (Claude Desktop, Cursor, Windsurf, etc.):

```json
{
  "mcpServers": {
    "outlook_mail": {
      "url": "https://gateway.pipeworx.io/outlook_mail/mcp"
    }
  }
}
```

### What this endpoint actually serves

`tools/list` at `https://gateway.pipeworx.io/outlook_mail/mcp` returns the tools in the table
above **plus the shared Pipeworx meta-tools** — `ask_pipeworx`,
`discover_tools`, `search_within`, `remember`/`recall` and the rest of the
gateway-wide set. So the tool count you see is larger than this table: a
single-pack endpoint currently lists roughly 30 shared tools alongside the
pack's own. The connection's `initialize` response states its exact scope, and
is the authoritative answer for a given day.

This is deliberate, not multiplexing by accident. The meta-tools are what let a
scoped connection answer a question this pack does not cover — via
`ask_pipeworx`, which routes across the whole catalog — without you adding a
second MCP server. There is currently no way to mount a pack endpoint without
them; if the extra schemas cost you more context than the routing is worth,
connect to the full gateway once rather than to several pack endpoints.

Or connect to the full Pipeworx gateway to get every pack's tools listed
directly, instead of just this one's:

```json
{
  "mcpServers": {
    "pipeworx": {
      "url": "https://gateway.pipeworx.io/mcp"
    }
  }
}
```

Both URLs reach the same gateway and the same 1481+ data sources. The
only difference is which pack's tools are listed **directly**; `ask_pipeworx`
reaches all of them from either one.

## Using with ask_pipeworx

Instead of calling tools directly, you can ask questions in plain English —
this works on the pack endpoint above as well as on the full gateway:

```
ask_pipeworx({ question: "your question about Outlook Mail data" })
```

The gateway picks the right tool and fills the arguments automatically.

## More

- [Docs and guides](https://pipeworx.io/docs)
- [pipeworx.io](https://pipeworx.io)

## License

MIT
