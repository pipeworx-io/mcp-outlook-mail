interface McpToolDefinition {
  name: string;
  description: string;
  inputSchema: {
    type: 'object';
    properties: Record<string, unknown>;
    required?: string[];
  };
}

interface McpToolExport {
  tools: McpToolDefinition[];
  callTool: (name: string, args: Record<string, unknown>) => Promise<unknown>;
  meter?: { credits: number };
  cost?: Record<string, unknown>;
  provider?: string;
}

/**
 * Outlook Mail (Microsoft 365) MCP Pack
 *
 * Requires OAuth connection — gateway injects credentials via _context.microsoft.
 * Read-only access to Outlook / Microsoft 365 email via Microsoft Graph v1.0.
 * Tools: list messages, get message, list folders, search messages, get profile.
 */


interface OutlookMailContext {
  microsoft?: { accessToken: string };
}

const API = 'https://graph.microsoft.com/v1.0';

/**
 * Fetch helper for Microsoft Graph.
 * - Returns { error: 'connection_required' } when no OAuth token is present.
 * - Returns { error: <status>, message: <body text> } on non-2xx responses.
 * - Otherwise returns the parsed JSON body.
 */
async function gFetch(
  ctx: OutlookMailContext,
  url: string,
  options: RequestInit = {},
): Promise<unknown> {
  if (!ctx.microsoft) {
    return {
      error: 'connection_required',
      message: 'Connect your Microsoft 365 account at https://pipeworx.io/account',
    };
  }
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${ctx.microsoft.accessToken}`,
      'Content-Type': 'application/json',
      ...(options.headers ?? {}),
    },
  });
  if (!res.ok) {
    const text = await res.text();
    return { error: res.status, message: text };
  }
  return res.json();
}

const tools: McpToolExport['tools'] = [
  {
    name: 'outlook_list_messages',
    description:
      'List email messages from an Outlook / Microsoft 365 mailbox folder (default: inbox). Returns compact message summaries (subject, sender, received time, preview, read status, attachment flag). Supports keyword search and unread-only filtering. Use to browse a user\'s Outlook email.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        folder: {
          type: 'string',
          description: 'Mail folder to read from (default: "inbox"). Use a well-known name like "inbox", "sentitems", "drafts", "deleteditems", or a folder ID from outlook_list_folders.',
        },
        top: {
          type: 'number',
          description: 'Maximum number of messages to return (default 25, max 100).',
        },
        search: {
          type: 'string',
          description: 'Optional free-text keyword search across the messages (subject, body, sender, etc.).',
        },
        unread_only: {
          type: 'boolean',
          description: 'If true, return only unread messages.',
        },
      },
      required: [],
    },
  },
  {
    name: 'outlook_get_message',
    description:
      'Get the full content of a single Outlook / Microsoft 365 email message by its ID, including subject, sender, recipients (to/cc), received time, and the full body (HTML or text). Use after outlook_list_messages or outlook_search_messages to read an email in full.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        id: {
          type: 'string',
          description: 'The ID of the email message to retrieve (from a list or search result).',
        },
      },
      required: ['id'],
    },
  },
  {
    name: 'outlook_list_folders',
    description:
      'List the mail folders in an Outlook / Microsoft 365 mailbox, with display names, unread counts, and total item counts. Use to discover folder names/IDs before listing messages.',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'outlook_search_messages',
    description:
      'Search across an entire Outlook / Microsoft 365 mailbox by keyword. Returns matching email summaries (subject, sender, received time, preview). Use to find emails by topic, sender, or content regardless of folder.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        query: {
          type: 'string',
          description: 'Free-text search query to match against email messages.',
        },
        top: {
          type: 'number',
          description: 'Maximum number of messages to return (default 25).',
        },
      },
      required: ['query'],
    },
  },
  {
    name: 'outlook_get_profile',
    description:
      'Get the signed-in user\'s Microsoft 365 / Outlook profile: display name, email address, user principal name, job title, and mobile phone. Use to identify whose mailbox is connected.',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
];

interface GraphMessage {
  id?: string;
  subject?: string;
  from?: unknown;
  toRecipients?: unknown;
  ccRecipients?: unknown;
  receivedDateTime?: string;
  bodyPreview?: string;
  body?: { contentType?: string; content?: string };
  isRead?: boolean;
  hasAttachments?: boolean;
}

function compactMessage(m: GraphMessage): Record<string, unknown> {
  return {
    id: m.id,
    subject: m.subject,
    from: m.from,
    receivedDateTime: m.receivedDateTime,
    bodyPreview: m.bodyPreview,
    isRead: m.isRead,
    hasAttachments: m.hasAttachments,
  };
}

async function callTool(name: string, args: Record<string, unknown>): Promise<unknown> {
  const context = (args._context ?? {}) as OutlookMailContext;
  delete args._context;

  switch (name) {
    case 'outlook_list_messages': {
      const folder = (args.folder as string) ?? 'inbox';
      const top = Math.min(100, Math.max(1, (args.top as number) ?? 25));
      const search = args.search as string | undefined;
      const unreadOnly = args.unread_only as boolean | undefined;

      const params = new URLSearchParams({
        $top: String(top),
        $select: 'id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments',
      });
      const headers: Record<string, string> = {};
      if (search) {
        params.set('$search', `"${search}"`);
        headers.ConsistencyLevel = 'eventual';
      }
      if (unreadOnly) {
        params.set('$filter', 'isRead eq false');
      }

      const result = await gFetch(
        context,
        `${API}/me/mailFolders/${encodeURIComponent(folder)}/messages?${params}`,
        { headers },
      );
      const value = (result as { value?: GraphMessage[] }).value;
      if (Array.isArray(value)) return value.map(compactMessage);
      return result;
    }
    case 'outlook_get_message': {
      const id = args.id as string;
      const params = new URLSearchParams({
        $select: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments',
      });
      const result = await gFetch(
        context,
        `${API}/me/messages/${encodeURIComponent(id)}?${params}`,
      );
      const m = result as GraphMessage;
      if (m && typeof m === 'object' && 'id' in m && !('error' in m)) {
        return {
          id: m.id,
          subject: m.subject,
          from: m.from,
          to: m.toRecipients,
          cc: m.ccRecipients,
          receivedDateTime: m.receivedDateTime,
          hasAttachments: m.hasAttachments,
          body: {
            contentType: m.body?.contentType,
            content: m.body?.content,
          },
        };
      }
      return result;
    }
    case 'outlook_list_folders': {
      const params = new URLSearchParams({
        $top: '100',
        $select: 'id,displayName,unreadItemCount,totalItemCount',
      });
      const result = await gFetch(context, `${API}/me/mailFolders?${params}`);
      const value = (result as { value?: unknown[] }).value;
      if (Array.isArray(value)) return value;
      return result;
    }
    case 'outlook_search_messages': {
      const query = args.query as string;
      const top = Math.min(100, Math.max(1, (args.top as number) ?? 25));
      const params = new URLSearchParams({
        $search: `"${query}"`,
        $top: String(top),
        $select: 'id,subject,from,receivedDateTime,bodyPreview',
      });
      const result = await gFetch(context, `${API}/me/messages?${params}`, {
        headers: { ConsistencyLevel: 'eventual' },
      });
      const value = (result as { value?: unknown[] }).value;
      if (Array.isArray(value)) return value;
      return result;
    }
    case 'outlook_get_profile': {
      const params = new URLSearchParams({
        $select: 'displayName,mail,userPrincipalName,jobTitle,mobilePhone',
      });
      return gFetch(context, `${API}/me?${params}`);
    }
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

export default { tools, callTool, meter: { credits: 1 }, provider: 'microsoft' } satisfies McpToolExport;
