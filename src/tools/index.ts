import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { Request } from 'express'
import {
  registerListEmails,
  registerReadEmail,
  registerSearchEmails,
  registerListFolders,
  registerGetAttachments
} from './mail'
import {
  registerSendEmail,
  registerReplyEmail,
  registerCreateDraft,
  registerMoveEmail,
  registerMarkEmail,
  registerGetProfile
} from './send'

/**
 * Static tool catalog for the public GET /tools endpoint.
 * Lets portals render an accurate tool picker BEFORE any account is
 * connected (the authenticated path is MCP tools/list on /mcp).
 * Keep in sync with registerAllTools below.
 */
export const TOOL_SUMMARIES = [
  { name: 'getProfile',     readOnly: true,  description: "Connected mailbox owner's profile (name, email)" },
  { name: 'listEmails',     readOnly: true,  description: 'List a folder’s messages with unread/sender filters' },
  { name: 'readEmail',      readOnly: true,  description: 'Full body of one message' },
  { name: 'searchEmails',   readOnly: true,  description: 'Full-text search across all folders (KQL)' },
  { name: 'listFolders',    readOnly: true,  description: 'Folder tree with unread counts' },
  { name: 'getAttachments', readOnly: true,  description: 'Attachment names, types, sizes' },
  { name: 'sendEmail',      readOnly: false, description: 'Send a new email' },
  { name: 'replyEmail',     readOnly: false, description: 'Reply / reply-all in-thread' },
  { name: 'createDraft',    readOnly: false, description: 'Create a draft without sending' },
  { name: 'moveEmail',      readOnly: false, description: 'Move message to folder / archive / trash' },
  { name: 'markEmail',      readOnly: false, description: 'Mark read/unread, flag for follow-up' }
]

export function registerAllTools(
  server: McpServer,
  req: Request,
  sessionId: string = 'default'
): void {
  registerGetProfile(server, req, sessionId)
  registerListEmails(server, req, sessionId)
  registerReadEmail(server, req, sessionId)
  registerSearchEmails(server, req, sessionId)
  registerListFolders(server, req, sessionId)
  registerGetAttachments(server, req, sessionId)
  registerSendEmail(server, req, sessionId)
  registerReplyEmail(server, req, sessionId)
  registerCreateDraft(server, req, sessionId)
  registerMoveEmail(server, req, sessionId)
  registerMarkEmail(server, req, sessionId)
  console.log('[MCP] 11 Outlook tools registered')
}
