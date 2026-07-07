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
