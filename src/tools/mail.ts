/**
 * Read-side mail tools: list, read, search, folders, attachments.
 * Every tool resolves the Graph context per call (PATH A bearer or PATH B session).
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import { Request } from 'express'
import { resolveGraphCtx, graphRequest, notAuthenticatedError, graphApiError } from '../lib/graph'

/** Trim a Graph message down to what an LLM actually needs. */
function slimMessage(m: any, includeBody = false) {
  return {
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.address,
    fromName: m.from?.emailAddress?.name,
    to: (m.toRecipients ?? []).map((r: any) => r.emailAddress?.address),
    cc: (m.ccRecipients ?? []).map((r: any) => r.emailAddress?.address),
    receivedDateTime: m.receivedDateTime,
    isRead: m.isRead,
    hasAttachments: m.hasAttachments,
    importance: m.importance,
    bodyPreview: m.bodyPreview,
    webLink: m.webLink,
    conversationId: m.conversationId,
    ...(includeBody && {
      bodyContentType: m.body?.contentType,
      body: m.body?.content
    })
  }
}

export function registerListEmails(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'listEmails',
    `Lists emails from an Outlook mail folder, newest first.

USE THIS FIRST to discover message ids — every other mail tool takes the ids this returns.
Returns metadata + a short bodyPreview only. Call readEmail for the full body.

FOLDERS: inbox (default), sentitems, drafts, deleteditems, archive, junkemail — or a folder id from listFolders.`,
    {
      folder: z.string().optional().describe("Well-known folder name or folder id. Default 'inbox'."),
      top: z.number().min(1).max(50).optional().describe('Max messages to return. Default 10, max 50.'),
      unreadOnly: z.boolean().optional().describe('Only unread messages when true.'),
      from: z.string().optional().describe('Filter: only messages whose sender address equals this email.')
    },
    async ({ folder, top, unreadOnly, from }) => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        const filters: string[] = []
        if (unreadOnly) filters.push('isRead eq false')
        if (from)       filters.push(`from/emailAddress/address eq '${from.replace(/'/g, "''")}'`)
        const params = new URLSearchParams({
          $top: String(top ?? 10),
          $orderby: 'receivedDateTime desc',
          $select: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,importance,bodyPreview,webLink,conversationId'
        })
        if (filters.length) params.set('$filter', filters.join(' and '))
        const data = await graphRequest<any>(ctx, 'GET', `/me/mailFolders/${encodeURIComponent(folder || 'inbox')}/messages?${params}`)
        const messages = (data.value ?? []).map((m: any) => slimMessage(m))
        return { content: [{ type: 'text', text: JSON.stringify({ count: messages.length, messages }) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}

export function registerReadEmail(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'readEmail',
    `Reads one email in full — headers plus complete body (HTML or text).

Use the id returned by listEmails or searchEmails. Marks nothing as read.`,
    {
      messageId: z.string().describe('Full Graph message id from listEmails/searchEmails.'),
      markAsRead: z.boolean().optional().describe('Also mark the message as read. Default false.')
    },
    async ({ messageId, markAsRead }) => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        const m = await graphRequest<any>(ctx, 'GET', `/me/messages/${encodeURIComponent(messageId)}`)
        if (markAsRead && !m.isRead) {
          await graphRequest(ctx, 'PATCH', `/me/messages/${encodeURIComponent(messageId)}`, { isRead: true })
        }
        return { content: [{ type: 'text', text: JSON.stringify(slimMessage(m, true)) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}

export function registerSearchEmails(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'searchEmails',
    `Full-text search across ALL mail folders (subject, body, sender, recipients).

USE searchEmails INSTEAD OF listEmails WHEN:
• Looking for a topic, order number, name, or phrase — not a specific folder
• You don't know which folder holds the message

Search syntax follows Outlook KQL basics: plain words, "quoted phrases",
from:someone@x.com, subject:invoice, hasattachment:true.`,
    {
      query: z.string().describe('Search text, e.g. \'invoice 4711\' or \'from:jane@acme.com subject:renewal\''),
      top: z.number().min(1).max(50).optional().describe('Max results. Default 10.')
    },
    async ({ query, top }) => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        const params = new URLSearchParams({
          $search: `"${query.replace(/"/g, '\\"')}"`,
          $top: String(top ?? 10),
          $select: 'id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,webLink,conversationId'
        })
        const data = await graphRequest<any>(ctx, 'GET', `/me/messages?${params}`)
        const messages = (data.value ?? []).map((m: any) => slimMessage(m))
        return { content: [{ type: 'text', text: JSON.stringify({ count: messages.length, messages }) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}

export function registerListFolders(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'listFolders',
    `Lists the mailbox folder tree with unread/total counts.

Use when the user references a custom folder ('check my Invoices folder') —
pass the returned folder id to listEmails/moveEmail.`,
    {},
    async () => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        const data = await graphRequest<any>(ctx, 'GET', '/me/mailFolders?$top=100&$select=id,displayName,unreadItemCount,totalItemCount,childFolderCount')
        const folders = (data.value ?? []).map((f: any) => ({
          id: f.id,
          name: f.displayName,
          unread: f.unreadItemCount,
          total: f.totalItemCount,
          hasChildren: (f.childFolderCount ?? 0) > 0
        }))
        return { content: [{ type: 'text', text: JSON.stringify({ folders }) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}

export function registerGetAttachments(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'getAttachments',
    `Lists the attachments on a message (name, type, size). Does NOT download
binary content — report the names/sizes to the user instead.`,
    {
      messageId: z.string().describe('Full Graph message id.')
    },
    async ({ messageId }) => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        const data = await graphRequest<any>(ctx, 'GET', `/me/messages/${encodeURIComponent(messageId)}/attachments?$select=id,name,contentType,size,isInline`)
        const attachments = (data.value ?? []).map((a: any) => ({
          id: a.id, name: a.name, contentType: a.contentType, size: a.size, isInline: a.isInline
        }))
        return { content: [{ type: 'text', text: JSON.stringify({ count: attachments.length, attachments }) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}
