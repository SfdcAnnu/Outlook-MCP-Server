/**
 * Write-side mail tools: send, reply, draft, move, mark read/unread.
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import { Request } from 'express'
import { resolveGraphCtx, graphRequest, notAuthenticatedError, graphApiError } from '../lib/graph'

const recipients = (addresses: string[]) =>
  addresses.map(a => ({ emailAddress: { address: a } }))

export function registerSendEmail(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'sendEmail',
    `Sends a new email from the connected mailbox.

RULES:
• Confirm recipients and content with the user before sending anything they did not explicitly dictate.
• bodyType 'html' allows formatting (<b>, <p>, lists); 'text' sends plain text.
• The message lands in the user's Sent Items.`,
    {
      to: z.array(z.string()).min(1).describe('Recipient email addresses.'),
      subject: z.string().describe('Email subject.'),
      body: z.string().describe('Email body content.'),
      bodyType: z.enum(['text', 'html']).optional().describe("Body format. Default 'text'."),
      cc: z.array(z.string()).optional().describe('CC addresses.'),
      bcc: z.array(z.string()).optional().describe('BCC addresses.')
    },
    async ({ to, subject, body, bodyType, cc, bcc }) => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        await graphRequest(ctx, 'POST', '/me/sendMail', {
          message: {
            subject,
            body: { contentType: bodyType === 'html' ? 'HTML' : 'Text', content: body },
            toRecipients: recipients(to),
            ...(cc?.length  && { ccRecipients: recipients(cc) }),
            ...(bcc?.length && { bccRecipients: recipients(bcc) })
          },
          saveToSentItems: true
        })
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, sentTo: to, subject }) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}

export function registerReplyEmail(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'replyEmail',
    `Replies to an existing email (keeps the thread). Use replyAll=true to
include everyone on the original message.

Confirm the reply text with the user before sending unless they dictated it.`,
    {
      messageId: z.string().describe('Full Graph message id of the email being replied to.'),
      comment: z.string().describe('The reply text (plain text).'),
      replyAll: z.boolean().optional().describe('Reply to all recipients. Default false (sender only).')
    },
    async ({ messageId, comment, replyAll }) => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        const action = replyAll ? 'replyAll' : 'reply'
        await graphRequest(ctx, 'POST', `/me/messages/${encodeURIComponent(messageId)}/${action}`, { comment })
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, action, messageId }) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}

export function registerCreateDraft(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'createDraft',
    `Creates a draft email WITHOUT sending it. The user reviews and sends from
Outlook themselves. Prefer this over sendEmail when the user says 'draft',
'prepare', or hasn't confirmed sending.`,
    {
      to: z.array(z.string()).min(1).describe('Recipient email addresses.'),
      subject: z.string().describe('Email subject.'),
      body: z.string().describe('Email body content.'),
      bodyType: z.enum(['text', 'html']).optional().describe("Body format. Default 'text'."),
      cc: z.array(z.string()).optional().describe('CC addresses.')
    },
    async ({ to, subject, body, bodyType, cc }) => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        const draft = await graphRequest<any>(ctx, 'POST', '/me/messages', {
          subject,
          body: { contentType: bodyType === 'html' ? 'HTML' : 'Text', content: body },
          toRecipients: recipients(to),
          ...(cc?.length && { ccRecipients: recipients(cc) })
        })
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, draftId: draft.id, webLink: draft.webLink }) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}

export function registerMoveEmail(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'moveEmail',
    `Moves a message to another folder. Destination can be a well-known name
(inbox, archive, deleteditems, junkemail) or a folder id from listFolders.`,
    {
      messageId: z.string().describe('Full Graph message id.'),
      destinationFolder: z.string().describe("Target folder: well-known name or folder id.")
    },
    async ({ messageId, destinationFolder }) => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        const moved = await graphRequest<any>(ctx, 'POST', `/me/messages/${encodeURIComponent(messageId)}/move`, {
          destinationId: destinationFolder
        })
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, newId: moved.id }) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}

export function registerMarkEmail(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'markEmail',
    `Marks a message read or unread, and/or flags it for follow-up.`,
    {
      messageId: z.string().describe('Full Graph message id.'),
      isRead: z.boolean().optional().describe('true = mark read, false = mark unread.'),
      flag: z.boolean().optional().describe('true = flag for follow-up, false = clear flag.')
    },
    async ({ messageId, isRead, flag }) => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        const patch: any = {}
        if (typeof isRead === 'boolean') patch.isRead = isRead
        if (typeof flag === 'boolean')   patch.flag = { flagStatus: flag ? 'flagged' : 'notFlagged' }
        if (Object.keys(patch).length === 0) {
          return { content: [{ type: 'text', text: JSON.stringify({ error: 'NOTHING_TO_DO', message: 'Pass isRead and/or flag.' }) }], isError: true }
        }
        await graphRequest(ctx, 'PATCH', `/me/messages/${encodeURIComponent(messageId)}`, patch)
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, applied: patch }) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}

export function registerGetProfile(server: McpServer, req: Request, sessionId: string) {
  server.tool(
    'getProfile',
    `Returns the connected mailbox owner's profile (name, email). Use to confirm
which account is connected before sending on the user's behalf.`,
    {},
    async () => {
      const ctx = await resolveGraphCtx(req, sessionId)
      if (!ctx) return notAuthenticatedError(process.env.BASE_URL!)
      try {
        const me = await graphRequest<any>(ctx, 'GET', '/me?$select=id,displayName,mail,userPrincipalName')
        return { content: [{ type: 'text', text: JSON.stringify({
          id: me.id, name: me.displayName, email: me.mail ?? me.userPrincipalName
        }) }] }
      } catch (err) {
        return graphApiError(err)
      }
    }
  )
}
