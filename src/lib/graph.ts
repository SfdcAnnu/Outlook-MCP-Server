import 'dotenv/config'
import axios, { AxiosRequestConfig } from 'axios'
import { Request } from 'express'
import {
  getSessionTokens,
  updateSessionTokens,
  getVaultTokens,
  updateVaultAccessToken
} from './session'

const GRAPH = 'https://graph.microsoft.com/v1.0'

/**
 * A resolved Graph context: the current access token plus a refresh()
 * hook the caller retries through when Graph returns 401.
 */
export interface GraphCtx {
  token: string
  /** Attempts a refresh; returns the new access token or null. */
  refresh: () => Promise<string | null>
}

async function refreshWithMicrosoft(refreshToken: string): Promise<{ access_token: string; refresh_token?: string } | null> {
  try {
    const tenant = process.env.MS_TENANT ?? 'common'
    const params = new URLSearchParams({
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: process.env.MS_CLIENT_ID!,
      client_secret: process.env.MS_CLIENT_SECRET!,
      scope: 'openid profile offline_access User.Read Mail.Read Mail.ReadWrite Mail.Send'
    })
    const res = await axios.post(
      `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
      params.toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    )
    return res.data
  } catch (err: any) {
    console.error('[Graph] refresh failed:', err.response?.data ?? err.message)
    return null
  }
}

export async function resolveGraphCtx(req: Request, sessionId: string): Promise<GraphCtx | null> {
  // PATH A: Bearer token in Authorization header
  const authHeader = req.headers['authorization'] as string | undefined
  if (authHeader?.startsWith('Bearer ')) {
    const accessToken = authHeader.slice(7)

    const vault = getVaultTokens(accessToken)
    if (vault) {
      console.log('[Graph] PATH A — vault hit, auto-refresh enabled')
      return {
        token: vault.accessToken,
        refresh: async () => {
          const tok = await refreshWithMicrosoft(vault.refreshToken)
          if (!tok) return null
          updateVaultAccessToken(accessToken, tok.access_token, tok.refresh_token)
          console.log('[Graph] PATH A — access token auto-refreshed')
          return tok.access_token
        }
      }
    }

    // No vault entry (server restarted, or a token we never proxied).
    // Use the token as-is; no refresh possible.
    console.log('[Graph] PATH A — stateless bearer (no vault entry)')
    return { token: accessToken, refresh: async () => null }
  }

  // PATH B: session store
  const session = getSessionTokens(sessionId)
  if (!session) {
    console.log('[Graph] PATH B — no session tokens for', sessionId)
    return null
  }
  console.log('[Graph] PATH B — using session token (auto-refresh enabled)')
  return {
    token: session.accessToken,
    refresh: async () => {
      const tok = await refreshWithMicrosoft(session.refreshToken)
      if (!tok) return null
      updateSessionTokens(sessionId, { accessToken: tok.access_token, refreshToken: tok.refresh_token })
      console.log('[Graph] PATH B — access token auto-refreshed')
      return tok.access_token
    }
  }
}

/**
 * Call Microsoft Graph with automatic single-retry on 401 (expired token).
 * `path` is relative to https://graph.microsoft.com/v1.0
 */
export async function graphRequest<T = unknown>(
  ctx: GraphCtx,
  method: 'GET' | 'POST' | 'PATCH' | 'DELETE',
  path: string,
  body?: unknown,
  extraConfig?: AxiosRequestConfig
): Promise<T> {
  const doCall = (token: string) => axios.request<T>({
    method,
    url: `${GRAPH}${path}`,
    data: body,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    ...extraConfig
  })

  try {
    const res = await doCall(ctx.token)
    return res.data
  } catch (err: any) {
    if (err.response?.status === 401) {
      const newToken = await ctx.refresh()
      if (newToken) {
        const res = await doCall(newToken)
        return res.data
      }
    }
    throw err
  }
}

// ── MCP tool result helpers ─────────────────────────────────────────

export function notAuthenticatedError(baseUrl: string) {
  return {
    content: [{
      type: 'text' as const,
      text: JSON.stringify({
        error: 'NOT_AUTHENTICATED',
        message: 'No Microsoft 365 connection found.',
        resolution: {
          claude_connector: 'Make sure the Outlook MCP connector is connected in Claude Settings.',
          standalone: `Visit ${baseUrl}/auth to authenticate.`
        }
      })
    }],
    isError: true
  }
}

export function graphApiError(err: unknown) {
  const e = err as { response?: { status?: number; data?: { error?: { code?: string; message?: string } } }; message?: string }
  const code    = e.response?.data?.error?.code ?? `HTTP_${e.response?.status ?? 'ERROR'}`
  const message = e.response?.data?.error?.message ?? e.message ?? String(err)
  console.error('[Graph] API Error:', code, message)

  let hint: string | undefined
  if (code === 'ErrorItemNotFound')          hint = 'The message id does not exist or was deleted. Call listEmails first to get fresh ids.'
  else if (code === 'ErrorAccessDenied')     hint = 'Missing Graph permission. The Azure app needs the Mail.Read / Mail.Send delegated scopes granted.'
  else if (code === 'ErrorInvalidIdMalformed') hint = 'Message ids must be the full Graph id from listEmails, not a truncated value.'

  return {
    content: [{ type: 'text' as const, text: JSON.stringify({ errorCode: code, message, ...(hint && { hint }) }) }],
    isError: true
  }
}
