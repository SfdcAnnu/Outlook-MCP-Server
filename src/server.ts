import 'dotenv/config'
import express from 'express'
import cors from 'cors'
import crypto from 'crypto'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js'
import wellknownRouter from './routes/wellknown'
import authRouter from './routes/auth'
import { registerAllTools, TOOL_SUMMARIES } from './tools/index'
import { unlinkSession, storeVaultTokens } from './lib/session'

const app = express()
app.use(cors({ origin: '*' }))
app.use(express.json())
app.use(express.urlencoded({ extended: true }))

// logging
app.use((req: any, res: any, next: any) => {
  console.log('[REQ]', req.method, req.path, 'auth:', req.headers['authorization'] ? 'YES bearer=' + (req.headers['authorization'] as string).slice(7, 20) + '...' : 'NONE')
  next()
})

app.get('/', (req: any, res: any) => {
  res.json({ name: 'outlook-mcp', version: '1.0.0', status: 'running' })
})

// Public, unauthenticated tool catalog — static metadata only, no execution.
// Portals use this to render tool pickers before a mailbox is connected;
// the authenticated equivalent is MCP tools/list on /mcp.
app.get('/tools', (req: any, res: any) => {
  res.json({ server: 'outlook-mcp', count: TOOL_SUMMARIES.length, tools: TOOL_SUMMARIES })
})

app.use('/', wellknownRouter)

const MS_AUTH_BASE = () =>
  `https://login.microsoftonline.com/${process.env.MS_TENANT ?? 'common'}/oauth2/v2.0`

const DEFAULT_SCOPE = 'openid profile offline_access User.Read Mail.Read Mail.ReadWrite Mail.Send'

/**
 * OAuth proxy — Claude's connector (and any MCP client doing DCR) hits these
 * endpoints. We FORCE our registered Azure client_id/secret because DCR hands
 * the client a random id Microsoft has never heard of. The client's
 * redirect_uri passes through untouched — it must be whitelisted on the
 * Azure app registration.
 */
app.get('/authorize', (req: any, res: any) => {
  const params = new URLSearchParams()
  Object.entries(req.query).forEach(([k, v]) => { if (v && k !== 'resource') params.set(k, v as string) })
  if (process.env.MS_CLIENT_ID) params.set('client_id', process.env.MS_CLIENT_ID)
  if (!params.get('scope')) params.set('scope', DEFAULT_SCOPE)
  const msUrl = `${MS_AUTH_BASE()}/authorize?${params.toString()}`
  console.log('[Auth] /authorize ->', msUrl)
  res.redirect(msUrl)
})

app.post('/token', async (req: any, res: any) => {
  const axios = (await import('axios')).default
  try {
    const params = new URLSearchParams(req.body)
    // Swap in the real Azure credentials — DCR clients send made-up ones.
    if (process.env.MS_CLIENT_ID)     params.set('client_id', process.env.MS_CLIENT_ID)
    if (process.env.MS_CLIENT_SECRET) params.set('client_secret', process.env.MS_CLIENT_SECRET)
    if (!params.get('scope')) params.set('scope', DEFAULT_SCOPE)
    console.log('[Auth] /token -> exchanging with Microsoft, grant:', params.get('grant_type'))
    const response = await axios.post(`${MS_AUTH_BASE()}/token`, params.toString(), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    })
    const data = response.data
    // Vault the tokens so the SERVER can refresh on the client's behalf.
    // This is what makes the server generic — any AI portal works without
    // implementing OAuth refresh itself.
    if (data.access_token && data.refresh_token) {
      storeVaultTokens({
        accessToken: data.access_token,
        refreshToken: data.refresh_token,
        issuedAt: Date.now()
      })
      console.log('[Auth] /token -> tokens vaulted (refresh enabled)')
    }
    res.json(data)
  } catch (err: any) {
    console.error('[Auth] /token error:', err.response?.data ?? err.message)
    res.status(400).json(err.response?.data ?? { error: 'token_exchange_failed' })
  }
})

app.use('/', authRouter)

// Session store for streamable HTTP
const sessions = new Map<string, StreamableHTTPServerTransport>()

app.post('/mcp', (req: any, res: any, next: any) => {
  const auth = req.headers['authorization']
  if (!auth || !auth.startsWith('Bearer ')) {
    res.status(401)
      .set('WWW-Authenticate', `Bearer resource_metadata="${process.env.BASE_URL}/.well-known/oauth-protected-resource"`)
      .json({ error: 'unauthorized' })
    return
  }
  next()
}, async (req: any, res: any) => {
  try {
    const sessionId = req.headers['mcp-session-id'] as string | undefined

    let transport: StreamableHTTPServerTransport

    if (sessionId && sessions.has(sessionId)) {
      // Reuse existing session
      transport = sessions.get(sessionId)!
      await transport.handleRequest(req, res, req.body)
      return
    }

    // New session
    const mcpServer = new McpServer({ name: 'outlook-mcp', version: '1.0.0' })
    registerAllTools(mcpServer, req, crypto.randomUUID())

    transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => crypto.randomUUID(),
      onsessioninitialized: (sid: string) => {
        console.log('[MCP] Session initialized:', sid)
        sessions.set(sid, transport)
      }
    })

    transport.onclose = () => {
      const sid = transport.sessionId
      if (sid) {
        sessions.delete(sid)
        unlinkSession(sid)
        console.log('[MCP] Session closed:', sid)
      }
    }

    await mcpServer.connect(transport)
    await transport.handleRequest(req, res, req.body)
  } catch (err) {
    console.error('[MCP] Error:', err)
    res.status(500).json({ error: 'internal_error' })
  }
})

app.get('/mcp', async (req: any, res: any) => {
  const sessionId = req.headers['mcp-session-id'] as string | undefined
  if (!sessionId || !sessions.has(sessionId)) {
    res.status(400).json({ error: 'invalid_session' })
    return
  }
  const transport = sessions.get(sessionId)!
  await transport.handleRequest(req, res)
})

app.delete('/mcp', async (req: any, res: any) => {
  const sessionId = req.headers['mcp-session-id'] as string | undefined
  if (sessionId) {
    sessions.delete(sessionId)
    unlinkSession(sessionId)
  }
  res.status(200).json({ success: true })
})

const PORT = parseInt(process.env.PORT ?? '3100', 10)
app.listen(PORT, () => console.log(`Outlook MCP server running on port ${PORT}`))
export default app
