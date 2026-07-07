/**
 * Session store for PATH B (standalone clients — n8n, Cursor, custom apps)
 * NOT used when Claude connector injects the token via Authorization header (PATH A)
 *
 * Two-layer store:
 *   Layer 1: sessionId (ephemeral, changes per SSE connection) → msUserId (stable)
 *   Layer 2: msUserId (stable Microsoft account id)            → Graph tokens
 *
 * For production: replace Maps with Upstash Redis.
 */

export interface GraphSession {
  accessToken: string
  refreshToken: string
  issuedAt: number
}

// Layer 1: ephemeral sessionId → stable msUserId
const sessionToUser = new Map<string, string>()

// Layer 2: stable msUserId → Graph tokens
const userTokenStore = new Map<string, GraphSession>()

// ── Token vault for PATH A (any AI portal / Claude connector) ──────────────
// Keyed by the access_token the client sends as Bearer. Captured at /token so
// the SERVER owns refresh — the client never has to implement OAuth refresh.
export interface VaultEntry {
  accessToken: string      // current live access token (updated on refresh)
  refreshToken: string
  issuedAt: number
}

const accessTokenVault = new Map<string, VaultEntry>()

// Called in /token after a successful authorization_code exchange
export function storeVaultTokens(entry: VaultEntry): void {
  accessTokenVault.set(entry.accessToken, entry)
}

// Called by resolveGraphToken() on every PATH A tool call
export function getVaultTokens(accessToken: string): VaultEntry | undefined {
  return accessTokenVault.get(accessToken)
}

// Called after a server-side refresh mints a new access token. The client
// keeps sending the ORIGINAL bearer, so we keep that key pointing at the
// same (mutated) entry, and also index the new token in case the client
// adopts it.
export function updateVaultAccessToken(originalAccessToken: string, newAccessToken: string, newRefreshToken?: string): void {
  const entry = accessTokenVault.get(originalAccessToken)
  if (!entry) return
  entry.accessToken = newAccessToken
  if (newRefreshToken) entry.refreshToken = newRefreshToken
  entry.issuedAt = Date.now()
  accessTokenVault.set(newAccessToken, entry)
}

// Called in /auth/callback after OAuth exchange
export function linkSession(sessionId: string, msUserId: string, session: GraphSession): void {
  sessionToUser.set(sessionId, msUserId)
  userTokenStore.set(msUserId, session)
}

// Called in /mcp when a returning user connects with a cookie
export function relinkSession(sessionId: string, msUserId: string): void {
  sessionToUser.set(sessionId, msUserId)
}

// Called by resolveGraphToken() in every tool handler
export function getSessionTokens(sessionId: string): GraphSession | undefined {
  const userId = sessionToUser.get(sessionId)
  if (!userId) return undefined
  return userTokenStore.get(userId)
}

// Called after token refresh to update stored access token
export function updateSessionTokens(sessionId: string, updated: { accessToken: string; refreshToken?: string }): void {
  const userId = sessionToUser.get(sessionId)
  if (!userId) return
  const existing = userTokenStore.get(userId)
  if (!existing) return
  userTokenStore.set(userId, {
    ...existing,
    accessToken: updated.accessToken,
    refreshToken: updated.refreshToken ?? existing.refreshToken,
    issuedAt: Date.now()
  })
}

export function getMsUserId(sessionId: string): string | undefined {
  return sessionToUser.get(sessionId)
}

export function hasUserTokens(msUserId: string): boolean {
  return userTokenStore.has(msUserId)
}

// When SSE disconnects — remove sessionId link but KEEP user tokens
// (other sessions from same user still valid)
export function unlinkSession(sessionId: string): void {
  sessionToUser.delete(sessionId)
}
