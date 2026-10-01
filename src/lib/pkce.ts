/**
 * Which client PKCE challenges /authorize can send to Microsoft as they are.
 *
 * An S256 challenge is BASE64URL(SHA256(verifier)) without padding (RFC 7636
 * §4.2): always 43 characters from A-Z, a-z, 0-9, "-" and "_". Microsoft
 * accepts that shape, and checks the client's own verifier at /token, so the
 * server keeps nothing between the two calls and any copy of it can answer
 * /token, also after a restart.
 *
 * Anything else goes through the two-leg mapping in server.ts (upstream issue
 * #266: Microsoft rejected the challenge claude.ai web sent). That includes an
 * S256 challenge sent without code_challenge_method: Microsoft would read it
 * as plain, while the mapping checks it as S256 at /token.
 */

const S256_CHALLENGE = /^[A-Za-z0-9_-]{43}$/;

export function isStandardS256Challenge(challenge: string, method: string | null): boolean {
  return method === 'S256' && S256_CHALLENGE.test(challenge);
}
