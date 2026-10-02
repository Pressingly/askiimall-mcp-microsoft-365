/**
 * Where the Graph client gets a token when a call carries none of its own.
 *
 * GraphClient takes the token from the call, then from the request context,
 * then from AuthManager. In plain --http mode every /mcp request runs inside
 * its caller's context, so the last step only answers a call that has no
 * caller, and it would answer with whatever the server holds itself: a mounted
 * MSAL cache, or a token set through setOAuthToken. With many users on one
 * server that is one account's data for everyone. So in that mode the client
 * gets an AuthManager whose getToken() refuses.
 *
 * stdio mode and --trust-proxy-auth run as one account by design and keep
 * AuthManager as it is.
 */
import type AuthManager from '../auth.js';
import type { CommandOptions } from '../cli.js';

export function graphTokenSource(
  authManager: AuthManager,
  options: Pick<CommandOptions, 'http' | 'trustProxyAuth'>
): AuthManager {
  if (!options.http || options.trustProxyAuth) {
    return authManager;
  }
  return new Proxy(authManager, {
    get(target, property, receiver) {
      if (property === 'getToken') {
        return async () => {
          throw new Error(
            "No access token for this request: in HTTP mode a Graph call uses only its caller's bearer token."
          );
        };
      }
      return Reflect.get(target, property, receiver);
    },
  });
}
