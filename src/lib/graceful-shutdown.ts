/**
 * Stop the HTTP server cleanly on SIGTERM or SIGINT.
 *
 * A rollout sends SIGTERM and waits a grace period before SIGKILL. In a
 * container the server is process 1, and Linux delivers SIGTERM to process 1
 * only when the program handles it; without this the server kept taking
 * requests until SIGKILL cut every one still running, a Graph call among them.
 *
 * On the first signal the listeners close (MicrosoftGraphServer.stop()): no new
 * connections, idle keep-alive connections dropped, requests in flight allowed
 * to finish. The process exits 0 when they have. A second signal exits 1 at
 * once, so a second Ctrl-C still ends a local run that is waiting on a slow call.
 *
 * `signals` and `exit` are the process's own unless a test passes others.
 */
import type { EventEmitter } from 'node:events';
import logger from '../logger.js';

export interface GracefulShutdownOptions {
  signals?: EventEmitter;
  exit?: (code: number) => void;
}

export function installGracefulShutdown(
  server: { stop(): Promise<void> },
  { signals = process, exit = (code) => process.exit(code) }: GracefulShutdownOptions = {}
): void {
  let stopping = false;

  const onSignal = (signal: string) => {
    if (stopping) {
      logger.warn(`Received ${signal} again: exiting without waiting for requests in flight`);
      exit(1);
      return;
    }
    stopping = true;
    logger.info(`Received ${signal}: closing listeners, finishing requests in flight`);
    server.stop().then(
      () => {
        logger.info('Requests in flight finished: exiting');
        exit(0);
      },
      (error: unknown) => {
        logger.error(`Shutdown failed: ${error instanceof Error ? error.message : String(error)}`);
        exit(1);
      }
    );
  };

  signals.on('SIGTERM', onSignal);
  signals.on('SIGINT', onSignal);
}
