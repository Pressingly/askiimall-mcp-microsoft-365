/**
 * The OData query options Graph accepts on the tools this fork serves, where
 * Microsoft's OpenAPI spec offers more.
 *
 * The spec is generated from Graph's OData metadata, and its converter adds $top,
 * $skip, $search, $filter, $count, $orderby, $select and $expand to every list
 * endpoint that carries no capability annotation saying otherwise; most carry none
 * (microsoftgraph/msgraph-metadata#298). Each workload behind Graph accepts its own
 * subset. A model reads every offered option as usable: it sent $orderby to channel
 * messages and $top to team channels, and Graph answered 400 "Query option 'X' is
 * not allowed". Elsewhere Graph accepts an option and ignores it ($filter, $search
 * and $orderby on file versions), so the model gets a list it did not ask for.
 *
 * A row lists the options from the API's reference page on learn.microsoft.com,
 * kept only where a live request with that option succeeded (2026-10-14). Tools
 * without a row keep the spec's options: mail accepts all eight. Tool-specific query
 * parameters and the synthetic `skiptoken` cursor are never touched.
 *
 * Read by queryParameterSchema (query-parameter-schema.ts), so tool registration,
 * get-tool-schema and execute-tool all apply it.
 */
import { z } from 'zod';

const ODATA_OPTIONS = new Set([
  'top',
  'skip',
  'search',
  'filter',
  'count',
  'orderby',
  'select',
  'expand',
]);

interface QueryOptionRow {
  options: readonly string[];
  /** Graph's own limit for $top; a larger value is a 400. */
  maxTop?: number;
}

const QUERY_OPTIONS: Record<string, QueryOptionRow> = {
  // https://learn.microsoft.com/graph/api/drive-list
  'list-drives': { options: ['top', 'orderby', 'select', 'expand'] },
  'list-sharepoint-site-drives': { options: ['top', 'orderby', 'select', 'expand'] },
  // https://learn.microsoft.com/graph/api/driveitem-list-children
  'list-folder-files': { options: ['top', 'orderby', 'select', 'expand'] },
  // https://learn.microsoft.com/graph/api/driveitem-search
  'search-onedrive-files': { options: ['top', 'orderby', 'select', 'expand'] },
  // https://learn.microsoft.com/graph/api/driveitem-delta
  'get-drive-delta': { options: ['top', 'select', 'expand'] },
  // https://learn.microsoft.com/graph/api/driveitem-list-thumbnails
  'list-drive-item-thumbnails': { options: ['select'] },
  // https://learn.microsoft.com/graph/api/driveitem-list-permissions
  'list-drive-item-permissions': { options: ['select'] },
  // https://learn.microsoft.com/graph/api/driveitem-list-versions: the page names no
  // options; Graph applies $top and ignores $filter, $search and $orderby.
  'list-drive-item-versions': { options: ['top', 'select'] },
  // https://learn.microsoft.com/graph/api/site-search: the page names no options;
  // Graph rejects $skip, $count and $expand, and $filter on site properties.
  'search-sharepoint-sites': { options: ['top', 'search', 'orderby', 'select'] },
  // https://learn.microsoft.com/graph/api/chat-list. Upstream caps $top at 50 with an
  // intersection; this row replaces it with the plain schema (capQueryTop).
  'list-chats': { options: ['top', 'filter', 'orderby', 'expand'], maxTop: 50 },
  // https://learn.microsoft.com/graph/api/chat-list-members
  'list-chat-members': { options: [] },
  // https://learn.microsoft.com/graph/api/chat-list-messages
  'list-chat-messages': { options: ['top', 'filter', 'orderby'], maxTop: 50 },
  // https://learn.microsoft.com/graph/api/channel-list
  'list-team-channels': { options: ['filter', 'select'] },
  // https://learn.microsoft.com/graph/api/channel-list-messages
  'list-channel-messages': { options: ['top', 'expand'], maxTop: 50 },
  // https://learn.microsoft.com/graph/api/chatmessage-list-replies
  'list-channel-message-replies': { options: ['top'], maxTop: 50 },
  // https://learn.microsoft.com/graph/api/channel-get
  'get-team-channel': { options: ['select'] },
  // https://learn.microsoft.com/graph/api/chatmessage-get: Graph rejects $expand
  'get-channel-message': { options: ['select'] },
};

/** True when the tool has a row and the OData option is not in it. */
export function isQueryOptionWithheld(toolName: string, bareName: string): boolean {
  const row = QUERY_OPTIONS[toolName];
  return row !== undefined && ODATA_OPTIONS.has(bareName) && !row.options.includes(bareName);
}

/**
 * The $top schema with Graph's limit for the tool, when its row sets one. The limit
 * goes on the provider's number schema itself, keeping its other bounds, so the
 * advertised JSON schema stays one `integer` with a `maximum`: mcpo builds its
 * request model from that schema and reads no `allOf`, so an intersection
 * (`.and()`) reached the model as an untyped field.
 */
export function capQueryTop(
  toolName: string,
  source: z.ZodTypeAny,
  schema: z.ZodTypeAny
): z.ZodTypeAny {
  const maxTop = QUERY_OPTIONS[toolName]?.maxTop;
  if (maxTop === undefined) return schema;
  return source instanceof z.ZodNumber ? source.max(maxTop) : source.and(z.number().max(maxTop));
}
