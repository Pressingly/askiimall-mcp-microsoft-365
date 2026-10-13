// Microsoft's OpenAPI metadata gives action parameters a `default` that is not
// always what Graph does when the parameter is left out. The tool schemas pass
// that default on to MCP clients, and a client that fills defaults in from the
// advertised schema then sends it on every call.
//
// /me/sendMail keeps a copy in Sent Items unless SaveToSentItems is false, but
// the metadata says `default: false`, so mail sent through send-mail never
// reached the sender's Sent Items.
export function fixGraphActionDefaults(openApiSpec) {
  const saveToSentItems =
    openApiSpec.components?.requestBodies?.sendMailRequestBody?.content?.['application/json']
      ?.schema?.properties?.SaveToSentItems;
  if (saveToSentItems) {
    saveToSentItems.default = true;
  }
}
