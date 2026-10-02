import fs from 'fs';

// The msgraph-metadata commit both specs come from. The generated client is not in
// git, so a branch here would let the same commit of this repo build different tools
// on different days. Move it forward in a commit of its own, for example when a new
// endpoint needs a newer spec. A copy already in openapi/ is reused: after changing
// this, run `npm run generate -- --force`.
const GRAPH_SPEC_COMMIT = '7b2914c8ad1340129f52aa785f13c074cb46fd7c'; // master, 2026-09-29

const DEFAULT_OPENAPI_URL = `https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/${GRAPH_SPEC_COMMIT}/openapi/v1.0/openapi.yaml`;

// Microsoft publishes a parallel /beta OpenAPI spec at the same path root. Endpoints
// flagged "apiVersion": "beta" in endpoints.json are generated from this spec instead.
export const BETA_OPENAPI_URL = `https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/${GRAPH_SPEC_COMMIT}/openapi/beta/openapi.yaml`;

export async function downloadGraphOpenAPI(
  targetDir,
  targetFile,
  openapiUrl = DEFAULT_OPENAPI_URL,
  forceDownload = false
) {
  if (!fs.existsSync(targetDir)) {
    console.log(`Creating directory: ${targetDir}`);
    fs.mkdirSync(targetDir, { recursive: true });
  }

  if (fs.existsSync(targetFile) && !forceDownload) {
    console.log(`OpenAPI specification already exists at ${targetFile}`);
    console.log('Use --force to download again');
    return false;
  }

  console.log(`Downloading OpenAPI specification from ${openapiUrl}`);

  try {
    const response = await fetch(openapiUrl);

    if (!response.ok) {
      throw new Error(`Failed to download: ${response.status} ${response.statusText}`);
    }

    const content = await response.text();
    fs.writeFileSync(targetFile, content);
    console.log(`OpenAPI specification downloaded to ${targetFile}`);
    return true;
  } catch (error) {
    console.error('Error downloading OpenAPI specification:', error.message);
    throw error;
  }
}
