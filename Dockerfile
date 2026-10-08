FROM node:24-alpine AS builder

WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run generate
RUN npm run build

FROM node:24-alpine AS release

WORKDIR /app

COPY --from=builder /app/dist /app/dist
COPY --from=builder /app/package*.json ./

ENV NODE_ENV=production
RUN npm ci --ignore-scripts --omit=dev

USER node

ENTRYPOINT ["node", "dist/index.js"]

# Askii runs one copy of this image for all its Microsoft 365 products; the
# tools grow with each product (test/gateway-tool-list.test.ts holds the list).
# --org-mode registers the SharePoint site tools, which carry work-account
# scopes only; it adds nothing to onedrive now that the sensitivity-labels tool
# (Files.Read.All, admin consent) is out of that preset. Never
# --allowed-scopes: /authorize must pass on the scope each AskiiMall service
# asks for.
CMD ["--http", "0.0.0.0:3000", "-v", "--preset", "onedrive,sharepoint", "--org-mode"]
