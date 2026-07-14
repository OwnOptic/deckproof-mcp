# syntax=docker/dockerfile:1

# ---- build ----
FROM node:22-alpine AS build
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY tsconfig.json ./
COPY src ./src
RUN npm run build

# ---- runtime ----
FROM node:22-alpine
WORKDIR /app
ENV NODE_ENV=production
COPY package.json package-lock.json ./
RUN npm ci --omit=dev && npm cache clean --force
COPY --from=build /app/dist ./dist

# This image's whole purpose is remote deployment, so always speak Streamable
# HTTP. Most PaaS platforms (Render, Fly, Railway, Cloud Run) inject their own
# PORT at runtime, which overrides this default.
ENV MCP_TRANSPORT=http
ENV PORT=3000
ENV HOST=0.0.0.0
EXPOSE 3000

USER node
CMD ["node", "dist/index.js"]
