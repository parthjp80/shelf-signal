# --- Build stage: compile the React frontend ---
FROM node:20-alpine AS build
WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .
RUN npm run build

# --- Runtime stage: small image with just the server + built assets ---
FROM node:20-alpine
WORKDIR /app
ENV NODE_ENV=production

COPY package*.json ./
RUN npm install --omit=dev

COPY --from=build /app/dist ./dist
COPY server ./server

# Data directory for persisted reports (mount a volume here)
RUN mkdir -p /app/data
VOLUME ["/app/data"]

EXPOSE 3000
ENV PORT=3000

CMD ["node", "server/server.js"]
