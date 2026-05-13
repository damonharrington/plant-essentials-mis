# ─── Stage 1: Build ──────────────────────────────────────────────────────────
FROM node:20-alpine AS builder

# Set working directory
WORKDIR /app

# Copy dependency manifests first (layer cache optimisation)
COPY package.json package-lock.json* ./

# Install dependencies (ci = clean, reproducible, faster than install)
RUN npm install

# Copy source
COPY . .

# Build production bundle
RUN npm run build

# ─── Stage 2: Serve ──────────────────────────────────────────────────────────
FROM nginx:1.25-alpine AS runner

# Remove default nginx static assets
RUN rm -rf /usr/share/nginx/html/*

# Copy built assets from builder stage
COPY --from=builder /app/dist /usr/share/nginx/html

# Copy custom nginx config
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Expose port
EXPOSE 80

# Nginx runs in foreground
CMD ["nginx", "-g", "daemon off;"]
