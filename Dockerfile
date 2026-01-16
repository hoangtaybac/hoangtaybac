# Node.js + Ruby + Inkscape for MathType conversion
FROM node:20-bookworm-slim

WORKDIR /app
ENV NODE_ENV=production

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    ruby \
    ruby-dev \
    build-essential \
    curl \
    inkscape \
    ca-certificates \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# Install Ruby gem for MathType → MathML
RUN gem install mathtype_to_mathml -N

# Copy package files first (for cache)
COPY package.json package-lock.json* ./

# Install Node dependencies (prefer ci if lock exists)
RUN if [ -f package-lock.json ]; then npm ci --omit=dev; else npm install --production; fi

# Copy application source
COPY server.js mt2mml.rb ./

# Railway will inject PORT; server already uses process.env.PORT || 3000
EXPOSE 3000

# Healthcheck -> use /ping (your server has it)
HEALTHCHECK --interval=30s --timeout=10s --start-period=10s --retries=3 \
  CMD curl -fsS http://localhost:${PORT:-3000}/ping || exit 1

# Start server
CMD ["node", "server.js"]
