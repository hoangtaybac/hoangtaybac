# Node.js + Ruby for MathType conversion (Ultra Optimized)
FROM node:20-slim

# Install Ruby and build dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    ruby \
    ruby-dev \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Install Ruby gems
RUN gem install pry --no-doc
RUN gem install mathtype_to_mathml --no-doc

WORKDIR /app

# Copy package files and install
COPY package.json ./
RUN npm install --production

# Copy application files (including batch script)
COPY server.js mt2mml.rb mt2mml_v2.rb mt2mml_batch.rb ./

# Environment with optimized defaults
ENV NODE_ENV=production
ENV PORT=8000
ENV CONCURRENCY_LIMIT=4
ENV BATCH_SIZE=10
ENV CACHE_MAX_SIZE=500
ENV CACHE_TTL_MS=1800000
ENV RUBY_TIMEOUT_MS=30000

EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=10s --retries=3 \
    CMD curl -f http://localhost:${PORT}/health || exit 1

CMD ["node", "server.js"]
