# Node.js + Ruby for MathType conversion
FROM node:20-slim

# Install Ruby and build dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    ruby \
    ruby-dev \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Install Ruby gem mathtype_to_mathml
RUN gem install pry --no-doc
RUN gem install mathtype_to_mathml --no-doc

WORKDIR /app

# Copy package files and install
COPY package.json ./
RUN npm install --production

# Copy application files (including debug script for MTEF XML)
COPY server.js mt2mml.rb mt2mml_v2.rb ./

# Environment
ENV NODE_ENV=production
ENV PORT=8000

EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=10s --retries=3 \
    CMD curl -f http://localhost:${PORT}/health || exit 1

CMD ["node", "server.js"]
