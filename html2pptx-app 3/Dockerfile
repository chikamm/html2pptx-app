# html2pptx-app
#
# Installs system Chromium via apt instead of letting Puppeteer download its
# own copy - this is both faster to build and avoids relying on Puppeteer's
# download CDN (storage.googleapis.com), which is blocked in some corporate/
# offline network environments (this was in fact required to build & test
# this very project - see README "Known limitations").

FROM node:20-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium \
    fonts-noto-cjk \
    fonts-ipafont-gothic \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

ENV PUPPETEER_SKIP_DOWNLOAD=true \
    PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium \
    NODE_ENV=production \
    PORT=8787

WORKDIR /app
COPY package.json ./
RUN npm install --omit=dev --no-audit --no-fund

COPY server ./server
COPY public ./public

EXPOSE 8787
CMD ["node", "server/index.js"]
