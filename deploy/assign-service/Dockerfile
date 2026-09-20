FROM mcr.microsoft.com/playwright:v1.54.0-jammy

WORKDIR /app

COPY code/UX-Watcher/package.json code/UX-Watcher/package-lock.json ./
RUN npm ci --omit=dev

COPY code/UX-Watcher/assign-sachbearbeiter.js \
     code/UX-Watcher/assign-server.js \
     code/UX-Watcher/google.js \
     ./

ENV NODE_ENV=production
ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright
ENV PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1
ENV ASSIGN_PORT=3080

EXPOSE 3080

CMD ["node", "assign-server.js"]
