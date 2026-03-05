FROM node:20-alpine

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY web ./web
COPY src ./src
COPY server.js ./server.js

ENV NODE_ENV=production
ENV USE_DEV_CERTS=false
ENV PORT=3100

EXPOSE 3100

CMD ["npm", "start"]
