FROM node:20-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer fonts-noto-core fonts-noto-ui-core \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY package*.json ./
RUN npm install --omit=dev
COPY convert-server.mjs .
ENV PORT=3001
EXPOSE 3001
CMD ["node", "convert-server.mjs"]
