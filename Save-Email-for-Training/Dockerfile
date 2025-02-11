FROM node:18-alpine

RUN apk add --no-cache sudo

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm ci

# Copy source code
COPY . .

# Expose port 3000 (adjust if your app uses a different port)
EXPOSE 3000
