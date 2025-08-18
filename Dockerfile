# Dockerfile
# Build stage
FROM node:18-alpine as builder

WORKDIR /app

# Copy package.json and package-lock.json (or yarn.lock)
COPY package.json /

# Install dependencies
RUN npm install --legacy-peer-deps

# Copy the project files into the docker image
COPY . .

# Build the application
RUN npm run build -- --env BASE_PATH=$BASE_PATH

# Production stage
FROM nginx:alpine
WORKDIR /usr/share/nginx/html

# Remove default nginx static assets
RUN rm -rf ./*

# Copy static assets from builder stage
COPY --from=builder /app/dist .

COPY nginx.conf /etc/nginx/conf.d/default.conf

# Containers run nginx with global directives and daemon off
CMD ["nginx", "-g", "daemon off;"]
