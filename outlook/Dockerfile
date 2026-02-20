# Stage 1: Build the React Application
FROM node:14.17.0-alpine AS build

# Set working directory
WORKDIR /app

# Copy package.json and install dependencies
COPY package*.json ./
RUN npm install

# Copy the rest of the application source code
COPY . .

# Expose a build argument for the domain so the manifest.xml can be safely compiled
ARG DOMAIN=localhost:3000
ENV DOMAIN=$DOMAIN

# Run the Webpack production build using the domain variable
RUN npm run build

# Stage 2: Serve the build output with Nginx
FROM nginx:alpine

# Remove default Nginx config and copy our custom one
RUN rm /etc/nginx/conf.d/default.conf
COPY nginx/default.conf /etc/nginx/conf.d/

# Copy the built React app from the build stage
COPY --from=build /app/dist /usr/share/nginx/html

# Expose port 80 (to be routed via proxy)
EXPOSE 80

# Start Nginx
CMD ["nginx", "-g", "daemon off;"]
