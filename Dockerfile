FROM python:3.13-alpine AS builder

WORKDIR /build
COPY generate_lighting_dashboard.py test.xlsx ./
RUN python generate_lighting_dashboard.py \
    --excel test.xlsx \
    --output lighting_dashboard.html

FROM nginx:alpine
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=builder /build/lighting_dashboard.html /usr/share/nginx/html/lighting_dashboard.html

