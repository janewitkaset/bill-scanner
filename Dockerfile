FROM node:20-alpine

# ติดตั้ง Python3 + openpyxl สำหรับเขียนไฟล์ Excel
RUN apk add --no-cache python3 py3-pip && \
    pip3 install openpyxl --break-system-packages

WORKDIR /app
COPY package*.json ./
RUN npm install --production
COPY . .

# สร้างโฟลเดอร์ data สำหรับเก็บ Excel
RUN mkdir -p data

EXPOSE 3000
CMD ["node", "index.js"]
