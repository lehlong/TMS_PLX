# Dockerfile

# Sử dụng hình ảnh Node.js
FROM node:18

# Tạo thư mục làm việc
WORKDIR /app

# Sao chép package.json và package-lock.json
COPY package*.json ./

# Cài đặt các dependencies
RUN npm install

# Sao chép toàn bộ mã nguồn
COPY . .

# Cài đặt Angular CLI toàn cục
RUN npm install -g @angular/cli

# Mở cổng 4200
EXPOSE 4200

# Lệnh để khởi động ứng dụng với live-reloading
CMD ["ng", "serve", "--host", "0.0.0.0"]
