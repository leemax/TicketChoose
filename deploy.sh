#!/bin/bash

# 零距离文旅船票分配系统 - 服务器部署脚本

echo "🚢 开始部署零距离文旅船票分配系统..."

# 1. 更新系统包
echo "📦 更新系统包..."
sudo apt update

# 2. 安装Node.js（如果未安装）
if ! command -v node &> /dev/null; then
    echo "📦 安装Node.js..."
    curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -
    sudo apt-get install -y nodejs
else
    echo "✅ Node.js已安装: $(node -v)"
fi

# 3. 安装PM2（如果未安装）
if ! command -v pm2 &> /dev/null; then
    echo "📦 安装PM2..."
    sudo npm install -g pm2
else
    echo "✅ PM2已安装"
fi

# 4. 克隆或更新代码
if [ -d "TicketChoose" ]; then
    echo "🔄 更新代码..."
    cd TicketChoose
    git pull origin main
else
    echo "📥 克隆代码..."
    git clone git@github.com:leemax/TicketChoose.git
    cd TicketChoose
fi

# 5. 安装依赖
echo "📦 安装依赖..."
npm install

# 6. 创建必要的目录
echo "📁 创建必要的目录..."
mkdir -p uploads temp output

# 7. 停止旧进程（如果存在）
echo "🛑 停止旧进程..."
pm2 delete ticket-system 2>/dev/null || true

# 8. 启动服务
echo "🚀 启动服务..."
pm2 start server.js --name ticket-system

# 9. 设置开机自启
echo "⚙️  设置开机自启..."
pm2 startup
pm2 save

# 10. 显示状态
echo ""
echo "✅ 部署完成！"
echo ""
pm2 status
echo ""
echo "📊 查看日志: pm2 logs ticket-system"
echo "🔄 重启服务: pm2 restart ticket-system"
echo "🛑 停止服务: pm2 stop ticket-system"
echo ""
echo "🌐 服务运行在: http://$(hostname -I | awk '{print $1}'):3000"
