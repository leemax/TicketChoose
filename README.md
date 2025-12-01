# PDF智能匹配系统

一个高效的SaaS应用，用于根据Excel表格数据筛选和打包PDF文件。

## 功能特点

- 📦 支持ZIP和RAR格式压缩包上传（最大200MB）
- 📊 Excel表格数据解析（支持.xlsx和.xls格式）
- 🔍 智能匹配：根据房号和姓名精确匹配PDF文件
- ⬇️ 一键下载匹配结果的压缩包
- 🎨 现代化、响应式UI设计
- 🌐 完整的中文界面

## 匹配逻辑

**PDF文件名格式**: `[前缀]-[房号]-[姓名]--登船凭证.pdf`
- 示例: `BDM03251202-2-4112-方珊玲--登船凭证.pdf`
- 提取房号: `4112`
- 提取姓名: `方珊玲`

**Excel表格要求**:
- 必须包含"房号"列（Column C）
- 必须包含"中文姓名"列（Column G）

系统会自动匹配房号和姓名都相符的PDF文件。

## 快速开始

### 安装依赖

```bash
npm install
```

### 启动服务器

```bash
npm start
```

服务器将在 `http://localhost:3000` 启动。

## 使用方法

1. **上传压缩包**: 上传包含PDF文件的ZIP或RAR压缩包
2. **上传Excel表格**: 上传包含房号和姓名的Excel文件
3. **下载结果**: 系统自动匹配并打包，点击下载按钮获取结果

## 技术栈

- **后端**: Node.js + Express
- **文件处理**: 
  - `adm-zip` - ZIP文件处理
  - `node-unrar-js` - RAR文件解压
  - `xlsx` - Excel文件解析
  - `multer` - 文件上传处理
- **前端**: HTML5 + CSS3 + JavaScript
- **设计**: 现代化玻璃态设计，渐变色彩，流畅动画

## 项目结构

```
TicketChoose/
├── server.js           # Express服务器和业务逻辑
├── package.json        # 项目依赖配置
├── public/             # 前端静态文件
│   ├── index.html      # 主页面
│   ├── styles.css      # 样式表
│   └── app.js          # 前端交互逻辑
├── uploads/            # 临时上传文件（自动创建）
├── temp/               # 解压临时文件（自动创建）
└── output/             # 输出文件（自动创建）
```

## API接口

### POST /api/upload-archive
上传压缩包并解压

**请求**: FormData with `archive` file
**响应**: 
```json
{
  "success": true,
  "sessionId": "1234567890",
  "message": "压缩文件上传并解压成功"
}
```

### POST /api/upload-excel
上传Excel并处理匹配

**请求**: FormData with `excel` file and `sessionId`
**响应**:
```json
{
  "success": true,
  "matched": 25,
  "total": 30,
  "downloadUrl": "/api/download/1234567890",
  "message": "成功匹配 25 个PDF文件"
}
```

### GET /api/download/:sessionId
下载匹配结果的ZIP文件

## 注意事项

- 文件大小限制：200MB
- 临时文件会在1小时后自动清理
- 确保Excel表格包含正确的列名："房号"和"中文姓名"
- PDF文件名必须符合指定格式才能被正确匹配

## 许可证

ISC
