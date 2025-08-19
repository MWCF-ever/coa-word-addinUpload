# AIMTA Document Processor - Word Add-in

一个基于Microsoft Word的Office加载项，用于处理AIMTA（Advanced In-Market Testing Analytics）批次分析文档，提供自动化的COA（Certificate of Analysis）文档处理和Word文档生成功能。

## 📋 项目概述

### 功能特性

- **SSO认证集成**: 使用Microsoft Entra ID进行单点登录
- **批次数据处理**: 自动从PDF文件中提取批次分析数据
- **Word文档生成**: 创建符合监管要求的完整AIMTA文档
- **多区域模板支持**: 支持CN、EU、US等不同地区的模板
- **实时数据同步**: 与后端API进行实时数据交换
- **Office.js集成**: 深度集成Microsoft Word功能

### 技术栈

- **前端框架**: React 17 + TypeScript
- **UI组件**: Microsoft Fluent UI (@fluentui/react)
- **认证**: Azure MSAL Browser
- **API通信**: Axios with 拦截器
- **构建工具**: Webpack 5 + Babel
- **Office集成**: Office.js API
- **样式**: CSS + Fluent UI

## 🏗️ 项目结构

```
src/
├── components/              # 通用组件
│   └── AuthGuard.tsx       # 认证守卫组件
├── contexts/               # React Context
│   └── AuthContext.tsx     # 认证状态管理
├── services/               # 服务层
│   ├── authService.ts      # 认证服务 (MSAL集成)
│   ├── httpInterceptor.ts  # HTTP拦截器
│   └── batchTemplete.ts    # 批次模板生成服务
├── taskpane/              # 任务面板主要组件
│   ├── components/
│   │   ├── App.tsx        # 主应用组件
│   │   ├── CompoundSelector.tsx  # 化合物选择器
│   │   ├── TemplateSelector.tsx  # 模板选择器
│   │   └── BatchDataTable.tsx    # 批次数据表格
│   ├── hooks/             # 自定义Hooks
│   │   ├── useAsyncProcessing.ts # 异步处理Hook
│   │   └── useWordDocument.ts    # Word文档操作Hook
│   ├── index.tsx          # 入口文件
│   ├── index.html         # HTML模板
│   └── taskpane.css       # 样式文件
├── types/                 # TypeScript类型定义
│   ├── index.ts          # 主要类型定义
│   └── global.d.ts       # 全局类型声明
├── utils/                # 工具函数
│   └── officeAddinHelpers.ts # Office加载项辅助函数
├── commands/             # Office命令
│   └── commands.ts       # 命令处理函数
└── polyfills.ts         # IE11兼容性填充
```

## 🚀 快速开始

### 环境要求

- Node.js 16+
- npm 或 yarn
- Microsoft Word (桌面版或在线版)
- 有效的BeiGene企业账户

### 安装和开发

```bash
# 克隆项目
git clone <repository-url>
cd coa-word-addin

# 安装依赖
npm install --legacy-peer-deps

# 开发模式 (本地HTTPS服务器)
npm run dev

# 构建生产版本
npm run build

# 验证manifest文件
npm run validate

# 启动Office调试
npm run start

# 停止Office调试
npm run stop
```

### 部署配置

项目使用**同域部署**配置，所有API请求使用相对路径：

```typescript
// API基础URL配置 (httpInterceptor.ts)
const finalBaseURL = baseURL || '/api/aimta';

// 认证重定向URI (authService.ts)
redirectUri: window.location.origin + "/user/login"
```

## 🔧 核心组件详解

### 1. 认证系统 (`AuthContext` + `AuthService`)

```typescript
// 支持的认证流程
1. SSO重定向登录 (优先)
2. MSAL静默认证 (fallback)
3. MSAL弹窗认证 (最后选择)
```

**关键特性**:
- 自动Token刷新
- 会话状态持久化
- 多种认证流程支持
- 错误恢复机制

### 2. 批次模板服务 (`BatchTemplateService`)

负责生成完整的AIMTA文档，包含：

- 页眉页脚设置
- 目录生成
- 多个数据表格
- 符合监管要求的格式

```typescript
// 使用示例
await createAIMTADocument(batchDataList, compoundCode);
```

### 3. HTTP拦截器 (`httpInterceptor`)

提供统一的API通信管理：

- 自动Token注入
- 401错误自动重试
- 网络错误处理
- Token监控和刷新

### 4. Word文档操作

通过Office.js API深度集成Word功能：

- 文档内容生成
- 表格创建和格式化
- 页眉页脚设置
- 样式应用

## 📊 数据流和状态管理

### 应用状态流程

```
1. 用户认证 (AuthContext)
   ↓
2. 选择化合物 (CompoundSelector)
   ↓
3. 选择模板区域 (TemplateSelector)
   ↓
4. 处理批次数据 (API调用)
   ↓
5. 生成Word文档 (BatchTemplateService)
```

### 关键数据类型

```typescript
interface BatchData {
    filename: string;
    batch_number: string;
    manufacture_date: string;
    manufacturer: string;
    test_results: Record<string, string>;
}

interface AppState {
    selectedCompound?: Compound;
    selectedTemplate?: Template;
    extractedData: ExtractedField[];
    isLoading: boolean;
    error?: string;
}
```

## 🔌 API集成

### 主要端点

- `GET /compounds` - 获取化合物列表
- `GET /templates` - 获取模板列表
- `GET /documents/check-cache` - 检查缓存数据
- `POST /documents/process-directory` - 处理文档目录
- `GET /health` - 健康检查

### 认证相关

- `GET /auth/status` - 认证状态检查
- `/user/login` - SSO登录重定向

## 🎨 样式和主题

使用Microsoft Fluent UI设计系统：

```typescript
// 主题配置
<FluentProvider theme={webLightTheme}>
  <AuthProvider>
    <App />
  </AuthProvider>
</FluentProvider>
```

## 🔧 配置文件详解

### `manifest.xml`
Office加载项的配置文件，定义：
- 加载项ID和版本
- 支持的Office应用
- 权限和域名配置
- 用户界面元素

### `webpack.config.js`
- 支持开发和生产环境
- HTTPS证书配置
- 代码分割和优化
- 相对路径配置（同域部署）

### `tsconfig.json`
- TypeScript编译配置
- Office.js类型支持
- 严格模式设置

## 🐛 调试和故障排除

### 调试工具

项目提供内置调试工具：

```javascript
// 在浏览器控制台中使用
window.isOfficeAddinEnvironment()    // 检查Office环境
window.testOfficeAddinConnection()   // 测试API连接
window.fullOfficeAddinTest()         // 完整诊断测试
```

### 常见问题

1. **认证失败**
   - 检查企业账户状态
   - 验证重定向URI配置
   - 查看浏览器控制台错误

2. **API连接问题**
   - 确认后端服务运行状态
   - 检查CORS配置
   - 验证网络连接

3. **Word文档生成错误**
   - 确保Office.js API可用
   - 检查数据格式正确性
   - 查看Word版本兼容性

## 🔒 安全考虑

### 认证安全
- 使用Microsoft Entra ID企业级认证
- Token自动刷新和过期处理
- 安全的会话管理

### 数据安全
- HTTPS强制通信
- 敏感数据不在前端缓存
- API请求自动认证

## 📈 性能优化

### 代码优化
- React.memo组件缓存
- useMemo/useCallback Hook优化
- 代码分割和懒加载

### 资源优化
- Webpack代码分割
- 图片和资源压缩
- 缓存策略配置

## 🚀 部署指南

### 生产环境部署

1. **构建项目**
```bash
npm run build
```

2. **部署到Web服务器**
   - 将`dist/`目录内容部署到Web服务器
   - 确保支持HTTPS
   - 配置正确的CORS策略

3. **Office Store发布**
   - 验证manifest.xml
   - 提交到Microsoft AppSource
   - 等待审核通过

### 企业内部部署

1. **内网服务器部署**
   - 配置内网域名
   - 设置SSL证书
   - 更新manifest.xml中的URL

2. **SharePoint集成**
   - 上传到SharePoint应用目录
   - 配置权限和访问控制

## 🔄 后续开发建议

### 功能扩展
- 支持更多文档模板
- 添加批次数据编辑功能
- 实现文档版本控制
- 支持多语言界面

### 技术升级
- 迁移到React 18
- 升级到最新Office.js API
- 实现PWA支持
- 添加离线功能

### 代码改进
- 增加单元测试覆盖
- 实现端到端测试
- 优化错误处理机制
- 改进日志记录

## 📝 开发规范

### 代码风格
- 使用TypeScript严格模式
- 遵循ESLint规则
- 使用Prettier格式化
- 组件和函数命名规范


