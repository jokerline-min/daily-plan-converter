# 夜将晨-日规划转换工具

将 Markdown 格式的日规划表格转换为 Excel 文件的在线工具。

## 功能特点

- **智能解析**：自动识别学生姓名、日期范围和本周核心目标
- **表格转换**：支持解析包含日期、星期和各学科任务的 Markdown 表格
- **Excel 生成**：生成格式化的 Excel 文件，包含合并单元格、字体样式等
- **实时预览**：输入内容后即时显示解析结果
- **示例加载**：内置示例数据，方便快速了解使用方法

## 输入格式

工具支持以下格式的 Markdown 日规划内容：

```markdown
**学生名字日规划（1.13 - 1.18）执行表**
**本周核心目标：** 语文提升古诗与文言文能力，数学紧跟期末复习...

| 日期 | 星期 | 语文 | 数学 | 英语 | 物理 | 化学 | 生物 |
| ---- | ---- | ---- | ---- | ---- | ---- | ---- | ---- |
| 1.13 | 周一 | 任务1<br>任务2 | ... | ... | ... | ... | ... |
| 1.14 | 周二 | ... | ... | ... | ... | ... | ... |
```

## 输出格式

生成的 Excel 文件包含：

- **标题区域**：学生姓名 + 日规划 + 日期范围
- **核心目标行**：本周核心目标内容
- **表头行**：日期、星期、各学科名称
- **数据行**：每天的学习任务，日期和星期列为红色粗体

## 技术栈

- **前端框架**：React 19 + TypeScript
- **UI 组件**：shadcn/ui + Tailwind CSS 4
- **Excel 生成**：xlsx-js-style
- **构建工具**：Vite

## 本地开发

### 环境要求

- Node.js 18+
- pnpm 10+

### 安装依赖

```bash
pnpm install
```

### 启动开发服务器

```bash
pnpm dev
```

### 构建生产版本

```bash
pnpm build
```

## 项目结构

```
daily-plan-converter/
├── client/
│   ├── src/
│   │   ├── components/    # UI 组件
│   │   ├── lib/
│   │   │   └── converter.ts  # 核心转换逻辑
│   │   ├── pages/
│   │   │   └── Home.tsx   # 主页面
│   │   ├── App.tsx        # 应用入口
│   │   └── index.css      # 全局样式
│   └── index.html
├── package.json
└── README.md
```

## 核心文件说明

### converter.ts

包含 Markdown 解析和 Excel 生成的核心逻辑：

- `parseMarkdown()` - 解析 Markdown 内容，提取学生信息和表格数据
- `downloadExcel()` - 生成并下载 Excel 文件
- `exampleMarkdown` - 示例 Markdown 内容

### Home.tsx

主页面组件，包含：

- Markdown 输入框
- 解析结果展示
- 下载按钮
- 使用说明

## 相关项目

- [周规划转换工具](https://github.com/jokerline-min/weekly-plan-converter) - 将周规划 Markdown 转换为 Excel

## 许可证

MIT License
