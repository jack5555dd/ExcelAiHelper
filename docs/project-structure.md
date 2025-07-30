# Excel COM 加载项项目结构

## 目录结构

```
Excel-COM-Addin/
│
├── src/                    # 源代码
│   ├── main/               # 主要代码
│   │   └── cs/             # C# 源文件
│   │       ├── Connect.cs  # 主要加载项实现
│   │       └── SimpleConnect.cs # 简化版加载项实现
│   │
│   ├── resources/          # 资源文件
│   │   └── reg/            # 注册表文件
│   │
│   └── scripts/            # 脚本文件
│       ├── build/          # 构建脚本
│       ├── install/        # 安装脚本
│       ├── register/       # 注册脚本
│       └── test/           # 测试脚本
│
├── bin/                    # 编译输出
│   └── Debug/              # 调试版本
│
├── config/                 # 配置文件
│   └── settings/           # 设置文件
│
└── docs/                   # 文档
    └── help/               # 帮助文档
```

## 主要组件

### 1. 源代码 (src/)

- **Connect.cs**: 主要COM加载项类，实现了Excel加载项的主要功能
- **SimpleConnect.cs**: 简化版加载项实现，作为备用或示例

### 2. 资源文件 (src/resources/)

- **reg/**: 包含注册表文件，用于COM组件注册

### 3. 脚本 (src/scripts/)

- **build/**: 构建项目的脚本
- **install/**: 安装加载项的脚本
- **register/**: 注册和取消注册COM组件的脚本
- **test/**: 用于测试加载项的脚本

### 4. 配置 (config/)

- **settings/**: 包含应用程序设置文件

### 5. 文档 (docs/)

- **help/**: 帮助文档，包括安装说明和故障排除指南 