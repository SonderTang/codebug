# CodeBug - 前端代码与缺陷数据采集系统

> ​**项目状态**: Active | ​**最后更新**: 2025-04-16 | ​**版本**: 1.0.0  
> [![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)  
> [![Build Status](https://img.shields.io/github/actions/workflow/status/yourusername/codebug/ci.yml)](https://github.com/yourusername/codebug/actions)

## 📂 项目结构
```text
code_bug_project/
├── config/                 # 配置管理
│   ├── database.conf      # 数据库连接配置
│   └── api_keys.conf      # API密钥管理
├── src/                   # 核心源码
│   ├── main.py            # Flask服务入口
│   ├── flask_app/         # Web服务模块
│   │   ├── __init__.py
│   │   └── routes.py      # API路由定义
│   ├── data_processing/   # 数据处理流水线
│   │   ├── bug_query.py   # 缺陷数据ETL模块
│   │   └── code_query.py  # 代码仓库爬虫模块
├── utils/                 # 工具函数库
│   ├── db_utils.py        # 数据库操作工具
│   └── excel_utils.py     # Excel报表生成工具
├── tests/                 # 单元测试(覆盖率≥85%)
│   ├── test_bug_query.py
│   └── test_code_query.py
└── docs/                  # 项目文档
    ├── development_guide.md  # 开发规范
    └── api_spec.md          # API接口文档

