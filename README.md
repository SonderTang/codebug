# codebug
用于前端代码和缺陷数据爬取

code_bug_project/
├── config/
│   ├── database.conf  # 数据库配置文件
│   ├── api_keys.conf  # API密钥配置文件
├── src/
│   ├── main.py  # 主程序入口
│   ├── flask_app/  # Flask应用相关代码
│   │   ├── __init__.py
│   │   ├── routes.py  # 路由定义
│   ├── data_processing/  # 数据处理相关代码
│   │   ├── __init__.py
│   │   ├── bug_query.py  # Bug查询相关函数
│   │   ├── code_query.py  # 代码查询相关函数
│   ├── utils/  # 工具函数
│   │   ├── __init__.py
│   │   ├── db_utils.py  # 数据库操作工具
│   │   ├── excel_utils.py  # Excel操作工具
├── tests/  # 单元测试
│   ├── test_bug_query.py
│   ├── test_code_query.py
├── docs/  # 项目文档
│   ├── README.md
│   ├── development_guide.md
├── requirements.txt  # 项目依赖
