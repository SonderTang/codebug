from flask_app import create_app
from dotenv import load_dotenv

load_dotenv() # 自动加载 .env 文件

app = create_app()

if __name__ == '__main__':
    app.run(port=5001, debug=True, host='0.0.0.0')