import os
import pymysql
BASE_URL = "http://127.0.0.1:5005"
PATH = "./data/测试用例.xlsx"
EXCEL_SHEET = "测试用例"

# MySQL数据库配置
HOST = "127.0.0.1",
PORT = 3306,
USER = "root",
PASSWORD = os.environ.get("DB_PASSWORD"),
DB = "test",
CHARSET = "utf8mb4",
CURSORCLASS = pymysql.cursors.DictCursor,
AUTOCOMMIT = True,
