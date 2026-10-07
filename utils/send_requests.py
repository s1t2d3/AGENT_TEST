import pymysql
import requests
import allure
from config.config import *
import logging

logger = logging.getLogger(__name__)
# 全局唯一的 session
http_session = requests.Session()

@allure.step("2.发送请求")
def send_http_request(request_data, use_session=True):
    """
    use_session=True  -> 用全局 session（带 Cookie，用于已登录用例）
    use_session=False -> 用一次性请求（不带 Cookie，用于未登录用例）
    """
    try:
        logger.info(f"开始发送请求：{request_data}")

        if use_session:
            response = http_session.request(**request_data)
        else:
            response = requests.request(**request_data)

        logger.info(f"请求成功：{response.status_code}")
        return response
    except Exception as e:
        logger.error(f"请求异常：{e}")
        raise e

# 执行sql查询语句(sql请求)
def send_jdbc_request(sql_statement,index=0):
    try:
        logger.info(f"开始执行sql查询语句：{sql_statement}")
        conn = pymysql.connect(
            host=HOST,
            port=PORT,
            user=USER,
            password=PASSWORD,
            db=DB,
            charset=CHARSET,
            cursorclass=CURSORCLASS,
            autocommit=AUTOCOMMIT
        )

        cursor = conn.cursor()
        cursor.execute(sql_statement)
        result = cursor.fetchall()
        logger.info(f"sql执行成功：{result[index]}")
        cursor.close()
        conn.close()

        return result[index]
    except Exception as e:
        logger.error(f"sql执行异常：{e}")
        raise e
