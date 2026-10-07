# 数据提取
import jsonpath
import allure
from utils.send_requests import send_jdbc_request
import logging

logger = logging.getLogger(__name__)

def json_extract(case, response,all):
    logger.info("开始数据提取")
    try:
        if case["jsonExData"]:
            with allure.step("4.数据提取"):
                for key, value in case["jsonExData"].items():
                    value = jsonpath.jsonpath(response.json(), value)[0]
                    all[key] = value
    except Exception as e:
        logger.error("数据提取失败")
        raise e

def jdbc_extract(case, response,all):
    logger.info("开始数据提取")
    try:

        if case["sqlExData"]:
            with allure.step("4.数据提取"):
                for key, value in case["sqlExData"].items():
                    sql_result = send_jdbc_request(value)
                    all[key] = sql_result[0]
    except Exception as e:
        logger.error("数据提取失败")
        raise e
