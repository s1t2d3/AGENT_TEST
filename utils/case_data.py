# 处理url
import logging
import allure
from config.config import BASE_URL

logger = logging.getLogger(__name__)

def url_handle(path):
    return BASE_URL + path

# 处理用例数据
@allure.step("1.处理用例数据")
def case_handle(case):
    try:
        logger.info(f"开始处理用例数据：{case['id']}")
        case["url"] = url_handle(case["path"])
        request_data = {
            "method": case["method"],
            "url": case["url"],
            "headers": case["headers"],
            "params": case["params"],
            "data": case["data"],
            "json": case["json"],
            "files": case["files"]
        }
        logger.info(f"用例数据处理完成:{request_data}")
        return request_data

    except Exception as e:
        logger.error(f"用例数据处理异常：{e}")
        raise e
