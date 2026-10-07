import pytest
from config.config import PATH
from utils.allure_util import allure_init
from utils.case_data import case_handle
from utils.data_extraction import json_extract, jdbc_extract
from utils.excel_util import excel_handle
import logging
from utils.asserts import assert_util
from utils.send_requests import send_http_request, send_jdbc_request
import json

# 配置日志
logger = logging.getLogger(__name__)


class TestRunner:

    all = {}

    @pytest.mark.parametrize("case", excel_handle(PATH))
    def test_case(self, case):

        # 获取全局变量
        all = self.all

        # 模板渲染：把 ${xxx} 替换成 all 里的真实值
        case_str = json.dumps(case, ensure_ascii=False)
        for k, v in all.items():
            case_str = case_str.replace("${" + k + "}", str(v))
        case = json.loads(case_str)

        print("all =", all)
        print("path =", case["path"])

        # allure报告初始化
        allure_init(case)

        logger.info(f"开始执行用例：{case['id']}, {case['feature']}")
        # 1.处理用例数据
        request_data = case_handle(case)

        # 2.发送请求
        if case["precondition"] == "未登录":
            response = send_http_request(request_data, use_session=False)
        else:
            response = send_http_request(request_data, use_session=True)

        # 3.断言
        assert_util(case, response)

        # 4.数据提取
        json_extract(case, response, all)
        # jdbc_extract(case, all)

        logger.info(f"用例执行完成：{case['id']}, {case['feature']}")