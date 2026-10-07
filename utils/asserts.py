import jsonpath
import allure
import logging
import json

logger = logging.getLogger(__name__)

# 解析sse响应
def parse_sse(response):
    events = []
    for line in response.text.splitlines():
        line = line.strip()
        if line.startswith("data:"):
            data = line[6:].strip()
            if data:
                events.append(json.loads(data))
    return events

@allure.step("3.断言")
def assert_util(case, response):
    logger.info(f"开始断言:{case['id']}")
    try:
        content_type = response.headers.get("Content-Type", "")

        # 只有真正的 SSE 才走 SSE 解析
        if "text/event-stream" in content_type:
            events = parse_sse(response)
            if not events:
                raise AssertionError(f"SSE 响应为空：{case['id']}")
            last_event = events[-1]
            actual = last_event.get("msg", "")
            assert case["expected"] in actual, \
                f"预期包含 {case['expected']}，实际 {actual}"
            logger.info(f"断言成功:预期结果({case['expected']}) 包含于 实际结果({actual})")
            return

        # 普通 JSON 响应
        if case["check"] and case["check"].startswith("$"):
            result = jsonpath.jsonpath(response.json(), case["check"])
            assert result[0] == case["expected"]
            logger.info(f"断言成功:预期结果({case['expected']}) == 实际结果({result[0]})")
            return

        # HTML 响应
        assert case["expected"] in response.text
        logger.info(f"断言成功:预期结果({case['expected']}) 在页面中找到")

    except Exception as e:
        logger.error(f"断言失败:{case['id']}")
        raise e

    # b.数据库断言
    # if case["sql_check"] and case["sql_expected"]:
    #
    #     # 执行sql查询语句
    #     sql_result = send_jdbc_request(case["sql_check"])
    #
    #     assert sql_result[0] == case["sql_expected"]