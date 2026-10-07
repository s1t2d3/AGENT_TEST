import allure

def allure_init(case):
    # allure初始化
    if case["feature"]:
        allure.dynamic.feature(case["feature"])

    # if case["story"]:
    #     allure.dynamic.story(case["story"])

    if case["title"]:
        allure.dynamic.title(case["id"] + "--" + case["title"])