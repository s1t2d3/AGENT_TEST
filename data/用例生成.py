# gen_cases.py
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ---------- 表头 ----------
HEADERS = [
    "模块", "用例ID", "用例标题", "优先级", "前置条件",
    "请求方法", "请求路径", "请求头", "url参数", "data参数",
    "json参数", "文件参数", "检验字段", "预期结果",
    "json提取", "sql提取", "备注", "是否执行用例"
]
EN_HEADERS = [
    "feature", "id", "title", "priority", "precondition",
    "method", "path", "headers", "params", "data",
    "json", "files", "check", "expected",
    "jsonExData", "sqlExData", "note", "is_true"
]

# ---------- 用例数据（按调整后的执行顺序） ----------
# 每条: [模块, 用例ID, 用例标题, 优先级, 前置条件, 方法, 路径, 请求头, url参数, data参数, json参数, 文件参数, 检验字段, 预期结果, json提取, sql提取, 备注, 是否执行]
CASES = [
    # ===== 注册 =====
    ["认证-注册", "AUTH_REG_001", "正常注册新用户", "高", "用户名不存在", "POST", "/api/auth/register", "", "", "", '{"username": "testuser", "password": "123456", "confirm_password": "123456"}', "", "$..msg", "注册成功", "", "", "", True],
    ["认证-注册", "AUTH_REG_002", "用户名为空", "高", "无", "POST", "/api/auth/register", "", "", "", '{"username": "", "password": "123456", "confirm_password": "123456"}', "", "$..msg", "用户名和密码不能为空", "", "", "", True],
    ["认证-注册", "AUTH_REG_003", "密码为空", "高", "无", "POST", "/api/auth/register", "", "", "", '{"username": "testuser", "password": "", "confirm_password": ""}', "", "$..msg", "用户名和密码不能为空", "", "", "", True],
    ["认证-注册", "AUTH_REG_004", "用户名长度不足2位", "中", "无", "POST", "/api/auth/register", "", "", "", '{"username": "a", "password": "123456", "confirm_password": "123456"}', "", "$..msg", "用户名至少2位", "", "", "边界值测试", True],
    ["认证-注册", "AUTH_REG_005", "用户名刚好2位", "中", "用户名不存在", "POST", "/api/auth/register", "", "", "", '{"username": "ab", "password": "123456", "confirm_password": "123456"}', "", "$..msg", "注册成功", "", "", "边界值测试", True],
    ["认证-注册", "AUTH_REG_006", "密码长度不足6位", "中", "无", "POST", "/api/auth/register", "", "", "", '{"username": "testuser", "password": "12345", "confirm_password": "12345"}', "", "$..msg", "密码长度至少6位", "", "", "边界值测试", True],
    ["认证-注册", "AUTH_REG_007", "密码刚好6位", "中", "用户名不存在", "POST", "/api/auth/register", "", "", "", '{"username": "testuser2", "password": "123456", "confirm_password": "123456"}', "", "$..msg", "注册成功", "", "", "边界值测试", True],
    ["认证-注册", "AUTH_REG_008", "两次密码不一致", "高", "无", "POST", "/api/auth/register", "", "", "", '{"username": "testuser", "password": "123456", "confirm_password": "654321"}', "", "$..msg", "两次输入的密码不一致", "", "", "", True],
    ["认证-注册", "AUTH_REG_009", "用户名已存在", "高", "已存在用户 testuser", "POST", "/api/auth/register", "", "", "", '{"username": "testuser", "password": "123456", "confirm_password": "123456"}', "", "$..msg", "用户名已存在", "", "", "", True],
    ["认证-注册", "AUTH_REG_010", "用户名前后有空格（自动trim）", "低", "用户名不存在", "POST", "/api/auth/register", "", "", "", '{"username": "  spaceuser  ", "password": "123456", "confirm_password": "123456"}', "", "$..msg", "注册成功", "", "", "检查strip处理", True],

    # ===== 登录 =====
    ["认证-登录", "AUTH_LOGIN_001", "正常登录", "高", "已存在用户 testuser/123456", "POST", "/api/auth/login", "", "", "", '{"username": "testuser", "password": "123456", "remember": false}', "", "$..msg", "登录成功", "", "", "", True],
    ["认证-登录", "AUTH_LOGIN_002", "登录并记住我", "中", "已存在用户 testuser/123456", "POST", "/api/auth/login", "", "", "", '{"username": "testuser", "password": "123456", "remember": true}', "", "$..msg", "登录成功", "", "", "7天有效期", True],
    ["认证-登录", "AUTH_LOGIN_003", "用户名为空", "高", "无", "POST", "/api/auth/login", "", "", "", '{"username": "", "password": "123456"}', "", "$..msg", "用户名和密码不能为空", "", "", "", True],
    ["认证-登录", "AUTH_LOGIN_004", "密码为空", "高", "无", "POST", "/api/auth/login", "", "", "", '{"username": "testuser", "password": ""}', "", "$..msg", "用户名和密码不能为空", "", "", "", True],
    ["认证-登录", "AUTH_LOGIN_005", "用户名不存在", "高", "无", "POST", "/api/auth/login", "", "", "", '{"username": "nouser", "password": "123456"}', "", "$..msg", "用户名或密码错误", "", "", "", True],
    ["认证-登录", "AUTH_LOGIN_006", "密码错误", "高", "已存在用户 testuser/123456", "POST", "/api/auth/login", "", "", "", '{"username": "testuser", "password": "wrongpwd"}', "", "$..msg", "用户名或密码错误", "", "", "", True],

    # ===== 已登录状态检查（必须放在登出前） =====
    ["认证-状态", "AUTH_CHECK_001", "已登录检查状态", "高", "已登录", "GET", "/api/auth/check", "", "", "", "", "", "$..msg", "已登录", "", "", "", True],

    # ===== 会话管理（已登录部分） =====
    ["会话管理", "SESSION_001", "获取所有会话（已登录）", "高", "已登录", "GET", "/api/sessions", "", "", "", "", "", "$..msg", "获取会话列表成功", "", "", "", True],
    ["会话管理", "SESSION_003", "创建新会话（带消息）", "高", "已登录", "POST", "/api/sessions", "", "", "", '{"messages": [{"role": "user", "content": "你好，请介绍一下Python"}]}', "", "$..msg", "创建会话成功", "$..data.session.id", "session_id", "", True],
    ["会话管理", "SESSION_004", "创建新会话（无消息）", "中", "已登录", "POST", "/api/sessions", "", "", "", '{"messages": []}', "", "$..msg", "创建会话成功", "$..data.session.id", "session_id", "", True],
    ["会话管理", "SESSION_005", "创建会话标题超长截断", "中", "已登录", "POST", "/api/sessions", "", "", "", '{"messages": [{"role": "user", "content": "这是一条非常非常长的消息用来测试标题截断功能是否正常工作"}]}', "", "$..msg", "创建会话成功", "$..data.session.id", "session_id", "边界值测试", True],
    ["会话管理", "SESSION_007", "更新会话消息", "高", "已登录，存在会话 session_id", "PUT", "/api/sessions/${session_id}", "", "", "", '{"messages": [{"role": "user", "content": "新消息"}]}', "", "$..msg", "更新会话成功", "", "", "", True],
    ["会话管理", "SESSION_008", "更新会话标题", "中", "已登录，存在会话 session_id", "PUT", "/api/sessions/${session_id}", "", "", "", '{"title": "自定义标题"}', "", "$..msg", "更新会话成功", "", "", "", True],
    ["会话管理", "SESSION_009", "更新不存在的会话", "中", "已登录", "PUT", "/api/sessions/nonexistent", "", "", "", '{"messages": []}', "", "$..msg", "会话不存在", "", "", "", True],
    ["会话管理", "SESSION_011", "删除会话", "高", "已登录，存在会话 session_id", "DELETE", "/api/sessions/${session_id}", "", "", "", "", "", "$..msg", "删除会话成功", "", "", "", True],
    ["会话管理", "SESSION_012", "删除不存在的会话", "中", "已登录", "DELETE", "/api/sessions/nonexistent", "", "", "", "", "", "$..msg", "会话不存在", "", "", "", True],

    # ===== 聊天（已登录部分） =====
    ["聊天", "CHAT_001", "正常流式聊天", "高", "已登录", "POST", "/api/chat/stream", "", "", "", '{"prompt": "你好", "session_id": "abc12345"}', "", "$..msg", "回复完成", "", "", "", True],
    ["聊天", "CHAT_002", "空问题", "高", "已登录", "POST", "/api/chat/stream", "", "", "", '{"prompt": "", "session_id": "abc12345"}', "", "$..msg", "问题不能为空", "", "", "", True],
    ["聊天", "CHAT_003", "仅空格问题", "中", "已登录", "POST", "/api/chat/stream", "", "", "", '{"prompt": "   ", "session_id": "abc12345"}', "", "$..msg", "问题不能为空", "", "", "strip后为空", True],
    ["聊天", "CHAT_005", "聊天保存到会话（首条消息更新标题）", "高", "已登录，存在空会话 session_id", "POST", "/api/chat/stream", "", "", "", '{"prompt": "测试消息", "session_id": "abc12345"}', "", "$..msg", "回复完成", "", "", "", True],
    ["聊天", "CHAT_006", "聊天异常处理", "中", "已登录，Agent执行异常", "POST", "/api/chat/stream", "", "", "", '{"prompt": "触发异常", "session_id": "abc12345"}', "", "$..msg", "发生错误", "", "", "异常场景", True],
    ["兼容性", "COMPAT_002", "不存在的session_id聊天", "中", "已登录", "POST", "/api/chat/stream", "", "", "", '{"prompt": "你好", "session_id": "nonexistent"}', "", "$..msg", "回复完成", "", "", "", True],
    ["兼容性", "COMPAT_003", "不传session_id聊天", "中", "已登录", "POST", "/api/chat/stream", "", "", "", '{"prompt": "你好"}', "", "$..msg", "回复完成", "", "", "", True],

    # ===== 页面（已登录） =====
    ["页面", "PAGE_002", "已登录访问首页", "高", "已登录", "GET", "/", "", "", "", "", "", "显示聊天页面", "", "", "", True],

    # ===== 安全（已登录） =====
    ["安全", "SEC_003", "越权访问他人会话", "高", "用户A登录，用户B存在会话", "GET", "/api/sessions", "", "", "", "", "", "$..msg", "获取会话列表成功", "", "", "安全测试", True],

    # ===== 登出（收尾） =====
    ["认证-登出", "AUTH_LOGOUT_001", "已登录用户登出", "高", "已登录", "POST", "/api/auth/logout", "", "", "", "", "", "$..msg", "退出成功", "", "", "", True],
    ["认证-登出", "AUTH_LOGOUT_002", "未登录用户登出", "低", "未登录", "POST", "/api/auth/logout", "", "", "", "", "", "$..msg", "退出成功", "", "", "", True],

    # ===== 未登录（登出后） =====
    ["认证-状态", "AUTH_CHECK_002", "未登录检查状态", "高", "未登录", "GET", "/api/auth/check", "", "", "", "", "", "$..msg", "未登录", "", "", "", True],
    ["会话管理", "SESSION_002", "获取所有会话（未登录）", "高", "未登录", "GET", "/api/sessions", "", "", "", "", "", "$..msg", "未登录", "", "", "", True],
    ["会话管理", "SESSION_006", "创建会话（未登录）", "高", "未登录", "POST", "/api/sessions", "", "", "", '{"messages": []}', "", "$..msg", "未登录", "", "", "", True],
    ["会话管理", "SESSION_010", "更新会话（未登录）", "高", "未登录", "PUT", "/api/sessions/${session_id}", "", "", "", '{"messages": []}', "", "$..msg", "未登录", "", "", "", True],
    ["会话管理", "SESSION_013", "删除会话（未登录）", "高", "未登录", "DELETE", "/api/sessions/${session_id}", "", "", "", "", "", "$..msg", "未登录", "", "", "", True],
    ["聊天", "CHAT_004", "聊天（未登录）", "高", "未登录", "POST", "/api/chat/stream", "", "", "", '{"prompt": "你好", "session_id": "abc12345"}', "", "$..msg", "未登录", "", "", "", True],
    ["页面", "PAGE_001", "未登录访问首页", "高", "未登录", "GET", "/", "", "", "", "", "", "显示登录页面", "", "", "", True],

    # ===== 无登录依赖的杂项 =====
    ["页面", "PAGE_003", "访问不存在的接口", "中", "无", "GET", "/api/notexist", "", "", "", "", "", "$..msg", "接口不存在", "", "", "", True],
    ["安全", "SEC_001", "SQL注入尝试（用户名）", "高", "无", "POST", "/api/auth/login", "", "", "", '{"username": "admin\' OR \'1\'=\'1", "password": "123456"}', "", "$..msg", "用户名或密码错误", "", "", "安全测试", True],
    ["安全", "SEC_002", "XSS尝试（用户名）", "高", "无", "POST", "/api/auth/register", "", "", "", '{"username": "<script>alert(1)</script>", "password": "123456", "confirm_password": "123456"}', "", "$..msg", "注册成功", "", "", "安全测试，检查前端输出编码", True],
    ["安全", "SEC_004", "超长用户名注册", "中", "无", "POST", "/api/auth/register", "", "", "", '{"username": "' + "a" * 1000 + '", "password": "123456", "confirm_password": "123456"}', "", "$..msg", "注册成功", "", "", "边界/安全测试，建议增加长度限制", True],
    ["兼容性", "COMPAT_001", "缺少Content-Type请求", "中", "无", "POST", "/api/auth/login", "", "", "", "", "", "$..msg", "用户名和密码不能为空", "", "", "健壮性测试，建议增加判空", True],
]


def build_excel(output_path="测试用例.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "测试用例"

    # 样式
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 第1行：中文表头
    ws.append(HEADERS)
    # 第2行：英文表头
    ws.append(EN_HEADERS)

    for col_idx in range(1, len(HEADERS) + 1):
        for row_idx in (1, 2):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border

    # 数据行从第3行开始
    for case in CASES:
        ws.append(case)

    # 数据行加边框
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=len(HEADERS)):
        for cell in row:
            cell.border = border

    # 列宽
    widths = [12, 16, 30, 8, 28, 8, 32, 10, 10, 10, 60, 10, 10, 22, 20, 12, 24, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    wb.save(output_path)
    print(f"已生成：{output_path}")
    print(f"用例总数：{len(CASES)}")


if __name__ == "__main__":
    build_excel()