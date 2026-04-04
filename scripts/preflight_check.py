#!/usr/bin/env python3
import json
import requests
import websocket

BASE_URL = "https://cst.uf-tree.com"

# 支持的浏览器 CDP 端口
BROWSERS = [
    {"name": "Edge", "port": 9223, "url": "http://localhost:9223/json"},
    {"name": "Chrome", "port": 18800, "url": "http://localhost:18800/json"},
]


def find_browser():
    """自动检测可用的浏览器，优先返回包含财税通页面的浏览器"""
    available = []
    for browser in BROWSERS:
        try:
            pages = requests.get(browser["url"], timeout=6).json()
            # 检查是否有财税通页面
            has_cst = any("cst.uf-tree.com" in p.get("url", "") for p in pages)
            available.append({**browser, "has_cst": has_cst})
        except Exception:
            continue

    if not available:
        return None

    # 优先返回有财税通页面的浏览器
    for b in available:
        if b["has_cst"]:
            return b

    # 如果都没有财税通页面，返回第一个可用的
    return available[0]


def get_auth():
    browser = find_browser()
    if not browser:
        raise RuntimeError("未检测到可用的浏览器。请按以下步骤操作：\n"
                          "1. 打开 Edge 浏览器:\n"
                          "   /Applications/Microsoft\\ Edge.app/Contents/MacOS/Microsoft\\ Edge --remote-debugging-port=9223 --remote-allow-origins=*\n"
                          "2. 或打开 Chrome 浏览器:\n"
                          "   /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=18800 --remote-allow-origins=*\n"
                          "3. 登录 https://cst.uf-tree.com")

    if not browser["has_cst"]:
        raise RuntimeError(f"{browser['name']} 中未发现财税通页面，请先登录 https://cst.uf-tree.com")

    print(f"✅ 检测到 {browser['name']} 浏览器 (端口 {browser['port']})")

    pages = requests.get(browser["url"], timeout=10).json()
    ws_url = None
    for p in pages:
        if "cst.uf-tree.com" in p.get("url", ""):
            ws_url = p.get("webSocketDebuggerUrl")
            break
    if not ws_url:
        raise RuntimeError(f"未找到财税通页面（浏览器：{browser['name']}），请先登录并保持浏览器打开")

    ws = websocket.create_connection(ws_url, timeout=10, suppress_origin=True)
    ws.send(json.dumps({
        "id": 1,
        "method": "Runtime.evaluate",
        "params": {"expression": "localStorage.getItem('vuex')", "returnByValue": True}
    }))

    value = None
    for _ in range(10):
        msg = json.loads(ws.recv())
        if msg.get("id") == 1:
            value = msg.get("result", {}).get("result", {}).get("value")
            break
    ws.close()

    if not value:
        raise RuntimeError("读取登录态失败")
    data = json.loads(value)
    token = data["user"]["token"]
    company_id = data["user"]["company"]["id"]
    return token, company_id


def check_get(url, headers, params, name):
    r = requests.get(url, headers=headers, params=params, timeout=12)
    j = r.json()
    ok = j.get("code") == 200 or j.get("success") is True
    print(("✅" if ok else "❌"), name)
    return ok


def check_post(url, headers, payload, name):
    r = requests.post(url, headers=headers, json=payload, timeout=12)
    j = r.json()
    ok = j.get("code") == 200 or j.get("success") is True
    print(("✅" if ok else "❌"), name)
    return ok


if __name__ == "__main__":
    token, company_id = get_auth()
    headers = {"x-token": token, "Content-Type": "application/json"}
    print(f"✅ 登录态可读: companyId={company_id}")

    checks = [
        check_post(f"{BASE_URL}/api/member/department/queryCompany", headers, {"companyId": company_id}, "queryCompany"),
        check_get(f"{BASE_URL}/api/member/department/queryDepartments", headers, {"companyId": company_id}, "queryDepartments"),
        check_get(f"{BASE_URL}/api/member/role/get/tree", headers, {"companyId": company_id}, "role/get/tree"),
        check_get(f"{BASE_URL}/api/bill/feeTemplate/queryFeeTemplate", headers, {"companyId": company_id, "status": 1, "pageSize": 1000}, "queryFeeTemplate"),
        check_get(f"{BASE_URL}/api/bpm/workflow/queryWorkFlow", headers, {"companyId": company_id, "size": 200}, "queryWorkFlow"),
        check_get(f"{BASE_URL}/api/bill/template/queryTemplateTree", headers, {"companyId": company_id}, "queryTemplateTree"),
    ]

    if all(checks):
        print("\n✅ PRECHECK PASS")
    else:
        print("\n❌ PRECHECK FAIL")
        raise SystemExit(1)
