#!/usr/bin/env python3
import json
import requests
import websocket

BASE_URL = "https://cst.uf-tree.com"
CDP_PORT = 9223


def get_auth():
    pages = requests.get(f"http://localhost:{CDP_PORT}/json/list", timeout=10).json()
    ws_url = None
    for p in pages:
        if "cst.uf-tree.com" in p.get("url", ""):
            ws_url = p.get("webSocketDebuggerUrl")
            break
    if not ws_url:
        raise RuntimeError("未找到财税通页面，请先登录并保持 Edge 打开")

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
