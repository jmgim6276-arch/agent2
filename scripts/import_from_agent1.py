#!/usr/bin/env python3
import argparse
import json
import time
from pathlib import Path

import pandas as pd
import requests
import websocket

BASE_URL = "https://cst.uf-tree.com"
CDP_PORT = 9223


def split_values(v):
    t = str(v).strip()
    if not t or t.lower() == "nan":
        return []
    for ch in ["，", "、", ";", "；"]:
        t = t.replace(ch, ",")
    return [x.strip() for x in t.split(",") if x.strip()]


def get_auth():
    pages = requests.get(f"http://localhost:{CDP_PORT}/json/list", timeout=10).json()
    ws_url = next((p.get("webSocketDebuggerUrl") for p in pages if "cst.uf-tree.com" in p.get("url", "")), None)
    if not ws_url:
        raise RuntimeError("未找到财税通页面，请先登录")

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

    data = json.loads(value)
    return data["user"]["token"], data["user"]["company"]["id"], data["user"].get("id")


def read_sheet_with_header(path: Path, sheet: str, header_key: str):
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    header_row = raw.index[raw.apply(lambda r: r.astype(str).str.contains(header_key).any(), axis=1)][0]
    return pd.read_excel(path, sheet_name=sheet, header=header_row)


def main():
    parser = argparse.ArgumentParser(description="导入 Agent1 三表到财税通")
    parser.add_argument("--xlsx", required=True, help="Agent1 生成的三表文件")
    parser.add_argument("--output", default="./agent2_import_report.json", help="导入报告输出路径")
    args = parser.parse_args()

    xlsx = Path(args.xlsx)
    token, company_id, _ = get_auth()
    h = {"x-token": token, "Content-Type": "application/json"}

    report = {
        "companyId": company_id,
        "xlsx": str(xlsx),
        "step1": {"ok": 0, "fail": []},
        "step2": {"relations_ok": 0, "relations_fail": [], "role_by_doc": {}, "leaf_by_doc": {}},
        "step25": {},
        "step3": {"ok": 0, "fail": [], "branch_fee_role": [], "branch_leaf_fee": [], "branch_skip": []},
    }

    # Base maps
    users = requests.post(f"{BASE_URL}/api/member/department/queryCompany", headers=h, json={"companyId": company_id}, timeout=15).json().get("result", {}).get("users", [])
    user_map = {u.get("nickName"): u.get("id") for u in users if u.get("nickName") and u.get("id")}
    deps = requests.get(f"{BASE_URL}/api/member/department/queryDepartments", headers=h, params={"companyId": company_id}, timeout=15).json().get("result", [])
    dep_map = {d.get("title"): d.get("id") for d in deps if d.get("title") and d.get("id")}

    # Step1
    df1 = read_sheet_with_header(xlsx, "01_添加员工", "是否导入")
    df1 = df1[df1["是否导入"].astype(str).str.strip() == "是"]
    for i, row in df1.iterrows():
        name = str(row.get("姓名", "")).strip()
        mobile = str(row.get("手机号", "")).strip()[:11]
        dept = str(row.get("二级部门", "")).strip()
        if not dept or dept.lower() == "nan":
            dept = str(row.get("一级部门名称", "")).strip()
        dep_id = dep_map.get(dept)
        if not (name and mobile and dep_id):
            report["step1"]["fail"].append({"row": int(i + 1), "reason": "姓名/手机号/部门缺失或无效"})
            continue
        payload = {"nickName": name, "mobile": mobile, "departmentIds": [dep_id], "companyId": company_id}
        r = requests.post(f"{BASE_URL}/api/member/userInfo/add", headers=h, json=payload, timeout=12).json()
        if r.get("code") == 200 or r.get("success"):
            report["step1"]["ok"] += 1
        else:
            msg = str(r.get("message", ""))
            if "已" in msg or "存在" in msg:
                report["step1"]["ok"] += 1
            else:
                report["step1"]["fail"].append({"row": int(i + 1), "reason": msg})

    # Fee templates tree
    fee_tree = requests.get(f"{BASE_URL}/api/bill/feeTemplate/queryFeeTemplate", headers=h, params={"companyId": company_id, "status": 1, "pageSize": 1000}, timeout=20).json().get("result", [])
    primary = {p.get("name"): p for p in fee_tree if p.get("parentId") == -1}
    child = {(p.get("id"), c.get("name")): c.get("id") for p in fee_tree for c in (p.get("children") or []) if p.get("id") and c.get("name") and c.get("id")}

    # Step2
    df2 = read_sheet_with_header(xlsx, "02_费用科目配置", "一级费用科目")
    df2 = df2[df2["是否执行"].astype(str).str.strip() == "是"].copy()
    for c in ["一级费用科目", "二级费用科目", "归属单据名称"]:
        df2[c] = df2[c].ffill()

    # ensure fee role group
    requests.post(f"{BASE_URL}/api/member/role/add/group", headers=h, json={"companyId": company_id, "name": "费用角色组"}, timeout=12)
    tree = requests.get(f"{BASE_URL}/api/member/role/get/tree", headers=h, params={"companyId": company_id}, timeout=12).json().get("result", [])

    def fee_roles_map():
        t = requests.get(f"{BASE_URL}/api/member/role/get/tree", headers=h, params={"companyId": company_id}, timeout=12).json().get("result", [])
        m = {}
        for cat in t:
            if cat.get("name") == "费用角色组":
                for rr in cat.get("children", []) or []:
                    if rr.get("name") and rr.get("id"):
                        m[rr["name"]] = rr["id"]
        return m

    fee_roles = fee_roles_map()
    has_people = {}

    for _, row in df2.iterrows():
        p = str(row.get("一级费用科目", "")).strip()
        s = str(row.get("二级费用科目", "")).strip()
        t3 = str(row.get("三级费用科目", "")).strip()
        doc = str(row.get("归属单据名称", "")).strip()
        people = split_values(row.get("单据适配人员（多人用中文逗号）", ""))
        if not (p and s and doc):
            continue

        has_people[doc] = has_people.get(doc, False) or bool(people)

        pid = (primary.get(p) or {}).get("id")
        sid = child.get((pid, s)) if pid else None
        if not sid:
            continue

        leaf_id = sid
        if t3 and t3.lower() != "nan":
            sec = requests.get(f"{BASE_URL}/api/bill/feeTemplate/getFeeTemplateById", headers=h, params={"id": sid, "companyId": company_id}, timeout=12).json().get("result", {})
            third = {c.get("name"): c.get("id") for c in (sec.get("children") or []) if c.get("name") and c.get("id")}
            if third.get(t3):
                leaf_id = third[t3]

        report["step2"]["leaf_by_doc"].setdefault(doc, [])
        if leaf_id not in report["step2"]["leaf_by_doc"][doc]:
            report["step2"]["leaf_by_doc"][doc].append(leaf_id)

        # 条件触发费用角色链路
        if people:
            rid = fee_roles.get(doc)
            if not rid:
                requests.post(f"{BASE_URL}/api/member/role/add", headers=h, json={"companyId": company_id, "name": doc, "dataType": "FEE_TYPE"}, timeout=12)
                fee_roles = fee_roles_map()
                rid = fee_roles.get(doc)
            uids = [user_map[n] for n in people if n in user_map]
            if rid and uids:
                rel = requests.post(
                    f"{BASE_URL}/api/member/role/add/relation",
                    headers=h,
                    json={"roleId": rid, "userIds": uids, "feeTemplateIds": [leaf_id], "companyId": company_id},
                    timeout=12,
                ).json()
                if rel.get("code") == 200:
                    report["step2"]["relations_ok"] += 1
                    report["step2"]["role_by_doc"][doc] = rid
                else:
                    report["step2"]["relations_fail"].append({"doc": doc, "message": rel.get("message")})

    # Step2.5
    wfs = requests.get(f"{BASE_URL}/api/bpm/workflow/queryWorkFlow", headers=h, params={"companyId": company_id, "size": 200}, timeout=12).json().get("result", []) or []
    workflow_id = None
    workflow_name = None
    for w in wfs:
        if "通用审批" in str(w.get("tpName", "")):
            workflow_id = w.get("id")
            workflow_name = w.get("tpName")
            break
    if not workflow_id and wfs:
        workflow_id = wfs[0].get("id")
        workflow_name = wfs[0].get("tpName")
    report["step25"] = {"workflowId": workflow_id, "workflowName": workflow_name, "count": len(wfs)}

    # Step3
    roles_vis = {}
    tree_all = requests.get(f"{BASE_URL}/api/member/role/get/tree", headers=h, params={"companyId": company_id}, timeout=12).json().get("result", [])
    for cat in tree_all:
        if cat.get("name") == "费用角色组":
            continue
        for rr in cat.get("children", []) or []:
            if rr.get("name") and rr.get("id"):
                roles_vis[rr["name"]] = rr["id"]

    groups = requests.get(f"{BASE_URL}/api/bill/template/queryTemplateTree", headers=h, params={"companyId": company_id}, timeout=12).json().get("result", []) or []
    group_map = {g.get("name") or g.get("title"): g.get("id") for g in groups if (g.get("name") or g.get("title")) and g.get("id")}

    df3 = read_sheet_with_header(xlsx, "03_单据表", "单据模板名称")
    df3 = df3[df3["是否创建"].astype(str).str.strip() == "是"].copy()
    df3["单据分组（一级目录）"] = df3["单据分组（一级目录）"].ffill()

    type_map = {"报销单": "EXPENSE", "借款单": "LOAN", "批量付款单": "PAYMENT", "申请单": "REQUISITION"}

    for _, row in df3.iterrows():
        group_name = str(row.get("单据分组（一级目录）", "")).strip()
        doc_type = str(row.get("单据大类（二级目录）", "")).strip()
        doc_name = str(row.get("单据模板名称", "")).strip()
        vis_type = str(row.get("可见范围类型", "")).strip()
        vis_obj = str(row.get("可见范围对象", "")).strip()

        if group_name not in group_map:
            requests.post(f"{BASE_URL}/api/bill/template/createTemplateGroup", headers=h, json={"name": group_name, "companyId": company_id}, timeout=12)
            time.sleep(0.4)
            groups = requests.get(f"{BASE_URL}/api/bill/template/queryTemplateTree", headers=h, params={"companyId": company_id}, timeout=12).json().get("result", []) or []
            group_map = {g.get("name") or g.get("title"): g.get("id") for g in groups if (g.get("name") or g.get("title")) and g.get("id")}

        targets = split_values(vis_obj)
        role_ids = [roles_vis[t] for t in targets if vis_type == "角色" and t in roles_vis]
        user_ids = [user_map[t] for t in targets if vis_type == "员工" and t in user_map]
        dep_ids = [dep_map[t] for t in targets if vis_type == "部门" and t in dep_map]

        payload = {
            "applyRelateFlag": True,
            "applyRelateNecessary": False,
            "businessType": "PRIVATE",
            "companyId": company_id,
            "componentJson": [],
            "departmentIds": dep_ids,
            "feeIds": [],
            "feeScopeFlag": False,
            "groupId": group_map.get(group_name),
            "icon": "md-pricetag",
            "iconColor": "#4c7cc3",
            "loanIds": [],
            "name": f"{doc_name}_{int(time.time()) % 100000}",
            "payFlag": True,
            "requestScope": False,
            "requisitionIds": [],
            "roleIds": role_ids,
            "status": "ACTIVE",
            "type": type_map.get(doc_type, "EXPENSE"),
            "userIds": user_ids,
            "userScopeFlag": True,
            "workFlowId": workflow_id,
        }
        if payload["type"] == "REQUISITION":
            payload["applyContentType"] = "TEXT"

        # 费用限制分支
        if payload["type"] in ("LOAN", "REQUISITION"):
            report["step3"]["branch_skip"].append(doc_name)
        elif has_people.get(doc_name, False) and report["step2"]["role_by_doc"].get(doc_name):
            # 费用角色限制：当前系统字段未完全公开，先记录分支命中
            report["step3"]["branch_fee_role"].append({"doc": doc_name, "feeRoleId": report["step2"]["role_by_doc"][doc_name]})
        else:
            fee_ids = report["step2"]["leaf_by_doc"].get(doc_name, [])
            if fee_ids:
                payload["feeIds"] = fee_ids
                payload["feeScopeFlag"] = True
            report["step3"]["branch_leaf_fee"].append({"doc": doc_name, "feeIds": fee_ids})

        cr = requests.post(f"{BASE_URL}/api/bill/template/createTemplate", headers=h, json=payload, timeout=15).json()
        if cr.get("code") == 200 and cr.get("success"):
            report["step3"]["ok"] += 1
        else:
            report["step3"]["fail"].append({"doc": doc_name, "message": cr.get("message")})

    Path(args.output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print("✅ 导入完成")
    print(json.dumps({
        "step1_ok": report["step1"]["ok"],
        "step2_relations_ok": report["step2"]["relations_ok"],
        "step3_ok": report["step3"]["ok"],
        "output": args.output,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
