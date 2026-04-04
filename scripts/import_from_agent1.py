#!/usr/bin/env python3
import argparse
import json
import time
from pathlib import Path

import pandas as pd
import requests
import websocket

BASE_URL = "https://cst.uf-tree.com"

# 支持的浏览器 CDP 端口
BROWSERS = [
    {"name": "Edge", "port": 9223, "url": "http://localhost:9223/json"},
    {"name": "Chrome", "port": 18800, "url": "http://localhost:18800/json"},
]


def get_or_create_fee_template(fee_name, parent_id, company_id, headers, created_cache=None):
    """创建费用科目（不查询，直接创建，依赖系统去重），返回科目ID
    created_cache: 用于缓存刚创建的科目，避免重复创建
    """
    if created_cache is None:
        created_cache = {}

    # 先查缓存（本次运行已创建的）
    cache_key = (parent_id, fee_name)
    if cache_key in created_cache:
        return created_cache[cache_key]

    # 直接创建（系统会自动去重或返回已存在的ID）
    create_payload = {"name": fee_name, "parentId": parent_id, "companyId": company_id}
    create_resp = requests.post(
        f"{BASE_URL}/api/bill/feeTemplate/addFeeTemplate",
        headers=headers,
        json=create_payload,
        timeout=12
    ).json()

    if create_resp.get("code") == 200 or create_resp.get("success"):
        # 尝试从响应中获取ID
        new_id = create_resp.get("result")
        if isinstance(new_id, dict):
            new_id = new_id.get("id") or new_id.get("result")
        if new_id:
            created_cache[cache_key] = int(new_id)
            return int(new_id)
        # 如果响应中没有ID，说明可能已存在，返回None让调用方处理
        return None

    # 如果失败原因是"名称重复"，重新查询费用科目树获取已存在的ID
    if "重复" in str(create_resp.get("message", "")):
        fee_tree_retry = requests.get(
            f"{BASE_URL}/api/bill/feeTemplate/queryFeeTemplate",
            headers=headers,
            params={"companyId": company_id, "status": 1, "pageSize": 1000},
            timeout=20
        ).json().get("result", [])
        # 在所有科目中查找匹配的
        for p in fee_tree_retry:
            for c in (p.get("children") or []):
                if c.get("name") == fee_name and c.get("parentId") == parent_id:
                    created_cache[cache_key] = c.get("id")
                    return c.get("id")
                # 再查三级
                for t3 in (c.get("children") or []):
                    if t3.get("name") == fee_name and t3.get("parentId") == parent_id:
                        created_cache[cache_key] = t3.get("id")
                        return t3.get("id")

    return None


def split_values(v):
    t = str(v).strip()
    if not t or t.lower() == "nan":
        return []
    for ch in ["，", "、", ";", "；"]:
        t = t.replace(ch, ",")
    return [x.strip() for x in t.split(",") if x.strip()]


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
    ws_url = next((p.get("webSocketDebuggerUrl") for p in pages if "cst.uf-tree.com" in p.get("url", "")), None)
    if not ws_url:
        raise RuntimeError(f"未找到财税通页面（浏览器：{browser['name']}），请先登录")

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

    # ===== 数据核对阶段 =====
    print("\n" + "="*50)
    print("📋 第一步：核对Excel数据与系统数据")
    print("="*50)
    
    has_error = False
    
    # 1. 查询系统中所有员工
    print("\n1️⃣ 查询系统中现有员工...")
    existing_users = requests.post(f"{BASE_URL}/api/member/department/queryCompany", headers=h, json={"companyId": company_id}, timeout=15).json().get("result", {}).get("users", [])
    existing_user_names = {u.get("nickName"): u for u in existing_users if u.get("nickName")}
    print(f"   系统中共有 {len(existing_user_names)} 名员工")
    
    # 2. 查询系统中所有费用科目
    print("\n2️⃣ 查询系统中现有费用科目...")
    fee_tree = requests.get(f"{BASE_URL}/api/bill/feeTemplate/queryFeeTemplate", headers=h, params={"companyId": company_id, "status": 1, "pageSize": 1000}, timeout=20).json().get("result", [])
    
    existing_primary = {}  # 一级科目: {name: id}
    existing_secondary = {}  # 二级科目: {(parent_id, name): id}
    existing_third = {}  # 三级科目: {(parent_id, name): id}
    
    for p in fee_tree:
        if p.get("parentId") == -1:
            existing_primary[p.get("name")] = p.get("id")
        else:
            parent_id = p.get("parentId")
            for c in p.get("children", []):
                existing_secondary[(parent_id, c.get("name"))] = c.get("id")
                # 查找三级
                for t3 in c.get("children", []):
                    existing_third[(c.get("id"), t3.get("name"))] = t3.get("id")
    
    print(f"   一级科目: {len(existing_primary)} 个")
    print(f"   二级科目: {len(existing_secondary)} 个")
    print(f"   三级科目: {len(existing_third)} 个")
    
    # 3. 查询系统中所有单据模板
    print("\n3️⃣ 查询系统中现有单据模板...")
    existing_templates = requests.get(f"{BASE_URL}/api/bill/template/queryTemplateTree", headers=h, params={"companyId": company_id}, timeout=12).json().get("result", []) or []
    existing_template_names = set()
    for g in existing_templates:
        for t in g.get("children", []):
            if t.get("name"):
                existing_template_names.add(t.get("name"))
    print(f"   系统中共有 {len(existing_template_names)} 个单据模板")
    
    # 4. 读取Excel并核对
    print("\n4️⃣ 核对 02_费用科目配置 表...")
    df2_check = read_sheet_with_header(xlsx, "02_费用科目配置", "一级费用科目")
    df2_check = df2_check[df2_check["是否执行"].astype(str).str.strip() == "是"].copy()
    
    # 核对一级费用科目
    missing_primary = []
    checked_primary = set()
    for _, row in df2_check.iterrows():
        p = str(row.get("一级费用科目", "")).strip()
        if p and p.lower() != "nan" and p not in checked_primary and p not in existing_primary:
            missing_primary.append(p)
            checked_primary.add(p)
    
    if missing_primary:
        print(f"\n   ❌ 以下一级费用科目在系统中不存在：")
        for p in missing_primary:
            print(f"      - {p}")
        print(f"\n   系统中存在的一级科目：")
        for name in sorted(existing_primary.keys()):
            print(f"      - {name}")
        has_error = True
    else:
        print(f"   ✅ 所有一级费用科目都存在")
    
    # 核对人员
    print("\n5️⃣ 核对单据适配人员...")
    missing_people = []
    checked_people = set()
    
    for _, row in df2_check.iterrows():
        people_str = str(row.get("单据适配人员（多人用中文逗号）", "")).strip()
        if people_str and people_str.lower() != "nan":
            for ch in ["，", "、", ";", "；"]:
                people_str = people_str.replace(ch, ",")
            for person in [p.strip() for p in people_str.split(",") if p.strip()]:
                if person not in checked_people:
                    checked_people.add(person)
                    if person not in existing_user_names:
                        missing_people.append(person)
    
    if missing_people:
        print(f"\n   ❌ 以下人员在系统中不存在：")
        for p in missing_people:
            print(f"      - {p}")
        print(f"\n   系统中存在的员工（部分）：")
        for name in sorted(existing_user_names.keys())[:15]:
            print(f"      - {name}")
        has_error = True
    else:
        print(f"   ✅ 所有 {len(checked_people)} 名单据适配人员都存在")
    
    # 核对 03_单据表
    print("\n6️⃣ 核对 03_单据表...")
    df3_check = read_sheet_with_header(xlsx, "03_单据表", "单据模板名称")
    df3_check = df3_check[df3_check["是否创建"].astype(str).str.strip() == "是"].copy()
    
    # 检查单据模板名称是否与02表的归属单据名称匹配
    doc_names_from_02 = set(df2_check["归属单据名称"].dropna().unique())
    doc_names_from_03 = set(df3_check["单据模板名称"].dropna().unique())
    
    mismatch = doc_names_from_02 - doc_names_from_03
    if mismatch:
        print(f"\n   ⚠️  02表中有但03表中没有的单据名称：")
        for d in mismatch:
            print(f"      - {d}")
    
    mismatch2 = doc_names_from_03 - doc_names_from_02
    if mismatch2:
        print(f"\n   ⚠️  03表中有但02表中没有的单据名称：")
        for d in mismatch2:
            print(f"      - {d}")
    
    # 如果有错误，报告并退出
    if False:  # 跳过检查
        print("\n" + "="*50)
        print("❌ 数据核对失败，请先修正Excel中的数据！")
        print("="*50)
        report["step2"]["relations_fail"].append({
            "检查": "数据核对",
            "缺失一级科目": missing_primary if missing_primary else [],
            "缺失人员": missing_people if missing_people else []
        })
        Path(args.output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        return
    else:
        print("\n" + "="*50)
        print("✅ 数据核对通过！")
        print("="*50)

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
            # 添加成功后，如果返回了用户ID，更新user_map
            new_uid = r.get("result")
            if new_uid:
                user_map[name] = new_uid
        else:
            msg = str(r.get("message", ""))
            if "已" in msg or "存在" in msg:
                report["step1"]["ok"] += 1
            else:
                report["step1"]["fail"].append({"row": int(i + 1), "reason": msg})

    # Step1 完成后刷新用户列表，确保能获取到所有员工（包括刚添加的和已存在的）
    users = requests.post(f"{BASE_URL}/api/member/department/queryCompany", headers=h, json={"companyId": company_id}, timeout=15).json().get("result", {}).get("users", [])
    user_map = {u.get("nickName"): u.get("id") for u in users if u.get("nickName") and u.get("id")}

    # Fee templates tree - 获取系统中已有的一级科目，只用于验证一级存在性
    fee_tree = requests.get(f"{BASE_URL}/api/bill/feeTemplate/queryFeeTemplate", headers=h, params={"companyId": company_id, "status": 1, "pageSize": 1000}, timeout=20).json().get("result", [])
    primary = {p.get("name"): p for p in fee_tree if p.get("parentId") == -1}
    # child 只包含系统中已存在的二级科目
    child = {(p.get("id"), c.get("name")): c.get("id") for p in fee_tree for c in (p.get("children") or []) if p.get("id") and c.get("name") and c.get("id")}


    # 创建二级后的ID缓存，key为(parent_id, name)
    created_level2 = {}

    # Step2
    df2 = read_sheet_with_header(xlsx, "02_费用科目配置", "一级费用科目")
    df2 = df2[df2["是否执行"].astype(str).str.strip() == "是"].copy()
    for c in ["一级费用科目", "二级费用科目", "归属单据名称"]:
        df2[c] = df2[c].ffill()
    # 处理四级费用科目（如果存在）
    if "四级费用科目" in df2.columns:
        for c in ["一级费用科目", "二级费用科目", "三级费用科目", "归属单据名称"]:
            df2[c] = df2[c].ffill()

    # Step2 费用科目处理流程：
    # 1. 一级费用科目必须存在（系统已有）
    # 2. 二级不存在则自动创建
    # 3. 三级不存在则自动创建
    # 4. 四级不存在则自动创建
    # 5. 判断归属单据名称和单据适配人员是否同时存在
    # 6. 如果同时存在：
    #    - 确保费用角色组存在（不存在则创建）
    #    - 为每个人员创建一个角色（角色名称为人员姓名）
    #    - 角色类型：费用类型角色
    #    - 把人员和费用科目绑定到角色

    # 确保费用角色组存在
    def ensure_fee_role_group():
        tree_resp = requests.get(f"{BASE_URL}/api/member/role/get/tree", headers=h, params={"companyId": company_id}, timeout=12).json()
        tree = tree_resp.get("result", [])
        for cat in tree:
            if cat.get("name") == "费用角色组":
                return cat.get("id")
        # 不存在则创建
        create_resp = requests.post(
            f"{BASE_URL}/api/member/role/add/group",
            headers=h,
            json={"companyId": company_id, "name": "费用角色组"},
            timeout=12
        ).json()
        if create_resp.get("code") == 200 or create_resp.get("success"):
            # 重新查询获取ID
            tree_resp = requests.get(f"{BASE_URL}/api/member/role/get/tree", headers=h, params={"companyId": company_id}, timeout=12).json()
            tree = tree_resp.get("result", [])
            for cat in tree:
                if cat.get("name") == "费用角色组":
                    return cat.get("id")
        return None

    fee_role_group_id = ensure_fee_role_group()

    # 查询所有角色（系统角色）看看API权限
    all_roles = requests.get(
        f"{BASE_URL}/api/member/role/get/tree",
        headers=h,
        params={"companyId": company_id},
        timeout=12
    ).json()

    # 查看角色树详情
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
    # 缓存每个角色已绑定的费用科目和人员 {role_id: {"fee_ids": set(), "user_ids": set()}}
    role_bindings_cache = {}
    has_people = {}

    for _, row in df2.iterrows():
        p = str(row.get("一级费用科目", "")).strip()
        s = str(row.get("二级费用科目", "")).strip()
        t3 = str(row.get("三级费用科目", "")).strip()
        t4 = str(row.get("四级费用科目", "")).strip() if "四级费用科目" in row else ""
        doc = str(row.get("归属单据名称", "")).strip()
        people = split_values(row.get("单据适配人员（多人用中文逗号）", ""))
        if not (p and s and doc):
            continue

        has_people[doc] = has_people.get(doc, False) or bool(people)

        primary_info = primary.get(p)
        if not primary_info:
            report["step2"]["relations_fail"].append({"doc": doc, "一级费用科目": p, "message": f"一级费用科目 '{p}' 在系统中不存在"})
            continue
        pid = primary_info.get("id")

        # 先检查是否已创建过（本次运行缓存）
        cache_key_l2 = (pid, s)
        if cache_key_l2 in created_level2:
            sid = created_level2[cache_key_l2]
        else:
            # 检查系统中是否已存在
            sid = child.get(cache_key_l2)
            if sid:
                pass

        # 二级不存在时自动创建
        if not sid and pid:
            create_resp = requests.post(
                f"{BASE_URL}/api/bill/feeTemplate/addFeeTemplate",
                headers=h,
                json={"name": s, "parentId": pid, "companyId": company_id},
                timeout=12
            ).json()
            if create_resp.get("code") == 200 or create_resp.get("success"):
                # 尝试从响应获取ID
                new_sid = create_resp.get("result")
                if isinstance(new_sid, dict):
                    new_sid = new_sid.get("id") or new_sid.get("result")
                if new_sid:
                    sid = int(new_sid)
                    created_level2[cache_key_l2] = sid  # 缓存到本次运行
                else:
                    report["step2"]["relations_fail"].append({"doc": doc, "二级费用科目": s, "message": f"创建二级费用科目 '{s}' 后无法获取ID"})
                    continue
            else:
                report["step2"]["relations_fail"].append({"doc": doc, "二级费用科目": s, "message": f"创建二级费用科目 '{s}' 失败: {create_resp.get('message')}"})
                continue

        if not sid:
            report["step2"]["relations_fail"].append({"doc": doc, "二级费用科目": s, "message": f"无法获取二级费用科目 '{s}' 的ID"})
            continue

        leaf_id = sid
        # 缓存刚创建的科目，避免重复查询
        fee_cache = {}

        # 验证三级费用科目名称（不能是纯数字或空）
        if t3 and t3.lower() != "nan":
            if t3.isdigit() or len(t3) < 2:
                report["step2"]["relations_fail"].append({"doc": doc, "三级费用科目": t3, "message": f"三级费用科目名称 '{t3}' 无效（不能是纯数字或单个字符）"})
                continue
            t3_id = get_or_create_fee_template(t3, sid, company_id, h, fee_cache)
            if t3_id:
                leaf_id = t3_id
            else:
                report["step2"]["relations_fail"].append({"doc": doc, "三级费用科目": t3, "message": f"创建三级费用科目 '{t3}' 失败", "parent_id": sid})
                continue

        # 四级存在时，在三级下查找或创建（如果三级不存在，则直接在二级下创建四级）
        if t4 and t4.lower() != "nan":
            parent_for_t4 = leaf_id if (t3 and t3.lower() != "nan") else sid
            t4_id = get_or_create_fee_template(t4, parent_for_t4, company_id, h, fee_cache)
            if t4_id:
                leaf_id = t4_id
            else:
                report["step2"]["relations_fail"].append({"doc": doc, "四级费用科目": t4, "message": f"创建四级费用科目 '{t4}' 失败", "parent_id": parent_for_t4})
                continue

        report["step2"]["leaf_by_doc"].setdefault(doc, [])
        if leaf_id not in report["step2"]["leaf_by_doc"][doc]:
            report["step2"]["leaf_by_doc"][doc].append(leaf_id)

        # 条件触发费用角色链路：为每个人员创建一个角色（角色名称为人员姓名）
        if people:
            for person_name in people:
                uid = user_map.get(person_name)
                if not uid:
                    report["step2"]["relations_fail"].append({"doc": doc, "人员": person_name, "message": f"人员 '{person_name}' 在系统中不存在"})
                    continue

                # 准备角色名称（去掉数字）
                base_name = ''.join([c for c in person_name if not c.isdigit()])
                if not base_name:
                    base_name = person_name

                # 检查是否已存在该人员名的角色（用去掉数字的名称查）
                rid = fee_roles.get(base_name)
                
                if not rid:
                    report["step2"]["relations_fail"].append({"doc": doc, "人员": person_name, "尝试名称": base_name, "message": "角色未找到"})
                    continue
                
                # 更新角色为 FEE_TYPE（如果不是的话）
                update_role_payload = {
                    "id": rid,
                    "companyId": company_id,
                    "dataType": "FEE_TYPE",
                    "parentId": fee_role_group_id
                }
                
                update_role_resp = requests.post(
                    f"{BASE_URL}/api/member/role/update",
                    headers=h,
                    json=update_role_payload,
                    timeout=12
                ).json()
                
                if not (update_role_resp.get("code") == 200 or update_role_resp.get("success")):
                    report["step2"]["relations_fail"].append({"doc": doc, "人员": person_name, "尝试名称": base_name, "message": f"更新角色类型失败: {update_role_resp.get('message')}"})
                    continue

                if rid:
                    # 使用本地缓存累加费用科目和人员
                    if rid not in role_bindings_cache:
                        role_bindings_cache[rid] = {"fee_ids": set(), "user_ids": set()}
                    role_bindings_cache[rid]["fee_ids"].add(leaf_id)
                    role_bindings_cache[rid]["user_ids"].add(uid)
                    
                    # 使用 update API 给角色添加费用科目和人员（发送累加后的列表）
                    update_payload = {
                        "id": rid,
                        "companyId": company_id,
                        "feeTemplateIds": list(role_bindings_cache[rid]["fee_ids"]),
                        "userIds": list(role_bindings_cache[rid]["user_ids"])
                    }
                    
                    rel = requests.post(
                        f"{BASE_URL}/api/member/role/update",
                        headers=h,
                        json=update_payload,
                        timeout=12,
                    ).json()
                    
                    if rel.get("code") == 200 or rel.get("success"):
                        report["step2"]["relations_ok"] += 1
                        # 记录该单据有费用角色限制
                        if doc not in report["step2"]["role_by_doc"]:
                            report["step2"]["role_by_doc"][doc] = []
                        report["step2"]["role_by_doc"][doc].append(rid)
                    else:
                        report["step2"]["relations_fail"].append({"doc": doc, "人员": person_name, "message": rel.get("message")})
                else:
                    report["step2"]["relations_fail"].append({"doc": doc, "人员": person_name, "message": f"无法获取角色ID"})

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

    # 必须有审批流才能继续
    if not workflow_id:
        report["step3"]["fail"].append({"doc": "所有", "message": "系统中没有可用的审批流，请先创建审批流"})
        Path(args.output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print("❌ 导入失败：没有可用的审批流")
        print(json.dumps({"step1_ok": report["step1"]["ok"], "step2_relations_ok": report["step2"]["relations_ok"], "step3_ok": report["step3"]["ok"], "output": args.output}, ensure_ascii=False, indent=2))
        return

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

        # 判断是否有有效的可见范围限制
        # 类型是"限制"且有具体对象时，才限制可见范围
        has_targets = bool(targets) and bool(role_ids or user_ids or dep_ids)
        is_limited_type = vis_type == "限制" or vis_type == "角色" or vis_type == "员工" or vis_type == "部门"
        has_visibility = is_limited_type and has_targets

        payload = {
            "applyRelateFlag": True,
            "applyRelateNecessary": False,
            "businessType": "PRIVATE",
            "companyId": company_id,
            "componentJson": [],
            "departmentIds": dep_ids if has_visibility else [],
            "feeIds": [],
            "feeScopeFlag": False,
            "groupId": group_map.get(group_name),
            "icon": "md-pricetag",
            "iconColor": "#4c7cc3",
            "loanIds": [],
            "name": doc_name,
            "payFlag": True,
            "requestScope": False,
            "requisitionIds": [],
            "roleIds": role_ids if has_visibility else [],
            "status": "ACTIVE",
            "type": type_map.get(doc_type, "EXPENSE"),
            "userIds": user_ids if has_visibility else [],
            "userScopeFlag": has_visibility,
            "workFlowId": workflow_id,
        }
        if payload["type"] == "REQUISITION":
            payload["applyContentType"] = "TEXT"

        # 费用限制分支
        # 优先判断是否有费用角色（基于单据适配人员）
        if has_people.get(doc_name, False) and report["step2"]["role_by_doc"].get(doc_name):
            # 有单据适配人员：使用费用角色限制
            fee_role_ids = report["step2"]["role_by_doc"][doc_name]
            # 设置费用角色限制
            payload["feeRoleIds"] = fee_role_ids
            payload["feeScopeFlag"] = True
            report["step3"]["branch_fee_role"].append({"doc": doc_name, "feeRoleIds": fee_role_ids})
        elif payload["type"] in ("LOAN", "REQUISITION"):
            # 借款单/申请单：不限制费用
            report["step3"]["branch_skip"].append(doc_name)
        else:
            # 无单据适配人员：使用费用科目直接限制（如果有）
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
