#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
API 数据导出 Excel 工具 - 版本2：多账号
通过账号密码登录各账号，分别拉取数据后合并为一个 Excel 导出。
配置见 config_v2.yaml。
"""

import copy
import os
import sys
from datetime import datetime
from pathlib import Path

import requests

# 复用 run.py 中的逻辑
from run import (
    apply_filters,
    export_excel,
    extract_list,
    fetch_data,
    get_nested,
    load_config,
    normalize_columns,
    row_to_values,
)

# 账号列在合并表中的虚拟字段名
ACCOUNT_FIELD = "_account"


def _replace_placeholders(obj, replacements: dict):
    """递归替换对象中的 {{key}} 为 replacements[key]."""
    if isinstance(obj, str):
        for k, v in replacements.items():
            obj = obj.replace("{{" + k + "}}", str(v))
        return obj
    if isinstance(obj, dict):
        return {key: _replace_placeholders(val, replacements) for key, val in obj.items()}
    if isinstance(obj, list):
        return [_replace_placeholders(item, replacements) for item in obj]
    return obj


def do_login(login_cfg: dict, account: dict) -> requests.Session:
    """
    使用账号密码登录，返回带鉴权的 Session。
    支持 auth_from: cookie（Session 自动带 Cookie）或 body（从响应取 token 设到请求头）。
    """
    url = login_cfg.get("url")
    if not url:
        raise ValueError("login.url 不能为空")
    method = (login_cfg.get("method") or "POST").upper()
    headers = dict(login_cfg.get("headers") or {})
    body = copy.deepcopy(login_cfg.get("body") or {})
    body = _replace_placeholders(body, {
        "username": account.get("username", ""),
        "password": account.get("password", ""),
    })

    session = requests.Session()
    if method == "POST":
        resp = session.post(url, json=body, headers=headers, timeout=30)
    else:
        resp = session.get(url, params=body, headers=headers, timeout=30)

    if not resp.ok:
        print(f"  登录失败 [{account.get('username', '?')}]: {resp.status_code} {resp.text[:200]}")
        resp.raise_for_status()
    else:
        print(f"  登录成功 [{account.get('username', '?')}]: {resp.status_code}")

    # 调试：打印登录响应 body 的顶层 key，便于核对 auth_headers 里的路径
    if login_cfg.get("debug_login_response"):
        try:
            keys = list((resp.json() or {}).keys())
            # print(f"  登录响应 body 顶层 key: {keys}")
        except Exception:
            pass

    # 登录成功后设置请求头：Content-Type、Cookie、x-access-token 等
    # auth_headers 中 key 为请求头名，value 为固定字符串或响应 body 路径（如 data.accessToken）
    resp_data = resp.json() if resp.text else {}
    auth_headers = login_cfg.get("auth_headers") or {}
    for name, value in auth_headers.items():
        if value is None or value == "":
            continue
        value = str(value).strip()
        # 先尝试从响应 body 按路径取值（支持单 key 如 Cookie，或多级如 data.accessToken）；取不到再当作文本
        from_body = get_nested(resp_data, value)
        if from_body is not None:
            session.headers[name] = str(from_body)
        else:
            session.headers[name] = value

    # 若 Cookie 未从 body 取到（如接口用 Set-Cookie 而非 body），用 Session 已存的 Cookie 拼成 Cookie 头，确保数据接口能带上
    if not (session.headers.get("Cookie") or "").strip():
        cookie_parts = []
        for c in session.cookies:
            cookie_parts.append(f"{c.name}={c.value}")
        if cookie_parts:
            session.headers["Cookie"] = "; ".join(cookie_parts)
        # if login_cfg.get("debug_login_response"):
        #     print(f"  已从 Set-Cookie 拼出 Cookie 头（{len(cookie_parts)} 项）")

    # 兼容旧配置：auth_from body 时从 auth_body_path 取 token 设到 auth_header_name
    auth_from = (login_cfg.get("auth_from") or "cookie").lower()
    if auth_from == "body":
        path = login_cfg.get("auth_body_path") or "data.token"
        token = get_nested(resp_data, path)
        if token is not None:
            header_name = login_cfg.get("auth_header_name") or "Authorization"
            header_value = (login_cfg.get("auth_header_value") or "Bearer {{token}}").replace("{{token}}", str(token))
            session.headers[header_name] = header_value

    return session


def main():
    config_path = os.environ.get("CONFIG", "config_v2.yaml")
    config = load_config(config_path)

    login_cfg = config.get("login")
    accounts = config.get("accounts") or []
    if not login_cfg or not accounts:
        print("错误：版本2 需在配置中提供 login 与 accounts（多账号列表）")
        sys.exit(1)

    api_cfg = config.get("api") or {}
    data_path = config.get("data_path") or ""
    columns = config.get("columns") or {}
    if not columns:
        print("错误：请配置 columns")
        sys.exit(1)

    columns_normalized = normalize_columns(columns)
    merge_add_account_column = config.get("merge_add_account_column", True)
    merge_account_header = config.get("merge_account_column_header") or "账号"
    if merge_add_account_column:
        columns_normalized = [(ACCOUNT_FIELD, merge_account_header, None)] + columns_normalized

    filters = config.get("filters") or []
    out_cfg = config.get("output") or {}
    out_dir = out_cfg.get("dir") or "./output"
    out_name = out_cfg.get("filename") or "export_data"

    all_rows = []
    for i, account in enumerate(accounts):
        if not isinstance(account, dict):
            continue
        username = account.get("username", "")
        label = account.get("label") or account.get("name") or username or f"账号{i+1}"
        print(f"正在处理账号: {label} ...")
        try:
            session = do_login(login_cfg, account)
            # print("api_cfg：", api_cfg)
            # print("session：", session)
            
            full_data = fetch_data(api_cfg, session=session)
            print("返回结果提示信息：", full_data.get("message"))
            rows = extract_list(full_data, data_path)
            for row in rows:
                if isinstance(row, dict):
                    row = copy.deepcopy(row)
                    row[ACCOUNT_FIELD] = label
                    all_rows.append(row)
            print(f"  获取 {len(rows)} 条")
        except Exception as e:
            print(f"  跳过账号 {label}: {e}")
            continue

    print(f"合并共 {len(all_rows)} 条")

    if filters:
        before = len(all_rows)
        all_rows = apply_filters(all_rows, filters)
        print(f"过滤后保留 {len(all_rows)} 条（过滤掉 {before - len(all_rows)} 条）")

    if not all_rows:
        print("没有数据可导出")
        sys.exit(0)

    Path(out_dir).mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(out_dir, f"{out_name}_{ts}.xlsx")
    export_excel(all_rows, columns_normalized, out_path)


if __name__ == "__main__":
    main()
