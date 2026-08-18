#!/usr/bin/env python3
"""造数平台 - Excel 导入模板数据生成 & 导入"""

import json
import os
import random
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path

import openpyxl
import requests
from flask import Flask, abort, jsonify, redirect, render_template, request, send_file

app = Flask(__name__)

# ============================================================
# 配置
# ============================================================
BASE_DIR = Path(r"E:\大而全导入数据\国内外包")
CONFIG_FILE = BASE_DIR / "platform" / "config.json"
COOKIE_FILE = BASE_DIR / "platform" / "cookie.txt"    # Cookie 独立文件
ZTOKEN_FILE = BASE_DIR / "platform" / "ztoken.txt"    # ztToken 独立文件（自动登录生成）
TEMPLATE_DIR = BASE_DIR  # 模板文件存放目录

# SSO 登录参数（CAS）
SSO_LOGIN_URL = "https://test-sso.eminxing.com/cas/login"
SSO_SERVICE = "https://uat-hrms.eminxing.com/api/hc/toLogin"
SSO_PUBLIC_KEY = "MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQCxVpeTp5fNeie2+oBzEQijvHtS0xsbyN00OTc3YdtWR8pE47gp+aDXFOb89tIaazwC56IX6QfaqpIlXIZ5iu77kwQK3N26CuCrmq4XMzzqxqu+HCsMdA7wVe5ZxcmgRNSfgKZMVF/b0ZbFddiwRwnN7atgFXMk31bFV1hZzum8OQIDAQAB"

# 默认配置
DEFAULT_CONFIG = {
    "import_api_url": "https://uat-hrms.eminxing.com/api/dw/costActivity/detail/float/import",
    "import_api_method": "POST",
    "import_api_headers": {},
    "import_api_field_name": "file",
    "import_api_extra_fields": {},
    "import_api_cookie": "",
    # HRMS 导入业务参数（float/import 接口要求）
    "import_activity_id": "",      # 活动详情ID = 成本活动的 accrualId（计提期）/ actualId（实发期）
    "import_cost_object": "CO01",  # CO01=员工成本明细, CO02=组织成本明细
    "import_float_type": "6",      # 6=特殊场景, 1=每期浮动科目, 5=线下计算
    "import_include_data": "",     # 仅 floatType=1 时使用 (0/1)
    "pay_password": "",            # 工资条密码（安全验证），留空则用 SSO 密码
    # SSO 自动登录账号
    "sso_username": "",
    "sso_password": "",
}


def load_config():
    """加载配置，Cookie/ztToken 优先从独立文件读取"""
    cfg = DEFAULT_CONFIG.copy()
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg.update(json.load(f))
    # Cookie 独立文件覆盖 config.json 中的 cookie
    if COOKIE_FILE.exists():
        with open(COOKIE_FILE, "r", encoding="utf-8") as f:
            raw = f.read().strip()
            # 去掉注释行（以 # 开头的行），保留有效 cookie
            lines = [l.strip() for l in raw.splitlines() if l.strip() and not l.strip().startswith("#")]
            cookie = " ".join(lines)  # 保持原始格式
            if cookie:
                cfg["import_api_cookie"] = cookie
    # ztToken 独立文件
    if ZTOKEN_FILE.exists():
        with open(ZTOKEN_FILE, "r", encoding="utf-8") as f:
            raw = f.read().strip()
            lines = [l.strip() for l in raw.splitlines() if l.strip() and not l.strip().startswith("#")]
            ztoken = " ".join(lines)
            if ztoken:
                cfg["ztoken"] = ztoken
    return cfg


def save_config(cfg):
    """保存配置，Cookie/ztToken 写入独立文件"""
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    # 分离 cookie 和 ztoken
    cookie = cfg.pop("import_api_cookie", "")
    ztoken = cfg.pop("ztoken", "")
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    # 回写 cookie 到独立文件
    if cookie.strip():
        with open(COOKIE_FILE, "w", encoding="utf-8") as f:
            f.write(cookie.strip())
    elif COOKIE_FILE.exists():
        COOKIE_FILE.unlink()  # 空 cookie 就删文件
    # 回写 ztoken 到独立文件
    if ztoken.strip():
        with open(ZTOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(ztoken.strip())


# ============================================================
# 工资条安全验证（接口返回 701「请输入密码」的门禁）
# ============================================================
HRMS_BASE = "https://uat-hrms.eminxing.com"


def _hrms_headers(config):
    """构造带 ztToken 和 Cookie 的请求头（HRMS 认证核心）"""
    headers = {"User-Agent": "Mozilla/5.0", "Accept-Language": "zh-CN"}
    ztoken = config.get("ztoken", "").strip()
    if ztoken:
        headers["ztToken"] = ztoken
    cookie = config.get("import_api_cookie", "").strip()
    if cookie:
        headers["Cookie"] = cookie
    return headers


def _hrms_post_with_relogin(url, payload, timeout=120):
    """POST HRMS，若返回 401（ztToken 过期）则自动重新登录后重试一次"""
    cfg = load_config()
    headers = _hrms_headers(cfg)
    headers["Content-Type"] = "application/json"
    resp = requests.post(url, headers=headers, json=payload, timeout=timeout, verify=False)
    if resp.status_code == 401:
        zt, ck, err = sso_login()
        if not err:
            cfg = load_config()
            headers = _hrms_headers(cfg)
            headers["Content-Type"] = "application/json"
            resp = requests.post(url, headers=headers, json=payload, timeout=timeout, verify=False)
    return resp


def check_pay_gate(config):
    """查询工资条密码门禁状态。返回 (ok, code, message)"""
    try:
        r = requests.get(f"{HRMS_BASE}/api/dw/password/checkExpire",
                         headers=_hrms_headers(config), timeout=15, verify=False)
        j = r.json()
        code = j.get("code")
        return code in (0, 200), code, j.get("message", "")
    except Exception as e:
        return False, None, str(e)


def unlock_pay_gate(config):
    """解除工资条密码门禁（对应前端「安全验证」弹窗的 slipPassword 流程）。
    返回 (ok, detail)。密码取 pay_password，留空则用 SSO 密码。"""
    ok, code, msg = check_pay_gate(config)
    if ok:
        return True, {"status": "already_unlocked", "code": code}

    password = (config.get("pay_password") or config.get("sso_password") or "").strip()
    if not password:
        return False, {"status": "no_password", "code": code, "message": msg or "需要工资条密码"}

    headers = dict(_hrms_headers(config))
    headers["Content-Type"] = "application/json"

    def _call(path, payload):
        try:
            r = requests.post(f"{HRMS_BASE}{path}", headers=headers, json=payload,
                              timeout=15, verify=False)
            j = r.json()
            return j.get("code"), j.get("message", "")
        except Exception as e:
            return None, str(e)

    # 1) 已有密码 → 直接验证解锁
    vcode, vmsg = _call("/api/dw/password/verify", {"password": password})
    if vcode in (0, 200):
        return True, {"status": "verified", "code": vcode}

    # 2) 未设置过工资条密码 → 自动设置为 SSO 密码（UAT 造数场景）
    scode, smsg = _call("/api/dw/password/set", {"password": password, "confirmPassword": password})
    if scode in (0, 200):
        ok2, code2, msg2 = check_pay_gate(config)
        if ok2:
            return True, {"status": "set_and_unlocked", "code": code2}
        return False, {"status": "set_ok_but_gate_closed", "code": code2, "message": msg2}
    return False, {"status": "unlock_failed", "verify_code": vcode, "set_code": scode,
                   "message": smsg or vmsg}


# ============================================================
# SSO 自动登录（CAS → HRMS 获取 ztToken）
# ============================================================
def sso_login(username=None, password=None):
    """执行 CAS SSO 登录，返回 (ztoken, cookie_str)"""
    import base64
    import re
    import urllib3
    from urllib.parse import urlparse, parse_qs
    from cryptography.hazmat.primitives.asymmetric import padding
    from cryptography.hazmat.primitives import serialization

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    cfg = load_config()
    username = username or cfg.get("sso_username", "").strip()
    password = password or cfg.get("sso_password", "")

    if not username or not password:
        return None, None, "未配置 SSO 账号密码，请先在接口配置中填写"

    # RSA 公钥加载
    pub_b64 = base64.b64encode(base64.b64decode(SSO_PUBLIC_KEY)).decode()
    key_pem = f"-----BEGIN PUBLIC KEY-----\n{pub_b64}\n-----END PUBLIC KEY-----"
    public_key = serialization.load_pem_public_key(key_pem.encode())

    def rsa_encrypt(text):
        encrypted = public_key.encrypt(text.encode("utf-8"), padding.PKCS1v15())
        return base64.b64encode(encrypted).decode()

    s = requests.Session()
    s.verify = False

    # Step 1: GET 登录页，拿 SESSION 和 execution
    login_page_url = f"{SSO_LOGIN_URL}?service={SSO_SERVICE}&locale=zh_CN"
    try:
        resp = s.get(login_page_url, timeout=30, allow_redirects=True)
    except Exception as e:
        return None, None, f"无法连接 CAS 登录页: {e}"

    m = re.search(r'name="execution" value="([^"]+)"', resp.text)
    execution = m.group(1) if m else "e1s1"

    # Step 2: POST 登录（RSA 加密密码）
    post_data = {
        "username": username,
        "password": rsa_encrypt(password),
        "execution": execution,
        "_eventId": "submit",
        "geolocation": "",
    }
    try:
        resp2 = s.post(SSO_LOGIN_URL, data=post_data, timeout=30, allow_redirects=False)
    except Exception as e:
        return None, None, f"CAS 登录请求失败: {e}"

    if resp2.status_code not in (301, 302, 303):
        # 登录失败，检查错误信息
        err_m = re.search(r'id="errorMessage"[^>]*>([^<]+)<', resp2.text)
        msg = err_m.group(1).strip() if err_m else f"登录失败 (HTTP {resp2.status_code})"
        return None, None, msg

    # Step 3: 跟随重定向获取 ticket
    loc = resp2.headers.get("Location", "")
    if loc.startswith("/"):
        loc = "https://test-sso.eminxing.com" + loc
    try:
        page = s.get(loc, timeout=30, allow_redirects=True)
    except Exception as e:
        return None, None, f"获取 ticket 失败: {e}"

    # Step 4: 从 URL 提取 ticket (JWT = ztToken)
    qs = parse_qs(urlparse(page.url).query)
    ztoken = qs.get("ticket", [""])[0]
    if not ztoken:
        return None, None, "未在重定向 URL 中找到 ticket"

    # Step 5: 汇总 Cookie（HRMS 域下）
    cookies = []
    for c in s.cookies:
        if "eminxing" in c.domain:
            cookies.append(f"{c.name}={c.value}")
    cookie_str = "; ".join(cookies)

    # 保存到文件
    with open(ZTOKEN_FILE, "w", encoding="utf-8") as f:
        f.write(ztoken)
    if cookie_str:
        with open(COOKIE_FILE, "w", encoding="utf-8") as f:
            f.write(cookie_str)

    return ztoken, cookie_str, None


# ============================================================
# 模板解析
# ============================================================
def template_display_meta(name):
    """从规范模板文件名解析展示信息与成本对象分组。
    格式: {方案}_{计提|实发}_{YYYYMM}_{浮动类型}_{成本对象}.xlsx
    返回 (co_group, display, month, period, ft)"""
    co_group = ""
    for label in ("员工成本明细", "组织成本"):
        if label in name:
            co_group = label
            break
    if not co_group:
        co_group = "未区分"

    period = "计提" if "_计提_" in name else ("实发" if "_实发_" in name else "")
    month = ""
    for p in name.replace(".xlsx", "").split("_"):
        if len(p) == 6 and p.isdigit() and p.startswith("20"):
            month = p
            break
    ft = ""
    for label in ("每期浮动", "线下计算", "不计成本", "特殊场景"):
        if label in name:
            ft = label
            break

    bits = []
    if period:
        bits.append(period + "期")
    if month:
        bits.append(month[:4] + "/" + month[4:])
    if ft:
        bits.append(ft)
    display = " · ".join(bits) if bits else name
    return co_group, display, month, period, ft


def list_templates():
    """列出目录下的所有 xlsx 模板"""
    templates = []
    for f in TEMPLATE_DIR.glob("*.xlsx"):
        if f.name.startswith("~$"):
            continue
        try:
            wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
            sheets = wb.sheetnames
            has_hidden = "hidden_head" in sheets
            data_sheet = next((s for s in sheets if s != "hidden_head"), sheets[0])
            # 读取列名
            ws = wb[data_sheet]
            headers = [str(ws.cell(row=1, column=c).value or "") for c in range(1, ws.max_column + 1)]
            wb.close()
            co_group, display, month, period, ft = template_display_meta(f.name)
            templates.append({
                "name": f.name,
                "path": str(f),
                "sheets": sheets,
                "has_hidden": has_hidden,
                "data_sheet": data_sheet,
                "columns": len(headers),
                "headers": headers,
                "co_group": co_group,
                "display": display,
                "month": month,
                "period": period,
            })
        except Exception as e:
            templates.append({"name": f.name, "path": str(f), "error": str(e),
                              "co_group": "未区分", "display": f.name})
    return templates


def parse_template(filepath):
    """解析模板，返回字段元数据"""
    wb = openpyxl.load_workbook(filepath)

    # 找数据 sheet
    data_sheet_name = None
    for sn in wb.sheetnames:
        if sn != "hidden_head":
            data_sheet_name = sn
            break

    ws = wb[data_sheet_name]
    ncols = ws.max_column

    # 解析 hidden_head
    col_metas = []
    enum_sources = {}

    if "hidden_head" in wb.sheetnames:
        hd = wb["hidden_head"]
        en_names = [c.value for c in hd[1]]
        cn_names = [c.value for c in hd[2]]
        types = [c.value for c in hd[3]]
        requireds = [c.value for c in hd[4]]
        dict_names = [c.value for c in hd[5]]
        formats = [c.value for c in hd[6]]
        precs = [c.value for c in hd[7]]

        for i in range(ncols):
            cn = str(cn_names[i] or "") if cn_names[i] is not None else ""
            col_type = str(types[i]) if types[i] is not None else "0"

            meta = {
                "index": i,
                "col_letter": _col_letter(i + 1),
                "en_name": str(en_names[i] or ""),
                "cn_name": cn,
                "type": col_type,
                "type_label": {"0": "自由输入", "1": "必填文本", "2": "选填文本", "3": "枚举下拉"}.get(col_type, "未知"),
                "required": str(requireds[i]) == "1" if requireds[i] is not None else False,
                "dict_name": str(dict_names[i] or ""),
                "format": str(formats[i] or ""),
                "precision": precs[i],
                "enum_values": [],
            }

            # 枚举字段 → 找对应 sheet（优先列名精确匹配，其次字典名，最后互为子串兜底）
            if col_type == "3":
                sn = _find_enum_sheet(wb, cn, meta["dict_name"], data_sheet_name)
                if sn:
                    vals = _read_enum(wb[sn])
                    meta["enum_values"] = vals
                    enum_sources[i + 1] = sn

            col_metas.append(meta)

    # 枚举 pools for generation
    enum_pools = {}
    for col_idx, sn in enum_sources.items():
        vals = _read_enum(wb[sn])
        if vals:
            enum_pools[col_idx] = vals

    wb.close()

    return {
        "data_sheet": data_sheet_name,
        "ncols": ncols,
        "columns": col_metas,
        "enum_pools": enum_pools,
    }


def _read_enum(ws, col=0, filter_test=True):
    vals = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        v = row[col]
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        # 过滤枚举 sheet 里的测试垃圾数据（如 测试徐、test、不要再测试了-001）
        if filter_test and ("测试" in s or s.lower() == "test" or s.lower().startswith("test")):
            continue
        vals.append(s)
    return vals


def _find_enum_sheet(wb, cn, dict_name, data_sheet_name):
    """按列名/字典名在模板 sheet 中定位枚举取值 sheet，找不到返回 None。
    匹配顺序：列名精确 → 字典名精确 → 列名与 sheet 名互为子串（如 工资发放标准 ↔ 工资发放）"""
    clean = (cn or "").strip().lstrip("*")
    if not clean:
        return None
    others = [sn for sn in wb.sheetnames if sn not in ("hidden_head", data_sheet_name)]
    if clean in others:
        return clean
    if dict_name and dict_name in others:
        return dict_name
    for sn in others:
        if len(sn) >= 2 and (sn in clean or clean in sn):
            return sn
    return None


def _col_letter(n):
    """1→A, 27→AA"""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ============================================================
# 数据生成引擎
# ============================================================
def generate_data(filepath, count, seed=None, fixed_values=None, unique_cols=None, month=None):
    """生成数据行，返回 (headers, rows)。
    month: 选中的成本活动月份（如 2026/08），用于自动固定 计薪所属月份 / 周期开始 / 周期结束 日期"""
    if seed is None:
        seed = int.from_bytes(os.urandom(4), "big")
    random.seed(seed)

    fixed_values = {int(k): v for k, v in (fixed_values or {}).items()}
    unique_cols = unique_cols or []  # list of column indices (0-based)

    # 活动月份 → 周期起止日期（首日/末日）
    month_start = month_end = None
    if month:
        try:
            import calendar
            y, m = int(str(month).split("/")[0]), int(str(month).split("/")[1])
            month_start = f"{y}/{m:02d}/01"
            month_end = f"{y}/{m:02d}/{calendar.monthrange(y, m)[1]:02d}"
        except Exception:
            pass

    # Read additional data pools
    wb = openpyxl.load_workbook(filepath)

    # 集团组织 pool（组织名称/组织ID/部门ID 均取真实组织，同一行保持名称+ID 对应）
    org_rows = []
    if "集团组织" in wb.sheetnames:
        org_ws = wb["集团组织"]
        for r in range(2, org_ws.max_row + 1):
            name = org_ws.cell(row=r, column=1).value
            oid = org_ws.cell(row=r, column=3).value
            if oid is not None and str(oid).strip():
                org_rows.append((str(name or "").strip(), str(oid).strip()))
    org_ids = [oid for _, oid in org_rows]

    # Parse template metadata
    data_sheet_name = None
    for sn in wb.sheetnames:
        if sn != "hidden_head":
            data_sheet_name = sn
            break

    ws = wb[data_sheet_name]
    ncols = ws.max_column
    headers = [str(ws.cell(row=1, column=c).value or "") for c in range(1, ncols + 1)]

    # Column metadata
    col_metas = []
    if "hidden_head" in wb.sheetnames:
        hd = wb["hidden_head"]
        cn_names = [c.value for c in hd[2]]
        types = [c.value for c in hd[3]]
        dict_names = [c.value for c in hd[5]]
        formats = [c.value for c in hd[6]]
        precs = [c.value for c in hd[7]]
        for i in range(ncols):
            col_metas.append({
                "cn_name": str(cn_names[i] or ""),
                "type": str(types[i]) if types[i] is not None else "0",
                "dict_name": str(dict_names[i] or ""),
                "format": str(formats[i] or ""),
                "precision": precs[i],
            })

    # Enum pools（列名/字典名匹配模板内枚举 sheet）
    enum_pools = {}
    for i in range(ncols):
        meta = col_metas[i]
        clean = meta["cn_name"].strip().lstrip("*")
        if not clean:
            continue
        sn = None
        if meta["type"] == "3":
            # 枚举下拉列：列名精确 → 字典名 → 互为子串兜底
            sn = _find_enum_sheet(wb, meta["cn_name"], meta["dict_name"], data_sheet_name)
        elif clean in wb.sheetnames and clean not in (data_sheet_name, "hidden_head"):
            # 非枚举列但模板存在同名 sheet（如 划分区域）→ 同样从该 sheet 取真实值
            sn = clean
        if sn:
            vals = _read_enum(wb[sn])
            if vals:
                enum_pools[i + 1] = vals

    wb.close()

    # Data generation helpers
    remarks_pool = [
        "正常发放，无异常情况", "已按流程审批通过", "待确认后调整",
        "本月新增项目补贴", "年度调整项，已备案", "特殊情况，已报备HRBP",
        "季度绩效奖金核算", "跨月调整差额", "系统自动计算，无需人工干预",
        "含加班补贴及交通补助", "个税专项附加扣除已更新", "社保基数年度调整",
        "薪资结构变更已审批", "转正调薪差额补发", "离职结算最终清算",
    ]
    countries = ["中国", "美国", "日本", "新加坡", "英国", "德国", "韩国", "澳大利亚", "巴西", "印度"]
    zones = ["华南仓", "华东仓", "华北仓", "西南仓", "华中仓", "海外仓"]
    surnames = ["张", "李", "王", "赵", "刘", "陈", "杨", "黄", "周", "吴"]
    given_names = ["明", "华", "强", "丽", "伟", "芳", "敏", "勇", "静", "涛", "磊", "洋"]

    # 标准工时 → 人数 区间的列索引（用于未命中规则的数值列兜底随机数值）
    std_idx = next((i for i in range(ncols) if "标准工时" in col_metas[i]["cn_name"]), None)
    ppl_idx = next((i for i in range(ncols) if "人数" in col_metas[i]["cn_name"]), ncols)

    used_combos = set()
    rows = []

    for _ in range(count):
        # Per-row context — tracks computed values that downstream columns depend on
        ctx = {}

        # ---- Phase 1: generate values for all columns ----
        row_data = [None] * ncols

        for c in range(ncols):
            cn = col_metas[c]["cn_name"]
            fmt = col_metas[c]["format"]
            prec = col_metas[c]["precision"]

            # Fixed values take priority
            if c in fixed_values:
                row_data[c] = fixed_values[c]
                continue

            # === Date columns ===
            if "计薪所属月份" in cn:
                if month:
                    row_data[c] = month
                else:
                    y = random.randint(2024, 2026)
                    m = random.randint(1, 12)
                    row_data[c] = f"{y}/{m:02d}"
            elif "周期开始日期" in cn:
                row_data[c] = month_start if month_start else f"{random.randint(2024,2026)}/{random.randint(1,12):02d}/{random.randint(1,28):02d}"
            elif "周期结束日期" in cn:
                row_data[c] = month_end if month_end else f"{random.randint(2024,2026)}/{random.randint(1,12):02d}/{random.randint(1,28):02d}"
            elif "入职日期" in cn:
                ey = random.randint(2018, 2025)
                row_data[c] = f"{ey}/{random.randint(1,12):02d}/{random.randint(1,28):02d}"
            elif "离职日期" in cn:
                if random.random() < 0.3:
                    row_data[c] = f"{random.randint(2020,2026)}/{random.randint(1,12):02d}/{random.randint(1,28):02d}"

            # === Identity columns ===
            elif "姓名" == cn:
                row_data[c] = random.choice(surnames) + random.choice(given_names) + random.choice(given_names)
            elif "工号" in cn:
                prefix = random.choice(["EMP", "YTO", "SZX", "BJ", "SH"])
                row_data[c] = f"{prefix}{random.randint(10000, 99999)}"
            elif "组织名称" in cn:
                if "org" not in ctx:
                    ctx["org"] = random.choice(org_rows) if org_rows else None
                row_data[c] = ctx["org"][0] if ctx["org"] else ""
            elif "组织ID" in cn or "部门ID" in cn:
                if "org" not in ctx:
                    ctx["org"] = random.choice(org_rows) if org_rows else None
                if ctx["org"]:
                    row_data[c] = ctx["org"][1]
                else:
                    row_data[c] = random.choice(org_ids) if org_ids else str(random.randint(100000, 999999))

            # === Location / classification ===
            elif "考勤地点" in cn or "发薪国家" in cn:
                row_data[c] = random.choice(countries)
            elif "成本分区" in cn:
                row_data[c] = random.choice(zones)
            elif "挂账分类" in cn:
                row_data[c] = random.choice(remarks_pool)
            elif "备注" in cn or "说明" in cn:
                row_data[c] = random.choice(remarks_pool)

            # === Numeric: rate & hours ===
            elif "汇率" in cn:
                val = round(random.uniform(6.5, 7.5), 4)
                row_data[c] = val
                ctx["rate"] = val
            elif "标准工时" in cn:
                val = round(random.uniform(160, 200), 2)
                row_data[c] = val
                ctx["std_hours"] = val
            elif "实际工时" in cn:
                base = ctx.get("std_hours", 170)
                row_data[c] = round(base + random.uniform(-20, 30), 2)

            # === Salary chain ===
            elif "应发工资_原币" in cn:
                val = round(random.uniform(5000, 50000), 2)
                row_data[c] = val
                ctx["gross_orig"] = val
            elif "应发工资_人民币" in cn:
                r = ctx.get("rate", 7.0)
                gross = ctx.get("gross_orig", 30000)
                row_data[c] = round(gross * r, 2)
            elif "实发工资_原币" in cn:
                gross = ctx.get("gross_orig", 30000)
                deductions = round(random.uniform(500, 5000), 2)
                val = round(gross - deductions, 2)
                row_data[c] = val
                ctx["net_orig"] = val
            elif "实发工资_人民币" in cn:
                r = ctx.get("rate", 7.0)
                net = ctx.get("net_orig", 25000)
                row_data[c] = round(net * r, 2)

            # === Labor cost chain ===
            elif "月度人力成本（不含年终奖）" in cn:
                gross = ctx.get("gross_orig", 30000)
                val = round(gross * random.uniform(1.2, 1.5), 2)
                row_data[c] = val
                ctx["cost_no_bonus"] = val
            elif "月度人力成本（含年终奖）" in cn:
                base = ctx.get("cost_no_bonus", 36000)
                row_data[c] = round(base * random.uniform(1.05, 1.15), 2)

            # === 人数 ===
            elif "人数" in cn:
                row_data[c] = random.randint(1, 50)

            # === 枚举列（从模板内对应 sheet 取真实值，必须早于数值兜底） ===
            elif c + 1 in enum_pools:
                row_data[c] = random.choice(enum_pools[c + 1])

            # === 标准工时→人数 区间内未命中规则的数值列：随机数值兜底 ===
            elif std_idx is not None and std_idx <= c < ppl_idx:
                if "工时" in cn or "天数" in cn or "天" in cn:
                    row_data[c] = round(random.uniform(0, 200), 2)
                else:
                    row_data[c] = round(random.uniform(100, 100000), 2)

        rows.append(row_data)

        # Unique constraint check
        if unique_cols:
            combo = tuple(row_data[i] for i in unique_cols if i < len(row_data))
            if combo in used_combos:
                pass  # Simple retry not implemented — rare with large pools
            used_combos.add(combo)

    return headers, rows


def write_excel(template_path, output_path, rows):
    """将数据行写入 Excel"""
    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)

    # 找数据 sheet
    data_sheet_name = None
    for sn in wb.sheetnames:
        if sn != "hidden_head":
            data_sheet_name = sn
            break

    ws = wb[data_sheet_name]

    # 清空旧数据
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # 写入新数据
    for i, row in enumerate(rows):
        for c, val in enumerate(row, 1):
            if val is not None:
                ws.cell(row=i + 2, column=c, value=val)

    wb.save(output_path)
    wb.close()


# ============================================================
# 导入 API 调用
# ============================================================
def call_import_api(filepath, config):
    """调用外部导入接口（含工资条密码门禁解锁 + 业务参数）"""
    url = config.get("import_api_url", "").strip()
    if not url:
        return {"success": False, "error": "未配置导入接口 URL"}

    method = config.get("import_api_method", "POST").upper()
    headers = dict(config.get("import_api_headers", {}))
    field_name = config.get("import_api_field_name", "file")
    cookie = config.get("import_api_cookie", "").strip()

    # 将 Cookie / ztToken 注入请求头（HRMS 认证核心）
    if cookie and "Cookie" not in headers:
        headers["Cookie"] = cookie
    ztoken = config.get("ztoken", "").strip()
    if ztoken and "ztToken" not in headers:
        headers["ztToken"] = ztoken

    # 业务表单参数：活动详情ID + 成本对象 + 浮动类型（+ 额外字段）
    data = dict(config.get("import_api_extra_fields", {}))
    if config.get("import_activity_id", "").strip() and "activityDetailId" not in data:
        data["activityDetailId"] = str(config.get("import_activity_id", "")).strip()
    if config.get("import_cost_object", "").strip() and "costObject" not in data:
        data["costObject"] = config.get("import_cost_object", "").strip()
    if config.get("import_float_type", "").strip() and "floatType" not in data:
        data["floatType"] = str(config.get("import_float_type", "")).strip()
    if "includeData" not in data:
        data["includeData"] = str(config.get("import_include_data", ""))
    data = data or None

    result = {"url": url}

    # 0) 工资条密码门禁（接口会返回 701「请输入密码」）
    if "password/" not in url and ztoken:
        unlocked, gate_detail = unlock_pay_gate(config)
        result["pay_gate"] = gate_detail
        if not unlocked and gate_detail.get("status") == "no_password":
            result["success"] = False
            result["error"] = "工资条密码门禁未解除，且未配置密码（pay_password / SSO 密码）"
            return result

    try:
        with open(filepath, "rb") as f:
            files = {field_name: (os.path.basename(filepath), f,
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            if method == "POST":
                resp = requests.post(url, files=files, data=data, headers=headers, timeout=120)
            elif method == "PUT":
                resp = requests.put(url, files=files, data=data, headers=headers, timeout=120)
            else:
                return {"success": False, "error": f"不支持的 HTTP 方法: {method}"}

        # 解析 JSON 响应（HRMS 统一返回 {code, message, data}）
        try:
            resp_json = resp.json()
        except Exception:
            resp_json = None

        result["status_code"] = resp.status_code
        result["is_auth_error"] = resp.status_code in (401, 403)
        result["response"] = (resp_json if resp_json else resp.text[:2000])

        if isinstance(resp_json, dict):
            biz_code = resp_json.get("code")
            result["biz_code"] = biz_code
            result["biz_message"] = resp_json.get("message", "")
            result["biz_data"] = resp_json.get("data")
            # 业务码 0/200 = 成功；701 = 密码门禁；其余为业务失败
            if str(biz_code) in ("0", "200"):
                result["success"] = True
            elif biz_code == 701:
                result["success"] = False
                result["error"] = "需要工资条密码验证（701）— 请检查 pay_password 配置"
            else:
                result["success"] = False
                result["error"] = resp_json.get("message", "") or f"业务失败 (code={biz_code})"
        else:
            result["success"] = resp.ok
            if not resp.ok:
                result["error"] = f"HTTP {resp.status_code}"
        return result
    except requests.exceptions.ConnectionError:
        return {"success": False, "error": f"无法连接到 {url}"}
    except requests.exceptions.Timeout:
        return {"success": False, "error": "请求超时"}
    except Exception as e:
        return {"success": False, "error": str(e)}


@app.route("/api/test-connection", methods=["POST"])
def api_test_connection():
    """测试导入接口连通性（发送最小 Excel 文件，验证认证）"""
    config = load_config()
    url = config.get("import_api_url", "").strip()
    if not url:
        return jsonify({"success": False, "error": "未配置导入接口 URL"})

    headers = dict(config.get("import_api_headers", {}))
    cookie = config.get("import_api_cookie", "").strip()
    if cookie and "Cookie" not in headers:
        headers["Cookie"] = cookie
    ztoken = config.get("ztoken", "").strip()
    if ztoken and "ztToken" not in headers:
        headers["ztToken"] = ztoken

    field_name = config.get("import_api_field_name", "file")

    # 生成一个最小的合法 xlsx 文件（空工作簿）用于测试
    import io
    from openpyxl import Workbook as _WB
    tmp_wb = _WB()
    tmp_wb.active.title = "Sheet1"
    buf = io.BytesIO()
    tmp_wb.save(buf)
    buf.seek(0)
    dummy_bytes = buf.read()
    buf.close()

    try:
        # 0) 工资条密码门禁检查/解锁
        gate_ok, gate_detail = unlock_pay_gate(config)

        # 1) 业务表单参数（与真实导入一致）
        data = dict(config.get("import_api_extra_fields", {}))
        if config.get("import_activity_id", "").strip() and "activityDetailId" not in data:
            data["activityDetailId"] = str(config.get("import_activity_id", "")).strip()
        if config.get("import_cost_object", "").strip() and "costObject" not in data:
            data["costObject"] = config.get("import_cost_object", "").strip()
        if config.get("import_float_type", "").strip() and "floatType" not in data:
            data["floatType"] = str(config.get("import_float_type", "")).strip()
        if "includeData" not in data:
            data["includeData"] = str(config.get("import_include_data", ""))

        files = {field_name: ("_test_connection.xlsx", dummy_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        method = config.get("import_api_method", "POST").upper()

        if method == "POST":
            resp = requests.post(url, files=files, data=data, headers=headers, timeout=30, allow_redirects=False)
        else:
            resp = requests.put(url, files=files, data=data, headers=headers, timeout=30, allow_redirects=False)

        # 尝试解析响应
        try:
            resp_json = resp.json()
            resp_text = json.dumps(resp_json, ensure_ascii=False)
        except:
            resp_json = None
            resp_text = resp.text[:500]

        return jsonify({
            "success": True,
            "status_code": resp.status_code,
            "is_auth_error": resp.status_code in (401, 403),
            "is_redirect_login": resp.status_code in (301, 302, 303) and "login" in (resp.headers.get("Location", "")).lower(),
            "method_used": method,
            "biz_code": resp_json.get("code") if isinstance(resp_json, dict) else None,
            "biz_message": resp_json.get("message", "") if isinstance(resp_json, dict) else "",
            "response_preview": resp_text[:500],
            "url": url,
            "pay_gate": gate_detail,
        })
    except requests.exceptions.ConnectionError:
        return jsonify({"success": False, "error": f"无法连接到 {url}"})
    except requests.exceptions.Timeout:
        return jsonify({"success": False, "error": "连接超时"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/auto-login", methods=["POST"])
def api_auto_login():
    """SSO 自动登录，实时获取 ztToken"""
    data = request.json or {}
    username = data.get("username", "").strip()
    password = data.get("password", "")

    ztoken, cookie_str, err = sso_login(username, password)
    if err:
        return jsonify({"success": False, "error": err})

    # 登录后顺手解除工资条密码门禁（导入接口要求）
    cfg = load_config()
    gate_ok, gate_detail = unlock_pay_gate(cfg)

    return jsonify({
        "success": True,
        "ztoken": ztoken,
        "ztoken_preview": ztoken[:40] + "..." if len(ztoken) > 40 else ztoken,
        "cookie_len": len(cookie_str) if cookie_str else 0,
        "saved_to": str(ZTOKEN_FILE),
        "pay_gate": gate_detail,
    })


@app.route("/api/auth-status", methods=["GET"])
def api_auth_status():
    """查看当前认证状态（工资条门禁未解锁时自动解锁，无需用户操作）"""
    cfg = load_config()
    ztoken = cfg.get("ztoken", "")
    has_cookie = bool(cfg.get("import_api_cookie", "").strip())

    # 工资条密码门禁状态：先查，未解锁则自动解锁；token 过期则自动重新登录
    pay_gate = {"checked": False}
    if ztoken:
        ok, code, msg = check_pay_gate(cfg)
        if ok:
            pay_gate = {"checked": True, "ok": True, "code": code,
                        "message": msg or "", "status": "already_unlocked"}
        else:
            unlocked, detail = unlock_pay_gate(cfg)
            if not unlocked:
                # ztToken 可能过期 → 自动重新登录后再解锁一次
                zt, ck, err = sso_login()
                if not err:
                    cfg = load_config()
                    unlocked, detail = unlock_pay_gate(cfg)
            pay_gate = {"checked": True, "ok": unlocked, "code": detail.get("code"),
                        "message": detail.get("message", ""), "status": detail.get("status")}

    return jsonify({
        "has_ztoken": bool(ztoken),
        "ztoken_preview": (ztoken[:40] + "...") if ztoken else "",
        "has_cookie": has_cookie,
        "pay_gate": pay_gate,
    })


@app.route("/api/activities", methods=["GET"])
def api_activities():
    """列出 HRMS 成本活动（用于选择活动详情ID）。
    可选 ?month=2026/08 只返回指定月份的活动"""
    cfg = load_config()
    ztoken = cfg.get("ztoken", "").strip()
    if not ztoken:
        return jsonify({"success": False, "error": "未登录，请先一键登录"}), 401

    month_filter = request.args.get("month", "").strip()

    try:
        resp = _hrms_post_with_relogin(f"{HRMS_BASE}/api/dw/costActivity/pageList",
                                       {"pageNum": 1, "pageSize": 60}, timeout=30)
        j = resp.json()
        if str(j.get("code")) not in ("0", "200"):
            return jsonify({"success": False, "error": j.get("message") or f"业务失败 code={j.get('code')}"}), 500
        records = j.get("data", {}).get("records", []) or []
        if month_filter:
            records = [r for r in records if str(r.get("costMonth") or "") == month_filter]
        items = [{
            "id": r.get("id"),
            "name": r.get("programName") or "",
            "month": r.get("costMonth") or "",
            "applyEmploymentType": r.get("applyEmploymentType") or "",
            "programCode": r.get("programCode") or "",
            "accrualId": r.get("accrualId"),
            "actualId": r.get("actualId"),
            "accrualStatus": r.get("accrualStatusName") or "",
            "actualStatus": r.get("actualStatusName") or "",
        } for r in records]
        return jsonify({"success": True, "items": items})
    except requests.exceptions.ConnectionError:
        return jsonify({"success": False, "error": "无法连接 HRMS"}), 502
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


# ============================================================
# Web 路由
# ============================================================
@app.route("/")
def index():
    config = load_config()
    templates = list_templates()
    # 按成本对象分组，供首页「选择模板」下拉框精确区分
    template_groups = []
    for g in ("员工成本明细", "组织成本", "未区分"):
        items = [t for t in templates if t.get("co_group") == g]
        if items:
            template_groups.append({"label": g, "items": items})
    default_template = ""
    if templates:
        default_template = templates[0]["name"]
    return render_template("index.html", templates=templates, config=config,
                           default_template=default_template, template_groups=template_groups)


@app.route("/<path:stray>")
def catch_all(stray):
    """兜底：地址栏误输入的杂散路径（如 /））重定向回首页；API 路径保持 404。"""
    if stray.startswith("api"):
        abort(404)
    return redirect("/")


@app.route("/api/templates")
def api_templates():
    return jsonify(list_templates())


@app.route("/api/template/<name>/columns")
def api_template_columns(name):
    filepath = TEMPLATE_DIR / name
    if not filepath.exists():
        return jsonify({"error": "模板不存在"}), 404
    meta = parse_template(str(filepath))
    return jsonify(meta["columns"])


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.json
    template_name = data.get("template", "")
    count = int(data.get("count", 10))
    seed = data.get("seed") or None
    fixed_values = data.get("fixed_values", {})
    unique_cols = data.get("unique_cols", [])
    month = data.get("month") or ""

    filepath = TEMPLATE_DIR / template_name
    if not filepath.exists():
        return jsonify({"error": "模板不存在"}), 404

    try:
        headers, rows = generate_data(str(filepath), count, seed, fixed_values, unique_cols, month)
        return jsonify({
            "headers": headers,
            "rows": rows,
            "count": len(rows),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/download", methods=["POST"])
def api_download():
    """生成 Excel 并返回下载"""
    data = request.json
    template_name = data.get("template", "")
    count = int(data.get("count", 10))
    seed = data.get("seed") or None
    fixed_values = data.get("fixed_values", {})
    unique_cols = data.get("unique_cols", [])
    month = data.get("month") or ""

    filepath = TEMPLATE_DIR / template_name
    if not filepath.exists():
        return jsonify({"error": "模板不存在"}), 404

    # 生成到临时文件
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_path = tmp.name
    tmp.close()

    try:
        _, rows = generate_data(str(filepath), count, seed, fixed_values, unique_cols, month)
        write_excel(str(filepath), tmp_path, rows)

        download_name = f"生成数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            tmp_path,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        # 延迟清理
        pass


@app.route("/api/import", methods=["POST"])
def api_import():
    """生成数据并调用导入接口"""
    template_name = request.form.get("template", "")
    count = int(request.form.get("count", 10))
    seed = request.form.get("seed") or None
    month = request.form.get("month") or ""

    filepath = TEMPLATE_DIR / template_name
    if not filepath.exists():
        return jsonify({"error": "模板不存在"}), 404

    config = load_config()

    # 生成临时文件
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_path = tmp.name
    tmp.close()

    try:
        # 解析 fixed_values (form 中以 fv_ 开头的字段)
        fixed_values = {}
        for key in request.form:
            if key.startswith("fv_"):
                col_idx = int(key[3:])
                fixed_values[col_idx] = request.form[key]

        _, rows = generate_data(str(filepath), count, seed, fixed_values, month=month)
        write_excel(str(filepath), tmp_path, rows)

        # 尚无 ztToken → 先自动登录
        if not config.get("ztoken", "").strip():
            ztoken, cookie_str, err = sso_login()
            if err:
                result = {"success": False, "error": f"未登录且自动登录失败: {err}"}
                result["rows_generated"] = len(rows)
                return jsonify(result), 401
            config = load_config()

        # 调用导入接口
        result = call_import_api(tmp_path, config)
        result["rows_generated"] = len(rows)

        # 401 时自动登录重试一次
        if result.get("status_code") == 401 or (result.get("response") and isinstance(result["response"], dict) and result["response"].get("code") == 401):
            ztoken, cookie_str, err = sso_login()
            if err:
                result["auto_login"] = f"认证失败且自动登录失败: {err}"
            else:
                config2 = load_config()  # sso_login 已保存 ztoken/cookie
                result = call_import_api(tmp_path, config2)
                result["rows_generated"] = len(rows)
                result["auto_login"] = "401 后已自动重新登录并重试"

        app.logger.info(
            "[IMPORT] template=%s activity=%s costObject=%s floatType=%s rows=%s success=%s biz_code=%s biz_message=%s error=%s",
            template_name, config.get("import_activity_id"), config.get("import_cost_object"),
            config.get("import_float_type"), result.get("rows_generated"), result.get("success"),
            result.get("biz_code"), result.get("biz_message"), result.get("error", ""),
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        try:
            os.unlink(tmp_path)
        except:
            pass


@app.route("/api/download-template", methods=["POST"])
def api_download_template():
    """从 HRMS 下载指定活动+成本对象的导入模板，保存到模板目录。
    文件名带用工类型+计提/实发+月份+成本对象（员工成本明细/组织成本），便于首页区分。
    body: {"activityDetailId": "283", "costObject": "CO01"}，缺省用配置值"""
    data = request.get_json(silent=True) or {}
    cfg = load_config()
    ztoken = cfg.get("ztoken", "").strip()
    if not ztoken:
        ztoken, cookie_str, err = sso_login()
        if err:
            return jsonify({"success": False, "error": f"未登录且自动登录失败: {err}"}), 401
        cfg = load_config()

    activity_id = str(data.get("activityDetailId") or cfg.get("import_activity_id") or "").strip()
    if not activity_id:
        return jsonify({"success": False, "error": "未配置活动详情ID（计提期/实发期）"}), 400
    cost_object = str(data.get("costObject") or cfg.get("import_cost_object") or "CO01").strip()
    float_type = str(data.get("floatType") or cfg.get("import_float_type", "6") or "6")

    # 确保工资条验证已解锁（否则 HRMS 返回「请输入密码」）
    try:
        ok, code, msg = check_pay_gate(cfg)
        if not ok:
            unlocked, detail = unlock_pay_gate(cfg)
            if not unlocked:
                zt, ck, err = sso_login()
                if not err:
                    cfg = load_config()
                    unlock_pay_gate(cfg)
    except Exception:
        pass

    body = {
        "activityDetailId": activity_id,
        "costObject": cost_object,
        "floatType": float_type,
        "includeData": str(cfg.get("import_include_data", "")),
    }

    # 查活动信息（用工类型名/月份/计提或实发），用于模板文件名
    prog_name, month_str, period = "", "", "计提"
    try:
        r = _hrms_post_with_relogin(f"{HRMS_BASE}/api/dw/costActivity/pageList",
                                    {"pageNum": 1, "pageSize": 200}, timeout=30)
        recs = (r.json().get("data") or {}).get("records") or []
        for x in recs:
            if str(x.get("accrualId")) == activity_id:
                prog_name = x.get("programName") or ""
                month_str = str(x.get("costMonth") or "").replace("/", "")
                period = "计提"
                break
            if str(x.get("actualId")) == activity_id:
                prog_name = x.get("programName") or ""
                month_str = str(x.get("costMonth") or "").replace("/", "")
                period = "实发"
                break
    except Exception:
        pass

    try:
        resp = _hrms_post_with_relogin(f"{HRMS_BASE}/api/dw/costActivity/detail/float/downloadTemplate",
                                       body, timeout=120)
    except Exception as e:
        return jsonify({"success": False, "error": f"下载失败: {e}"}), 502

    ct = resp.headers.get("Content-Type", "")
    if resp.status_code != 200 or "excel" not in ct:
        try:
            j = resp.json()
            msg = j.get("message") or json.dumps(j, ensure_ascii=False)[:200]
        except Exception:
            msg = resp.text[:200]
        return jsonify({"success": False, "error": f"模板下载失败 (HTTP {resp.status_code}): {msg}"}), 502

    ft = float_type
    ft_label = {"1": "每期浮动", "4": "不计成本", "5": "线下计算", "6": "特殊场景"}.get(ft, f"类型{ft}")
    co_label = {"CO01": "员工成本明细", "CO02": "组织成本"}.get(cost_object, cost_object)
    if prog_name:
        fname = f"{prog_name}_{period}_{month_str}_{ft_label}_{co_label}.xlsx"
    else:
        fname = f"导入模板_活动{activity_id}_{ft_label}_{co_label}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    dest = TEMPLATE_DIR / fname
    dest.write_bytes(resp.content)
    return jsonify({"success": True, "filename": fname, "size": len(resp.content),
                    "saved_to": str(dest),
                    "activity_name": prog_name, "period": period, "month": month_str,
                    "float_type": ft_label, "cost_object": cost_object,
                    "cost_object_label": co_label})


@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    if request.method == "POST":
        cfg = request.json
        save_config(cfg)
        return jsonify({"success": True})
    return jsonify(load_config())


# ============================================================
# 启动
# ============================================================
if __name__ == "__main__":
    print("=" * 55)
    print("  📊 造数平台 v1.0")
    print(f"  模板目录: {TEMPLATE_DIR}")
    print(f"  配置文件: {CONFIG_FILE}")
    print(f"  访问地址: http://127.0.0.1:5050")
    print("=" * 55)
    app.run(host="0.0.0.0", port=5050, debug=False)
