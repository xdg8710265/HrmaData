# HrmaData — HRMS 造数平台

从 HRMS 下载的 Excel 导入模板生成测试数据（枚举/组织等字段按模板内 sheet 取真实值），并通过 CAS SSO 自动登录导入 HRMS 成本活动明细接口。

## 功能

- 解析 HRMS 下载的导入模板（`hidden_head` 列元数据 + 枚举 sheet + 集团组织 sheet）
- 按模板规则生成测试数据：枚举列取模板内枚举值、组织名称/组织ID 取集团组织真实数据且同行对应、工资链列保持勾稽关系
- 一键生成预览 / 下载 Excel / 生成并导入 HRMS
- CAS SSO 自动登录（RSA 加密密码，token 过期自动重登）
- 工资条安全验证自动解锁（701 门禁）

## 快速开始

```bash
pip install -r requirements.txt
cp config.example.json config.json   # 填入 sso_username / sso_password（及 pay_password）
python start_server.py               # 或直接 python app.py
```

访问 http://127.0.0.1:5050/

## 配置说明

| 键 | 说明 |
|---|---|
| `import_api_url` | HRMS 导入接口地址 |
| `import_activity_id` | 成本活动期间 ID（计提期 accrualId / 实发期 actualId），页面选活动后自动填入 |
| `import_cost_object` | CO01=员工成本明细 / CO02=组织成本明细 |
| `import_float_type` | 1=每期浮动科目 2=读取员工循环类数据 3=读取上期实发 4=不计成本 5=线下计算 6=特殊场景 |
| `import_include_data` | 仅 floatType=1 时使用 |
| `pay_password` | 工资条密码（安全验证），留空则用 SSO 密码 |
| `sso_username` / `sso_password` | CAS 登录账号 |

敏感文件（`config.json` / `ztoken.txt` / `cookie.txt`）不入库，见 `.gitignore`。

## 目录结构

```
platform/
├── app.py              # Flask 主程序（模板解析/生成/导入/SSO）
├── start_server.py     # 脱离会话的后台启动器（含健康检查）
├── templates/
│   └── index.html      # 前端页面（Jinja2 + 原生 JS）
├── config.example.json # 配置模板（复制为 config.json 使用）
└── requirements.txt    # Python 依赖
```

## 分支规范

- `master`：稳定分支，**必须通过 PR 且经确认后才能合并**，不允许直接推送
- `dev`：日常开发分支，本地以此为主，所有修改先在 dev 上提交
