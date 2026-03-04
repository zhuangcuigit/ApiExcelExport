# API 数据导出 Excel 工具

根据配置的接口地址、请求参数和列映射，请求接口获取数据并导出为 Excel 表格。

## 使用步骤

1. **复制配置**  
   将 `config.example.yaml` 复制为 `config.yaml`（或直接修改 `config.yaml`）。

2. **填写配置**  
   - **api.url**：接口完整地址（必填）  
   - **api.method**：`GET` 或 `POST`  
   - **api.cookie**（可选）：接口鉴权用 Cookie，填写浏览器复制出的整段；也可用环境变量 `COOKIE` 传入（更安全）  
   - **api.headers**：可选，如 `Content-Type` 等  
   - **api.params**：GET 时为 query 参数，POST 时默认作为 JSON body  
   - **data_path**：接口返回的 JSON 中，列表数据所在路径，如 `data.list`；若直接返回数组可留空  
   - **columns**：要导出的列；列名支持对象路径（如 `member.salesmanName`），时间列可配置 `format` 格式化  
   - **filters**（可选）：过滤规则，如某列不能为空，只保留满足条件的行  

3. **安装依赖并运行**  
   ```bash
   cd api_excel_export
   py -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   python run.py
   ```

4. **查看结果**  
   Excel 文件会生成在 `output` 目录（可在配置中修改），文件名带时间戳。

---

## 版本2：多账号登录合并导出

多账号、账号密码登录后，**分别**请求数据接口，**合并**为**一个** Excel 导出（可选增加一列「账号」标识来源）。

1. **配置**  
   使用 `config_v2.yaml`（可复制 `config_v2.example.yaml` 后修改）。

2. **必填项**  
   - **login**：登录接口地址、method、body；body 中用 `{{username}}`、`{{password}}` 占位符，运行时按每个账号替换。  
   - **accounts**：账号列表，每项为 `username`、`password`，可选 `label`（Excel 中“账号”列显示名）。  
   - **api**：数据接口（与单账号一致），请求时使用登录后的 Session，无需再填 cookie。  
   - **merge_add_account_column**：是否在合并表中增加一列“账号”（默认 true）。  
   - 其余 **data_path、columns、filters、output** 与单账号用法相同。

3. **登录成功后设置请求头**  
   - **login.auth_headers**：登录成功后要带上的请求头（后续数据接口会带上）。  
   - key 为请求头名（如 `Content-Type`、`Cookie`、`x-access-token`），value 为固定字符串（如 `"application/json"`）或**响应 body 路径**（如 `"data.accessToken"` 表示从响应 JSON 的 `data.accessToken` 取值）。  
   - 示例：`Content-Type: "application/json"`、`Cookie: "data.cookie"`、`x-access-token: "data.accessToken"`（路径需按实际登录接口返回字段修改）。

4. **登录鉴权方式**  
   - **auth_from: cookie**：登录后由 Session 自动带 Cookie，无需额外配置。  
   - **auth_from: body**：从登录响应 JSON 取 token，需配置 `auth_body_path`（如 `data.accessToken`）、`auth_header_name`、`auth_header_value`（如 `Bearer {{token}}`）。

5. **运行**  
   ```bash
   python run_v2.py
   ```  
   也可指定配置：`set CONFIG=my_v2.yaml && python run_v2.py`

6. **结果**  
   所有账号的数据合并到同一张表，Excel 中多一列「账号」标识每行来自哪个账号（若未关闭 `merge_add_account_column`）。

## 配置说明

| 配置项 | 说明 |
|--------|------|
| api.url | 接口地址 |
| api.method | GET / POST |
| api.cookie | Cookie 鉴权（整段字符串）；不填则读取环境变量 `COOKIE` |
| api.headers | 请求头 |
| api.params | 请求参数（GET 为 query，POST 为 body） |
| api.body | 仅 POST 时可用，作为 JSON body，优先级高于 params |
| data_path | 数据列表在 JSON 中的路径，如 `data.records` |
| columns | 列映射：字段支持对象路径（如 `a.b`）；值可为表头字符串或 `{ header, format }` 做时间格式化 |
| filters | 过滤规则列表，见下方 |
| output.dir | 导出目录，默认 `./output` |

### 过滤规则 filters

每条规则为一个对象，需指定 `column`（字段名，支持嵌套）和 `rule`：

| rule | 说明 |
|------|------|
| not_empty | 该列不能为空（`null`、空字符串、仅空白均视为空，不满足的行会被排除） |
| equals | 该列值等于 `value`（需在规则中配置 `value`） |
| in | 该列值在 `values` 列表中（需在规则中配置 `values` 数组） |

示例：

```yaml
filters:
  - column: name
    rule: not_empty
  - column: status
    rule: equals
    value: "正常"
  # - column: type
  #   rule: in
  #   values: [ "A", "B" ]
```

多条规则为“且”关系：行需同时满足所有规则才会保留。
| output.filename | 文件名前缀，会自动加时间戳和 .xlsx |

### 列配置 columns（对象路径与时间格式）

- **列名（key）**：接口返回的字段名，支持**对象路径**从嵌套对象取值，如 `user.name`、`member.salesmanName`。
- **列值（value）**：简写为表头字符串（如 `name: "姓名"`），或对象 `{ header: "表头", format: "时间格式" }` 用于时间列。
- **时间格式**：`format` 使用 strftime，如 `"%Y-%m-%d %H:%M:%S"`、`"%Y-%m-%d"`；接口返回的 ISO 字符串或时间戳会自动解析后按该格式输出。

```yaml
columns:
  id: "ID"
  realName: "用户名称"
  member.salesmanName: "业务员"
  registerDate:
    header: "创建时间"
    format: "%Y-%m-%d %H:%M:%S"
```

## 示例

接口返回：

```json
{
  "code": 0,
  "data": {
    "list": [
      { "id": 1, "name": "张三", "status": "正常" },
      { "id": 2, "name": "李四", "status": "正常" }
    ]
  }
}
```

配置 `data_path: "data.list"`，`columns: { id: "ID", name: "姓名", status: "状态" }`，即可导出包含 ID、姓名、状态三列的 Excel。

## 注意事项

- 接口需返回 JSON。鉴权推荐用 **Cookie**：在 `config.yaml` 的 `api.cookie` 中填写，或设置环境变量 `COOKIE`（避免把 Cookie 写进配置文件）。也可在 `api.headers` 里配置 `Authorization` 等。  
- 列顺序以 `columns` 中键的顺序为准。  
- 若某条数据缺少某字段，该单元格会为空。
