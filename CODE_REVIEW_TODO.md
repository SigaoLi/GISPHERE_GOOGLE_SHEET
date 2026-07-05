# GISource 代码审查待办

> 生成日期：2026-06-29
> 审查范围：`Google_Sheet/` 全部 13 个 `.py`（与 `GISPHERE_GOOGLE_SHEET-main/` 当前逐字节一致）
> 状态约定：☐ 待处理 / ☑ 已确认无需改 / ✅ 已修复

---

## A. 疑似 Bug（优先确认）

### ✅ A1. 删除行的行号换算两处不一致，off-by-one —— 2026-07-05 已修复
`delete_rows_from_sheet` 把传入值直接当作 `deleteDimension.startIndex`（0-based，表头=0，第一条数据=1）。两个调用点算法不同：

- `main.py:281` `load_and_clean_data`：`rows_to_delete_sheet = [x + 2 for x in ...]`（注释称 +1 表头 / +1 因 0 开始）
- `main.py:624` `update_google_sheets`：`rows_to_delete = [x + 1 for x in ...]`

`deleteDimension` 为 0-based，删第一条数据应是 `index+1` → `update_google_sheets` 的 `+1` 正确，`load_and_clean_data` 的 `+2` 看起来多删了相邻的下一行。两处用同一 helper 却差 1，至少一处错。
**风险**：过期/重复清理时误删紧邻的有效行，行数多时不易察觉。
**待办**：修 `main.py:281` 改为 `x + 1`。

**✔ 2026-07-05 已在 Content_test 实测证实**：追加标记行 T1..T4 后，用 `+2` 公式删 T2 实际删掉的是 T3（下一行）；用 `+1` 公式精确命中目标。**`main.py:281` 确为 off-by-one bug。**
线上从未暴露的原因（同日核实）：① 条目通常在过期前就经步骤 7（`+1` 正确路径）移入 Filled，当前 Unfilled 0 条过期，该分支极少触发；② Deadline 升序排列，过期行成片出现在顶部，整块删偏一行时删掉的大多仍是过期行——症状仅为"块首留一行过期（下次运行再触发）+ 块尾误删一行有效数据"，非常隐蔽。

**修复内容**：`main.py:281` 改为 `[x + 1 for x in all_rows_to_delete]`，注释同步更正。集成测试（Content_test 重定向跑真函数 `load_and_clean_data`）：2 行过期标记被精确删除、非过期行零误删、行数吻合。

### ✅ A2. `update_university_info` 写回范围写死 `A:Z` —— 2026-07-05 已修复（锚点写法）
`main.py:351`：
```python
range_name = f'Unfilled!A{row + 2}:Z{row + 2}'
update_data = [unfilled_data.iloc[row].tolist()]
```
表列数几乎肯定 >26，但范围只到 Z，而 values 长度等于实际列数 → `values.update` 超范围会报错（"tried writing to column AA"）。
（注：此处 `row+2` 用 1-based A1 记法、表头占 1 行，是正确的；与 A1 的 delete 偏移本就不同，注意区分。）

**✔ 2026-07-05 已实测证实**：Unfilled 实际 31 列（末列 AE）。在 Content_test 上复现：`A900:Z900` 写 31 个值 → HTTP 400 "tried writing to column [AA]"。且该函数在 `main()` 大 try 内，一旦触发整个运行直接崩溃——至今没炸说明 `modified_rows` 每次都为空（University_CN 总是已填），属于未引爆的雷。
**推荐修法（按工程化程度排序）**：
1. **锚点写法（最推荐）**：`range_name = f'Unfilled!A{row + 2}'`——只给起始单元格，Sheets API 自动按 values 实际宽度向右写入，列数对齐由 `len(unfilled_data.columns)` 决定，永不写死。已实测可行。
2. 动态末列：`column_index_to_letter(len(unfilled_data.columns) - 1)` → `A{r}:AE{r}`。已实测可行，但列数变化时语义上多一层换算。
3. 顺手做 C2：多行合并为一次 `values.batchUpdate`（`data=[{range, values}, ...]`），锚点写法同样适用。

**修复内容**：`main.py:351` 改为 `f'Unfilled!A{row + 2}'`（方案 1）。集成测试（假 DB 返回 `Univ_A2TEST`）：真函数走完写回，University_CN/Country_CN 正确落表、无 400、None 填充列按 API 语义跳过不覆盖。

### ✅ A4.（2026-07-05 修 A2 时新发现）Deadline 原地转 Timestamp 导致写回崩溃/冲掉原文 —— 同日已修复
`load_and_clean_data` 原来在 `main.py:229` 把非 Soon 行的 `Deadline` **原地**转成 `pd.Timestamp/NaT`；仅当发生删除时才重新 fetch 恢复成字符串。后果有二（均本地演示证实）：
1. 无删除的运行中若有大学需补 CN 名，写回 body 含 `Timestamp` → `json.dumps` 抛 `TypeError`（NaT 同理）；
2. 无法解析的文本（如 `rolling`）被 coerce 成 NaT，写回时会**冲掉单元格原文**——比崩溃更隐蔽的数据损失。

**修复内容**：不再原地转换。改用临时序列 `deadline_parsed = pd.to_datetime(unfilled_data['Deadline'].where(mask_not_soon), errors='coerce')` 做过期判断，`unfilled_data` 全程保持表格原始字符串（'Soon'、'rolling'、日期文本原样）。下游兼容性已核：步骤 4（main.py:444-449）本就自行 astype(str)+to_datetime 在副本上算，`selected_row` 取自原始 df（main.py:479），`data_processor.py:123-139` 对 str/datetime 双分支——"全字符串"正是发生过删除的运行的既有状态。
**集成测试通过**：混合 Soon/rolling/过期/未来行时只删过期行；无删除运行返回的 df 无 Timestamp/NaT；`update_university_info` 写回 3 行成功且 'Soon'/'rolling' 原文完好。

### ☑ A3. `'1'` 的类型判断不一致 —— 已确认当前无影响
- `data_processor.py:28` `create_job_title`：`row[column] in [1, '1', 1.0]`
- `data_processor.py:211+` `generate_abbreviation`、`:307` 标签：`== '1'`（仅认字符串）

**✔ 2026-07-05 已实测**：`values().get` 默认 `FORMATTED_VALUE`，返回的所有单元格一律为 `str`（对真实 Unfilled 全表逐格验证，类型集合 = {str}），int 1 不可能经此路径出现，`== '1'` 当前安全。
残余风险（低）：① 单元格显示格式改成数字带小数会返回 `'1.00'`；② 改用复选框会返回 `'TRUE'`；③ 改动 `valueRenderOption` 为 UNFORMATTED 会返回 int。均属"将来有人改表/改代码"的防御性问题，可在做 D 系列时顺手抽 `_is_one(v) = str(v).strip() in ('1','1.0')`。
**⚠ 顺带发现的数据问题**：Unfilled 的 `PostDoc` 列存在值 `'2'`（疑为录入者把名额数填进了岗位列），两处判断都只认 `'1'`，该行会被漏判为"非 PostDoc"——这是数据约定问题而非代码 bug，需和录入约定核对（名额应填 `Number_Places`）。

---

## B. 健壮性

### ✅ B1. 空表 / 空返回未防御 —— 2026-07-05 已修复
`main.py` 原 212/218 直接 `unfilled_raw[0]`；`fetch_data` 返回 `[]` 时 `IndexError`。
**修复内容**：新增 `main.fetch_dataframe(range_name)` helper（fetch→表头→adjust→DataFrame 四行样板 ×3 处收敛为一），整表拉取为空抛明确 `RuntimeError`；仅表头零数据行是合法状态返回空 df。三种路径均实测。

### ✅ B2. `database.py` schema 限定不一致 —— 2026-07-05 已修复
原 `clean_university_names` / `get_gisource_data` / `check_universities_exist` 硬编码 `TEST.` 前缀，其余用裸表名。
**修复内容**：删除三处 `TEST.` 前缀，库名统一由 `sql_credentials.txt` 的 `database` 决定（用户已确认 TEST 是真库名、仅当年未改名）。实测：`SELECT DATABASE()`=TEST，裸表名与 `TEST.` 前缀逐表行数等价，三个只读函数实跑正常。迁库（如上 Azure）时只需改各人 ini 一行。

---

## C. 效率

### ✅ C1. 凭据与 service 每次请求都重建 —— 2026-07-05 已修复
**修复内容**：`google_sheets.authorize_credentials` 加进程内 `_cached_creds`（`valid` 失效自动落回原刷新/重授权流程）；`google_http.build_google_service` 按 `(api, version)` 缓存并校验 creds 身份，新增 `invalidate_service_cache()` 在 `execute_with_retry` 重试前调用——保留"瞬时错误后重建连接"的原有语义。实测缓存命中 5ms→0.001ms，重试失效后返回新对象。

### ✅ C2. `update_university_info` 逐行写回 —— 2026-07-05 已修复
**修复内容**：`google_sheets.py` 新增 `batch_update_data_in_sheet(updates)`（`values.batchUpdate`），`update_university_info` 攒齐所有修改行一次发出，范围沿用 A2 锚点写法。实测 3 行写回 = 1 次 API 调用、全部正确落表。多人共享项目配额下降低 429 风险。

### ☑ C3. `config.py` 导入即探测代理 —— 决定不做
多地区多人运行环境核实：国内开代理者探测列表首位即命中（约 0.3–0.8s），海外直连者 6 端口秒拒近零开销——各类环境现状已近最优，缓存仅省约 0.5s 却让代理分岔逻辑背上回归风险，收益/风险比不划算。已在 README 补充说明：显式设 `GOOGLE_API_PROXY` 可跳过探测。

---

## D. 代码质量 / 未来兼容

### ✅ D1. pandas 弃用 API + 全局压警告 —— 2026-07-05 已修复
- `applymap` → 纯 Python 逐格转换 `[[convert_value(v) for v in selected_row.iloc[0]]]`（不用 `DataFrame.map`：它要求 pandas≥2.1，多人环境版本不一，直接绕开版本敏感 API）
- `utils.is_date`、`main.load_and_clean_data` 的 `to_datetime` 格式推断告警 → `warnings.catch_warnings()` **局部定向**静音（试解析混合文本属预期行为）
- **删除全局 `warnings.filterwarnings('ignore')`**——删除后立刻暴露一条真实 UserWarning 并被定向处理，验证了"警告雷达"价值。实测全流程运行无 Future/DeprecationWarning、Timestamp 正确转字符串。

### ✅ D2. 裸 `except:` —— 2026-07-05 已修复
`data_processor.py:136、474` → `except Exception`（复查 google_docs.py 并无裸 except，原清单有误）。实测 `KeyboardInterrupt` 可穿透 `parse_deadline_for_sort`，正常兜底逻辑不变。

### ✅ D3. 依赖清理 —— 2026-07-05 已修复
requirements.txt：移除 `pycountry`（从未被 import，且其无中文国家名数据，不能替代 `COUNTRY_DICTIONARY`）；顺带移除 `configparser`（Python 2 backport，py3 标准库自带）；为 pandas/numpy/google-api 等关键依赖加版本上界，防多人环境装到不兼容大版本。

### ✅ D4. 步骤编号 —— 2026-07-05 已修复
补 `print("步骤 8: 生成职位缩写...")`（选择补空号而非重排 9/10，避免改动日志 step id 影响历史日志可比性）。

---

## 状态总结（2026-07-05）
全部 12 项关闭：A1✅ A2✅ A3☑ A4✅（审查后新发现）B1✅ B2✅ C1✅ C2✅ C3☑（决定不做）D1✅ D2✅ D3✅ D4✅。

> **尚未同步到 `GISPHERE_GOOGLE_SHEET-main/`**（改动涉及 main.py / google_sheets.py / google_http.py / database.py / utils.py / data_processor.py / requirements.txt / README.md），待本轮全部改动定稿后一次同步。
>
> ⚠ 测试铁律（2026-07-05 事故教训）：跑会写表的测试前，必须重定向 main 命名空间**全部**写入函数（update/batch_update/append/delete），并在测试首尾快照校验真实 Unfilled 逐字节未变。
