# 在 DSH-HARNESS 中通过插件读写 Excel VBA 工程的可行性论证

- 结论日期：2026-10-18
- 验证环境：Windows / Excel 16.0.20430.20182 / Python 3.14.4 / Node 24.15.0 / DSH `desktop` profile
- 交付物：Host 插件 `dsh-vba-bridge`（本目录），已安装并在本会话完成端到端实测

---

## 1. 结论

**可行。** 读、写、保存后校验、Excel 编译执行、离线只读、导出与体检这些链路
都已跑通，不是纸面推演：

1. **读**：Excel COM 能拿到工程名、保护状态、全部组件（标准模块、工作表、
   `ThisWorkbook`、类、窗体）及引用；离线 OVBA 解析在无 Excel 时也能取出同样的源码。
2. **写**：通过 COM 替换/追加/新建模块源码，Excel 保存并重新打开后源码一致
   （SHA-256 比对 `verified: true`）；注入的函数在 Excel 里被编译并执行成功，
   `DshDouble(21)` 返回 `42`。UserForm 的代码区（事件处理）同样可重写。
3. **导出/过程/体检**：`vba_export_modules` 导出源码或原生 `.frm` + `.frx`；
   `vba_list_procedures` 返回过程起止行与参数；`vba_doctor` 检查 Python、
   pywin32、Excel COM、AccessVBOM 与离线解析器。
4. **集成**：插件以 Host bundle 安装（`application: applied`），注册出
   `vba_inspect` / `vba_read_module` / `vba_write_module` /
   `vba_export_modules` / `vba_list_procedures` / `vba_doctor` 六个原生工具，
   本会话直接调用通过；工作簿数据（12 张表、数组公式）在保存后零差异。

成立需要三个前置条件，缺一不可：

| 条件 | 本机状态 | 缺失后果 |
| --- | --- | --- |
| Windows + 桌面 Excel | ✅ Office16 `EXCEL.EXE` | 只能离线读，不能写 |
| `AccessVBOM = 1`（信任对 VBA 工程对象模型的访问） | ✅ 已开启（非本次修改） | COM 拒绝访问，报 `VBA_TRUST_DISABLED` |
| VBA 工程无密码、文件未被占用 | ✅ 测试样本均满足 | 写被拒绝，报 `PROJECT_PROTECTED` / `FILE_LOCKED` |

不适用的场景也说清楚：**无 Excel 的无人值守服务器写入**（需要另做 OVBA 重写器，
见 §8）、**密码保护的 VBA 工程**（COM 无法编辑）、**需要保留数字签名的工程**
（任何改动都会使签名失效）。

---

## 2. 技术对象：VBA 工程不是普通文本

`.xlsm` 本质是 ZIP，VBA 工程在 `xl/vbaProject.bin`，是一个 OLE 复合文档。
对测试样本解包后看到的结构：

```text
xl/vbaProject.bin  (32,256 bytes, OLE/CFB)
├─ PROJECT                工程元数据（工程名、模块表、保护状态等）
├─ PROJECTwm              模块名映射
├─ VBA/_VBA_PROJECT       工程级版本信息
├─ VBA/dir                模块目录：每个模块的类型、流名、TextOffset
├─ VBA/模块1              [性能缓存 p-code][MS-OVBA 压缩的源码]
├─ VBA/ThisWorkbook       同上（文档模块）
└─ VBA/Sheet1 … Sheet12   同上
```

两个关键事实决定了实现路线：

1. **源码是压缩的**，位于模块流中由 `VBA/dir` 的 `MODULEOFFSET`（TextOffset）
   指向的位置；不是可以直接字符串替换的明文。
2. **模块流前还有 p-code 性能缓存**。即使改写了压缩源码，还要让 Excel 认可
   缓存/偏移的一致性，否则可能不生效或触发修复。

因此只有两条正经路线：**让 Excel 自己改**（COM/VBIDE），或**自己实现 OVBA +
OLE 重写**。前者是本插件的主引擎，后者只实现到“读”。

---

## 3. 路线对比

| 路线 | 读 | 写 | 保真度 | 依赖 | 结论 |
| --- | --- | --- | --- | --- | --- |
| **Excel COM / VBIDE** | 完整（组件、源码、引用、保护） | 完整（增删改模块、保存） | Excel 自己写，包内其他内容由 Excel 维护 | Windows + Excel + AccessVBOM | **采用** |
| **离线 OVBA 解析** | 完整（源码、模块名、类型） | 需自写压缩器 + 重写 CFB + 处理 p-code | 需自行保证 | 纯 Python（oletools） | 读采用，写作研究项 |
| LibreOffice Basic | 可读 ODF 宏 | 不维护 `vbaProject.bin`，导出会丢 VBA | 低 | LibreOffice | 排除 |
| Open XML SDK / 直接改 zip | 只能搬运二进制 | 不能改源码 | 低 | .NET | 排除 |
| VBA 自举（宏改宏） | 可 | 可 | 中 | 需先能运行宏，鸡生蛋 | 排除 |

补充：**MCP 不是替代方案**。MCP 只是传输层，底层仍要接 COM 或 OVBA；而 DSH 的
Host 插件可以直接注册原生工具，少一层进程/协议，配置和生命周期也更简单。

---

## 4. 实测证据

所有证据 JSON 在 [evidence/](../evidence/)，可复现命令见附录 A。

### 4.1 环境

| 项 | 值 |
| --- | --- |
| Excel | `C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE`，16.0.20430.20182 |
| VBA 信任 | `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`（`VBAWarnings = 1`） |
| Python | 3.14.4，`pywin32` 可用 |
| 离线解析 | oletools 0.60.2 精简子集（`vendor/`，约 1.8 MB / 90 文件；去掉 GPL 的 pcodedmp/xxxswf 与未用到的 cryptography） |
| Node | v24.15.0 |
| 测试样本 | `fixtures/sample.xlsm`（39,195 B，14 个组件）、`fixtures/sample2.xlsm`（69,298 B，15 个组件） |

### 4.2 读取

`vba_inspect` 对 `sample.xlsm` 的输出（节选，完整见
[evidence/inspect-com.json](../evidence/inspect-com.json)）：

```text
VBA project: VBAProject  [engine: com]
Excel version: 16.0
Protection: none
Components (14):
- ThisWorkbook (document, 0 lines)
- Sheet1 … Sheet12 (document, 0 lines)
- 模块1 (std, 9 lines)
References: VBA 4.2, Excel 1.9, stdole 2.0, Office 2.8
```

`vba_read_module` 读出的 `模块1` 与离线 oletools 提取的源码逐字一致（COM 不含
隐藏的 `Attribute VB_*` 行，离线引擎包含；这是两个引擎的已知差异，已在工具输出
中标注）。

### 4.3 写入与保存校验

| 操作 | 结果 |
| --- | --- |
| `replace` 重写 `模块1` | 9 → 8 行，`verified: true`，SHA-256 `0c33d055…` → `9e835062…` |
| `append` 追加 `DshPong` | 13 行，`verified: true` |
| `create=true` 新建 `DshNew` | `created: true`，5 行，组件数 15 |
| 重写文档模块 `ThisWorkbook` | 5 行，`verified: true`；离线复读确认 7 条 `VB_*` 属性仍在 |
| 通过插件工具写入 `ToolSmoke` / `DshToolMod` | `verified: true`，备份文件生成 |

校验方式：保存后**重新启动一个 Excel 实例、以只读方式打开**，读取模块源码并
按 LF 归一化后比对 SHA-256。只有一致才返回 `verified: true`。

实测发现一个实现细节：本机 Excel 的 `CodeModule.AddFromString` 会把新代码插到
模块**开头**而不是末尾，所以 `append` 不依赖它，而是“读出旧代码 → 拼接 → 整体
重写”。

### 4.4 Excel 接受并编译执行

把 `DshToolMod` 模块（含 `Public Function DshDouble(n As Long) As Long`）写入
`run-test.xlsm` 后，用 `AutomationSecurity=1` 打开并调用：

```json
{
  "modules": ["模块1", "DshToolMod"],
  "DshDouble(21)": 42,
  "compiled_and_ran": true
}
```

这证明写入的不是“看起来像源码的字节”，而是 Excel 能编译、能执行的合法 VBA。

### 4.5 工作簿数据完整性

对保存过 VBA 的 `perf.xlsm` 与原始 `sample.xlsm` 用 openpyxl 逐单元格比较
（12 张表、数组公式按公式文本归一化）：

```json
{"sheets_equal": true, "sheet_count": 12, "real_diffs": [], "diff_count": 0}
```

注意：Excel 保存会重写整个包，字节数从 39,195 变为 37,537（-4.2%），这是
Excel 自己的规范化结果，不代表内容丢失；但生产使用仍应保留备份并做业务校验。

### 4.6 离线读取（无 Excel）

`engine=ovba` 不启动 Excel，直接解析 ZIP/OLE。对同一文件输出的组件名、源码与
COM 一致（离线结果多出 `Attribute` 行，行数统计因此不同，组件类型也只能近似判断）：

```text
Components (15): ThisWorkbook … 模块1, ToolSmoke
offline read: attributes are included in code, p-code is not interpreted
```

### 4.7 失败路径（fail-closed）

| 输入 | 返回 |
| --- | --- |
| `.xlsx`（无 VBA 容器） | `NOT_MACRO_ENABLED` + “use .xlsm/.xlsb/.xls” |
| 不存在的路径 | `FILE_NOT_FOUND` |
| 不存在的模块 | `MODULE_NOT_FOUND` + 候选模块列表 |
| 存在 `~$` 锁文件 | `FILE_LOCKED` + 关闭工作簿提示 |
| 无 AccessVBOM | `VBA_TRUST_DISABLED` + 注册表路径与勾选项 |
| 密码保护工程 | `PROJECT_PROTECTED` |

错误以工具错误返回，模型能看到代码、原因与处理建议，不会“静默失败”。

### 4.8 性能

单次 COM 调用的耗时拆解（`poc_timing.py`，样本 14 个组件）：

| 阶段 | 秒 |
| --- | --- |
| `DispatchEx("Excel.Application")` | 2.67 |
| `Workbooks.Open` | 2.74 |
| 取 `VBProject` | 0.36 |
| 遍历 14 个组件（含 `CountOfLines`） | 2.00 |
| `Close` + `Quit` | 0.14 |
| **合计** | **8.5** |

端到端：`inspect` 8.9–20.5 s；`write` 改为在同一 Excel 实例里复开校验后，
本机实测约 5.3 s（此前“保存 + 重启第二个 Excel”为 12–35 s）。对交互式 agent
可接受；高频场景可做常驻 worker（§8）。

### 4.9 窗体、导出、过程与体检

UserForm 在 VBA 里同时是“设计器 + 代码模块”：控件布局在 `.frm` / `.frx`，
事件处理在它自己的 `CodeModule`。对真实加载项 `MyVBALibrary.xlam` 的
`FRM_PictureViewer`（type=3，45 行）实测：

- `vba_list_procedures` 返回 7 个过程，包括 `UserForm_Activate`、
  `btnCancel_Click`、`imgBigPicture_MouseUp`、`Property Let pictureFileName`；
- `vba_read_module` 读出完整代码（含中文注释）；
- `vba_write_module mode=append` 写入注释后 `verified: true`，行数 45 → 47；
- 写完后 `vba_export_modules mode=component` 仍导出
  `FRM_PictureViewer.frm`（1,687 B，头部 `VERSION 5.00 / Begin ...`）与
  `.frx`（2,584 B，`OleObjectBlob` 完整），manifest 标记 `design_included: true`。

导出两种模式（17 个组件实测）：

| mode | 产物 | 用途 |
| --- | --- | --- |
| `code` | `.bas` / `.cls` / `.frm`（UTF-8，仅代码） | grep、diff、喂给模型分析 |
| `component` | `VBComponent.Export` 原生文件，窗体带 `.frx` | 备份、迁移、重新导入 |

边界：本插件只处理窗体**代码**，不能修改控件布局、属性页或 `.frx` 设计数据。
`vba_doctor` 在没有 `file` 参数时只查环境；带 `file` 时额外检查锁文件与
`VBProject` 访问。

### 4.10 发布包与测试

`npm pack` 产物（2026-10-18）：

| 指标 | 值 |
| --- | --- |
| tarball | 0.88 MB |
| 解包 | 约 2.3 MB / 100 个文件 |
| `.pyc` | 0（`prepack` 清理 + `files` 白名单） |
| vendor | 约 1.8 MB，移除 GPL 的 `pcodedmp`/`xxxswf` 与未使用的 `cryptography` |

`npm test` 6 项 Node 单元测试、`npm run test:py` 5 项 Python 测试通过；把 tarball
装进空目录后六个工具全部注册、`vba_doctor` 返回 `READY`。`private` 已移除，
`license`/`engines.dsh`/`os`/`keywords` 已补齐，仓库地址已指向
`https://github.com/alexltr/dsh-vba-bridge`；远端仓库尚未创建，npm 尚未发布。

---

## 5. 插件如何接入 DSH-HARNESS

### 5.1 形态

一个 **Host-only bundle**，不需要 Client/UI：

```text
package.json      dsh.bundle.patch -> cordis.patch.yml
cordis.patch.yml  insert: { id: vba-bridge, name: dsh-vba-bridge, config: {...} }
index.js          export const inject = ['tools']; export function apply(ctx, config)
bridge/vba_bridge.py
```

安装由 `plugin_manager install_bundle` 完成，profile 里得到：

```text
dsh-vba-bridge 0.1.0 <- F:\MyProjects\dsh\vba_bridge
application: applied
```

安装后本会话立即出现三个工具的 schema；后续增加导出/过程/体检工具后，通过禁用/启用
`vba-bridge` 行重载（`application: applied`），六个工具随即可用，无需重启应用。

### 5.2 工具契约

每个工具用 `ctx.tools.register({...})` 注册，符合 `ToolDefinition`：

- `name` / `description`：模型可见；
- `parameters`：标准 JSON Schema（受 DSH 的 schema 子集约束：
  `type/properties/required/additionalProperties/enum` 等）；
- `output.schema` + `output.render(args, value)`：把结构化结果渲染成
  `[{ type: 'text', text }]`，模型看到的是人可读摘要/源码；
- `execute(args, exec)`：执行并返回可 JSON 序列化的值；`exec.signal` 支持中止；
- `timeoutMs`、`isConcurrencySafe: () => false`。

### 5.3 调用链

```text
Agent 工具调用
  → DSH tools registry（schema 校验、权限/审批、超时、日志）
  → index.js execute()
  → spawn(python, bridge/vba_bridge.py, --json-stdin)
  → vba_bridge.py：COM（Excel）或 oletools（离线）
  → stdout JSON → render → 模型
```

- **串行化**：插件内部用一条 Promise 链把调用排队，避免多请求同时操作一个
  Excel 实例/同一个文件。
- **超时/中止**：默认 240 s；`AbortSignal` 触发时杀掉子进程。
- **协议**：stdin 传 JSON 请求，stdout 只输出一行 JSON 结果；stderr 作为诊断
  尾部附在错误里。大段代码走 stdin，避免命令行长度和转义问题。

### 5.4 安全设计

| 措施 | 说明 |
| --- | --- |
| `AutomationSecurity = 3` | 桥接期间禁用宏，读写源码不会顺带执行工作簿里的宏 |
| 不提供“运行宏”工具 | 运行宏等于任意代码执行，v1 不做；需要时单独开关 |
| 写入前备份 | 同目录 `*.YYYYmmdd-HHMMSS.bak.xlsm` |
| 保存后复开校验 | 不一致返回 `verified: false`，模型可据此回滚 |
| 锁文件检测 | 发现 `~$` 前缀文件直接拒绝，避免和用户正在编辑的实例打架 |
| `allowWrite` 配置 | 可一键把插件变成只读 |
| 删除/改名不做 | 只做“读、替换、追加、新建”，减少误伤面 |

### 5.5 为什么是 Python 子进程

- Node 没有成熟稳定的 COM 绑定；`pywin32` 是 Windows 上操作 VBIDE 的事实标准。
- 同一个 Python 进程里还能用 oletools 做离线回退，一套协议两个引擎。
- 代价是多一次进程启动（约 0.3 s，不含 Excel 启动），可接受。

---

## 6. 风险与边界

| 风险 | 影响 | 缓解 | 状态 |
| --- | --- | --- | --- |
| `AccessVBOM=1` 是全局开关 | 本机任何程序都能读写 VBA，等于扩大代码执行面 | 仅在可信机器开启；插件按“可信代码”管理；不需要时关闭 | 本机已开启（非本次修改） |
| 组策略锁定信任中心 | 用户改不了，COM 永久拒绝 | 报 `VBA_TRUST_DISABLED`，提示管理员策略 | 已覆盖 |
| VBA 工程有密码 | 无法编辑 | 检测 `VBProject.Protection`，报 `PROJECT_PROTECTED` | 已覆盖（未构造样本实测） |
| 文件被 Excel 打开 / 残留 `~$` | 写入失败或争用 | 预检锁文件 + `FILE_LOCKED` + 提示关闭 | 已覆盖 |
| Excel 保存重写整个包 | 字节变化、可能丢自定义部件、签名失效 | 备份 + `verified` + 业务数据比对；签名场景明确不支持 | 已实测字节变化与数据零差异 |
| 写坏文件 | 工程损坏 | `.bak` + 保存后复开校验；失败时不删除备份 | 已实现 |
| 宏执行 | 打开工作簿时执行恶意代码 | 桥接期 `AutomationSecurity=3`；写入的代码何时运行由用户宏设置决定 | 已实现 |
| UserForm 设计器 | 不能改控件布局、属性页、`.frx` 设计数据 | 只读写窗体 `CodeModule`；`mode=component` 原样导出设计文件 | 明确不支持设计编辑 |
| 插件绕过 DSH 文件沙箱 | Python 以宿主用户权限读写任意路径 | 只安装可信插件；后续加 `allowedRoots` 白名单 | 待办 |
| 延迟 5–20 s | 交互体验差、并发成本高 | 常驻 worker；写入已改为同实例复开校验 | 部分优化 |
| 并发/多 Excel 实例 | 文件锁、COM 竞争 | 插件内串行；跨进程不保证 | 部分覆盖 |
| 无 Excel 的服务器 | 不能写 | 离线只读；写入需 OVBA 重写器 | 不支持 |
| `.xlsb` / `.xls` | 未单独实测 | COM 同一套 `VBProject` API 适用；按扩展名放行 | 待验证 |
| OneDrive/网络盘 | 同步冲突、延迟、锁 | 建议本地副本操作后再同步 | 提示 |
| Office 更新/位数 | COM 接口变化 | 只用稳定 VBIDE 成员；32/64 位由 COM 透明处理 | 低风险 |
| 中文路径/模块名 | 编码问题 | 实测通过（`模块1`、中文文件名） | 已验证 |

---

## 7. 适用性判断

- **单机、可信、有人值守的自动化**（本机场景）：**可以生产使用**。读、写、
  备份、校验闭环完整，失败可诊断。
- **团队共享环境**：先做治理——谁有权开 `AccessVBOM`、插件只装可信版本、
  写操作默认备份、关键工作簿走版本库。技术上没有阻塞。
- **CI/无 Excel 服务器**：读写分离——读可用离线引擎；写不建议，除非接受
  “自研 OVBA 重写器 + 大量兼容性测试”的成本。
- **需要签名/合规审计的工程**：编辑会使签名失效，应走“改副本 → 重新签名”流程。

---

## 8. 后续路线

| 优先级 | 事项 | 价值 |
| --- | --- | --- |
| P0 | `Config` schema + `allowedRoots` 路径白名单 | 可运维、可审计 |
| P0 | 常驻 Excel worker（复用实例） | 单次调用从 8–20 s 降到约 3–5 s |
| P1 | 写入结果附模块级 diff（旧/新） | 模型和用户都能复核 |
| P1 | **已完成**：`vba_export_modules`（源码 / `.frm`+`.frx`）、`vba_list_procedures`、`vba_doctor` | 分析、备份、迁移、安装体检 |
| P1 | 剩余：`import .bas/.cls`、删除/重命名模块 | 覆盖真实 VBA 工程维护 |
| P1 | `vba_run_macro`（显式 opt-in + 审批） | 测试/自动化，但属于代码执行，默认关闭 |
| P2 | 离线 OVBA 写入：压缩器 + CFB 重写 + p-code 处理 + Excel 验收测试 | 摆脱 Excel 依赖，工作量大、风险高 |
| P2 | 孤儿 Excel 进程清理、OneDrive 场景适配 | 稳定性 |

离线写入的技术难点已经定位清楚：`VBA/dir` 的 `MODULEOFFSET`、模块流的
`[p-code][压缩源码]` 布局、以及 OLE 流长度变化后的复合文档重写。oletools 只读，
没有现成 Python 写库；这是一项独立工程，不建议在 v1 里冒险。

---

## 附录 A：复现命令

```powershell
# 1. 读
python bridge\vba_bridge.py inspect fixtures\sample.xlsm
python bridge\vba_bridge.py read fixtures\sample.xlsm 模块1

# 2. 写（先复制，别动样本）
Copy-Item fixtures\sample.xlsm work\repro.xlsm
python bridge\vba_bridge.py write work\repro.xlsm ReproMod --code-file work\perf.bas --create

# 3. 离线读
python bridge\vba_bridge.py inspect work\repro.xlsm --engine ovba

# 4. 导出（源码 / 原生窗体）
python bridge\vba_bridge.py export test_samples\MyVBALibrary.xlam work\export-code --mode code
python bridge\vba_bridge.py export test_samples\MyVBALibrary.xlam work\export-component --mode component --overwrite

# 5. 过程清单（含 UserForm 事件）
python bridge\vba_bridge.py procedures test_samples\MyVBALibrary.xlam FRM_PictureViewer

# 6. 编译执行（需允许宏；脚本内 AutomationSecurity=1）
python poc_run2.py

# 7. 不装插件直接测插件逻辑
node test_local.mjs
```

## 附录 B：证据文件

| 文件 | 内容 |
| --- | --- |
| [evidence/inspect-com.json](../evidence/inspect-com.json) | COM 读取 14 个组件与引用 |
| [evidence/read-com.json](../evidence/read-com.json) | `模块1` 源码 |
| [evidence/write-com.json](../evidence/write-com.json) | 新建模块、前后 SHA-256、`verified: true` |
| [evidence/inspect-ovba.json](../evidence/inspect-ovba.json) | 同一文件的离线解析 |
| [evidence/inspect-xlam.json](../evidence/inspect-xlam.json) | 真实 `.xlam` 的 17 个组件与 8 个引用 |
| [evidence/read-xlam-module.json](../evidence/read-xlam-module.json) | 加载项模块源码（含中文注释） |
| [evidence/write-xlam.json](../evidence/write-xlam.json) | `.xlam` 写入、哈希与 `verified: true` |
| [evidence/inspect-xlam-ovba.json](../evidence/inspect-xlam-ovba.json) | `.xlam` 的离线解析 |
| [evidence/procedures-form.json](../evidence/procedures-form.json) | `FRM_PictureViewer` 的 7 个过程 |
| [evidence/write-form.json](../evidence/write-form.json) | 写入 UserForm 代码区、`verified: true` |
| [evidence/export-component.json](../evidence/export-component.json) | 17 个组件原生导出、`.frm` + `.frx` |
| [evidence/run-macro.json](../evidence/run-macro.json) | `DshDouble(21) = 42` |

## 附录 C：关键实现位置

| 能力 | 位置 |
| --- | --- |
| 工具注册 / 参数 schema / render | [index.js](../index.js) |
| COM 读、写、备份、复开校验 | [bridge/vba_bridge.py](../bridge/vba_bridge.py)（`com_inspect` / `com_read` / `com_write`） |
| 导出（源码 / 原生） | 同上（`com_export`） |
| 过程清单 / 体检 | 同上（`parse_procedures` / `com_procedures` / `com_doctor`） |
| 离线 OVBA 读 | 同上（`ovba_inspect` / `ovba_read`） |
| 错误码映射 | 同上（`_map_com_error`） |
| 安装清单与配置 | [package.json](../package.json)、[cordis.patch.yml](../cordis.patch.yml) |
| 单元测试 | [tests/unit.test.mjs](../tests/unit.test.mjs)、[tests/test_bridge.py](../tests/test_bridge.py) |
