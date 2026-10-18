# dsh-vba-bridge

让 DSH 会话里的 AI 直接读写 **Excel VBA 工程**：列模块、读源码、改代码、导出
`.frm`/`.frx`、列过程，并在保存后重新打开校验。核心走 Excel 自己的 COM/VBIDE，
所以改完的宏由 Excel 编译、保存和验收，而不是硬拆 `vbaProject.bin`。

Windows · 桌面 Excel · DSH Host 插件 · MIT

GitHub: <https://github.com/alexltr/dsh-vba-bridge>

![VBA Bridge 讲解图](docs/screenshot.png)

## 工具

| 工具 | 作用 |
| --- | --- |
| `vba_doctor` | 安装后先跑：检查 Python、pywin32、Excel COM、`AccessVBOM`、离线解析器、临时目录，可选一个工作簿 |
| `vba_inspect` | 工程名、保护状态、全部组件（模块/工作表/ThisWorkbook/窗体）与引用 |
| `vba_read_module` | 读取单个模块源码，含 UserForm 代码区 |
| `vba_write_module` | 替换/追加/新建模块，原地保存，复开校验，生成 `.bak` |
| `vba_export_modules` | 导出全部组件：源码（`.bas/.cls/.frm`）或原生 `.frm` + `.frx`，写 `_manifest.json` |
| `vba_list_procedures` | 列出 `Sub`/`Function`/`Property` 的参数、起止行号，含窗体事件处理 |

相对路径按 **DSH 会话工作目录** 解析；绝对路径原样使用。

## 安装

发布到 npm 后：

```text
dsh plugin add dsh-vba-bridge
```

当前开发目录（本地链接安装）：

```text
plugin_manager action=install_bundle target=F:\MyProjects\dsh\vba_bridge
```

安装后先调用 `vba_doctor`，它会把“缺 Python / 缺 pywin32 / 没开 AccessVBOM / Excel
起不来”逐项列出来并给修复建议。

## 前置条件

1. **Windows + 桌面版 Microsoft Excel**（本机实测 16.0.20430.20182）。
2. **Python 3.10+**，并安装 `pywin32`：
   `pip install pywin32`。默认使用 PATH 里的 `python`，可用配置 `pythonPath` 或
   环境变量 `DSH_VBA_PYTHON` 指定。
3. Excel 信任中心开启 **“信任对 VBA 工程对象模型的访问”**：
   `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`。
   这是允许本机程序读写 VBA 的全局开关，属于安全决策，不要在不可信的机器上随意打开。
4. 目标文件是 `.xlsm` / `.xlsb` / `.xls` / `.xlam`；`.xlsx` 没有 VBA 容器，会被拒绝。

## 用法

直接用自然语言即可：

- “看一下 `D:\book.xlsm` 里有哪些 VBA 模块”
- “把 `模块1` 的源码读出来”
- “把这段代码写进 `模块1`，先备份”
- “把 `MyVBALibrary.xlam` 的全部模块导出到 `F:\dump`”
- “列出 `FRM_PictureViewer` 里有哪些事件处理过程”

手动参数示例：

```json
{"file": "D:\\book.xlsm"}
{"file": "D:\\book.xlsm", "module": "模块1"}
{"file": "D:\\book.xlsm", "module": "模块1", "code": "Sub Hello()\r\n    MsgBox \"hi\"\r\nEnd Sub\r\n", "mode": "replace"}
{"file": "D:\\book.xlsm", "outDir": "D:\\dump", "mode": "component", "overwrite": true}
{"file": "D:\\book.xlsm", "module": "FRM_PictureViewer"}
```

### 写入语义

- `mode=replace`：重写模块代码区；隐藏的 `VB_*` 属性由 Excel 保留，文档模块可安全重写。
- `mode=append`：读出旧代码 → 拼接 → 整体重写（不依赖 `AddFromString` 的插入位置）。
- `create=true`：模块不存在时新建标准模块。
- 每次写入先复制 `原名.YYYYmmdd-HHMMSS.bak.xlsm`，保存后在同一 Excel 实例里重新
  打开并比对源码 SHA-256，返回 `verified: true/false`。

### 窗体（UserForm）

- 窗体的代码就是它的 `CodeModule`：`UserForm_Activate`、`btnCancel_Click`、
  `imgBigPicture_MouseUp` 等事件处理都能读写，和普通模块一致。
- 只改代码区，**不改**控件布局、属性页或 `.frx` 设计数据。
- `mode=component` 导出 `.frm` + `.frx`，`.frm` 头部保留 `VERSION 5.00 / Begin ...`，
  `.frx` 的 `OleObjectBlob` 完整，可用于备份/迁移。
- `mode=code` 的 `.frm` 只有代码，不能直接重新导入。

### 导出

- `mode=code`：UTF-8 源码，适合 grep、diff、喂给模型分析。
- `mode=component`：调用 `VBComponent.Export`，窗体会多出 `.frx`。
- 输出目录写 `_manifest.json`（类型、行数、字节数、SHA-256、跳过/失败）。
- `overwrite` 默认 `false`，已存在的文件跳过并报告。

## 配置

安装行 id 是 `vba-bridge`，默认配置在 [cordis.patch.yml](cordis.patch.yml)：

| 键 | 默认 | 说明 |
| --- | --- | --- |
| `pythonPath` | 空 → `python`（其次 `DSH_VBA_PYTHON`） | 带 `pywin32` 的 Python |
| `bridgePath` | 空 → 包内 `bridge/vba_bridge.py` | 自定义桥接脚本 |
| `timeoutMs` | `240000` | 单次工具调用超时 |
| `allowWrite` | `true` | `false` 时 `vba_write_module` 直接拒绝 |
| `maxOutputChars` | `200000` | 工具输出截断上限 |

修改后重载插件（禁用/启用）或重启 Harness 生效；也可以在该 profile 的 patch 层
按行 id `vba-bridge` 覆盖配置。当前版本尚未导出 `Config` schema，所以这些键通过
patch 配置而不是设置页表单修改。

## 安全边界

- 桥接期间 `AutomationSecurity=3`，读写源码不会顺带执行工作簿里的宏。
- 不提供“运行宏”工具；运行宏等于任意代码执行，需要时另行评估。
- 插件以宿主用户权限运行 Python，**绕过 DSH 文件沙箱**；只安装可信版本，只对可信工作簿使用。
- 写入默认生成 `.bak` 并复开校验；校验失败会返回 `verified: false`，不会静默成功。
- 密码保护的 VBA 工程无法编辑；任何改动都会使数字签名失效。
- 保存由 Excel 重写整个包，字节会变化（本机样本 39,195 → 37,537 字节）；
  生产使用请保留备份并做业务校验。

## 错误码

| 代码 | 含义 | 处理 |
| --- | --- | --- |
| `NOT_MACRO_ENABLED` | `.xlsx` 等无 VBA 容器 | 换 `.xlsm`/`.xlsb`/`.xls` |
| `FILE_NOT_FOUND` | 路径不存在 | 检查路径/相对路径基准 |
| `FILE_LOCKED` | 被 Excel 占用或残留 `~$` 文件 | 关闭工作簿/清理锁文件 |
| `VBA_TRUST_DISABLED` | 未开 AccessVBOM | 信任中心勾选 |
| `PROJECT_PROTECTED` | VBA 工程有密码 | 在 VBE 中取消密码 |
| `MODULE_NOT_FOUND` | 模块名不匹配 | 按提示候选名重试 |
| `EXCEL_UNAVAILABLE` | COM 无法启动 Excel | 安装桌面 Excel、检查策略 |
| `PYWIN32_MISSING` | 当前 Python 无 pywin32 | `pip install pywin32` |
| `OLETOOLS_MISSING` | 离线解析器缺失 | 重装 bundle |
| `ENGINE_READONLY` | 指定 `engine=ovba` 做写入/导出 | 改用 COM |
| `INVALID_OUTDIR` | `outDir` 指向文件 | 传目录路径 |

## 开发与测试

```powershell
npm test          # Node 单元测试：工具注册、schema、相对路径、元数据
npm run test:py   # Python 单元测试：过程解析、文件名清洗
```

E2E 需要 Excel + `AccessVBOM=1`，见 [docs/feasibility.md](docs/feasibility.md) 与
`test_local.mjs`。桥接脚本可以直接跑：

```powershell
python bridge\vba_bridge.py inspect fixtures\sample.xlsm
python bridge\vba_bridge.py read fixtures\sample.xlsm 模块1
python bridge\vba_bridge.py export fixtures\sample.xlsm work\export --mode component
python bridge\vba_bridge.py procedures fixtures\sample.xlsm 模块1
```

## 项目结构

```text
vba_bridge/
├─ index.js                 # 六个工具与路径解析
├─ package.json             # bundle 清单与市场元数据
├─ cordis.patch.yml         # 插入 vba-bridge 行
├─ locale/{en,zh}.json      # 插件管理器显示名/描述
├─ icon.svg
├─ bridge/vba_bridge.py     # JSON 协议桥接：COM + 离线 OVBA
├─ vendor/                  # 精简 oletools 子集（离线只读）
├─ tests/                   # Node + Python 单元测试
├─ docs/feasibility.md      # 可行性论证
├─ docs/screenshot.png      # 讲解图
└─ CHANGELOG.md
```

## 限制

- 写入/导出依赖桌面 Excel；无 Excel 只能离线读，且离线模式的组件类型是近似值、
  行数包含隐藏的 `Attribute` 行。
- 窗体只支持代码，不支持设计器编辑。
- 当前不运行宏。
- 当前没有 `Config` schema / 设置页表单，也没有 `allowedRoots` 路径白名单；
  配置通过 profile patch 完成。
- `.xlsb` / `.xls` 未单独实测（`.xlsm` / `.xlam` 已实测）。
- 单次调用通常 5–15 秒，取决于 Excel 启动、杀毒软件和是否复开校验。

## License

MIT；第三方组件见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。
