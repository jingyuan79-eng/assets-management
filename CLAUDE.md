# 资产驾驶舱 · 项目说明

> **以实际代码为准，不要假设本文档中的结论仍然正确。**
> 每次分析前先读代码或核对线上状态。本文档里标注了「⚠️ 已核对」的条目是
> 2026-08-29 逐条对过代码的；其余部分未复核。

本仓库：`index.html`（前端全部）、`AppsScript.gs`（后端全部）、本文档，以及 10 套测试（见第七节）。

---

## 一、项目概况

个人资产管理 PWA。前端单文件 HTML，后端 Google Apps Script，数据存 Google Sheet。

| | |
|---|---|
| 线上地址 | assets-management-wine.vercel.app |
| 前端 | 单文件 `index.html`（无框架、无构建、无依赖） |
| 后端 | Google Apps Script Web App（`doGet` 读+写 / `doPost` 供 Shortcut） |
| 数据库 | Google Sheet（名称 `Aseset management`） |
| 记账入口 | iPhone 快捷指令「记账」（POST 写 Ledger） |
| 托管 | GitHub `jingyuan79-eng/assets-management` → Vercel 自动部署 |

### 部署流程

**前端**：改 `index.html` → 传 GitHub 覆盖 → Vercel 约 30 秒生效。

**后端**：粘贴代码 → 保存 → Deploy → **Manage deployments** → 铅笔 ✏️ → Version 选 **New version** → Deploy（网址不变）。

> 踩过的坑：`selfTest` 跑通只证明**代码保存了**，不证明**部署生效了**。
> 验证部署是否生效的唯一可靠方法：**浏览器直接打开 `/exec` 网址**，看到 JSON 才算成功。

**首次部署后必做**：在编辑器里手动运行一次 `installTriggers`，安装每日凌晨 3 点的自动任务（储蓄计息 + 国债付息）。

⚠️ 前后端有耦合，**必须同时上线**：前端的补记去重依赖后端返回的 70 天流水窗口（见第二节第 9 条）。

---

## 二、核心设计原则

这几条是整个系统的骨架，改动前务必理解。

### 1. Sheet 是唯一真相源

资产数据的权威值全在 Sheet。localStorage 存两类东西：**配置**，以及**上次同步结果的缓存**（缓存只用于加速首屏，任何时候都会被服务器数据覆盖，见第五节）。

### 2. 锚点模型（流动现金）

```
流动现金 = 最近一条锚点的金额 + 该锚点当天及之后的所有 Ledger 流水
```

- 锚点存在 `Anchor` 表，**只增不改**
- 「对账」时输入银行真实余额 → 落一条新锚点 → 误差归零
- 好处：不依赖每笔账都记对；换设备、清缓存、改代码都不影响

### 3. 记账起始日

= `Anchor` 表最早那条锚点**所在月的 1 号**。

这之前发生的任何事**一律不补录**（工资、固定支出、国债利息、定期转账），因为那些钱已经含在初始余额里了。

### 4. key 防重复

所有自动产生的流水都带唯一 key 写在 Ledger 的 E 列，服务器端查重：

| key 格式 | 用途 |
|---|---|
| `fix:<配置id>:<日期>` | 固定支出 |
| `pay:<日期>` | 工资 |
| `xfer:<配置id>:<日期>` | 定期转账 |
| `bond:<国债id>:<日期>` | 国债半年息 |
| `bondmat:<国债id>` / `bondint:<国债id>` | 国债到期本金 / 利息 |

**这是防重复的唯一机制**。本地状态丢失、多设备同时打开、重复点击，都不会产生重复记账。

### 5. 分类前缀决定资金性质

Ledger 的 `category` 列靠前缀区分四种性质：

| 前缀 | 对流动现金 | 计入 |
|---|---|---|
| 无（11 个支出分类） | 减 | 支出 |
| `Income ·` | 加 | 收入 |
| `Transfer ·` | 减 | 投资转出 |
| `Redeem ·` | 加 | 本金回流 |

**净收入 = 收入 − 支出**，不含转账和回流。
**月末现金 = 期初 + 净收入 − 投资转出 + 本金回流**。

前后端各有一份 `catKind` / `catSign`，规则必须保持一致。

### 6. 锚点式结息（储蓄）

任何中途变动都是「**先按旧状态结息 → 起始日设为今天 → 再改**」：改利率、转入、转出都走这条路径。这让「转账抹掉已滚利息」这类 bug 结构性不可能发生。

### 7. 乐观更新 ⚠️ 已核对

所有编辑操作：**本地先改 → 界面立刻刷新 → 后台写 Sheet → 失败自动撤回**。

写请求经 `apiWrite` 串行排队（`sheetQueue` promise 链）。**注意：没有固定间隔**——全文件 `setTimeout` / `setInterval` 计数为 0，纯靠 promise 链排队。（旧文档写的「间隔 120ms」不成立。）

多数写操作**直接返回最新列表**，前端据此更新，不再二次拉取。但以下情况仍会触发 `refreshFromSheet()` 全量重拉：Marcus 储蓄转入/转出、删除失败后的状态恢复。

**改动界面时的铁律**：本地改完数据后必须触发重绘。曾出现过 `adjustBrokerCash` 改完券商现金却不重绘，导致买入股票后股数立刻更新、现金要等下一次交互才变。凡是修改 `DATA` 的函数，末尾都要 `recompute()` + `refreshOpenDetail()`。

### 8. 服务器端锁

`LockService` 包住所有读写入口。Apps Script 允许并发执行，没有锁的话「先查 key、再写入」会被两个请求同时穿过。

### 9. 70 天流水窗口 ⚠️ 已核对

后端返回的 `ledger` 明细包含：**当月 + 上月的全部**，再加**最近 70 天**的任何流水。

这个 70 天不是随便定的——前端 `dueDates()` / `bookFloor()` 判断「这笔要不要补记」时用的是 **60 天**窗口，而去重靠 `ledgerHasKey()` 在 `ledgerRows` 里查。**两个窗口必须让 70 > 60**，否则落在缝隙里的 key 前端查不到，每次打开都会把早就写过的固定支出、工资重发一遍（服务器再逐条回 `skipped`）。

实测：对齐之前平均每次打开白跑 4.5 个写请求，月初最多 13 个；对齐后为 0。

`ledgerMonths` 仍然只声明 `[本月, 上月]`——70 天必然完整覆盖这两个月（最坏 31+30=61 天），前端 `rebuildMonthTotals()` 只重算这两个月，多出来的旧行会被自动过滤。

### 10. 首屏用缓存，但必须标明时点 ⚠️ 已核对

打开时先用 `sheetCache_v1` 把现金/储蓄/国债填满再渲染，不空等网络。

**原则：可以先显示旧数字，但必须让人一眼看出是旧的。** 同步徽标会显示数据时点（当天「上次同步 HH:MM」，隔天「数据为 M/D」），新数据到达后转为「已同步 HH:MM」。**绝不让过期余额冒充实时余额。**

---

## 三、Sheet 表结构

### Ledger（流水，只增不清）

`A:Date  B:category  C:amount  D:note  E:key`

金额一律正数，性质由分类前缀决定。

### Ledger_monthly（自动重算，勿手改）

`年月 | 收入 | 支出 | 净收入 | 投资转出 | 本金回流 | 月末流动现金 | Payroll | Bonus | 11个支出分类 | 未匹配`

`未匹配` 是报警灯，正常恒为 0。

注意：这张表只被 `writeMonthly` 写和比对，**从不作为数据源读取**——汇总每次都从 Ledger 原表重算。

### Anchor（对账锚点，只增不改）

`A:date  B:amount  C:note`

第一行是记账起点。

### Stock（股票持仓，双向同步）

`A:Symbol  B:Category  C:Share  D:Cost  E:Price`

- `E:Price` 由 `=GOOGLEFINANCE(A2,"price")` 自动拉取，**代码绝不写这一列**（Cash 行除外）
- `Cash` 行代表券商可用现金，金额借用 E 列，Category 写 `Gold.coin.cash`
- **ASMI 必须写作 ASMIY**，否则 GOOGLEFINANCE 拉不到价

Category → 小程序分组映射见代码里的 `STOCK_TAXONOMY`。

### Savings（储蓄）

`A:ID  B:Name  C:Type  D:Balance  E:Rate  F:Last Post  G:Next update  H:Maturity  I:Status`

- `Type`：`CD`（定期）/ `OS`（活期）
- `Rate` 按 **APY** 处理，`4` / `4.1` / `0.04` 三种写法都认
- `Balance` 是已入息后的余额，`Last Post` 是上次入息日——这两个值构成完整状态
- `Status`：`active` / `closed`
- ID 留空会自动生成

### Bond（国债）

`A:ID  B:Name  C:Type  D:Start  E:Term  F:Rate  G:Principal  H:Balance  I:LastPost  J:Status`

- `Type`：`T-Note` / `I-Bond` / `T-Bill`
- `LastPost` = 已经付息/复利到哪一期，**填错会导致重复补发或漏发**
- 到期日由 `Start + Term` 算出，不用填

### Notices（自动事件通知，脚本自建）

`A:id  B:date  C:text  D:dismissed`

只记录**无声的资金搬动**（CD 到期转 OS、国债到期回流）。利息不发通知，因为收入明细里看得到。

### 其他

`HSA` / `Retire` / `Property` 表已建但**尚未接入**，这三块目前仍在 localStorage。后端 `readAll` 已经返回 `hsa` / `retire` 字段，但前端还没读。

---

## 四、计算分工

### Apps Script 算

| 内容 | 说明 |
|---|---|
| 储蓄计息 | 每月 1 号入账，APY→日利率按日复利，`(1+APY)^(1/365)-1` |
| CD 到期 | 结息到到期日 → 本息自动转入 OS → `closed` → 发通知 |
| 国债付息 | T-Note 每半年 `本金×利率÷2` → 写 Ledger 收入 → 进流动现金 |
| I-Bond 复利 | 每半年 `余额×(1+利率/2)`，滚进本金不产生现金 |
| 国债到期 | 本金记 `Redeem · 国债`，利息记 `Income · 国债利息` |
| 流动现金余额 | 锚点 + 之后流水 |
| 月度汇总 | 写 `Ledger_monthly`，内容没变就不重写 |
| 对账 | 算差额、可选补一笔流水、落新锚点 |
| 所有 key 查重 | |

### 小程序算

| 内容 | 说明 |
|---|---|
| 当月各分类支出 | 从流水明细现算，改一笔立刻反映 |
| 股票市值 / 占比 | 用 Sheet 的股价（上次同步时的值，非实时行情） |
| 房贷摊销、401k 增长 | 纯本地 |
| 固定支出 / 工资 / 定期转账**该不该记** | 按配置和频率算出应记日期，再交给服务器查重 |
| 国债下一次付息金额与日期 | 显示用 |

### 两边都算但不冲突的

| 项目 | 为什么不冲突 |
|---|---|
| 流动现金 | Sheet 是权威值，本地只在编辑瞬间按增量临时调整，同步后被覆盖 |
| 当月分类支出 | 同一批流水、同一套规则（`catKind` 前后端一致） |
| 储蓄余额 | 小程序完全不算，只显示 |

### 计算引擎为什么留在读路径里 ⚠️ 已核对

`readAllInner` 每次都调 `runSavings` / `runBonds`，而 `dailyJob` 凌晨 3 点已经跑过同样的逻辑。这**不是**重复记账：

- 实测（dailyJob 已跑过的当天）：Savings 读 4 写 0，Bond 读 4 写 0，表格内容一个字节没变。引擎是幂等的。
- 这两个函数同时承担「读取 Savings/Bond 表并格式化成前端字段」的职责，**不能整段跳过**，否则这些读一次也省不了。
- 留在读路径里还能自愈：定时任务失败、或在 1 号凌晨 3 点前打开 App，读操作会顺手补上。

~~真正多余的只有每行 1 次的 `getValue`~~ **已于 v4 消除**：改用 `var needId = !r.id;` 在覆盖空 ID 之前先记下标记，不再回读单元格。实测 Savings 读 4→1、Bond 读 4→2（Bond 保留的 1 次是 `bondSheet` 的表头检查）。

### 性能瓶颈在哪（实测）⚠️ 已核对

打开慢**不是计算慢**。去掉 I/O 的净计算耗时：

| 流水行数 | 净计算耗时 |
|---|---|
| 813 行（约 2 年） | 2.59 ms |
| 4059 行（约 10 年） | 12.12 ms |

几秒的等待里，计算约占 0.05%。大头是 Apps Script 容器冷启动（0.5~3 秒，无法优化）、HTTPS 往返、以及 16 次 Sheets 读取往返。

**推论：把结果预先算好写进 Sheet 的 cell 并不能加速**——省下的是毫秒级计算，多出的每个 cell 读取却是一次真实往返。优化方向应该是「减少往返次数」和「首屏用缓存」，不是「减少计算」。

---

## 五、localStorage

### 配置（丢了要重设）

| 键 | 内容 | 丢失后果 |
|---|---|---|
| `fixedExpenses_v1` | 固定支出配置 | 要重设；**重设后 id 变化，key 跟着变，可能重复记账** |
| `incomeConfig_v1` | Payroll 设置 | 重设即可（key 防重，不会重复发） |
| `recurringTransfers_v1` | 定期转账配置 | 同上 |
| `k401_v1` / `property_v1` | 401k、房贷 | 要重填 |

### 缓存（丢了无影响，同步后覆盖）

| 键 | 内容 |
|---|---|
| `stockGroups_v2` / `brokerCash_v1` | 股票持仓、券商现金 |
| `hsaList_v1` | HSA（这一项目前也只有本地，尚未接 Sheet） |
| `sheetCache_v1` ⚠️ | 上次同步的 savings / bond / cash / ledger / monthly / ledgerMonths / notices |

`sheetCache_v1` 由 `saveSheetCache()` 在每次同步成功后写入，`loadSheetCache()` 在启动时注水。`setItem` 必须包 try/catch——隐私模式下会抛异常，不能中断启动。

---

## 六、界面结构

- **总览**：投资总额、四色配置条、五大类卡片（按金额排序）、转账按钮、退休、房产、自动事件通知条
- **股票**：按分类分组，点个股可增持/减持/改成本/删除；券商现金行点击可添加分红
- **储蓄 / 国债**：Sheet 驱动，可增删改；国债页底部有 `Next payment` 表（只列 T-Note）
- **现金**：收入 / 支出两个 tab
  - 收入：Payroll、Bonus、本月收入 / 净收入 / 对账、本月收入明细
  - 支出：11 个分类方块 + 本月合计，点方块看明细可改可删；`Bill & utilities` 等 9 类内嵌固定支出设置
- **转账**：两张余额卡并排；下面一个表单，顶部切换「一次性 / 定期」，共用 From/To/金额

### 支出分类（必须与 Shortcut 菜单完全一致）

```
Bill & utilities / Auto & Gas / Pet / Grocery & Household / Dining
Shopping / Health & Beauty / Travel & Entertainment / Home
Gifts & Families / Other or unexpected
```

固定支出可设在除 `Gifts & Families` 和 `Other or unexpected` 之外的 9 类里。

---

## 七、测试

10 套测试，**334 项断言**，用 jsdom + 假的 Sheets API 跑。**改代码后必须全部跑一遍。**

```bash
npm install     # 只需一次，唯一外部依赖是 jsdom
npm test        # 等价于 bash run-all.sh
```

全部通过时退出码为 0；有失败会打印失败项并以 1 退出。

| 文件 | 断言数 | 覆盖 | 读取的源文件 |
|---|---|---|---|
| `gastest.mjs` | 20 | Apps Script 各 action | `AppsScript.gs` |
| `savtest.mjs` | 20 | 储蓄计息、到期、改利率 | `AppsScript.gs` |
| `bondtest.mjs` | 34 | 国债付息、复利、到期、不补录历史 | `AppsScript.gs` |
| `e2etest.mjs` | 22 | **真实后端 → 真实前端**，防字段契约不一致 | 两者 |
| `test2.mjs` | 33 | 前端模块 | `index.html` |
| `test4.mjs` | 16 | 前端模块 | `index.html` |
| `test5.mjs` | 27 | 前端模块、定期转账迁移 | `index.html` |
| `test6.mjs` | 65 | 锚点现金模型 | 两者 |
| `uitest.mjs` | 54 | 界面布局与网络往返次数 | 两者 |
| `covtest.mjs` | 43 | 70 天窗口、整行回写、首屏缓存 | 两者 |

10 套全部通过 `readFileSync` 直接读仓库里的真实源文件，所以改动会立即被检验。测试只断言**公开行为**，不引用内部函数名——重构内部实现不会误伤。

> 教训：后端测后端、前端测前端各自通过，但**中间的接口从没对接过**，曾漏掉一次字段名不一致（后端返回 `bal`，前端读 `balance`），导致储蓄全显示 0。`e2etest.mjs` 就是为此加的。

### ⚠️ 已知覆盖盲区

原先 70 天窗口、首屏缓存、整行回写三处无断言，**已由 `covtest.mjs` 补上**（43 项）。

仍未覆盖 / 已知限制：

- **隐私模式下拿不到数据** —— 目前只有 `saveSheetCache` 包了 try/catch，
  `saveBrokerCash` / `persistStock` 等写入点没有。localStorage 一抛异常，
  `rebuildStockGroups` 就中断，整个同步落进 catch：界面起得来、徽标会显示
  「同步失败」，但一条数据都拿不到。`covtest.mjs` 末尾如实断言了这个当前行为，
  修好之后请把那条断言改成「仍能拿到数据」。
- **股票价格是上次同步的值**，不是实时行情；没有断言覆盖这个预期。

### 补充验证手段

这 9 套之外，本轮还用过以下方法，覆盖盲区时可以复用：

1. **Apps Script 假表测试台** —— 用 Node 造假的 `SpreadsheetApp` / `Utilities` / `LockService` 全局，`eval` 真实的 `AppsScript.gs`，即可对纯逻辑做断言，并统计每张表的读写往返次数。
2. **改动前后等价对比** —— 保留一份改动前的 `.gs`，用同一份数据分别跑 `readAllInner`，逐字段比对输出。本轮靠这个方法抓到过一个 `cash.balance` 差 $89 的回归。
3. **前端 stub 页面** —— 复制 `index.html`，在主 `<script>` 前插入一段替身脚本覆盖 `window.fetch` 并预置 localStorage，用本地 HTTP 服务打开（**不能用 `file://` 或 `data:`，localStorage 会被禁用**）。可验证首屏渲染、缓存注水、徽标状态。
4. **语法检查** —— `node --check`。前端需先把 `<script>` 内容抽出来。

### 排查工具（在 Apps Script 编辑器里手动运行）

| 函数 | 用途 |
|---|---|
| `selfTest` | 打印各表行数、字段完整性检查 |
| `cleanBackfill` | 删掉记账起始日之前被误补录的自动流水 |
| `cleanDuplicateKeys` | 删掉 key 重复的行 |
| `installTriggers` | 安装每日自动任务 |

---

## 八、代码中的暗礁

改动这些地方之前务必读完对应条目。

### 1. `writeSavRow` / `writeBondRow` 是整行回写

两者都从 `r.raw`（读取时留存的原始整行）复制，只覆盖自己负责计算的列，其余列原样带回。这是为了把 5 次 / 4 次 `setValue` 合并成 1 次 `setValues`。

**后果：绝不能在调用 `writeSavRow` 之前用 `setValue` 单独改某一列**——会被随后的整行回写覆盖掉。要改列，改 `r.raw[COL-1]`。`updateSaving` 就是这么做的。

特别注意 `Rate` 列：`readSavRows` 读出来的 `r.rate` 是归一化后的小数（`4.1` → `0.041`），而 `r.raw` 里是用户填的原值。回写走 `raw`，所以用户填的 `4.1` 不会被改写成 `0.041`。

### 2. 日期解析有两套，不能混用

- `s2d(v)` → `new Date("2026-08-01T00:00:00")`，按**本地**午夜解析
- `readAllInner` 里的 `new Date(d)`（当单元格是字符串时）→ 按 **UTC** 午夜解析

对文本格式的日期，两者**可能差一天**。`sumSignedRows` 特意逐字照搬 `sumLedgerSigned` 的 `s2d` + `d2s` 写法，就是为了不让现金余额在锚点边界上多算或漏算一行。

**v4 已统一**：`readAllInner` 里字符串日期也改走 `s2d`，与 `sumLedgerSigned` / `sumSignedRows` 同一套解析。非 ISO 文本（如 `8/1/2026`）仍用 `new Date()` 兜底，避免破坏原有的宽松解析。

实测（日期列为文本时）：`2026-08-01` 的支出正确计入 `2026-08`、`2026-07-31` 计入 `2026-07`，且 `monthly` 与 `cash.balance` 结果自洽。改之前这两条路径会在月份边界互相矛盾。

### 3. `ASOF` 的时区（v4 已修）

原先是 `const ASOF = new Date().toISOString().slice(0,10)`，取 **UTC** 日期。亚利桑那是 UTC-7，本地下午 5 点后右上角就显示成明天。实测旧公式在本地 17:30 / 23:45 都会跳到次日。

现在：初值用 `fmtD(todayD())`（本地当天），同步成功后换成后端返回的 `serverDate`（按 America/Phoenix 生成，权威值）。`serverDate` 后端一直在返回，此前前端从未使用。

### 4. `DATA` 里的硬编码数字

**国债和现金的种子已于 v4 清除**（与 `DATA.cd` 一致：宁可显示 $0，也不闪一下写死的旧数字）。实测首次安装的第一屏现在是 股票/储蓄/国债/现金 全部 $0，同步后才出数。

**但退休和 HSA 的数字保留，且这不是「过期缓存」**：

| | 可编辑 | 从 Sheet 同步 | 性质 |
|---|---|---|---|
| HSA `3458` / `2500` | ✅ `editSimple` | ❌ | 存 `hsaList_v1`，种子是初始值 |
| 401k `baseMv:102000` | ✅ | ❌ | 存 `k401_v1`，按 `rate` 本地滚动 |
| **Roth IRA `7515`** | ❌ | ❌ | **只能改代码才能更新** |

这三处的数字是你的真实资产，只是没有 Sheet 备份——删掉等于抹掉数据，不是修 bug。`Roth IRA` 最成问题：既不可编辑也不同步，值会一直停在写死的数字上。正解是接入 Sheet 或做成可编辑，见第九节第 1 条。

### 5. `apiWrite` 是全局串行队列

固定支出、工资、定期转账的补记全部经由同一条 `sheetQueue`。N 笔补记 = N 次串行的 Apps Script 往返。**因此绝不能把界面渲染排在补记链之后**——`syncStockFromSheet` 现在是「先渲染，补记放后台，写完再刷一次」。

---

## 九、已知缺口 / 待办

按优先级：

1. **HSA、退休、房贷仍在 localStorage** — 唯一还没有 Sheet 备份的部分（后端已返回 `hsa`/`retire`，前端未接）
2. **固定支出配置在 localStorage** — 丢了重设会导致 key 变化、可能重复记账，建议搬进 Sheet
3. **隐私模式下拿不到数据** — localStorage 抛异常会中断整个同步，写入点未全部包 try/catch，见第七节
4. **信用卡时间差** — 记账即扣款，算出的余额比银行 App 早约一个月，靠对账抹平
5. **股票分红只加券商现金，不进 Ledger** — 年底统计分红（1099-DIV）查不到
6. **储蓄利息不进 Ledger** — 净收入统计不含这块
7. **T-Bill 期限下拉只有整年** — 真实 T-Bill 是 4/8/13/26/52 周
8. **只有当月、上月或最近 70 天的流水能在小程序里改** — 更早的要去 Sheet 改，改完点「重算历史」
9. **`Roth IRA` 的 `7515` 既不可编辑也不同步** —— 只能改代码更新，见第八节第 4 条

---

## 十、沟通偏好

- 投资相关用中文，代码用英文
- 要量化论证：给具体金额和百分比，不要泛泛而谈
- 结论必须有数据支撑，不接受"应该差不多"
- 不要擅自增删需求，也不要在界面上暴露后台实现细节（如"写入 Sheet""GOOGLEFINANCE"）
- 改完代码要跑测试并报告通过数；跑不了的部分要明说跑不了，不要用「应该没问题」代替
- 区分「实测的」和「推断的」。性能数字尤其如此——没测过就说没测过
