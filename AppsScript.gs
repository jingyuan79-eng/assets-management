/**
 * 资产驾驶舱 · Google Apps Script 后端  v4  (2026-08-29)
 *
 * 部署：Extensions → Apps Script → 全选删除旧代码 → 粘贴本文件 → 保存
 *      → Deploy → Manage deployments → 铅笔 → Version 选 "New version" → Deploy
 * 部署设置：Execute as = Me，Who has access = Anyone
 * 时区：America/Phoenix（Arizona 不用夏令时）
 *
 * ── 改动记录（每次改代码请在最上方补一条，旧条目保留）──────────────
 *
 * v4  2026-08-29  清理与正确性
 *   · 修：monthly / expense / monthEnd 的日期解析改用 s2d（本地午夜），
 *         与 sumLedgerSigned 一致。原来用 new Date(字符串) 按 UTC 解析，
 *         文本格式的日期会在月份边界差一天。非 ISO 文本仍有兜底
 *   · 优化：runSavings / runBonds 改用 needId 标记，省掉每行 1 次
 *          「ID 是否为空」的读取（每次打开约 5 次读）
 *   · 删：readAnchor —— v3 重构留下的兼容包装，无人调用
 *   · 删：readAllInner 里 now / ym / prev / prevYm 的重复赋值
 *
 * v3  2026-08-27  性能（前端需同时上线，见下方 70 天窗口）
 *   · Anchor 整表只读一次：新增 readAnchorRows / anchorLatest /
 *     anchorBookStart，原先一次请求要读 4 遍
 *   · Ledger 全表只读一次：新增 sumSignedRows，computeCash 支持传入
 *     已读数据；原先 readAllInner 和 computeCash 各扫一遍全表
 *   · runSavings / runBonds 返回内存中已回写的 rows，不再整表重读
 *   · writeSavRow / writeBondRow 改为整行 setValues：5 次 / 4 次往返 → 各 1 次。
 *     不由本函数计算的列从 raw 原样带回（注意：调用前不可单独 setValue）
 *   · updateSaving 配合改为写 r.raw，4 次往返 → 1 次
 *   · ledger 明细窗口加宽到 70 天，盖过前端 60 天的补记判断窗口，
 *     消除每次打开重发已写过 key 的空请求（实测平均 4.5 次/打开 → 0）
 *   · 实测：平常日打开 表读取 22→16 写入 6→3；结息日 19→13 写入 21→6
 *
 * v2  2026-08-23  记账双向读写
 *   · Ledger 支持 改 / 删 / 按指定日期补记（小程序双向读写）
 *   · Ledger 增加 D:note（备注）、E:key（防重复键）两列
 *   · 固定支出自动记账用 key 去重，多设备打开也不会重复写
 *   · 写操作同时支持 GET（能拿到回执）和 POST（Shortcut 继续用 POST）
 *   · doGet 顺带返回 Savings / Bond / HSA / Retire 表（若存在），为后续接入预留
 */

var TZ = "America/Phoenix";

// 支出分类（与小程序、Shortcut 保持一致）；不在此列的归入「未匹配」
var CATS = ['Bill & utilities','Auto & Gas','Pet','Grocery & Household','Dining',
            'Shopping','Health & Beauty','Travel & Entertainment','Home',
            'Gifts & Families','Other or unexpected'];

// ==================== 入口 ====================

// Shortcut 用 POST；小程序也可用 POST（拿不到回执）
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    return jsonOut(handleAction(data));
  } catch (err) {
    return jsonOut({ status: "error", message: err.toString() });
  }
}

// 无 action 参数 = 读取全部；有 action = 执行写操作并返回回执
function doGet(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    if (p.action) return jsonOut(handleAction(p));
    return jsonOut(readAll());
  } catch (err) {
    return jsonOut({ status: "error", message: err.toString() });
  }
}

// ==================== 写操作总分发 ====================
// 加脚本锁：Apps Script 的 Web App 允许并发执行，没有锁的话
// 「先查 key 有没有、再写」这种读-判-写会被两个请求同时穿过，产生重复行。
function handleAction(data) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); }
  catch (e) { return { status: "error", message: "服务器忙，请稍后重试" }; }
  try { return handleActionInner(data); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

function handleActionInner(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = data.action || "ledger";

  switch (action) {
    // ---------- 记账 ----------
    case "ledger":      return addLedger(ss, data, true);   // Shortcut：日期用今天
    case "addLedger":   return addLedger(ss, data, false);  // 小程序补记：可指定日期
    case "autoLedger":  return autoLedger(ss, data);        // 固定支出：带 key 去重
    case "updateLedger":return updateLedger(ss, data);
    case "deleteLedger":return deleteLedger(ss, data);

    // ---------- 股票 ----------
    case "reconcile":   return reconcile(ss, data);
    case "recalc":      return { status: "success", type: "recalc", data: readAllInner(true) };
    case "adjustCash":  return adjustCash(ss, data);

    case "addBond":      return addBond(ss, data);
    case "updateBond":   return updateBond(ss, data);
    case "deleteBond":   return deleteBond(ss, data);

    case "addSaving":    return addSaving(ss, data);
    case "updateSaving": return updateSaving(ss, data);
    case "adjustSaving": return adjustSaving(ss, data);
    case "deleteSaving": return deleteSaving(ss, data);
    case "dismissNotice":return dismissNotice(ss, data);

    case "updateStock": return updateStock(ss, data);
    case "addStock":    return addStock(ss, data);
    case "deleteStock": return deleteStock(ss, data);
    case "updateCash":  return updateCash(ss, data);
  }
  return { status: "error", message: "unknown action: " + action };
}

// ==================== Ledger ====================
// 列：A:Date  B:category  C:amount  D:note  E:key
function ledgerSheet(ss) {
  var lg = ss.getSheetByName("Ledger");
  if (!lg) throw new Error("找不到 Ledger 表");
  // 首次运行时补上 D/E 两列表头（一次读、必要时一次写）
  var head = lg.getRange(1, 4, 1, 2).getValues()[0];
  if (head[0] === "" || head[1] === "") {
    lg.getRange(1, 4, 1, 2).setValues([[head[0] || "note", head[1] || "key"]]);
  }
  return lg;
}

function addLedger(ss, data, useToday) {
  var lg = ledgerSheet(ss);
  var date = useToday
    ? Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd")
    : (data.date || Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd"));
  var category = (data.category || "UNCATEGORIZED").toString().trim();
  var amount = parseFloat(data.amount) || 0;
  var note = (data.note || "").toString();
  lg.appendRow([date, category, amount, note, ""]);
  return { status: "success", type: "addLedger", row: lg.getLastRow(),
           date: date, category: category, amount: amount };
}

// 固定支出：同一个 key 只写一次（key 形如 "fix:网费:2026-08"）
function autoLedger(ss, data) {
  var key = (data.key || "").toString().trim();
  if (!key) return { status: "error", message: "autoLedger 缺少 key" };
  var lg = ledgerSheet(ss);
  var last = lg.getLastRow();
  if (last > 1) {
    var keys = lg.getRange(2, 5, last - 1, 1).getValues();   // 只读 E 列，不拉整表
    for (var i = 0; i < keys.length; i++) {
      if ((keys[i][0] || "").toString().trim() === key) {
        return { status: "skipped", type: "autoLedger", key: key, row: i + 2 };
      }
    }
  }
  SpreadsheetApp.flush();                     // 确保读到的是最新状态
  var date = data.date || Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
  lg.appendRow([date, (data.category || "Bill & utilities").toString().trim(),
                parseFloat(data.amount) || 0, (data.note || "").toString(), key]);
  SpreadsheetApp.flush();
  return { status: "success", type: "autoLedger", key: key, row: lg.getLastRow() };
}

function updateLedger(ss, data) {
  var lg = ledgerSheet(ss);
  var row = parseInt(data.row, 10);
  if (!(row > 1)) return { status: "error", message: "非法行号: " + data.row };

  // 安全校验：行号可能因为并发删除而错位，比对旧值确认改的是同一笔
  if (data.expectAmount != null) {
    var cur = parseFloat(lg.getRange(row, 3).getValue()) || 0;
    if (Math.abs(cur - parseFloat(data.expectAmount)) > 0.005) {
      return { status: "error", message: "数据已变化，请刷新后重试（行 " + row + "）" };
    }
  }
  // 一次性写 A:D 四列，比逐格 setValue 少三次往返
  var old = lg.getRange(row, 1, 1, 4).getValues()[0];
  var next = [
    data.date != null && data.date !== "" ? data.date : old[0],
    data.category ? data.category.toString().trim() : old[1],
    data.amount != null ? (parseFloat(data.amount) || 0) : old[2],
    data.note != null ? data.note.toString() : old[3]
  ];
  lg.getRange(row, 1, 1, 4).setValues([next]);
  return { status: "success", type: "updateLedger", row: row };
}

function deleteLedger(ss, data) {
  var lg = ledgerSheet(ss);
  var row = parseInt(data.row, 10);
  if (!(row > 1)) return { status: "error", message: "非法行号: " + data.row };
  if (data.expectAmount != null) {
    var cur = parseFloat(lg.getRange(row, 3).getValue()) || 0;
    if (Math.abs(cur - parseFloat(data.expectAmount)) > 0.005) {
      return { status: "error", message: "数据已变化，请刷新后重试（行 " + row + "）" };
    }
  }
  lg.deleteRow(row);
  return { status: "success", type: "deleteLedger", row: row };
}

// ==================== 流动现金（Chase）====================
// 模型：流动现金 = 最近一个锚点金额 + 锚点当天及之后的所有 Ledger 净额
// 锚点金额存的是「那天 00:00 的余额」，所以对账当天再补记的账也能正确计入。
// 分类前缀决定性质：Income · 加，Transfer · 减（不算支出），其余减（算支出）。
function catKind(cat) {
  var c = (cat || "").toString().trim();
  if (c.indexOf("Income") === 0) return "income";     // 真收入：工资、利息、分红
  if (c.indexOf("Redeem") === 0) return "redeem";     // 本金回流：加现金，但不是收入
  if (c.indexOf("Transfer") === 0) return "transfer"; // 投出去：减现金
  return "expense";
}
function catSign(cat) {
  var k = catKind(cat);
  return (k === "income" || k === "redeem") ? 1 : -1;
}

function anchorSheet(ss) {
  var sh = ss.getSheetByName("Anchor");
  if (!sh) { sh = ss.insertSheet("Anchor"); sh.getRange(1, 1, 1, 3).setValues([["date", "amount", "note"]]); }
  return sh;
}
// Anchor 表整表读一次。一次请求里 readAnchor / bookStart 会各要一遍，
// 分开读就是三四次往返，所以统一从这里出，调用方之间传数组复用。
function readAnchorRows(ss) {
  var sh = ss.getSheetByName("Anchor");
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    if (!v[i][0] || v[i][1] === "" || v[i][1] == null) continue;
    var d = s2d(v[i][0]);
    var amt = parseFloat(v[i][1]);
    if (!d || !isFinite(amt)) continue;
    out.push({ date: d, amount: amt, note: (v[i][2] || "").toString() });
  }
  return out;
}
// 取日期最晚的一条锚点
function anchorLatest(rows) {
  var best = null;
  for (var i = 0; i < rows.length; i++) {
    if (!best || rows[i].date >= best.date) best = rows[i];
  }
  return best;
}
// 记账起始日 = 最早一条锚点所在月的 1 号
function anchorBookStart(rows) {
  var first = null;
  for (var i = 0; i < rows.length; i++) {
    if (!first || rows[i].date < first) first = rows[i].date;
  }
  return first ? new Date(first.getFullYear(), first.getMonth(), 1) : null;
}
function addAnchorRow(ss, dateStr, amount, note) {
  anchorSheet(ss).appendRow([dateStr, Math.round(amount * 100) / 100, note || ""]);
}

// 对账：输入真实余额 → 算差额 →（可选）补一笔流水 → 落一个新锚点
function reconcile(ss, data) {
  var real = parseFloat(data.balance);
  if (!isFinite(real)) return { status: "error", message: "请输入有效余额" };
  var today = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
  var before = computeCash(ss);
  var diff = Math.round((real - before.balance) * 100) / 100;
  var mode = (data.mode || "adjust").toString();

  if (mode !== "adjust" && Math.abs(diff) > 0.005) {
    var lg = ledgerSheet(ss);
    var cat = (data.category || (diff < 0 ? "Other or unexpected" : "Income · 调整")).toString().trim();
    lg.appendRow([today, cat, Math.abs(diff), "对账差额", ""]);
    SpreadsheetApp.flush();
  }
  // 锚点存「当天 00:00 的余额」= 真实余额 − 当天已记的净额
  var todaySum = sumLedgerSigned(ss, today, today);
  addAnchorRow(ss, today, real - todaySum, data.note || ("对账 · 差 " + (diff >= 0 ? "+" : "") + diff));
  SpreadsheetApp.flush();
  return { status: "success", type: "reconcile", diff: diff, balance: real };
}

// 把 Ledger 里 [from, to] 区间的行按符号求和（空表示不限）
function sumLedgerSigned(ss, from, to) {
  var lg = ss.getSheetByName("Ledger");
  if (!lg || lg.getLastRow() < 2) return 0;
  var v = lg.getRange(2, 1, lg.getLastRow() - 1, 3).getValues();
  var sum = 0;
  for (var i = 0; i < v.length; i++) {
    if (!v[i][0]) continue;
    var d = s2d(v[i][0]); if (!d) continue;
    var ds = d2s(d);
    if (from && ds < from) continue;
    if (to && ds > to) continue;
    sum += catSign(v[i][1]) * (parseFloat(v[i][2]) || 0);
  }
  return Math.round(sum * 100) / 100;
}

// 记账起始日 = 最早的一条锚点。这之前发生的事都已经含在初始余额里，一律不补录。
function bookStart(ss) { return anchorBookStart(readAnchorRows(ss)); }

// sumLedgerSigned 的内存版：lv 是 getDataRange() 拿到的原始整表（含表头行）。
// 解析方式必须和 sumLedgerSigned 逐字一致 —— readAllInner 里另一处用的是
// new Date(字符串)（按 UTC 解析），和这里的 s2d（按本地午夜）差一天，
// 混用会让锚点边界上的流水被多算或漏算。
function sumSignedRows(lv, from, to) {
  var sum = 0;
  for (var i = 1; i < lv.length; i++) {
    if (!lv[i][0]) continue;
    var d = s2d(lv[i][0]); if (!d) continue;
    var ds = d2s(d);
    if (from && ds < from) continue;
    if (to && ds > to) continue;
    sum += catSign(lv[i][1]) * (parseFloat(lv[i][2]) || 0);
  }
  return Math.round(sum * 100) / 100;
}

// arows / lv 传进来就复用，不传就自己去读（reconcile 等单次写操作走这条）
function computeCash(ss, arows, lv) {
  var a = anchorLatest(arows || readAnchorRows(ss));
  if (!a) return { balance: 0, anchorDate: "", anchorAmount: 0, hasAnchor: false };
  var ad = d2s(a.date);
  var since = lv ? sumSignedRows(lv, ad, null) : sumLedgerSigned(ss, ad, null);
  return {
    balance: Math.round((a.amount + since) * 100) / 100,
    anchorDate: ad, anchorAmount: a.amount, anchorNote: a.note, hasAnchor: true
  };
}

// 券商现金按「增量」调整，避免用过期的本地值整体覆盖 Sheet
function adjustCash(ss, data) {
  var sh = ss.getSheetByName("Stock");
  var row = stockRowOf(sh, "Cash");
  if (row < 0) {
    row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 2).setValues([["Cash", "Gold.coin.cash"]]);
    sh.getRange(row, 5).setValue(0);
  }
  var cur = parseFloat(sh.getRange(row, 5).getValue()) || 0;
  var next = Math.round((cur + (parseFloat(data.delta) || 0)) * 100) / 100;
  sh.getRange(row, 5).setValue(next);
  return { status: "success", type: "adjustCash", amount: next };
}

// ==================== Savings（储蓄）====================
// 表结构 A:ID B:Name C:Type D:Balance E:Rate F:Last Post G:Next update H:Maturity I:Status
// 规则（全部在这里算，小程序只负责显示）：
//   · 余额只在「Next update」那天变，利息 = 从 Last Post 到当天的实际天数
//   · Rate 按 APY 处理，换算成日利率 (1+APY)^(1/365)-1，按日复利，年化正好等于 APY
//   · 任何中途变动（改利率 / 转入 / 转出）：先结息到当天 → Last Post 设为当天 → 再改
//   · 手动改余额：以手动值为准，不结息，只重设锚点
//   · CD 到期：结息到到期日 → 本息自动转入 OS 账户 → Status 置 closed
var SAV_COL = { ID:1, NAME:2, TYPE:3, BAL:4, RATE:5, POST:6, NEXT:7, MAT:8, STATUS:9 };

function savSheet(ss) {
  var sh = ss.getSheetByName("Savings");
  if (!sh) throw new Error("找不到 Savings 表");
  return sh;
}
function d2s(d) { return Utilities.formatDate(d, TZ, "yyyy-MM-dd"); }
function s2d(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return new Date(d2s(v) + "T00:00:00");
  var t = v.toString().trim();
  if (!t) return null;
  return new Date(t.slice(0, 10) + "T00:00:00");
}
function todaySv() { return new Date(d2s(new Date()) + "T00:00:00"); }
function firstOfNextMonth(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 1); }
function dayDiff(a, b) { return Math.round((b - a) / 86400000); }
// 4 / 4.1 / 0.04 都能识别：>1 视为百分数
function normRate(v) {
  var r = parseFloat(v) || 0;
  return r > 1 ? r / 100 : r;
}
// APY -> 日利率，按日复利，一年正好是 APY
function accrue(bal, apy, days) {
  if (days <= 0 || bal <= 0 || apy <= 0) return 0;
  var daily = Math.pow(1 + apy, 1 / 365) - 1;
  return bal * (Math.pow(1 + daily, days) - 1);
}

function readSavRows(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 9).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    if (!v[i][1]) continue;                       // 没名字的行跳过
    var st = (v[i][8] || "").toString().trim().toLowerCase();
    out.push({
      row: i + 2,
      raw: v[i].slice(0, 9),          // 原始单元格，回写时原样带上不归本函数管的列
      id: (v[i][0] || "").toString().trim(),
      name: v[i][1].toString().trim(),
      type: (v[i][2] || "OS").toString().trim().toUpperCase(),
      bal: parseFloat(v[i][3]) || 0,
      rate: normRate(v[i][4]),
      post: s2d(v[i][5]) || todaySv(),
      mat: s2d(v[i][7]),
      closed: (st === "closed" || st === "matured")
    });
  }
  return out;
}
// A:I 一次写完。逐格 setValue 是 5 次往返，合成一次能省 4 次。
// NAME / TYPE / RATE / MAT 四列不由本函数计算，从 raw 原样带回，绝不改动。
function writeSavRow(sh, r) {
  var out = r.raw.slice(0, 9);
  out[SAV_COL.ID - 1]     = r.id;
  out[SAV_COL.BAL - 1]    = Math.round(r.bal * 100) / 100;
  out[SAV_COL.POST - 1]   = d2s(r.post);
  out[SAV_COL.NEXT - 1]   = r.closed ? "" : d2s(firstOfNextMonth(r.post));
  out[SAV_COL.STATUS - 1] = r.closed ? "closed" : "active";
  sh.getRange(r.row, 1, 1, 9).setValues([out]);
}
function newSavId() { return "sv_" + Utilities.getUuid().slice(0, 8); }

// 结息到指定日期（不写表，只改内存对象）
function settleTo(r, when) {
  var days = dayDiff(r.post, when);
  if (days <= 0) return 0;
  var itr = accrue(r.bal, r.rate, days);
  r.bal += itr;
  r.post = when;
  return itr;
}

// 主引擎：补齐所有该入的月息 + 处理到期。可重复调用，结果一致。
function runSavings(ss) {
  var sh = savSheet(ss);
  var rows = readSavRows(sh);
  var today = todaySv();
  var notices = [];
  var osRow = null;
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i].closed && rows[i].type === "OS") { osRow = rows[i]; break; }
  }

  rows.forEach(function (r) {
    if (r.closed) return;
    // 先记下「原本就没有 ID」，下一行会把它填上。之前是靠回读单元格判断，
    // 每行多一次读；没到结息日时 changed 为 false，恰好每次打开都要付这笔钱。
    var needId = !r.id;
    if (needId) r.id = newSavId();
    var changed = false;

    // 按月推进：每次结到「下月 1 号」，不超过到期日
    while (true) {
      var nxt = firstOfNextMonth(r.post);
      if (r.mat && nxt > r.mat) break;
      if (nxt > today) break;
      settleTo(r, nxt);
      changed = true;
    }

    // 到期：最后一段按实际天数结息，本息转入 OS
    if (r.mat && today >= r.mat) {
      settleTo(r, r.mat);
      var payout = Math.round(r.bal * 100) / 100;
      r.closed = true;
      changed = true;
      if (osRow && osRow.id !== r.id) {
        settleTo(osRow, r.mat);            // OS 先结息到那一刻，再入账（锚点规则）
        osRow.bal += payout;
        osRow.post = r.mat;
        writeSavRow(sh, osRow);
        notices.push(r.name + " 已于 " + d2s(r.mat) + " 到期，本息 $" +
                     Math.round(payout).toLocaleString("en-US") + " 已转入 " + osRow.name);
      } else {
        notices.push(r.name + " 已于 " + d2s(r.mat) + " 到期，本息 $" +
                     Math.round(payout).toLocaleString("en-US") + "（未找到 OS 账户，请手动处理）");
      }
    }
    if (changed || needId) writeSavRow(sh, r);
  });

  notices.forEach(function (t) { addNotice(ss, t); });
  // rows 已经是回写后的状态，不必再整表读一遍
  return savOut(rows);
}

// 输出给小程序的格式（字段名必须和 index.html 的 rebuildSavings 一致）
function savOut(rows) {
  return rows.map(function (r) {
    return {
      id: r.id,
      name: r.name,
      type: r.type,
      balance: Math.round(r.bal * 100) / 100,
      rate: r.rate,                                  // 小数形式，如 0.034
      lastPost: r.post ? d2s(r.post) : "",
      nextUpdate: r.closed ? "" : d2s(firstOfNextMonth(r.post)),
      maturity: r.mat ? d2s(r.mat) : "",
      status: r.closed ? "closed" : "active"
    };
  });
}

// ---- 小程序发起的写操作 ----
function findSav(rows, id) {
  for (var i = 0; i < rows.length; i++) if (rows[i].id === id) return rows[i];
  return null;
}
function addSaving(ss, data) {
  var sh = savSheet(ss);
  var today = todaySv();
  var bal = parseFloat(data.balance) || 0;
  var id = newSavId();
  var row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 9).setValues([[
    id, (data.name || "").toString().trim(), (data.type || "CD").toString().toUpperCase(),
    Math.round(bal * 100) / 100, normRate(data.rate), d2s(today),
    d2s(firstOfNextMonth(today)), data.maturity || "", "active"
  ]]);
  // 从来源账户扣款（先结息再扣，不吃掉已滚利息）
  if (data.srcId) {
    var r = findSav(readSavRows(sh), data.srcId);
    if (r) { settleTo(r, today); r.bal -= bal; writeSavRow(sh, r); }
  }
  return { status: "success", type: "addSaving", id: id, savings: savOut(readSavRows(sh)) };
}
function updateSaving(ss, data) {
  var sh = savSheet(ss);
  var r = findSav(readSavRows(sh), data.id);
  if (!r) return { status: "error", message: "找不到账户 " + data.id };
  var today = todaySv();

  if (data.balance != null && data.balance !== "") {
    r.bal = parseFloat(data.balance) || 0;      // 手动改余额：以手动值为准，不结息
    r.post = today;
  } else if (data.rate != null && data.rate !== "" && normRate(data.rate) !== r.rate) {
    settleTo(r, today);                          // 改利率：先按旧利率结息，再换新利率
    r.post = today;
  }
  // 改到 raw 上，交给 writeSavRow 一次写完（分开 setValue 会被随后的整行回写覆盖）
  if (data.rate != null && data.rate !== "") r.raw[SAV_COL.RATE - 1] = normRate(data.rate);
  if (data.name) r.raw[SAV_COL.NAME - 1] = data.name.toString().trim();
  if (data.maturity != null) r.raw[SAV_COL.MAT - 1] = data.maturity || "";
  writeSavRow(sh, r);
  return { status: "success", type: "updateSaving", id: r.id, savings: savOut(readSavRows(sh)) };
}
// 转入/转出：先结息到今天，再加减，锚点重设
function adjustSaving(ss, data) {
  var sh = savSheet(ss);
  var r = findSav(readSavRows(sh), data.id);
  if (!r) return { status: "error", message: "找不到账户 " + data.id };
  var today = todaySv();
  settleTo(r, today);
  r.bal += parseFloat(data.delta) || 0;
  r.post = today;
  writeSavRow(sh, r);
  return { status: "success", type: "adjustSaving", id: r.id, savings: savOut(readSavRows(sh)) };
}
function deleteSaving(ss, data) {
  var sh = savSheet(ss);
  var r = findSav(readSavRows(sh), data.id);
  if (!r) return { status: "error", message: "找不到账户 " + data.id };
  sh.deleteRow(r.row);
  return { status: "success", type: "deleteSaving", id: data.id, savings: savOut(readSavRows(sh)) };
}

// 清理历史上重复写入的固定支出（同一个 key 只保留最早那行）
// 用法：在编辑器里手动运行一次 cleanDuplicateKeys
function cleanDuplicateKeys() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lg = ledgerSheet(ss);
  var last = lg.getLastRow();
  if (last < 2) { Logger.log("Ledger 为空"); return; }
  var keys = lg.getRange(2, 5, last - 1, 1).getValues();
  var seen = {}, dup = [];
  for (var i = 0; i < keys.length; i++) {
    var k = (keys[i][0] || "").toString().trim();
    if (!k) continue;
    if (seen[k]) dup.push(i + 2); else seen[k] = true;
  }
  dup.sort(function (a, b) { return b - a; });          // 从下往上删，行号才不会错位
  dup.forEach(function (r) { lg.deleteRow(r); });
  Logger.log(dup.length ? ("已删除 " + dup.length + " 行重复：原行号 " + dup.join(",")) : "没有发现重复");
}

// ==================== Bond（国债）====================
// 表结构 A:ID B:Name C:Type D:Start E:Term F:Rate G:Principal H:Balance I:LastPost J:Status
// 规则：
//   T-Note  每半年付息 = 本金 × 利率 ÷ 2 → 直接打进流动现金（Ledger 记 Income · 国债利息）
//   I-Bond  每半年复利 = 余额 ×(1 + 利率/2)，不产生现金；改利率只影响之后的复利
//   T-Bill  持有期不动，到期一次性拿回 本金 + 本金×利率×年数
//   到期：本金记 Redeem · 国债（加现金但不算收入），利息部分记 Income · 国债利息，Status → closed
var BOND_COL = { ID:1, NAME:2, TYPE:3, START:4, TERM:5, RATE:6, PRIN:7, BAL:8, POST:9, STATUS:10 };

function bondSheet(ss) {
  var sh = ss.getSheetByName("Bond");
  if (!sh) throw new Error("找不到 Bond 表");
  var h = sh.getRange(1, 9, 1, 2).getValues()[0];
  if (h[0] === "" || h[1] === "") {
    sh.getRange(1, 9, 1, 2).setValues([[h[0] || "LastPost", h[1] || "Status"]]);
  }
  return sh;
}
function addMonthsD(d, n) { var x = new Date(d.getFullYear(), d.getMonth() + n, d.getDate()); return x; }
function addYearsD(d, n) { return new Date(d.getFullYear() + n, d.getMonth(), d.getDate()); }

function readBondRows(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 10).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    if (!v[i][1]) continue;
    var st = (v[i][9] || "").toString().trim().toLowerCase();
    var start = s2d(v[i][3]);
    if (!start) continue;
    var term = parseInt(v[i][4], 10) || 1;
    var prin = parseFloat(v[i][6]) || 0;
    out.push({
      row: i + 2,
      raw: v[i].slice(0, 10),         // 原始单元格，回写时原样带上不归本函数管的列
      id: (v[i][0] || "").toString().trim(),
      name: v[i][1].toString().trim(),
      type: (v[i][2] || "T-Note").toString().trim(),
      start: start, term: term,
      rate: normRate(v[i][5]),
      principal: prin,
      bal: (v[i][7] === "" || v[i][7] == null) ? prin : (parseFloat(v[i][7]) || 0),
      post: s2d(v[i][8]) || start,
      closed: (st === "closed" || st === "matured"),
      mat: addYearsD(start, term)
    });
  }
  return out;
}
// A:J 一次写完（原本 4 次 setValue = 4 次往返）。
// NAME / TYPE / START / TERM / RATE / PRIN 不由本函数计算，从 raw 原样带回。
function writeBondRow(sh, r) {
  var out = r.raw.slice(0, 10);
  out[BOND_COL.ID - 1]     = r.id;
  out[BOND_COL.BAL - 1]    = Math.round(r.bal * 100) / 100;
  out[BOND_COL.POST - 1]   = d2s(r.post);
  out[BOND_COL.STATUS - 1] = r.closed ? "closed" : "active";
  sh.getRange(r.row, 1, 1, 10).setValues([out]);
}
function newBondId() { return "bd_" + Utilities.getUuid().slice(0, 8); }

// 往 Ledger 写一笔现金流，靠 key 防重复
function ledgerOnce(ss, key, dateStr, cat, amt, note) {
  if (!(amt > 0)) return false;
  var res = autoLedger(ss, { key: key, date: dateStr, category: cat, amount: amt, note: note });
  return res && res.status === "success";
}

// bookIn 传进来就复用（readAllInner 已经算过），不传就自己去读
function runBonds(ss, bookIn) {
  var sh = bondSheet(ss);
  var rows = readBondRows(sh);
  var today = todaySv();
  var book = (bookIn !== undefined) ? bookIn : bookStart(ss);   // 记账起始日；更早的一律跳过

  rows.forEach(function (r) {
    if (r.closed) return;
    var needId = !r.id;
    if (needId) r.id = newBondId();
    var changed = false;

    // 半年一次的付息 / 复利
    var guard = 0;
    while (guard++ < 200) {
      var next = addMonthsD(r.post, 6);
      if (next > r.mat || next > today) break;
      var ds = d2s(next);
      if (book && next < book) {                     // 记账开始之前：只推进进度，不补录
        r.post = next; changed = true; continue;
      }
      if (r.type === "I-Bond") {
        r.bal = r.bal * (1 + r.rate / 2);            // 用当下这一期的利率复利
      } else if (r.type === "T-Note") {
        var coupon = Math.round(r.principal * r.rate / 2 * 100) / 100;
        // 只写流水、不发通知：利息在「本月收入明细」里看得到，通知反而是噪音
        ledgerOnce(ss, "bond:" + r.id + ":" + ds, ds, "Income · 国债利息", coupon, r.name + " 半年息");
      }
      r.post = next;
      changed = true;
    }

    // 到期
    if (today >= r.mat) {
      var md = d2s(r.mat);
      if (book && r.mat < book) { r.closed = true; writeBondRow(sh, r); return; }  // 记账前就到期了
      var interest = 0, principal = r.principal;
      if (r.type === "I-Bond") { interest = Math.max(0, r.bal - r.principal); }
      else if (r.type === "T-Bill") { interest = Math.round(r.principal * r.rate * r.term * 100) / 100; }
      ledgerOnce(ss, "bondmat:" + r.id, md, "Redeem · 国债", principal, r.name + " 到期本金");
      if (interest > 0) {
        ledgerOnce(ss, "bondint:" + r.id, md, "Income · 国债利息", interest, r.name + " 到期利息");
      }
      addNotice(ss, r.name + " 已于 " + md + " 到期，本息 $" +
                Math.round(principal + interest).toLocaleString("en-US") + " 已存入流动现金");
      r.closed = true;
      changed = true;
    }
    if (changed || needId) writeBondRow(sh, r);
  });
  // rows 已经是回写后的状态，不必再整表读一遍
  return bondOut(rows);
}

function bondOut(rows) {
  return rows.map(function (r) {
    return {
      id: r.id, name: r.name, type: r.type,
      start: d2s(r.start), term: r.term, rate: r.rate,
      principal: Math.round(r.principal * 100) / 100,
      balance: Math.round(r.bal * 100) / 100,
      lastPost: d2s(r.post), maturity: d2s(r.mat),
      status: r.closed ? "closed" : "active"
    };
  });
}

function findBond(rows, id) {
  for (var i = 0; i < rows.length; i++) if (rows[i].id === id) return rows[i];
  return null;
}
function addBond(ss, data) {
  var sh = bondSheet(ss);
  var row = sh.getLastRow() + 1;
  var id = newBondId();
  var start = (data.start || d2s(todaySv())).toString().slice(0, 10);
  var prin = parseFloat(data.principal) || 0;
  var name = (data.name || "").toString().trim();
  var type = (data.type || "T-Note").toString().trim();

  // 防重复：同名 + 同起始日 + 同本金的活跃记录已存在就不再建
  var exist = readBondRows(sh);
  for (var i = 0; i < exist.length; i++) {
    if (!exist[i].closed && exist[i].name === name &&
        d2s(exist[i].start) === start && Math.abs(exist[i].principal - prin) < 0.01) {
      return { status: "error", message: "已存在同样的国债：" + name + "（" + start + "）" };
    }
  }
  // 进度直接推到「今天之前最近一次付息日」，历史利息不补录
  var post = s2d(start), today0 = todaySv(), guard0 = 0;
  while (guard0++ < 200) {
    var nx = addMonthsD(post, 6);
    if (nx > today0) break;
    post = nx;
  }
  sh.getRange(row, 1, 1, 10).setValues([[
    id, name, type, start, parseInt(data.term, 10) || 1, normRate(data.rate),
    prin, (data.balance != null && data.balance !== "") ? parseFloat(data.balance) : prin,
    d2s(post), "active"
  ]]);
  if (data.srcId) {                       // 从储蓄账户扣款
    var r = findSav(readSavRows(savSheet(ss)), data.srcId);
    if (r) { settleTo(r, todaySv()); r.bal -= prin; writeSavRow(savSheet(ss), r); }
  }
  return { status: "success", type: "addBond", id: id, bond: bondOut(readBondRows(sh)) };
}
function updateBond(ss, data) {
  var sh = bondSheet(ss);
  var r = findBond(readBondRows(sh), data.id);
  if (!r) return { status: "error", message: "找不到国债 " + data.id };
  if (data.name) sh.getRange(r.row, BOND_COL.NAME).setValue(data.name.toString().trim());
  if (data.type) sh.getRange(r.row, BOND_COL.TYPE).setValue(data.type.toString().trim());
  if (data.start) sh.getRange(r.row, BOND_COL.START).setValue(data.start);
  if (data.term) sh.getRange(r.row, BOND_COL.TERM).setValue(parseInt(data.term, 10) || 1);
  if (data.rate != null && data.rate !== "") sh.getRange(r.row, BOND_COL.RATE).setValue(normRate(data.rate));
  if (data.principal != null && data.principal !== "") sh.getRange(r.row, BOND_COL.PRIN).setValue(parseFloat(data.principal) || 0);
  // 余额手动填就以手动为准（比如 I-Bond 利率改晚了，自己按官网数字校正）
  if (data.balance != null && data.balance !== "") sh.getRange(r.row, BOND_COL.BAL).setValue(parseFloat(data.balance) || 0);
  return { status: "success", type: "updateBond", id: r.id, bond: bondOut(readBondRows(sh)) };
}
function deleteBond(ss, data) {
  var sh = bondSheet(ss);
  var r = findBond(readBondRows(sh), data.id);
  if (!r) return { status: "error", message: "找不到国债 " + data.id };
  sh.deleteRow(r.row);
  return { status: "success", type: "deleteBond", id: data.id, bond: bondOut(readBondRows(sh)) };
}

// 清理记账起始日之前被误补录的自动流水（国债利息、固定支出、工资、定期转账）
// 用法：在编辑器里手动运行一次 cleanBackfill
function cleanBackfill() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var book = bookStart(ss);
  if (!book) { Logger.log("Anchor 表还没有起始行，先填一行再运行"); return; }
  var bs = d2s(book);
  var lg = ledgerSheet(ss);
  var last = lg.getLastRow();
  if (last < 2) { Logger.log("Ledger 为空"); return; }
  var v = lg.getRange(2, 1, last - 1, 5).getValues();
  var del = [];
  for (var i = 0; i < v.length; i++) {
    var key = (v[i][4] || "").toString();
    if (!key) continue;                       // 手动记的账不动
    var d = s2d(v[i][0]);
    if (!d) continue;
    if (d2s(d) < bs) del.push(i + 2);
  }
  del.sort(function (a, b) { return b - a; });
  del.forEach(function (r) { lg.deleteRow(r); });
  Logger.log(del.length ? ("已删除 " + del.length + " 行 " + bs + " 之前的自动补录")
                        : ("没有发现 " + bs + " 之前的自动补录"));
}

// ==================== 通知（自动发生的事）====================
function noticeSheet(ss) {
  var sh = ss.getSheetByName("Notices");
  if (!sh) { sh = ss.insertSheet("Notices"); sh.getRange(1, 1, 1, 4).setValues([["id", "date", "text", "dismissed"]]); }
  return sh;
}
function addNotice(ss, text) {
  var sh = noticeSheet(ss);
  var last = sh.getLastRow();
  if (last > 1) {
    var ex = sh.getRange(2, 3, last - 1, 1).getValues();
    for (var i = 0; i < ex.length; i++) if ((ex[i][0] || "") === text) return;  // 同样的通知不重复
  }
  sh.appendRow(["nt_" + Utilities.getUuid().slice(0, 8), d2s(new Date()), text, ""]);
}
function readNotices(ss) {
  var sh = ss.getSheetByName("Notices");
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    if (!v[i][0] || v[i][3]) continue;
    out.push({ id: v[i][0], date: v[i][1] ? d2s(s2d(v[i][1])) : "", text: v[i][2] });
  }
  return out;
}
function dismissNotice(ss, data) {
  var sh = ss.getSheetByName("Notices");
  if (!sh) return { status: "success" };
  var v = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (v[i][0] === data.id) { sh.getRange(i + 2, 4).setValue("y"); break; }
  }
  return { status: "success", type: "dismissNotice", id: data.id };
}

// ==================== 每日定时任务（在编辑器里手动运行一次 installTriggers）====================
function dailyJob() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  runSavings(ss);
  runBonds(ss);
}
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "dailyJob") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("dailyJob").timeBased().atHour(3).everyDays(1).create();
  Logger.log("已安装每日 3 点自动计息任务");
}

// ==================== Stock ====================
function stockRowOf(sh, sym) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var col = sh.getRange(2, 1, last - 1, 1).getValues();      // 只读 A 列
  var target = (sym || "").toString().trim().toUpperCase();
  for (var i = 0; i < col.length; i++) {
    if ((col[i][0] || "").toString().trim().toUpperCase() === target) return i + 2;
  }
  return -1;
}

// 只改 C(股数) / D(成本)，绝不碰 E(GOOGLEFINANCE 公式)
function updateStock(ss, data) {
  var sh = ss.getSheetByName("Stock");
  var sym = (data.symbol || "").toString().trim();
  var row = stockRowOf(sh, sym);
  if (row < 0) return { status: "error", message: "not found: " + sym };
  sh.getRange(row, 3, 1, 2).setValues([[parseFloat(data.shares), parseFloat(data.cost)]]);
  return { status: "success", type: "updateStock", symbol: sym, row: row };
}

function addStock(ss, data) {
  var sh = ss.getSheetByName("Stock");
  var sym = (data.symbol || "").toString().trim();
  if (stockRowOf(sh, sym) > 0) return { status: "error", message: "already exists: " + sym };
  var newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, 4).setValues([[sym, data.category || "",
                                           parseFloat(data.shares), parseFloat(data.cost)]]);
  sh.getRange(newRow, 5).setFormula('=GOOGLEFINANCE(A' + newRow + ',"price")');
  return { status: "success", type: "addStock", symbol: sym, row: newRow };
}

function deleteStock(ss, data) {
  var sh = ss.getSheetByName("Stock");
  var sym = (data.symbol || "").toString().trim();
  var row = stockRowOf(sh, sym);
  if (row < 0) return { status: "error", message: "not found: " + sym };
  sh.deleteRow(row);
  return { status: "success", type: "deleteStock", symbol: sym };
}

// 券商现金：存放在 Cash 行的 E 列
function updateCash(ss, data) {
  var sh = ss.getSheetByName("Stock");
  var row = stockRowOf(sh, "Cash");
  if (row < 0) {
    row = sh.getLastRow() + 1;
    sh.getRange(row, 1).setValue("Cash");
    sh.getRange(row, 2).setValue("Gold.coin.cash");
  }
  sh.getRange(row, 5).setValue(parseFloat(data.amount) || 0);
  return { status: "success", type: "updateCash", amount: parseFloat(data.amount) || 0 };
}

// ==================== 读取 ====================
function readAll() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) {}
  try { return readAllInner(false); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

function readAllInner(force) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- Stock ----
  var sh = ss.getSheetByName("Stock");
  var sv = sh.getDataRange().getValues();
  var stock = [];
  for (var i = 1; i < sv.length; i++) {
    if (!sv[i][0]) continue;
    stock.push({ symbol: sv[i][0], category: sv[i][1],
                 shares: sv[i][2], cost: sv[i][3], price: sv[i][4] });
  }

  // ---- Ledger：当月汇总 + 近 70 天明细 + 每月总计 ----
  var expense = {};
  var ledger = [];
  var monthly = {};      // { "2026-08": {income, expense, transfer, payroll, bonus} }
  var allRows = [];      // 用于推算每个月末的流动现金
  var now = new Date(), ym = ymOf(now);
  var prev = new Date(now.getFullYear(), now.getMonth() - 1, 1), prevYm = ymOf(prev);
  // 明细窗口要盖过小程序 dueDates / bookFloor 的 60 天补记窗口，否则它查不到
  // 已经写过的 key，每次打开都会把旧的固定支出、工资重发一遍（服务器再逐条 skip）。
  var cutD = new Date(); cutD.setDate(cutD.getDate() - 70);
  var cutoff = d2s(cutD);
  var monthCats = {};    // { "2026-08": {Dining: 403.96, ...} }
  var lg = ss.getSheetByName("Ledger");
  if (lg) {
    var lv = lg.getDataRange().getValues();

    for (var j = 1; j < lv.length; j++) {
      var d = lv[j][0];
      if (!d) continue;
      // 注意：Apps Script 中 d instanceof Date 不可靠，必须用 toString.call
      // 字符串一律走 s2d（本地午夜），和 sumLedgerSigned / sumSignedRows 保持同一套
      // 解析。原先用 new Date(字符串) 是按 UTC 解析，文本日期会在月份边界差一天。
      var dateObj = (Object.prototype.toString.call(d) === '[object Date]') ? d : s2d(d);
      if (!dateObj || isNaN(dateObj.getTime())) dateObj = new Date(d);   // 非 ISO 文本兜底
      if (isNaN(dateObj.getTime())) continue;

      var ds = Utilities.formatDate(dateObj, TZ, "yyyy-MM-dd");
      var rowYm = ds.slice(0, 7);
      var cat = (lv[j][1] || "UNCATEGORIZED").toString().trim();
      var amt = parseFloat(lv[j][2]) || 0;

      // 每月统计：收入 / 支出 / 投资转出 分开算
      var kind = catKind(cat);
      if (!monthly[rowYm]) monthly[rowYm] = { income: 0, expense: 0, transfer: 0, redeem: 0, payroll: 0, bonus: 0 };
      var M = monthly[rowYm];
      if (kind === "income") {
        M.income += amt;
        if (cat.indexOf("Payroll") >= 0) M.payroll += amt;
        else if (cat.indexOf("Bonus") >= 0) M.bonus += amt;
      } else if (kind === "transfer") {
        M.transfer += amt;
      } else if (kind === "redeem") {
        M.redeem += amt;            // 本金回流：只影响现金，不进收入也不进支出
      } else {
        M.expense += amt;
        if (!monthCats[rowYm]) monthCats[rowYm] = {};
        var col = CATS.indexOf(cat) >= 0 ? cat : "未匹配";
        monthCats[rowYm][col] = (monthCats[rowYm][col] || 0) + amt;
        if (rowYm === ym) expense[cat] = (expense[cat] || 0) + amt;
      }
      allRows.push({ ds: ds, ym: rowYm, signed: catSign(cat) * amt });
      if (rowYm === ym || rowYm === prevYm || ds >= cutoff) {
        ledger.push({ row: j + 1, date: ds, category: cat, amount: amt,
                      note: (lv[j][3] || "").toString(), key: (lv[j][4] || "").toString() });
      }
    }
  }

  // 月末流动现金：从锚点起按日累加，锚点之前的月份留空（没有依据）
  var arows = readAnchorRows(ss);      // Anchor 整表只读这一次，下面全部复用
  var anchor = anchorLatest(arows);
  var monthEnd = {};
  if (anchor) {
    var ad = d2s(anchor.date);
    allRows.sort(function (a, b) { return a.ds < b.ds ? -1 : a.ds > b.ds ? 1 : 0; });
    var run = anchor.amount;
    var curYm2 = null;
    allRows.forEach(function (r) {
      if (r.ds < ad) return;
      if (curYm2 && r.ym !== curYm2) monthEnd[curYm2] = run;
      curYm2 = r.ym;
      run += r.signed;
      monthEnd[r.ym] = run;
    });
  }
  var cash = computeCash(ss, arows, lv);   // 复用已读的 Anchor 和 Ledger 原始数据
  var bs = anchorBookStart(arows);
  cash.bookStart = bs ? d2s(bs) : "";
  writeMonthly(ss, monthly, monthCats, monthEnd, !!force);

  // 明确告诉小程序：ledger 里这几个月的明细是完整的，可以据此重算总计
  var ledgerMonths = [ym, ymOf(prev)];
  var monthlyOut = {};
  Object.keys(monthly).forEach(function (k) {
    monthlyOut[k] = { income: r2(monthly[k].income), expense: r2(monthly[k].expense),
                      net: r2(monthly[k].income - monthly[k].expense),   // 净收入：不含转账与回流
                      transfer: r2(monthly[k].transfer), redeem: r2(monthly[k].redeem),
                      endCash: (monthEnd[k] == null ? null : r2(monthEnd[k])) };
  });
  return { status: "success", stock: stock, expense: expense, ledger: ledger,
           monthly: monthlyOut, cash: cash, ledgerMonths: ledgerMonths,
           savings: safeRun(function () { return runSavings(ss); }, []),
           notices: safeRun(function () { return readNotices(ss); }, []),
           bond: safeRun(function () { return runBonds(ss, bs); }, []),
           hsa: readTab(ss, "HSA"), retire: readTab(ss, "Retire"),
           serverDate: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd") };
}

// 把每月汇总写进 Ledger_monthly（内容没变就不写，避免每次打开都产生一次写操作）
function writeMonthly(ss, monthly, monthCats, monthEnd, force) {
  var sh = ss.getSheetByName("Ledger_monthly") || ss.getSheetByName("Ledger-monthly");
  if (!sh) sh = ss.insertSheet("Ledger_monthly");
  var head = ["年月", "收入", "支出", "净收入", "投资转出", "本金回流", "月末流动现金",
              "Payroll", "Bonus"].concat(CATS).concat(["未匹配"]);
  var rows = [head];
  Object.keys(monthly).sort().forEach(function (ym) {
    var m = monthly[ym];
    var c = monthCats[ym] || {};
    var r = [ym, r2(m.income), r2(m.expense), r2(m.income - m.expense),
             r2(m.transfer), r2(m.redeem),
             (monthEnd[ym] == null ? "" : r2(monthEnd[ym])),
             r2(m.payroll), r2(m.bonus)];
    CATS.forEach(function (k) { r.push(r2(c[k] || 0)); });
    r.push(r2(c["未匹配"] || 0));
    rows.push(r);
  });
  if (!force) {
    var cur = sh.getDataRange().getValues();
    if (JSON.stringify(cur) === JSON.stringify(rows)) return;
  }
  sh.clear();
  sh.getRange(1, 1, rows.length, head.length).setValues(rows);
}
function r2(n) { return Math.round((n || 0) * 100) / 100; }

function ymOf(d) { return Utilities.formatDate(d, TZ, "yyyy-MM"); }

// ==================== 月度支出汇总 ====================
// 把 Ledger 全部流水按年月汇总，写进 Monthly 表并返回 {"2026-08":4123.45, ...}
// 小程序据此按月扣现金，跨月不打开也不会漏扣。
// 通用读表：第一行当表头，返回对象数组。表不存在或为空时返回 []
function readTab(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var head = v[0].map(function (h) { return (h || "").toString().trim(); });
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    var o = { row: i + 1 };
    for (var c = 0; c < head.length; c++) {
      if (!head[c]) continue;
      var cell = v[i][c];
      if (Object.prototype.toString.call(cell) === '[object Date]') {
        cell = Utilities.formatDate(cell, TZ, "yyyy-MM-dd");
      }
      o[head[c]] = cell;
    }
    out.push(o);
  }
  return out;
}

// ==================== 工具 ====================
// 单个模块出错不该拖垮整个 doGet（否则股票、支出会一起读不出来）
function safeRun(fn, fallback) {
  try { return fn(); }
  catch (e) { Logger.log("模块出错，已跳过: " + e.toString()); return fallback; }
}
// 把返回对象序列化成 JSON 响应。所有 doGet / doPost 出口都依赖它。
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== 自检（在编辑器里手动运行一次）====================
function selfTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = readAll();
  Logger.log("Stock 行数: " + r.stock.length);
  Logger.log("当月类别: " + JSON.stringify(r.expense));
  Logger.log("可编辑流水: " + r.ledger.length + " 条");
  Logger.log("月度汇总: " + JSON.stringify(r.monthly));
  var sv = r.savings || [];
  Logger.log("Savings: " + sv.length + " 个账户");
  if (sv.length) {
    Logger.log("  首个账户字段: " + JSON.stringify(sv[0]));
    var need = ["id","name","type","balance","rate","lastPost","nextUpdate","maturity","status"];
    var miss = need.filter(function (k) { return !(k in sv[0]); });
    Logger.log(miss.length ? "  ❌ 缺少字段: " + miss.join(",") : "  ✅ 字段完整");
  }
  Logger.log("Bond/HSA/Retire: " + r.bond.length + "/" + r.hsa.length + "/" + r.retire.length);
}
