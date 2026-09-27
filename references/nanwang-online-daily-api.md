
# 南网在线 H5 每日电费查询接口笔记（已验证）

> 状态（2026-09-26）：登录闭环、单户查询、自动抓取脚本、飞书提醒均已跑通并部署定时任务。户号绑定双通道均已实测打通：「手机号查询」单次勾选 10 户绑定成功；「用户编号+代扣银行卡后6位」逐户绑定 6 户防溺水异动电表全部成功。**单账号上限 26 户已确认**；已批量解绑 15 户电信非防溺水户，当前账号保留 11 户（9 个防溺水监控 + 2 个居民私户）。**meteringPointId 可传空字符串**（无需浏览器 UI 收集）。**unBindElectricityUsers 接口已跑通**（批量解绑，参数 `{bindingIds:[...]}`）。防溺水 894 户统一代扣银行卡后6位=081121。

## 需求与目标形态

- 输入：本机供电清单 Excel 第一列"用户编号"（16 位文本，如 `0308000018184408`；202608 月清单约 1706 个唯一编号）
- 输出：每户每日电量（千瓦时）+ 估算电费，落盘 Excel
- 形态：Python 脚本 + 每日定时任务（已部署，每日 09:00）
- 账号：用户已有南网在线账号（手机号 13377597528，钟*凤；2026-09-26 两轮绑定后共 24 户 = 2 居民户 + 16 个电信站址户 + 6 个防溺水异动电表）

## 平台与认证机制（已实测确认）

- H5 端：`https://95598.csg.cn`（Vue 2.6.10 + axios 0.19 单页应用；登录页 `#/sz/login/login`）
- 请求风格：POST + JSON；响应格式 `{"sta":"00","message":"success","data":...}`，`sta="00"` 表示成功，`sta="04"` 表示未登录/token 失效
- 认证头：`x-auth-token`（登录后下发，存储在 `sessionStorage.token`）；另需 `custNumber` 头（值在 `sessionStorage.custnumber`）
- 登录方式：手机号+短信验证码（验证码按钮 120 秒冷却）；登录后 token 存 sessionStorage，cookie 设 `is-login=true`
- 首页有免登录"电费查询"快捷入口：用户编号 + 预留手机号 + 验证码 → 立即查询（适合单户抽查，不适合批量自动化）

## 已验证接口清单

| 用途 | 接口 | 加密 | 说明 |
|---|---|---|---|
| 已绑定户号列表 | `POST /eleCustNumber/queryBindEleUsers` | 否（明文） | **获取绑定户号的核心接口**，body 传 `{}`，返回含 bindingId/eleCustNumber/areaCode/eleAddress/userName 等 |
| 每日电量+温度（广东） | `POST /charge/queryDayElectricAndTemperature` | 否（明文） | **湛江/廉江等广东电网用户用这个**，body 传 `{areaCode, eleCustId, yearMonth, meteringPointId}` |
| 每日电量明细（按计量点） | `POST /charge/queryDayElectricByMPoint` | 是（need-crypto） | 页面日历加载时调用，请求体和响应均加密 |
| 批量每日电量 | `POST /charge/queryDayElectricList` | 是（need-crypto） | 批量查询多户，请求体 param 密文，响应 data 密文 |
| 计量点信息 | `POST /charge/queryMeteringPoint` | 是（need-crypto） | 获取 meteringPointId，body 传 `{bindingId, areaCode}`，但需加密头 |
| 用电日历（深圳） | `POST /charge/queryElectricityCalendar` | 否 | 前端函数名 queryDayElectricAndTemperatureSz，仅深圳 |
| 电费账单 | `POST /charge/selectElecBill`、`selectElecBillDetails` | 否 | 月账单与明细，另有 SZ/GD 分省版本 |

### 核心查询链路（已跑通）

```
1. 登录 → sessionStorage.token + sessionStorage.custnumber
2. POST /eleCustNumber/queryBindEleUsers {} → 绑定户号列表（含 bindingId）
3. 对每户：POST /charge/queryDayElectricAndTemperature
   {areaCode:"030800", eleCustId:"<bindingId>", yearMonth:"202609", meteringPointId:"<pointId>"}
   → 返回 [{date:"-09-24", power:"29.15", minTemperature, maxTemperature, averageTemperature}, ...]
```

### 关键参数说明

- **eleCustId** = 绑定户号列表中的 `bindingId`（如 `J6CFGiYb`），不是用户编号
- **meteringPointId**（即 pointId）= 计量点内部 ID，需通过浏览器 UI 切换户号获取（queryMeteringPoint 接口加密，Python 直调返回 sta=04）
- **areaCode** = `030800`（湛江地区代码）
- **yearMonth** = `202609` 格式
- **date** 响应格式 = `-09-24`（需拼接年份：`yearMonth[:4] + date` → `2026-09-24`）

### meteringPointId 获取方法（已验证）

queryMeteringPoint 接口带 `need-crypto` 加密（AES-128-CBC + ZeroPadding），Python 无法直接调用。获取方式：

1. Playwright 打开 `https://95598.csg.cn/#/gd/fee/feeService/calendar`
2. 点击"切换"按钮打开户号选择弹窗
3. 依次点击每个户号的 radio label（`ant-modal-wrap` 内的 `label` 元素，按 eleCustNumber 文本匹配）
4. 等待 3 秒后读取 Vuex store `currentUserNumInfo.pointId`
5. 收集所有户号的 pointId 存入 config.json 映射表

```javascript
// Playwright 自动化收集 pointId 示例
for (const num of telecomNumbers) {
  await page.evaluate(() => {
    const btns = [...document.querySelectorAll('button')].filter(b => b.textContent.includes('切换'));
    const modal = document.querySelector('.ant-modal-wrap');
    if (!modal && btns.length) btns[0].click();
  });
  await page.waitForTimeout(1200);
  await page.evaluate((num) => {
    const modal = document.querySelector('.ant-modal-wrap');
    const labels = [...modal.querySelectorAll('label')];
    const el = labels.find(l => l.textContent.includes(num));
    if (el) el.click();
  }, num);
  await page.waitForTimeout(3000);
  const state = await page.evaluate(() => {
    const app = document.querySelector('#app');
    const store = app.__vue_app__?.config.globalProperties.$store || app.__vue__?.$store;
    return { pointId: store.state.currentUserNumInfo.pointId, eleCustNumber: store.state.currentUserNumInfo.eleCustNumber };
  });
  // state.pointId 即为该户号的 meteringPointId
}
```

## Token 管理与续期

- token 存储在浏览器 `sessionStorage.token`，`custNumber` 在 `sessionStorage.custnumber`
- Python 脚本从 config.json 读取 token/custNumber，直接用 urllib 调用明文接口
- token 失效时接口返回 `sta=04`，脚本检测后弹出 Windows 桌面提醒（PowerShell MessageBox）
- 续期方法：浏览器打开 95598.csg.cn 登录 → F12 Console 执行 `copy(sessionStorage.getItem('token') + '|' + sessionStorage.getItem('custnumber'))` → 粘贴到 `nw_token_update.py` 更新 config.json；Playwright 会话中也可直接从网络面板任意请求的 `x-auth-token` 请求头提取新 token 写入 config.json（2026-09-26 两次续期均用此法）
- token 有效期未精确测定：Python 侧 token 实测跨日有效（9/26 00:02 定时任务抓取成功）；浏览器登录态存于 sessionStorage，浏览器重启即丢失，次日 UI 操作需重新短信登录属正常现象而非 token 过期

## 已部署自动化脚本

脚本目录：`<南网每日电费工作目录>\`

| 文件 | 说明 |
|---|---|
| `nw_daily_fetch.py` | 主抓取脚本：token 检测 → 拉绑定户号 → 逐户查每日电量 → 落盘 Excel + 飞书提醒 |
| `config.json` | 配置文件：token、custNumber、areaCode、飞书 webhook、提醒阈值、telecom_users 映射表（户号↔bindingId↔pointId） |
| `nw_token_update.py` | Token 过期后的快速更新工具（3 步操作） |

### 脚本核心逻辑

1. **token 检测**：调 `queryBindEleUsers`，sta=04 则弹窗提醒 + 退出
2. **抓取**：对 config.json 中 telecom_users 逐户调 `queryDayElectricAndTemperature`
3. **落盘**：openpyxl 写 Excel（无 openpyxl 则降级 CSV），表头：用户编号/用电地址/日期/用电量(千瓦时)/估算电费(元)/最低温/最高温/均温
4. **飞书提醒**：只看最新日期数据，日电量 > 阈值 或 日估算电费 > 阈值 → 推送飞书卡片（红色 header + 明细）
5. **防重复**：`.last_alert_YYYY-MM-DD.json` 标记文件，同日已推送则跳过
6. **月初容错**：本月数据未发布（T+3）时自动回退抓上月
7. **电价估算**：默认 0.696 元/度，可用 `--price` 参数或 config 覆盖

### 定时任务

- Windows 计划任务名：`南网每日电费抓取`
- 触发：每日 09:00
- 命令：`python nw_daily_fetch.py`
- 工作目录：`<南网每日电费工作目录>`

### 飞书提醒配置

- webhook：复用投资系统飞书机器人 `https://open.feishu.cn/open-apis/bot/v2/hook/9a3f0eb9-...`
- 阈值：config.json `alert_threshold.daily_fee`（默认 1.5 元）、`alert_threshold.daily_power`（默认 1.5 度）
- 推送格式：interactive 卡片，红色 header "⚡ 南网电费超标提醒"，含触发户数统计 + 每户明细

## 防溺水监控电表清单

- 数据文件：`<打标台账目录>\防溺水电表898唯一打标.xlsx`（Sheet: "防溺水电表898个"）
- 另有副本：`<防溺水台账目录>\防溺水电表898唯一打标.xlsx`（894 行，更干净）
- **防溺水监控电表 = 894 个**（电表用途列 = "防溺水监控"，用户编号为 15 位 `3082...`，补前导 0 后为 16 位 `030820...`）
- 文件中另有 381 个 16 位编号为基站/治安监控等非防溺水用途，不在本次范围
- 供电单位分布：良垌 211、石岭 179、河唇 129、石城 101、吉水 94、新民 90、雅塘 80 等
- 已提取标准清单：`<南网每日电费工作目录>\防溺水894用户清单.json`（用户编号已补零 16 位 + 出厂编号 + 用电地址 + 供电单位），批量绑定与抓取范围扩展直接复用

### 批量绑定实测（2026-09-26）

**「手机号查询」批量绑定已跑通**：户号绑定页（`#/gd/my/my/houseNumBinding`）选「手机号查询」→ 立即查询 → 列出登录手机号名下全部户号（已绑+未绑，带勾选框）→ 勾选未绑户号 → 立即绑定。实测单次勾选 10 户绑定成功（账号 8→18 户），适合同一预留手机号下户号的一次性批量绑定。

**绑定方式对比（含实测结论）**：
- 「手机号查询」：查的是登录手机号名下的户号，不是按用户编号绑定；实测该手机号名下仅 16 个电信站址/基站户，**防溺水 894 户均不在其中**，绑定防溺水户须先确认其在供电局登记的预留手机号
- 「用户编号验证」：仅支持"代扣银行卡后 6 位"验证。**2026-09-26 实测打通**：用户提供统一代扣银行卡后 6 位后，6 户防溺水异动电表逐一绑定全部成功——只要掌握银行卡后 6 位该通道即可用，操作流程见下节
- 首页「电费查询」：用户编号 + 预留手机号 + 短信验证码，免绑定直查，但每户每次需验证码，不适合每日自动化

### 「用户编号+代扣银行卡后6位」逐户绑定流程（Playwright 实测打通 2026-09-26）

适用场景：目标户号不在登录手机号名下（如防溺水监控电表），但掌握其代扣银行卡后 6 位（电信户多为公司统一代扣卡）。实测 6/6 户绑定成功，每户约 30 秒。

单户操作步骤（Playwright）：
1. 导航到 `https://95598.csg.cn/#/gd/my/my/houseNumBinding`（每户绑定完成后须重新导航整页）
2. 点击「用户编号 使用用户编号验证」切换验证方式（定位含"使用用户编号验证"文本的元素，点其 parentElement）
3. 用 Playwright fill 填「请输入用户编号」（16 位补零）与「请输入代扣银行卡后6位」
4. 点「立即查询」→ 出现"查询到1个结果"，核对户名/地址
5. 勾选结果行复选框 → 点「立即绑定」→ 出现"已绑户号"成功页 → 点「完 成」
6. 重新导航回绑定页，开始下一户

关键细节：
- 查询结果带"已绑定"标记的户直接跳过，不重复操作
- 绑定完成后调 `queryBindEleUsers` 一次拿全所有新户的 bindingId（实测账号 18→24 户）；pointId 仍需浏览器 UI 切换户号收集
- **Vue 表单必须用 Playwright fill 填值**：JS 原生 setter（`Object.getOwnPropertyDescriptor` 的 set + dispatchEvent）赋值后界面显示有值，但 Vue 数据模型未更新，点「立即查询」不发出任何请求（network 面板无 queryBindInfo）；从结果页点「返 回」后表单状态也可能残留，须重新导航
- 页面偶发 `ant-message-loading` 加载提示遮挡目标元素致 Playwright click 超时，改用 evaluate 直接 `el.click()` 绕过

### 批量绑定 894 户可行性评估（2026-09-26）

- 单户约 30 秒，894 户约需 7.5 小时不间断操作；单账号绑定上限未知（实测 24 户正常）
- 浏览器登录态存 sessionStorage，中途过期需重新短信登录（登录陷阱见上）
- 建议路径：① 按月优先绑定异动户（8 月异动 6 户已完成），逐步扩绑；② 若确认 894 户统一代扣银行卡，可写 Playwright 连续绑定脚本（保持登录态、跳过已绑定户）

**推荐路径**（2026-09-26 银行卡通道打通后更新）：
1. 首选「用户编号+代扣银行卡后6位」逐户绑定（已实测打通），按月优先绑异动户，逐步扩至 894 户
2. 若确认 894 户统一代扣银行卡，可自动化连续绑定（约 30 秒/户）
3. 绑定后调 queryBindEleUsers 拿 bindingId，再用 Playwright 收集 pointId，加入 config.json 映射表
4. 脚本自动扩展抓取范围
5. 规模化根本解法仍是企业客户/集团客户通道（南网集团电费集中支付、企业网厅或与供电局对接数据接口），个人账号绑定仅作过渡

### 单账号绑定上限（实测 2026-09-26）

- **单账号最多绑定 26 户**（`bindPopup/queryBindInfo` 返回 `"你已绑定N户，最多可选择26户"`）
- 实测账号从 24 户绑到 26 户后无法继续绑定；894 户全量绑定需约 35 个账号或多轮解绑/重绑
- 当前账号 13377597528 绑定 11 户（9 个防溺水监控 + 2 个居民私户），已解除 15 个电信非防溺水户

### 户号解绑（实测 2026-09-26）

**接口**：`POST /eleCustNumber/unBindElectricityUsers`，明文，body 传 `{bindingIds: ["id1","id2",...]}`

- 支持批量解绑：一次传多个 bindingId 即可全部解绑（实测 14 户一次性解绑成功）
- **必须使用最新 bindingId**：bindingId 在每次登录/token 刷新后可能变化，解绑前必须先调 `queryBindEleUsers` 获取当前有效的 bindingId；使用过期 bindingId 返回 `sta:02, "系统异常，请重试"`
- UI 无解绑入口：管理模式（个人中心→管理）仅有"添加到分组/分组管理/设为置顶"，无"解除绑定"按钮；解绑只能通过 API 调用
- 接口发现方法：从 chunk-25f15f87.js 中找到 `batchUnBind` 函数调用 `unBindElectricityUsers({bindingIds:n})`

### bindingId 不稳定性（实测 2026-09-26）

- bindingId 在 token 刷新/重新登录后会变化（如 `J6CFGiYb` → `goDmteyE` → 解绑后不存在）
- **任何绑定/解绑/查询操作前，必须先调 `queryBindEleUsers` 获取最新 bindingId**，不可复用 config.json 中的旧值
- config.json 中的 bindingId 仅作参考，脚本运行时应动态从接口获取最新值

## 探查方法（受限网络下提取 SPA 接口清单）

1. Playwright 打开目标站，先看网络请求面板已发生的 XHR，确认请求头（x-auth-token）与响应格式
2. 本机 `Invoke-WebRequest` 直接下载 JS 包失败（网络受限）时，改用 **Playwright evaluate 在页面上下文内 `fetch('/js/app.xxx.js')`**（同源请求不受限），再：
   - 正则 `url:"/xxx"` 提取全量接口路径
   - 按关键词（calendar/daily/electricRecord/login）取上下文，定位参数与前端函数名
3. JS 文件名带构建版本号（如 `app.1.6.230.1788437401584.js`），从网络面板静态请求列表获取
4. 加密接口识别：JS 中带 `headers:{"need-crypto":!0}` 的接口请求体和响应体均加密（AES-128-CBC + ZeroPadding，密钥来自 webpack 模块 47e3）

## 登录流程陷阱（实测 2026-09-26）

- **协议复选框必须先勾选再点登录**：登录页有"用户协议"复选框，若未勾选直接点"登录"会弹出"需要您进行授权"弹窗，弹窗关闭后验证码可能已过期。正确流程：填手机号 → 获取验证码 → 填验证码 → **勾选协议复选框** → 点登录。实测因未先勾选导致 3 次验证码浪费。
- **验证码有效期短**：从获取到提交窗口约 60 秒（按钮显示倒计时），协议弹窗打断后验证码极易过期，返回 `sta:0001003, "验证码校验失败"`。
- **地区默认非湛江**：登录后首页地区可能显示"佛山市"等默认值，须手动"切换地区"到湛江市。

## sendMsg 接口行为辨析（实测 2026-09-26）

- `POST /center/sendMsg` 请求体仅含 `{areaCode, phoneNumber, smsMsgType, vcType, infoType}`，**不含用户编号**，因此**不校验手机号是否为该户号的预留手机号**——只要手机号格式正确就会发送短信。
- 真正的预留手机号校验发生在 `POST /charge/queryChargesWithCode`（首页"电费查询"入口的查询接口）：若填写的手机号与户号预留手机号不匹配，返回 `sta:02, "请输入供电局预留的手机号"`。
- **结论**：sendMsg 成功 ≠ 预留手机号匹配；不能以"收到验证码"判断该手机号可绑定/查询该户号。

## 防溺水 894 户户名归属（实测 2026-09-26）

从原始电费清单 `电信电量电费202608.xlsx`（含"用户名称"列）匹配 894 户，结果：
- 613 户用户名称 = "中国电信股份有限公司湛江分公司"
- 87 户 = "中国电信股份有限公司湛江分公司(防溺水监控)"
- 84 户 = "中国电信股份有限公司湛江分公司（监控）"
- 68 户 = "中国电信股份有限公司廉江分公司"
- 其余为上述名称的括号/后缀变体
- **全部 894 户户名均为中国电信**，理论上可用营业执照走对公认证绑定，绕开手机号限制

## 日电量历史数据可用范围（实测 2026-09-26）

- `queryDayElectricAndTemperature` 可查到 **2025 年 4 月起**的日级数据（2025-01~03 返回 `sta:09, "无此用户数据"` 或超时）
- 即日级历史最多约 17 个月，更早的历史仅有月度清单数据
- 月度清单（`合并后的电费数据.xlsx`）覆盖 2022 年 12 月至今，可用于长周期趋势分析

## 未绑定户号的异动分析方法（实测 2026-09-26）

当日级数据因户号未绑定无法获取时，可用月度电费清单做异动分析：
1. 从 `合并后的电费数据.xlsx`（58229 行，含 2022-12 至今全部月度记录）中按用户编号提取目标户
2. 计算每户月度电量/电费序列，取前 6 个月中位数为"基线月均"
3. 环比增长率 = (当月 - 上月) / 上月 × 100%
4. 基线增长率 = (当月 - 基线月均) / 基线月均 × 100%
5. 突变判定：基线增长率 ≥ 50% 为严重异动
6. 实测 6 户防溺水异动电表：2025-01~2026-06 月均 8~25 度平稳运行，2026-07 起突然飙升（最高增幅 +13,273%），8 月合计异增电费 3,649.82 元

## 关键坑位

- **地区默认深圳**：H5 顶栏默认"深圳市"，湛江用户须"切换地区"到广东省，否则路由到深圳专用接口
- **广东/深圳接口分叉**：同一功能两套端点（用电日历最典型），湛江廉江属广东电网，勿用 Sz 后缀函数对应的端点
- **T+3 发布**：前端文案明确"用电日历 T+3 发布"（如 1 月 1 日用电量最迟 1 月 4 日可查）；"每日电费"非实时，定时任务与告警阈值按 T+3 设计
- **用户编号前导零**：清单 Excel 中为 16 位文本（0308 开头），CSV/浮点导出会丢前导零变 15 位（308 开头）；调接口前必须归一化补零到 16 位
- **meteringPointId 可传空字符串**：queryDayElectricAndTemperature 实测传 `meteringPointId: ""` 即可正常返回数据（2026-09-26 验证 9 户防溺水户全部成功），无需通过浏览器 UI 收集 pointId；该 ID 仍可通过 Vuex `currentUserNumInfo.pointId` 获取但非必需
- **eleCustId ≠ 用户编号**：eleCustId 是 bindingId（如 J6CFGiYb），不是 16 位用户编号
- **DNS 偶发抖动**：公司内网 DNS 偶发解析失败（getaddrinfo failed），重试即可恢复；浏览器不受影响（可能走不同网络路径）
- **绑定规模瓶颈**：实测单账号最多绑定 **26 户**（`queryBindInfo` 返回上限值）；894 户全量绑定需约 35 个账号或多轮解绑/重绑；规模化取决于企业通道或代扣银行卡统一性（银行卡通道已实测打通，894 户统一代扣银行卡后6位=081121）
- **Vue 表单填充**：南网在线为 Vue SPA，自动化填表必须用 Playwright fill/type；JS 原生 setter 赋值不触发 Vue 数据更新，提交后无请求发出
- **绑定页状态残留**：户号绑定页从结果页返回或绑定成功后，须重新导航整页再操作下一户，否则表单状态异常