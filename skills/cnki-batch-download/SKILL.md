---
name: cnki-batch-download
description: 按清单批量下载知网（CNKI）文献。先问清用户是个人账号、知网机构用户登录、还是学校 WebVPN/图书馆代理，再引导其自行登录，绝不索要密码。清单可为参考文献列表、标题+作者+年份、知网链接或 Excel/CSV。逐篇下载 PDF（无 PDF 退 CAJ），断点续跑，最后给出成功/失败清单与付费提示。Use when the user hands over a list of papers and wants them downloaded from CNKI in bulk.
argument-hint: "[清单文件路径] [--dir 保存目录] [--workdir 状态目录]"
---

# 知网批量下载

一句话：**问清登录方式 → 读清单 → 逐篇打开知网页面 → 点下载 → 检查文件真的存下来了 → 出报告。**

不需要用户给密码。本 skill 用 chrome-devtools MCP 驱动用户**自己已经登录好**的浏览器。

## 命令前缀：`<python>` 是什么

下文所有 `<python> batch.py ...` 都指：**用能跑 openpyxl 的 Python 执行本 skill 目录下的 `scripts/batch.py`**。

```bash
# macOS / Linux（把 SKILL_DIR 换成本 skill 的实际目录）
python3 "$SKILL_DIR/scripts/batch.py" --help

# Windows
python "<SKILL_DIR>\scripts\batch.py" --help
```

Windows 上如果 PATH 里的 python 没装 openpyxl（读 `.xlsx` 清单会报 `ModuleNotFoundError: No module named 'openpyxl'`），改用 DSH 自带解释器：

```
C:\Users\<用户名>\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe
```

其余清单格式（`.txt` / `.csv` / `.json` / 纯题名）不需要 openpyxl，任意 Python 3.9+ 都能跑。

---

## 第 0 步：先问清楚「你怎么登录知网」（**不要跳过，不要假设机构**）

很多用户不隶属任何学校：他们用**个人账号自己充值**下载。也有人是**知网官方机构用户登录**（单位给账号，或人在校园网 IP 内）。这两类都**不需要**改网址，也不要一上来就引导 VPN。

先问一句：

> 你在浏览器里是怎么登录知网的？1) 个人账号，自己注册充值 2) 单位/学校给的机构账号（知网右上「机构登录」）3) 学校 WebVPN（要先登 VPN 门户，之后网址会变长）4) 图书馆代理 / 路径前缀 VPN

| 用户选 | `--login` | 要不要改网址 | 下载收费 |
|---|---|---|---|
| 个人账号（手机号/邮箱注册） | `personal` | 不改，直接 `cnki.net` | 按篇扣余额 |
| 知网官方机构用户登录 | `institution` | 不改，点右上「机构登录」 | 单位付，免费 |
| 学校 WebVPN（域名编进网址） | `webvpn` | 改，形如 `kns-cnki-net-443.v.学校域名` | 学校付，免费 |
| 路径前缀 VPN / ezproxy | `prefix_vpn` / `ezproxy` | 改，多一段 `/https/<串>/` 或 `/login?url=` | 学校付，免费 |

### 分支 A：个人账号 或 机构用户登录（不需要改网址）

```
<python> batch.py init --login personal --download-dir "D:\论文"
<python> batch.py init --login institution --download-dir "D:\论文"
```

然后**主动把登录页打开给用户**，让他自己登：

- `navigate_page` 到 `entry_url`（personal → `https://www.cnki.net/`，institution → `https://kns.cnki.net/`）。
- 告诉用户：「浏览器已打开知网，请你用**自己的账号登录**。个人账号：手机号/邮箱；机构账号：点右上角『机构登录』。登录完回我一句『继续』。」
- **绝不索要账号密码**，也不要尝试自动填表。
- 用户回「继续」后，`evaluate_script` 检查一次是否已登录（见第 2 步的 JS）。

⚠️ 个人账号下载是**按篇扣费**的。跑起来大概率在第 1 篇就遇到付费墙，这是**正常结果**，不是故障：

- 标 `failed`，reason `账户无下载权限:需付费(单篇￥N)`，继续下一条。
- 汇总后如实告诉用户「共 N 篇，其中 M 篇需要付费，单价合计约 ￥X」。**不替用户付款、不绕过。**
- 如果用户说「我不想付费了，我有学校账号」→ 直接重跑 `init --login webvpn`（或 `institution`）换通道，再 `parse` 同一条清单续跑。

### 分支 B：学校 WebVPN / 路径前缀 VPN / ezproxy（需要改网址）

```
<python> batch.py init --login webvpn --download-dir "D:\论文"
```

这种方式**必须知道学校那条知网网址**，脚本会追问，或直接给：

```
<python> batch.py init --login webvpn --from-url "https://kns-cnki-net-443.v.xx.edu.cn/kns8s/search" --download-dir "D:\论文"
```

脚本会把网址归到四种模式之一（识别结果和用户选的不一致时会提示，以网址为准）：

| 网址长什么样 | 判为 | 含义 |
|---|---|---|
| `https://kns-cnki-net-443.v.学校域名/...` | `host` | WebVPN 把域名编进网址 |
| `https://vpn.学校域名/https/<一串>/kns8s/...` | `prefix` | 路径前缀 VPN |
| `https://libproxy.学校域名/login?url=...` | `ezproxy` | 图书馆代理 |
| `https://kns.cnki.net/...` 直接能开 | `direct` | 校园网 / CARSI / 机构 IP |

然后请用户**自己登录 VPN 门户**（这一步只能人工：扫码、短信、二次验证），登录完再继续。

### 换通道随时改

`init --login ...` 可以反复跑，后一次覆盖前一次。清单和状态不受影响（`items.json` 与配置无关），换完直接续跑。

## 第 1 步：自检

```
<python> batch.py doctor
```

输出中文检查表，含「登录方式」一行。`结论: 配置可用` 就继续。

- 「浏览器登录态」显示「提示」是正常的：命令行没有浏览器 cookie。
- 「入口网址可达」失败 → 网址写错了，或 VPN 域名当前不可用。

## 第 2 步：在浏览器里确认登录态

`navigate_page` 到配置里的 `entry_url`，再 `evaluate_script`：

```javascript
() => {
  const body = document.body.innerText;
  return {
    url: location.href,
    has_download_box: !!document.querySelector('.download-btns'),
    not_logged: /未登录/.test(body),
    who: (body.match(/[\u4e00-\u9fff]{2,12}(大学|学院|研究院|图书馆)/) || [''])[0],
    personal: /个人中心|退出登录|我的图书馆/.test(body)
  };
}
```

| 看到什么 | 含义 | 动作 |
|---|---|---|
| 学校统一身份认证页 | 还没登 VPN | 请用户自己登录（扫码/验证码），回「继续」 |
| 页头有学校名 / 单位名 | 机构权限生效 | 开始 |
| `personal: true` 或页头有账号昵称 | 个人/机构账号已登录 | 开始，但注意付费墙 |
| `not_logged: true` | 没登 | 请用户登，不要自己填表 |

---

## 清单解析

用能跑 openpyxl 的 Python（PATH 上的 `python` 可能是 3.9 且没有 openpyxl）：

```
<python> = %USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe
```

（Windows 下 `%USERPROFILE%` 就是 `C:\Users\<你的用户名>`；PowerShell 里可直接写 `$env:USERPROFILE\...`。）

Windows 之外的机器用 `python3`。

```
<python> batch.py parse "<清单文件>" --workdir "<状态目录>"
```

保存目录默认取配置里的 `download_dir`，要换就加 `--download-dir`。

支持的清单形态（自动判别，用户不用改格式）：

| 形态 | 例子 |
|---|---|
| GB/T 7714 | `[1] 张三, 李四. 题名[J]. 刊名, 2020, 43(5): 1-10.` |
| 无作者 GB/T | `题名[J]. 刊名, 2018, 40(2): 1-10.` |
| APA | `Smith, J. (2022). Title. Journal, 43(5), 1-10.` |
| 标题+作者+年份 | `题名 \| 作者 \| 2020`，制表符或全角冒号也行 |
| Markdown 表格 | `\| 题名 \| 作者 \| 2020 \|` |
| 知网链接 | `https://kns.cnki.net/kcms2/article/abstract?v=...` |
| Excel / CSV / JSON | 首行表头含"题名/标题/作者/年份/链接"等，自动识别列 |

输出 `items.json`（`key/title/authors/year/url/kind/n`）和 `state.json`。
`needs_search` = 还需要检索定位的条数。

解析后**先看前 3 条题名对不对**。若大量条目 `kind: "line"` 且题名里混了刊名年份，说明格式超出支持范围，请用户改成 Excel 或直接给链接，不要硬跑。

## 网址换成学校能访问的形态

只有 `host` / `prefix` / `ezproxy` 三种模式需要。`direct` 模式跳过这一步。

```
<python> batch.py proxy --workdir "<状态>"
```

可重复运行，已经换过的不会重复改（第二次返回 `rewritten: 0`）。参数默认读配置，临时改就用 `--mode/--school-domain/--prefix/--ezproxy`。

## 去重：已经下过的不要重下

```
<python> batch.py existing --workdir "<状态>" --dir "<保存目录>" --threshold 0.82
```

命中会自动标 `skipped` 并记文件名。

## 主循环

```
<python> batch.py next --workdir "<状态>"
```

返回下一条 `pending`。逐条做 A→D，最多 3 次重试。

### A. 打开详情页

- `url` 非空 → `navigate_page` 直达。**不要点搜索结果里的链接**，会开新标签页。
- `url` 为空 → 在知网内部检索。**注意：不能把 `?q=` 拼在网址上**（实测返回 0 行），必须三步：

  1. `navigate_page` 到本校检索页。
     - `host` 模式：`https://kns-cnki-net-443.v.<学校域名>/kns8s/search`
     - `prefix` 模式：`<prefix>/kns8s/search`
     - `direct` 模式：`https://kns.cnki.net/kns8s/search`
     - 最省事的写法：直接 `navigate_page` 到家里首页（`entry_url`），页面上有检索框。
     - **`/kns8s/` 只存在于 `kns-` 主机上**，`www-` 主机会返回「页面不存在」。
  2. `evaluate_script` 填词并点检索。点按钮会整页跳转，脚本里**不要 await 结果**：

     ```javascript
     () => {
       const i = document.querySelector('input.search-input');
       if (!i) return { error: 'no_search_box' };
       i.value = '题名';
       i.dispatchEvent(new Event('input', { bubbles: true }));
       document.querySelector('input.search-btn')?.click();
       return { submitted: true };
     }
     ```

  3. 下一个 `evaluate_script` 轮询取首条（30×500ms）：

     ```javascript
     async () => {
       for (let n = 0; n < 30; n++) {
         const row = document.querySelector('.result-table-list tbody tr');
         if (row) {
           const a = row.querySelector('td.name a.fz14');
           return { title: a?.innerText?.trim().replace(/\s+/g, '') || '', href: a?.href || '' };
         }
         await new Promise(r => setTimeout(r, 500));
       }
       return { error: 'no_result', url: location.href };
     }
     ```

  4. `batch.py sim "<检索题名>" "<清单题名>"`，`score >= 0.9` 才继续，并写回网址：
     `<python> batch.py mark --workdir "<状态>" --key <key> --status pending --url "<href>"`
     小于 0.9 → 标 `failed`，reason `检索未匹配(sim=0.xx)`。
     检索页面的链接已经是学校代理形态，不用再 `proxy`。

### B. 取下载基线并点击

先取当前时间戳当基线（**不要用 `scan --since 0`**，会吐出几百个文件）：

```
<python> -c "import time;print(repr(time.time()))"
```

然后一次 `evaluate_script` 搞定：等加载、查验证码、查登录、查标题、点 PDF：

```javascript
async () => {
  for (let n = 0; n < 30; n++) {
    if (document.querySelector('.download-btns')) break;
    await new Promise(r => setTimeout(r, 500));
  }
  const cap = document.querySelector('#tcaptcha_transform_dy');
  if (cap && cap.getBoundingClientRect().top >= 0) return { error: 'captcha' };
  if (/未登录/.test(document.body.innerText)) return { error: 'not_logged_in' };

  const title = document.querySelector('.brief h1')?.innerText?.trim()
    ?.replace(/\s*网络首发\s*$/, '') || '';

  // 下载区是 .download-btns；里面每个 a 的 id 都是 cajDown（重复 id，不能按 id 选）
  const anchors = [...document.querySelectorAll('.download-btns a')];
  const pdf = anchors.find(a => a.innerText.includes('PDF'));
  const caj = anchors.find(a => a.innerText.includes('CAJ'));
  const pick = pdf || caj;
  if (!pick) return { error: 'no_download', title, buttons: anchors.map(a => a.innerText.trim()) };

  if (window.__cnkiClicked) return { status: 'already_clicked', title };
  window.__cnkiClicked = true;      // 防连点：连点会下出「xxx (1).pdf」重复副本
  pick.removeAttribute('target');   // 必须，否则开新标签页，后面接管不到
  pick.click();
  return { status: 'downloading', format: pdf ? 'PDF' : 'CAJ', title };
}
```

### C. 判断是不是付费墙（很多用户会卡在这）

点完下载后立刻检查是否被送去收费页：

```javascript
() => {
  const t = document.body.innerText;
  const fee = location.href.includes('bar/fee') || /账户余额不足|请充值或开通会员/.test(t);
  const m = t.match(/单篇下载\s*￥\s*([\d.]+)/);
  return { paywall: fee, price: m ? m[1] : null };
}
```

`paywall: true` = 账号没有这篇的权限，**不是**下载失败，也不是验证码。标 `failed`，reason `账户无下载权限:需付费(单篇￥N)`，继续下一条。连续 3 条都是这样 → 暂停，告诉用户账号权限问题（换 VPN、换机构账号、或充值）。

### D. 校验落盘

下载是异步的，大文件（40 MB 学位论文）可能十几秒才写完。**不要靠猜时间**，用轮询命令等：

1. 点下载前先记基线（epoch 秒）：

   ```
   <python> -c "import time;print(repr(time.time()))"
   ```

2. 点完下载立刻调用（默认最多等 15 秒，每秒看一次；有文件就提前返回）：

   ```
   <python> batch.py waitnew --dir "<保存目录>" --since <基线> --timeout 15
   ```

   - `count ≥ 1` → 拿到 `files[0].path`，成功，`mark --status ok --file "<path>"`。
   - `count == 0`（退出码 1）→ 页面还在详情页就**再点一次**下载（脚本里的 `window.__cnkiClicked` 每页只允许点一次，所以先 `location.reload()` 或重新打开详情页），再 `waitnew` 一次。仍为空 → 标 `failed`，reason 写 `下载无文件`。
   - `count == 0` 但 `small_or_partial` 有内容 → 文件在写但还没写完，把 `--timeout` 放到 30 再等一次；仍 `ok:false` 才判 `文件0字节`。
3. 页面标题与清单题名 `sim < 0.75` → **不要下载**，标 `failed`，reason `页面标题不符(sim=0.xx)`，防止下错文献。
4. 拿到的文件名里出现 ` (1)`、`(2)` 后缀 = 之前多点了，`mark` 时选没有后缀的那个；多余的副本请用户自己删（不要代删用户文件）。

## 出错怎么处理

| 现象 | 含义 | 动作 |
|---|---|---|
| `error: captcha` | 滑块验证码 | **暂停整批**，请用户手动划过，回"继续" |
| `error: not_logged_in` | 掉登录 | **暂停整批**，请用户重新登录，回"继续" |
| `paywall: true` | 账号没权限 | 标失败继续；连续 3 条则暂停并说明 |
| `error: no_download` | 页面无下载按钮（可能无全文） | 标失败，继续 |
| `error: no_search_box` | 检索页没加载出来 | 回首页重试一次，仍失败则跳过该条 |
| `error: no_result` | 检索无结果 | 换简称再搜一次，仍无则标 `failed: 检索无结果` |
| 连续 3 条失败 | 权限到期或风控 | 暂停，报告，请用户确认 |

## 节流

- 每条之间 `sleep 1–2 秒`。
- 每 5 条报一次进度（`已完成 x / 剩余 y`），不要每篇刷屏。

## 收尾

```
<python> batch.py report --workdir "<状态>"
```

写出 `report.md`（已下载 / 已跳过 / 失败 / 待处理四段）和 `failed.md`（可勾选）。把路径给用户，失败项按原因归类说明：权限 / 未检索到 / 验证码中断 / 已存在。

## 中断后继续

```
<python> batch.py parse "<清单>" --workdir "<同一状态目录>"
```

`parse` 会按 key 继承旧状态并返回 `"resumed": true`，已下载的不会重下。

## 状态目录

```
<workdir>/
  items.json   清单 + 逐条状态（权威数据）
  state.json   进度快照（人工看）
  report.md    报告
  failed.md    失败清单
```

不要手改 `items.json` 的 `status`，用 `mark`（支持 `--key` / `--n` / `--title` 定位）。

## 常见问题（用户问到时照这个答）

| 用户说 | 原因 | 怎么办 |
|---|---|---|
| 「提示账户余额不足」 | 个人账号按篇扣费，余额不够 | 充值后重跑该条，或换学校 VPN / 机构登录 |
| 「我登了 VPN 还是收费」 | 打开的是 `cnki.net` 而不是代理域名 | 让用户从 VPN 门户里点进知网，把那条网址粘给 `init --from-url` |
| 「页面不存在」 | 在 `www-` 主机上访问了 `/kns8s/` | 换 `kns-` 主机，或从首页检索框走 |
| 「一直转圈」 | VPN 会话过期 | 重新登录 VPN，回"继续" |
| 「下错文献了」 | 检索到同名的其他文章 | `sim < 0.75` 本来会拦；确认后重跑该条 |
| 「我不用学校，就自己买」 | 用户是个人账号 | `init --login personal`，付费墙按条记录价格，不代付 |
| 「我单位给了知网账号」 | 机构用户，不是 VPN | `init --login institution`，让用户点右上「机构登录」 |
| 「我既没单位也没买」 | 无任何权限 | 只能先注册个人账号或找单位；不绕过 |

## 边界

- 只下用户账号有权限的文献，不绕过付费墙、不破解、不代登录、**不代付款**。
- 个人账号会如实报告每篇的付费金额，买不买由用户决定。
- 学位论文可能只有 CAJ，PDF 按钮不存在时自动退 CAJ，如实报告格式。
- 验证码与风控只能由用户手动过。
- 不做：批量导出题录到 Zotero（用 `cnki-export`）、期刊目录（用 `cnki-journal-toc`）。
- Windows 上 OneDrive/杀软偶尔锁文件，脚本已内置重试；`report` 报 `PermissionError` 就重跑一次。
