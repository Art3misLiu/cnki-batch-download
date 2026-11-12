# CNKI Batch Download · 知网批量下载

**给一份清单，让 AI 把里面的知网文献全部下载下来 —— 一个零依赖、断点续跑的 Skill**

[中文 Skill](skills/cnki-batch-download/SKILL.md) · [更新记录](CHANGELOG.md) · [示例清单](examples/)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/Art3misLiu/cnki-batch-download/pulls)

---

### 安装

。本仓库既是 DSH 插件，也是一个可直接使用的 Skill。它需要你自己已登录的浏览器，**不接收也不保存你的知网账号密码**。

**方式一：装成 DSH 插件（推荐，带自动更新）**

```bash
dsh plugin --profile desktop add Art3misLiu/cnki-batch-download
```

或从 npm（若已发布）：

```bash
dsh plugin --profile desktop add cnki-batch-download
```

装完重启 DSH，技能自动出现在技能目录里。

**方式二：只装 skill 文件**

```bash
git clone --depth 1 https://github.com/Art3misLiu/cnki-batch-download.git /tmp/cnki
cp -r /tmp/cnki/skills/cnki-batch-download "$HOME/.dsh/skills/"
```

Windows PowerShell：

```powershell
git clone --depth 1 https://github.com/Art3misLiu/cnki-batch-download.git $env:TEMP\cnki
Copy-Item -Recurse "$env:TEMP\cnki\skills\cnki-batch-download" "$env:USERPROFILE\.dsh\skills\"
```

装完重开一个对话，对 AI 说：

> 帮我批量下载这些知网文献：《题名1》《题名2》《题名3》

它会先问你「怎么登录知网」，再要你的清单，然后一篇篇下。全程你在自己的浏览器里操作登录，它只负责点下载按钮。

前置条件三样：DSH（或同类工具）+ `chrome-devtools` MCP（用真实浏览器点页面）+ Python 3.9 以上（`openpyxl` 只在读 `.xlsx` 清单时需要）。

## 这是什么

写论文的人都有这一步：参考文献攒了三十篇，知网让你一篇篇点开详情页、点下载、等落盘、改名。半小时没了，中间还夹着两次滑块验证码。

这个过程不该由人来做。它有三个特点：**重复、有固定步骤、出错能重试**——正好是 AI 加浏览器能干的活。

所以这个 Skill 做的事情很窄：

> **你给清单，你登录，它把文件放进你的下载目录，最后给你一张「成功 / 跳过 / 失败」的账单。**

它不绕过付费墙，不代你付钱，不碰你的密码。没有权限的文献，它只告诉你「这篇要 ￥3」。

## 效果对比

| | 手动 | 本 Skill |
|---|---|---|
| 30 篇文献 | 点 90 次、改名 30 次、约 40 分钟 | 一句指令，约 5 分钟（含网络等待） |
| 中途验证码 | 自己发现、自己划 | 立刻暂停整批，等你划完说「继续」 |
| 下载到一半断网 | 不知道下到第几篇 | 状态文件记着，重跑从断点继续 |
| 本地已有这篇 | 靠记忆，常常重复下 | 自动跳过（连「副标题被截掉 + `_作者`」的命名差异也认） |
| 某篇没权限 | 点开才发现 | 只记单价 `需付费(单篇￥3)`，剩下的接着下 |
| 下完以后 | 一屏文件，不知道漏了谁 | `report.md` + 可勾选的 `failed.md` |

## 快速开始

### 方式一：直接对话（推荐）

装好 Skill 后，把清单发过去就行。清单可以是任何一种：

| 你手里的东西 | 例子 |
|---|---|
| 参考文献列表（GB/T 7714） | `[1] 张三, 李四. 基于深度学习的图像识别方法研究[J]. 计算机学报, 2020, 43(5): 1-10.` |
| APA 参考文献 | `Smith, J., & Lee, K. (2022). Graph neural networks for citation analysis. Journal of Informetrics, 16(3), 101-120.` |
| 题名 + 作者 + 年份 | `目的论视角下国产游戏英译研究 / 张智 / 2023` |
| 只有题名 | 一行一篇 |
| 知网链接列表 | `https://kns.cnki.net/kcms2/article/abstract?v=...` |
| Excel / CSV / JSON | 表头含「题名、作者、年份、链接」任意几列 |

样例在 [`examples/`](examples/)：`refs.txt`、`list.csv`、`list.xlsx`、`list.json`。

### 方式二：命令行自检

`scripts/batch.py` 也可以单独跑，用来在正式开工前确认配置对不对（Windows 写 `python`，macOS / Linux 写 `python3`）：

```bash
python scripts/batch.py init --login personal --download-dir "D:/论文"   # 一次性配置
python scripts/batch.py doctor                                          # 六项自检
python scripts/batch.py parse "清单.xlsx" --workdir "状态"               # 解析清单
python scripts/batch.py existing --workdir "状态" --dir "D:/论文"        # 标出本地已有
python scripts/batch.py next --workdir "状态"                            # 看下一条待处理
python scripts/batch.py report --workdir "状态"                          # 出报告
```

## 五种登录方式

第一次用会先问这一句，因为「能下到什么」完全取决于它：

| 你的情况 | 说明 | 要不要给网址 |
|---|---|---|
| 个人账号，自己买 | 按篇付费，￥1–3 一篇 | 不要 |
| 知网机构用户登录 | 单位给的个人账号，能直连 | 不要 |
| 学校 WebVPN | 学校域名编进网址，形如 `https://kns-cnki-net-443.v.xx.edu.cn/` | 要一条能打开知网的网址 |
| 路径前缀型 VPN | 形如 `https://vpn.xx.edu.cn/https/7772.../kns8s/search` | 同上 |
| 图书馆 ezproxy | 形如 `https://libproxy.xx.edu.cn/login?url=...` | 同上 |

后三种只需要你粘一条「在浏览器里能打开知网」的网址，它自己认出是哪种，之后把所有知网链接自动改写成同一通道——不用你懂什么叫 host 映射。

## 仓库结构

```
cnki-batch-download/
├── package.json          # DSH 插件清单（dsh.bundle 指向 cordis.patch.yml）
├── index.js              # 6 行插件入口：把 skills/ 注册成技能源
├── cordis.patch.yml      # 插件挂载点
├── skills/
│   └── cnki-batch-download/
│       ├── SKILL.md      # 给 AI 读的说明书（安装后生效的就是它）
│       └── scripts/
│           └── batch.py  # 清单解析 + 断点续跑 + 落盘轮询，只用标准库
├── README.md             # 你正在看的这份
├── CHANGELOG.md
├── LICENSE               # MIT
└── examples/
    ├── refs.txt          # GB/T 7714 + APA 混排
    ├── list.csv
    ├── list.xlsx
    └── list.json
```

`batch.py` 只做「不需要浏览器」的部分：解析清单、比对本地文件、记状态、出报告。点页面、点下载由 AI 通过 `chrome-devtools` MCP 完成。这样职责清楚，知网改版时通常只用改 `SKILL.md` 里的选择器。

## 产物

```
状态/
├── items.json      # 每条：题名、作者、年份、网址、状态、文件、失败原因
├── state.json      # 计数镜像
├── report.md       # 已下载 / 已存在跳过 / 失败 / 待处理，四段
└── failed.md       # 可勾选的失败清单，改好条件后重跑就补下
```

## 合规声明

- **不接收、不保存、不代填账号密码。** 登录全部发生在你自己的浏览器里，Skill 只读页面状态（有没有下载按钮、有没有「未登录」字样）。
- **只下载你账号本来就有权限的文献。** 遇到付费墙只记录单价，不代付、不绕过、不尝试机构外访问。
- **不破解验证码。** 检测到滑块就暂停整批，等你手动划完。
- **节流。** 每篇之间停 1–2 秒，避免给知网服务器造成压力。
- 请遵守你所在单位与知网的使用协议，下载的文献仅用于个人学习与研究。

## 常见问题

**Q：会不会封号？**
A：它做的是「你自己点会做的事」，只是自动化，并且节流。但任何自动化都不保证零风险，请自行判断。

**Q：为什么需要 chrome-devtools MCP？**
A：知网的下载按钮要求真实浏览器会话和登录态。用 HTTP 请求模拟拿不到 PDF。

**Q：验证码老是弹？**
A：连续快速下载会触发风控。把间隔拉长，或者分几批下。

**Q：篇数很多（100+）怎么办？**
A：分批。每批 30–50 篇跑一次，出一次报告。状态文件是累积的，重跑不会重复。

**Q：知网改版了，按钮点不到？**
A：`SKILL.md` 里的选择器需要更新，通常是 `.download-btns a` 这一层。欢迎 PR。

## License

MIT。欢迎 PR，欢迎转发给正在肝参考文献的同门。
