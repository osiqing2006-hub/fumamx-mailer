# 孚盟MX 群发助手 / Fumamx Mailer

> 一个给外贸业务员用的 Windows 小工具：把「按每个客户当地时区，给几百个客户逐个定时发开发信」这件事，做成双击就能跑。
>
> A Tkinter desktop tool that drives the Fumamx MX web UI over the Chrome DevTools Protocol, scheduling personalized cold emails to each recipient's **local** working hour. Zero third-party dependencies — no Playwright, no Selenium, no browser download.

<!-- 截图位置：把界面截图放到 docs/screenshot.png 后取消下面这行注释 -->
<!-- ![screenshot](docs/screenshot.png) -->

---

## English summary

Cold-emailing a few hundred overseas customers one by one is tedious and error-prone: you have to pick a template, type a subject, set a scheduled time that matches *their* local morning, and then remember which of the 200 rows already went out when the run breaks halfway.

This tool wraps the Fumamx MX web client (a Chinese CRM/EDM SaaS) with a 6-tab Tkinter GUI and automates the clicking. It talks to Edge/Chrome directly through the DevTools Protocol using a hand-written RFC 6455 WebSocket client, so there is nothing to `pip install` beyond the standard library. It converts timezones per recipient, normalises city names to countries (Dubai → UAE), sends in resumable batches, retries failures, and writes the send date back into your existing customer spreadsheet without touching a single existing cell.

It is **not** a bulk-blast tool: it drives the vendor's own UI and its own scheduled-send queue, at a pace that respects the vendor's rate limits. Read [使用边界](#使用边界请务必读) before you run it on anyone's customer list.

---

## 目录

- [这东西解决什么问题](#这东西解决什么问题)
- [六个页面，对应六步操作](#六个页面对应六步操作)
- [技术上值得说的几点](#技术上值得说的几点)
- [快速开始](#快速开始)
- [客户表格式](#客户表格式)
- [目录结构](#目录结构)
- [几个不太显然的设计选择](#几个不太显然的设计选择)
- [已知限制](#已知限制)
- [使用边界（请务必读）](#使用边界请务必读)
- [License](#license)

---

## 这东西解决什么问题

孚盟 MX 自带群发，但它的群发是「同一时刻，全体一起发」。开发信不适合这么发——欧美客户在凌晨收到邮件，打开率会很难看。真正想要的效果是：每个客户都在**他自己当地的上午九点**收到。

手动做是这样的：打开客户表，找到第 1 行的国家，算出当地 9 点是北京几点，回孚盟写信页选模板、填主题、设定时、提交，等页面反应，切回写信页，找第 2 行……200 个客户，一个下午就没了。

而且这个过程会断。登录态过期、网络抖一下、孚盟限流、模板选择器改版，任何一种都会让批次停在中途，然后你面对的问题变成：「刚才那些到底发出去没有？」

这个工具就是把这套动作固化下来，并且重点解决三件事：

1. **时区**：客户表里写「国家」，程序换成 IANA 时区，按你要的当地钟点换算成孚盟能接受的时间，两个时间都写进日志，可以逐条核对。
2. **断点**：分批发，每批发多少自己填；每张分表记住发到第几位，中断了下次从断点接着走，不用重头数。
3. **回写**：发完把发送日期按原表**已有记录的格式**（`26.8.24` 这种）追加到客户那一行的下一个空列，不新建行、不覆盖任何已有单元格，原表另存一份。

---

## 六个页面，对应六步操作

界面是 tab 式，从左到右就是操作顺序。

| 页 | 干什么 |
|---|---|
| ① 连接孚盟 | 用独立 profile 启动一个 Edge 窗口，你在里面手工登录一次。之后登录态存在 profile 里，不用反复登。程序不接触账号密码。 |
| ② 客户表 | 选 xlsx / csv，选工作表，自动识别「国家 / 邮箱 / 客户名 / 公司」四列，认错了可以手改。下面预览前 40 行，只有填了邮箱的行会被发。 |
| ③ 发送设置 | 模板、主题、封间等待、工作时段、演练模式、跳过已发邮箱、发送前清掉模板自带的抄送。 |
| ④ 定时发送 | 按客户当地时间排定时任务，写进孚盟的定时队列。支持本批数量、断点续跑、标为未完成。 |
| ⑤ 立即发送 | 不走孚盟定时队列，程序自己到点一个个点发送。适合要盯着发的小批量。 |
| ⑥ 历史发送记录 | 按「模板 × 主题 × 分表」查每个客户发过几次、什么时候发的，可以单独重发某几条。 |

几个细节：

- **演练模式**：勾上之后只填表、不点发送，可以先把一整批走一遍看效果，一个客户都不会真的收到。
- **窗口不抢焦点**：发送过程中程序会弹轻提示，提示框不抢输入焦点。你可以在旁边打字，不会被打断。
- **发送前会核对**：不是点完「提交」就记成功。孚盟的定时提交成功 ≠ 投递成功，程序会在发件箱里再巡检一遍，把显示失败的挑出来提醒你。
- **失败重试**：单封失败自动重试 2 次，仍失败的进「待人工补发清单」，可以单独对这批再跑一轮。

---

## 技术上值得说的几点

### 零依赖，是因为不想让人装 300MB 的浏览器内核

`core/cdp.py` 是一个手写的 Chrome DevTools Protocol 客户端，包含一个最小化的 RFC 6455 WebSocket 实现（握手、掩码、分帧、ping/pong 都有）。用标准库的 `socket` + `struct` 直接和 Edge 的 `--remote-debugging-port` 对话。

为什么不用 Playwright / Selenium：

1. 它们都要下载各自的浏览器内核，几百 MB，第一次用还得联网初始化；
2. 打包成 exe 体积直接翻几倍；
3. 这个工具真正需要的能力只有四件——打开页面、找元素、点或输入、读文本。CDP 的 `Runtime.evaluate` + `Input.dispatchKeyEvent` 完全够。

代价是要自己处理连接保活、页面切换、元素不可见、点击被遮挡这些情况，`cdp.py` 里 700 行大部分是在干这个。

### 复用你已经登录的浏览器，而不是去存密码

程序用独立的 `--user-data-dir` 启动 Edge，你在弹出的窗口里手工登录一次，Cookie 就留在那个 profile 目录里。程序全程不读、不存、不上传任何账号密码——它连密码框都不碰，只检查「当前页面是不是已经登录后的状态」。

这一点是刻意的：任何要你填账号密码的第三方工具都值得警惕，包括这个。

### 城市名归一，是把脏数据挡在入口

外贸客户表里的「国家」经常写的是城市：迪拜、纽约、香港、圣保罗。而孚盟的「国家地区」下拉只认标准国名，搜「迪拜」必然失败。

`core/geo_alias.py` 收了 200 多条城市/别名到国家的映射，支持在 `data/geo_alias.json` 里自己加。香港、澳门、台湾统一换算成「中国香港 / 中国澳门 / 中国台湾」，搜索时再去掉前缀去比对——这里没有任何模糊空间。

### 时区换算，宁可不换也不换错

`core/tzmap.py` 覆盖 161 个国家和地区。优先用标准库 `zoneinfo`（能正确处理夏令时），拿不到 tzdata 时回落到内置的固定偏移表。**认不出的国家返回 `None`，调用方保持原值、不换算**——宁可发成你填的那个时间，也不猜一个错的。

这一条是踩坑踩出来的：早期版本把「用户填的 09:00」当成了各客户的当地时间直接写进孚盟，结果英国客户和客户相差 6~12 小时，看着都写着 9:00，实际发出去的时间五花八门。现在两种基准（北京时间 / 客户当地时间）可以选，两个时间都打日志。

### Excel 读写是自己写的

`core/xlsxlight.py` 只用 `zipfile` + `xml.etree` 读写 xlsx。因为需要处理的表格就几百行、格式简单，为了这个装 openpyxl 不划算。

唯一的例外是**回写客户表**（`core/sheet_writeback.py`）——要保留原表的样式、合并单元格、列宽，这部分用了 openpyxl。所以它是可选依赖：不装也能正常发送，只是「回写」按钮会提示缺模块。

---

## 快速开始

需要 **Windows + Python 3.10 以上**（用到 `zoneinfo`、`str | None` 类型语法）。macOS / Linux 理论上能跑（tkinter + CDP 都是跨平台的），但没测过，路径和浏览器启动逻辑可能需要微调。

```bash
git clone https://github.com/osiqing2006-hub/fumamx-mailer.git
cd fumamx-mailer

# 可选：让时区换算支持夏令时（不装也能跑，用内置偏移表）
pip install tzdata

# 可选：要用「回写客户表」功能才需要
pip install openpyxl

# 启动
python app.py
```

Windows 上也可以直接双击 `启动.bat`（自动找 `py` / `pythonw`，找不到会提示去装 Python）。

第一次使用的顺序：

1. ① 页点「打开孚盟MX 并登录」，在弹出的窗口里手工登录；
2. 点「读取模板列表」，下拉里选一个模板，再「读取历史主题」；
3. ② 页选你的客户表，确认四列对应关系；
4. ③ 页先勾上**演练模式**，跑一遍看看填得对不对；
5. 没问题了取消演练，④ 页排定时或 ⑤ 页立即发。

发送结果默认落在程序目录下的 `output/`，可以用环境变量 `FUMAMX_OUT_DIR` 改。

---

## 客户表格式

一行一个客户，至少要有**邮箱**列。列名支持中英文，识别逻辑在 `app.py` 的 `COUNTRY_KEYS` / `MAIL_KEYS` / `NAME_KEYS` / `COMPANY_KEYS`：

| 字段 | 能识别的列名 |
|---|---|
| 国家 | 国家、地区、country、国别 |
| 邮箱 | 邮箱、email、e-mail、mail、邮件 |
| 客户名 | 名字、姓名、联系人、客户名称、客户名、name、contact、客户 |
| 公司 | 公司名称、公司、company、企业、单位 |

邮箱解析能处理全角 `＠`、中文分号/顿号、换行、首尾引号、`张三 <a@b.com>` 这种带名字的写法、`mailto:` 前缀、一行里塞多个地址。

`samples/customers_sample.csv` 是一份假数据样例，含各种脏数据（空邮箱、城市当国家、全角符号），可以拿来试。

---

## 目录结构

```
fumamx-mailer/
├── app.py                  # Tkinter 界面 + 主流程（6 个 tab）
├── 启动.bat                # Windows 启动脚本
├── core/
│   ├── cdp.py              # 手写 CDP 客户端（含 RFC 6455 WebSocket）
│   ├── mx_web.py           # 孚盟页面操作层：登录态判断、选模板、填表、提交
│   ├── engine.py           # 单封邮件的发送编排
│   ├── scheduler.py        # 定时任务队列 + 三层去重
│   ├── tzmap.py            # 国家 → 时区
│   ├── geo_alias.py        # 城市/别名 → 国家
│   ├── xlsxlight.py        # 极简 xlsx 读写（纯标准库）
│   ├── sheet_writeback.py  # 发送记录回写客户表（需 openpyxl）
│   ├── result_export.py    # 结果导出
│   ├── send_stats.py       # 发送流水统计
│   └── failure.py          # 失败原因归类
├── samples/                # 假数据样例
└── data/                   # 运行时生成：流水、去重台账（已 gitignore）
```

---

## 几个不太显然的设计选择

**去重做了三层。** 发送前查已发台账；提交前查「同邮箱 + 同主题 + 同日」的待发任务；真正点发送前再查一次。因为中途停过、手动补过、浏览器卡过之后，单靠一层判断一定会重复发。

**「新一轮」是个单独的开关。** 默认情况下，发过的邮箱直接跳过——这对「一张表只发一轮」是对的。但要给老客户换主题再发一轮时，这个逻辑就成了障碍。所以有个「新一轮发送」勾选：勾上之后，只有「同一邮箱 + 同一主题」才算发过，换个主题照样发。

**回写不新建行。** 客户表是人工维护的，行号、合并单元格、隐藏列都有意义。程序只做一件事：找到这个客户所在的行，找到这行最后一条记录的列，从下一列写日期。列不够就照同样的格式往右扩，扩了哪些列会告诉你。

**发送节奏有两套模式，默认那套是慢的。** 默认开着「随机间隔」：每封之间 45–110 秒随机，连发 5–8 封休息 3–7 分钟，另外每封有 8% 概率多停 30–90 秒。这个默认值不是为了好看——09-29 那轮 114 封、封间中位 17 秒、还都带 1.9M 附件，结果孚盟服务器直接 Connection reset，发件箱一片「发送失败」。把间隔拉开是唯一有效的解法。

关掉之后走另一套：封间等待 `step_delay`，默认 1.2 秒，只够孚盟保存上一封、切回写信页。适合你坐在旁边盯着的几十封小批量。想清楚后果再关。

---

## 已知限制

- **强依赖孚盟 MX 的页面结构。** 选择器写在 `core/mx_web.py` 里，孚盟前端一改版就可能失效。代码里的 `mx_web.py` 是这一堆文件里最长的一个（2000 行），大部分是各种兜底和重试，原因就在这里。
- **只测过 Windows + Edge。** Chrome 应该也能用（同样是 CDP），但 `browser_exe` 的探测顺序是按 Windows 写的。
- **时区固定偏移表不含历史规则变更**，只在拿不到 tzdata 时兜底用。
- **不定时发大量邮件。** 工具本身不限制数量，但孚盟会限流，收件方也会。真要发几千封，先想清楚域名预热和退信率。

---

## 使用边界（请务必读）

这个工具是**在你自己的账号下，操作你自己已经登录的 SaaS 界面**。它不破解接口、不逆向私有 API、不伪造浏览器指纹、不切换代理 IP、不伪造发件人身份。

需要说明的一点：默认开启的「随机间隔」会随机化发送间隔，让请求模式不那么像自动化脚本。写这份文档时考虑过要不要把它说成纯粹的「性能优化」——不是。它的目的确实包含**避免被平台的风控识别出来**。这个做法的性质介于「合理的速率限制」和「规避平台检测」之间，取决于你怎么用、用多狠。我把它留在代码里，因为它在真实使用中是必要的（不打爆对方服务器，也避免自己的账号被临时封），但它属于你自己需要判断风险的那一类功能。

用它之前请确认：

1. **收件人是你有正当理由联系的人。** 展会上交换过名片的、官网留过询盘的、有过业务往来的。买来的名单群发，在任何司法辖区都不好过。
2. **遵守当地反垃圾邮件法规。** 美国 CAN-SPAM 要求真实发件人信息、可用的退订方式；欧盟/英国的 GDPR 和 PECR 对 B2B 冷邮件也有明确要求；加拿大的 CASL 更严。这些是你作为发件人的责任。
3. **遵守孚盟的服务条款。** 用脚本操作他们的网页界面，可能不符合其 ToS 的某些条款。是否使用、风险由你自行判断。
4. **控制频率。** 默认节奏是为了不打爆对方系统、也避免你自己的账号被风控，不是「可以调快的上限」。

仓库作者不对使用本工具造成的账号封禁、投诉、退信或法律后果负责。代码按 MIT 协议「按现状」提供。

---

## License

[MIT](LICENSE) © 2026 osiqing2006
