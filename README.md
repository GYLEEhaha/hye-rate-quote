# AQ国际 内部运价查询

面向货代内部员工的运价参考查询工具（v1）。**打开页面需先登录**（共用账号，由管理员私下告知）。请勿在公开场合传播链接与账号。

> 说明：登录校验在浏览器端完成（哈希比对），可防随手翻源码看到明文密码，但不是服务端鉴权；正式对外请改用服务端登录。


## 国内访问（免翻墙优先）

GitHub Pages（`*.github.io`）在大陆常不稳定。可优先试：

1. **jsDelivr 中国镜像（资源）**：JS/CSS/JSON 可用  
   `https://cdn.jsdmirror.com/gh/GYLEEhaha/hye-rate-quote@main/`  
   注意：jsDelivr 对 **HTML 强制 `text/plain`**，不能直接当网站入口。
2. **XHTML 入口（若 CDN 返回可渲染类型）**：  
   `https://cdn.jsdmirror.com/gh/GYLEEhaha/hye-rate-quote@main/index.xhtml`  
   （推送后稍等缓存；若仍是纯文本则不可用。）
3. 原 Pages：`https://gyleehaha.github.io/hye-rate-quote/`（需可访问 GitHub 时）

登录账号由管理员告知（共用账号 `AQGJ`）。

## 如何打开（本机）

```bash
cd hye-rate-quote
./start.sh
# 或: python3 -m http.server 8765
```

浏览器访问：**http://127.0.0.1:8765/**

## 发布到 GitHub Pages（公开网址）

1. 新建 GitHub 仓库，将本目录推上去（**不要**提交 `data/rates.xlsx`，见 `.gitignore`）。
2. 确认已包含解析后的 JSON：`data/parsed/`、`index.html`、`static/`。
3. 可运行检查脚本：

```bash
./scripts/prepare_pages.sh
```

4. 仓库 Settings → Pages → Source 选 **Deploy from a branch**，分支 `main`，文件夹 **/ (root)**。
5. 稍等后访问 `https://<用户名>.github.io/<仓库名>/`。

路径均为相对路径（`static/`、`data/parsed/`），从仓库根目录发布即可，无需 `/docs`。

## 如何更新 Excel 价格

1. 用新的报价表 **覆盖** 文件：`data/rates.xlsx`（仅本机保留，勿提交到公开仓库）
2. 重新解析：

```bash
python3 scripts/parse_rates.py
```

3. 将更新后的 `data/parsed/*.json` 提交并推送，Pages 会随之更新。

依赖：`openpyxl`（`pip3 install openpyxl`）。

## 能查什么

| 渠道大类 | 主要来源表 | 常用条件 |
|---------|-----------|---------|
| 海派头程 | 海派头程报价汇总 | 交货仓、包税 |
| 集运 | 集运价格汇总 | 仓、分区/邮编、重量档 |
| 海派 | HYE海派渠道汇总 | 仓、包税/自税、邮编、重量 |
| 海卡 | 快/普/特惠海卡 + 价格汇总 | **FBA仓库代码**、仓、税别、KG/CBM |
| 商私卡 | 美西/美中美东商私卡 | 仓、税别、分区/区号 |
| 超大件 / 小包裹 / 岛屿海派 | 对应表 | 仓、邮编分区 |

另可查看：**船期表**、**海运品名附加费**。

## 目录结构

```
hye-rate-quote/
  index.html              # 查询页面（含登录）
  static/                 # CSS / JS
  data/rates.xlsx         # 源报价 Excel（本机；不进公开仓库）
  data/parsed/            # 解析后的 JSON（需发布，供查询）
  scripts/parse_rates.py
  scripts/prepare_pages.sh
  README.md
```

## 注意

- 页面展示价格 **仅供内部参考**，不作为最终结算依据。
- Excel 含合并单元格与多产品区块，解析已尽量兼容；若某渠道异常，请对照原表或反馈需补的 sheet。
- 偏远邮编库、索赔说明、系统下单导航等参考表 v1 未接入查询。
