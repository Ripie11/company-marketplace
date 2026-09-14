# 深圳水务集团数字员工专家市场（内部）

本仓库是一个 **CodeBuddy / WorkBuddy 插件市场（Plugin Marketplace）**，用于在集团内部统一分发数字员工与专家能力。

## 仓库结构

```
company-marketplace/
├── .codebuddy-plugin/
│   └── marketplace.json          # 市场清单（必需，市场入口）
├── plugins/
│   └── water-office-writer/      # 专家包：文秘专员（v1.6.0）
│       ├── .codebuddy-plugin/plugin.json
│       ├── agents/water-office-writer.md
│       ├── avatars/expert.png
│       ├── USER_GUIDE.md / README.md
│       └── skills/
│           ├── zhibi-suiwen/             # 执笔随文（写得出）
│           └── knowledge-base-organizer/ # 知识库整理（找得到）
└── README.md
```

> ⚠️ 仓库根的 `output/` 是**生成报告的输出目录**（含内部经营数据），**不属于插件、
> 不入离线 zip、禁止外发**。分发时一律排除。

## 领导称谓口径（务必统一）

| 领导 | 称谓 |
|------|------|
| 龚利民 | 党委书记、董事长 |
| 方琳 | **总裁**（「执行总裁」为历史误称，勿再使用） |

乐享知识库主空间对应目录：`01-龚利民（董事长）` / `02-方琳（总裁）`。
旧目录名 `02-方琳（执行总裁）` 仅作为归档兼容别名保留在
`knowledge-base-organizer` 的映射表中，**新建目录一律用总裁口径**。

## 离线分发（无 Git / 无内网 HTTP 时）

```bash
cd 专家分发
zip -r water-office-writer-1.6.0-marketplace.zip company-marketplace/ \
    -x "company-marketplace/output/*" -x "*__pycache__*" -x "*.pyc"
```

把 zip 解压到员工可访问的共享盘目录，再按下面"三步安装"用**目录路径**方式添加市场：

```
/plugin marketplace add Z:/digital-employee/company-marketplace
```

> 版本升级后请重新打包，并把文件名里的版本号一并更新。

## 一、员工如何安装（三步）

在 WorkBuddy / CodeBuddy 中依次执行：

```
/plugin marketplace add <本仓库地址> 
/plugin install water-office-writer@szwater-experts
/reload-plugins
```

`<本仓库地址>` 按托管方式不同，取值为以下之一：

| 托管方式 | 地址写法 | 示例 |
|---|---|---|
| 集团 Git（GitLab/Gitea 等） | HTTPS 或 SSH | `https://git.szwater.local/digital-employee/company-marketplace.git` |
| 集团内网 HTTP 服务 | marketplace.json 直链 | `https://apps.szwater.local/plugin-marketplace/marketplace.json` |
| 本地共享盘（临时） | 目录路径 | `Z:/digital-employee/company-marketplace` |

## 二、管理员统一推送（推荐）

避免让每位员工手工安装。管理员在**项目级**配置 `.codebuddy/settings.json`：

```json
{
  "extraKnownMarketplaces": {
    "szwater-experts": {
      "source": {
        "source": "git",
        "url": "https://git.szwater.local/digital-employee/company-marketplace.git"
      }
    }
  },
  "enabledPlugins": {
    "water-office-writer@szwater-experts": true
  }
}
```

成员信任该仓库后，客户端会自动提示安装该市场与插件。

## 三、更新与版本管理

- 修改专家内容后，同步更新 `marketplace.json` 与 `plugin.json` 中的 `version`
- 用 `git tag` 打版本（如 `v1.0.0`），员工侧可精确锁定：`/plugin marketplace add <url>#v1.0.0`
- 员工侧刷新：`/plugin marketplace update szwater-experts`

## 四、安全与合规

- 插件与市场以**用户权限执行任意代码**，仅应从可信来源安装
- 组织可通过**托管市场限制**收窄允许添加的市场范围
- 本市场内容含集团内部公文风格档案，**禁止公开发布**，仓库须设为私有

## 五、本地测试（开发用）

```
/plugin marketplace add ./company-marketplace
/plugin install water-office-writer@szwater-experts
/reload-plugins
```
