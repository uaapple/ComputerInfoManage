# ITK China 电脑信息管理系统

这是一个用于管理 ITK China 电脑资产、人员归属、库存电脑和盘点记录的项目仓库。

当前仓库同时包含两套形态：

- `电脑信息管理工具.ps1`：PowerShell + Windows Forms 桌面版，当前仍是主要可用版本
- `web-server.js` + `web/`：Node.js 轻量 Web 预览版，用于验证后续内网 Web 化方向

## 项目简介

桌面版当前已覆盖的核心能力：

- 电脑信息管理：电脑名称、序列号、固定资产号、型号、MAC 地址、归属人、备注
- 人员名单管理：姓名、拼音、邮箱、部门、员工类型、Mentor
- 库存电脑管理：无归属人的电脑单独管理，可直接分配给人员
- 归属历史：查看每台电脑的归属变更记录
- 搜索与导出：支持搜索和 UTF-8 CSV 导出
- 邮件盘点：按规则生成待发送对象并调用默认邮件客户端
- NAS 云端备份：上传/拉取 JSON 数据，并带简单版本保护

Web 预览版当前仅覆盖：

- 品牌区与 Logo 展示
- 首页仪表盘统计卡片
- 在用电脑 / 库存电脑 / 人员名单三类列表视图
- 搜索栏与详情弹窗

它还不是正式替代版，不具备完整编辑、权限、并发保护和数据库能力。

## 仓库结构

```text
LaptopsInfo/
├─ 电脑信息管理工具.ps1
├─ 启动电脑信息管理工具.bat
├─ web-server.js
├─ 启动Web版预览.bat
├─ web/
│  ├─ index.html
│  ├─ styles.css
│  └─ app.js
├─ data/
│  ├─ computers.json
│  ├─ colleagues.json
│  ├─ inventory_mail_batches.json
│  ├─ cloud_backup_config.json
│  └─ cloud_backup_sync_state.json
├─ ITK_Logo_RGB.jpg
├─ PROJECT_CONTEXT.md
└─ docs/
   └─ requirements.md
```

## 本地运行

### 桌面版

双击：

`启动电脑信息管理工具.bat`

或执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\电脑信息管理工具.ps1
```

### Web 预览版

前提：

- 已安装 Node.js

执行：

```powershell
node .\web-server.js
```

或双击：

`启动Web版预览.bat`

默认监听：

- `0.0.0.0:8099`

本机访问：

- `http://127.0.0.1:8099`

内网访问示例：

- `http://10.36.76.95:8099`

## 部署说明

当前 Web 预览版面向公司内网 Windows 虚拟机部署，目标机器为：

- `10.36.76.95`

最小部署步骤：

1. 在虚拟机安装 Node.js
2. 将整个仓库复制到虚拟机目录
3. 执行 `启动Web版预览.bat`
4. 放行防火墙端口 `8099`

PowerShell 放行示例：

```powershell
New-NetFirewallRule -DisplayName "ITK Web 8099" -Direction Inbound -Protocol TCP -LocalPort 8099 -Action Allow
```

## 当前重要业务规则

- 归属人内部存储为人员 ID，界面显示中文名
- 员工类型分为正式员工 / 实习生
- 实习生必须绑定一位正式员工 Mentor
- 删除人员前，如果仍被用作 Mentor，则不允许删除
- 删除人员后，其名下电脑自动转入库存
- 固定资产号和 MAC 地址支持“暂无”，展示统一为 `N/A`
- 编辑已有电脑信息时需要输入授权密码，当前密码为 `1`
- 邮件盘点规则：
  - 正式员工名下电脑发给本人
  - 实习生名下电脑发给对应 Mentor

## 数据文件

桌面版和当前 Web 预览版都直接读取以下 JSON 文件：

- `data/computers.json`
- `data/colleagues.json`
- `data/inventory_mail_batches.json`

桌面版云端同步相关配置文件：

- `data/cloud_backup_config.json`
- `data/cloud_backup_sync_state.json`

## 当前状态

- 桌面版：可用，仍是当前主版本
- Web 预览版：已完成首页视觉和基础数据读取验证，但仍是预览性质
- 推荐继续阅读：
  - [PROJECT_CONTEXT.md](/c:/Users/zguan/Documents/IT_Topic/LaptopsInfo/PROJECT_CONTEXT.md)
  - [requirements.md](/c:/Users/zguan/Documents/IT_Topic/LaptopsInfo/docs/requirements.md)

## 已知问题

- `README` 之前遗留过旧描述，现已按当前状态更新
- Web 预览版存在编码脏数据问题，部分中文在服务端字符串中可能显示异常
- Web 预览版尚未接入数据库、登录权限、编辑能力和真正的并发控制
- 桌面版与 Web 预览版仍共享 JSON 文件，不适合多人同时写入
