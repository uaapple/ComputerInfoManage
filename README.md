# ITK China 电脑信息管理系统

这是一个用于管理 ITK China 电脑资产、人员归属、库存电脑和盘点记录的内部工具仓库。

当前仓库同时包含两套形态：

- `电脑信息管理工具.ps1`：PowerShell + Windows Forms 桌面版，仍然是当前主版本
- `web-server.js` + `web/`：Node.js 轻量 Web 版，已经支持电脑、人员、库存的增删改查和邮件盘点

## 项目简介

桌面版当前已覆盖的核心能力：

- 电脑信息管理：电脑名称、序列号、固定资产号、型号、MAC 地址、归属人、备注
- 人员名单管理：中文名、拼音、邮箱、部门、员工类型、Mentor
- 库存电脑管理：无归属人的电脑单独管理，可直接分配给人员
- 归属历史：查看每台电脑的归属变更记录
- 搜索与导出：支持搜索和 UTF-8 CSV 导出
- 邮件盘点：按业务规则生成待发送对象并调起默认邮件客户端
- NAS 云备份：上传 / 拉取 JSON 数据，并带简单版本保护

Web 版当前已覆盖：

- 首页仪表盘统计卡片
- 在用电脑 / 库存电脑 / 人员名单三类列表视图
- 关键字段搜索
- 详情面板与详情弹窗
- 电脑 / 人员 / 库存的新增、编辑、删除 / 转库存
- 邮件盘点候选人预览、勾选、批量生成邮件草稿
- 型号、归属人排序

Web 版已经不再是“只读预览页”，但也还不是最终正式版。它仍然缺少登录、角色权限、数据库和并发写保护。

## 仓库结构

```text
ComputerInfoManage/
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

### Web 版

前提：

- 已安装 Node.js

执行：

```powershell
node .\web-server.js
```

如果当前终端里 `node` 不在 PATH，也可以直接执行：

```powershell
& "C:\Program Files\nodejs\node.exe" .\web-server.js
```

或双击：

`启动Web版预览.bat`

默认监听：

- `0.0.0.0:8099`

本机访问：

- `http://127.0.0.1:8099`

## 当前重要业务规则

- 归属人在数据内部以人员 ID 存储，界面显示中文名
- 员工类型分为正式员工 / 实习生
- 实习生必须绑定一位正式员工 Mentor
- 删除人员前，如果仍被作为 Mentor 引用，则不允许删除
- 删除人员后，其名下电脑会自动转入库存
- 固定资产号和 MAC 地址支持为空，统一展示为 `N/A`
- 编辑电脑或库存电脑时需要输入授权密码，当前密码为 `1`
- 邮件盘点规则：
  - 正式员工名下电脑发给本人
  - 实习生名下电脑发给对应 Mentor
  - 同一个 Mentor 的相关电脑会合并到同一封邮件

## 数据文件

桌面版和当前 Web 版都直接读写以下 JSON 文件：

- `data/computers.json`
- `data/colleagues.json`
- `data/inventory_mail_batches.json`

桌面版云端同步相关配置文件：

- `data/cloud_backup_config.json`
- `data/cloud_backup_sync_state.json`

## 当前状态

- 桌面版：可用，仍是当前生产主版本
- Web 版：已支持电脑 / 人员 / 库存 CRUD、详情、排序、邮件盘点，但仍基于 JSON 文件存储
- 当前推荐继续阅读：
  - [PROJECT_CONTEXT.md](./PROJECT_CONTEXT.md)
  - [docs/requirements.md](./docs/requirements.md)

## 已知限制

- Web 版尚未接入数据库，仍直接读写 JSON
- Web 版尚未实现登录和角色权限控制
- 桌面版与 Web 版共享同一批 JSON 文件，不适合多人同时写入
- 旧数据和部分历史字符串曾存在中文编码脏数据问题，后续仍需继续清理
