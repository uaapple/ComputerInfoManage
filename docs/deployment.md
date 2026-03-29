# Web 部署说明

## 目标

推荐把 Web 版部署到一台 Windows 虚拟机上长期运行，并把“代码”和“数据”分开：

- 代码：可重复覆盖部署
- 数据：独立保留，不随发版包被覆盖

推荐目录：

```text
D:\ITK-ComputerInfoManage\
├─ app\
│  └─ current\
├─ data\
├─ backups\
├─ logs\
└─ packages\
```

## 推荐发布流程

### 1. 本机打包

在仓库根目录执行：

```powershell
.\scripts\package-web-release.ps1
```

也可以直接双击：

```text
scripts\package-web-release.bat
```

执行后会在 `release\` 下生成：

- 一个发布目录
- 一个对应的 zip 包

如果你希望首包带上当前 JSON 数据，可以执行：

```powershell
.\scripts\package-web-release.ps1 -IncludeSeedData
```

### 2. 将发布包拷到虚拟机

建议把 zip 包复制到虚拟机本地目录，例如：

```text
D:\ITK-Deploy\
```

### 3. 在虚拟机上部署

如果是 zip 包：

```powershell
.\scripts\deploy-web-release.ps1 -PackagePath "D:\ITK-Deploy\ComputerInfoManage-web-xxx.zip"
```

也可以直接把 zip 拖到下面这个文件上：

```text
scripts\deploy-web-release.bat
```

如果是已经解压后的目录：

```powershell
.\scripts\deploy-web-release.ps1 -PackageRoot "D:\ITK-Deploy\ComputerInfoManage-web-xxx"
```

部署脚本会自动：

- 创建 `app / data / backups / logs / packages`
- 初始化缺失的 JSON 数据文件
- 停掉旧服务（如果已安装）
- 备份旧版本
- 替换为新版本
- 重启服务（如果已安装）

## 推荐服务方式

建议使用 [NSSM](https://nssm.cc/) 把 Node 服务注册成 Windows Service。

先在虚拟机上准备：

- Node.js
- `nssm.exe`

然后执行：

```powershell
.\scripts\install-web-service.ps1 -InstallRoot "D:\ITK-ComputerInfoManage" -NssmPath "D:\tools\nssm\nssm.exe"
```

如果你已经把 `install-web-service.bat` 里的参数改成自己的环境，也可以直接双击它。

安装后可用：

```powershell
.\scripts\start-web-service.ps1
.\scripts\stop-web-service.ps1
```

也可以直接双击：

```text
scripts\start-web-service.bat
scripts\stop-web-service.bat
```

## 重要说明

### 1. 数据目录不会跟随版本替换

服务启动时会读取环境变量 `DATA_DIR`，因此生产数据固定存放在：

```text
D:\ITK-ComputerInfoManage\data
```

这意味着你以后升级代码时，不会把生产 JSON 数据覆盖掉。

### 2. 首次部署和后续升级流程不同

首次部署：

1. 打包
2. 拷到虚拟机
3. 运行部署脚本
4. 安装服务

后续升级：

1. 打包新版本
2. 拷到虚拟机
3. 再次运行部署脚本

### 3. 端口

默认监听端口是：

- `8099`

如需对内网开放，需要在虚拟机上放行防火墙端口。
