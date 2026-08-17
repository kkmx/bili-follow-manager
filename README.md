# B站清关注工具

一个本地运行的 B站关注列表清理与分析工具，适合关注数已到 5000 上限、需要批量整理的情况。支持图形界面和命令行两种方式。

> 本项目仅供个人整理自己的账号，请遵守 B站用户协议。取关不可逆，请先备份并确认名单。

## 主要功能

- 扫码登录，按账号独立保存数据
- 拉取完整关注列表
- 识别已注销账号
- 分析每个 UP：粉丝数、主要分类、热度 Top 视频、关注时间、关注时长、最新投稿/动态时间
- 生成 Excel 报告：全部 UP、按分类拆分、已注销、仅建议取关
- 取关管理：在界面内勾选、全选/反选、按天数批量选择
- 增量同步：取关后快速移除已取关账号、新增账号完整分析、现有账号更新最新动态
- 风控自动冷却、断点续传

## 安装依赖

需要 Python 3.9 或更高版本。

Windows 双击 `install_dependencies.bat`，或在命令行执行：

```bash
python -m pip install -r requirements.txt
```

## 图形界面（推荐）

Windows 双击：

```
启动B站清关注工具.bat
```

或命令行执行：

```bash
python bili_follow_manager_gui.py
```

界面里的“操作指引”按钮包含完整使用步骤。

## 命令行

```bash
python bili_follow_manager.py login
python bili_follow_manager.py fetch
python bili_follow_manager.py analyze --only-new
python bili_follow_manager.py export --days 180 --refresh-plan
python bili_follow_manager.py incremental-sync --export
python bili_follow_manager.py unfollow --from-selected
python bili_follow_manager.py unfollow-cancelled
python bili_follow_manager.py status
python bili_follow_manager.py logout
```

## 数据位置

每个账号的数据独立保存在：

```
data/accounts/<你的UID>/
    cookies.json
    followings.json
    analysis.json
    selected_unfollow.json
    bili_follow_report.xlsx
    unfollow_plan.xlsx
```

这些目录已通过 `.gitignore` 排除，不会上传到 GitHub。

## 打包为 Windows 可执行文件

图形界面已经可以打包成免 Python 的 exe：

```bash
python -m pip install pyinstaller
pyinstaller --onefile --windowed --name BiliFollowManager --add-data "partition_map.json;." bili_follow_manager_gui.py
```

生成文件在 `dist/BiliFollowManager.exe`，双击即可运行。首次运行会在 exe 所在目录创建 `data` 文件夹保存账号数据。

## 上传到 GitHub

1. 在 GitHub 新建空仓库，例如 `bili-follow-manager`，不要勾选 README/license。
2. 在本项目目录打开终端执行：

```bash
git init
git add .
git commit -m "B站关注列表清理与分析工具"
git branch -M main
git remote add origin https://github.com/你的用户名/bili-follow-manager.git
git push -u origin main
```

3. 之后别人可以：

```bash
git clone https://github.com/你的用户名/bili-follow-manager.git
cd bili-follow-manager
python -m pip install -r requirements.txt
python bili_follow_manager_gui.py
```

## 常见问题

- 遇到 412/-352 风控：等待 10～30 分钟再继续，工具会自动冷却重试。
- 二维码未弹出：查看日志里的链接，在浏览器中打开扫码。
- 报告/计划打不开：先点“生成报告”生成对应文件。
- 想清空关注：把未更新天数设为 0 天，但互相关注和特别关注会保留，且取关不可逆。
