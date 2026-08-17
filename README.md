# B站清关注工具

一个本地运行的 B站关注列表清理与分析工具，适合关注数已到 5000 上限、需要批量整理的情况。提供图形界面和命令行两种用法。

> ⚠️ 本项目仅供整理**自己的账号**，请遵守 B站用户协议。取关不可逆，请先备份并确认名单。

## ✨ 功能

- 扫码登录，多账号数据独立保存
- 拉取完整关注列表，识别已注销账号
- 分析每个 UP：粉丝数、主要分类、热度 Top 视频、关注时间、关注时长、最新投稿/动态时间
- 生成 Excel 报告：全部 UP、按分类拆分、已注销账号、仅建议取关
- 取关管理：界面内勾选、全选/反选、按距今天数批量选择
- 增量同步：移除已取关账号、新增账号完整分析、可选更新最新动态
- 风控自动冷却、断点续传、可强制中断任务

## 📦 下载发布版

在 Releases 页面下载 `BiliFollowManager_发布版.zip`，解压后双击 `BiliFollowManager.exe` 即可使用，**不需要安装 Python**。

## 🐍 从源码运行

需要 Python 3.9 或更高版本。

```bash
git clone https://github.com/你的用户名/bili-follow-manager.git
cd bili-follow-manager
python -m pip install -r requirements.txt
python bili_follow_manager_gui.py
```

Windows 也可以双击 `install_dependencies.bat` 安装依赖，再双击 `启动B站清关注工具.bat` 启动图形界面。

## 🖥️ 命令行

```bash
python bili_follow_manager.py login          # 扫码登录
python bili_follow_manager.py fetch          # 拉取关注列表
python bili_follow_manager.py analyze --only-new   # 分析
python bili_follow_manager.py export --days 180 --refresh-plan   # 生成报告
python bili_follow_manager.py incremental-sync --export   # 增量同步
python bili_follow_manager.py unfollow --from-selected    # 批量取关
python bili_follow_manager.py unfollow-cancelled         # 取关已注销账号
python bili_follow_manager.py status
python bili_follow_manager.py logout
```

## 📁 数据位置

每个账号的数据独立保存在 `data/accounts/<你的UID>/` 下，包括 Cookie、关注列表、分析结果和 Excel 报告。该目录已被 `.gitignore` 排除，不会上传到 GitHub。

## 🔨 打包为 exe

```bash
python -m pip install pyinstaller
pyinstaller --onefile --windowed --name BiliFollowManager --add-data "partition_map.json;." bili_follow_manager_gui.py
```

生成文件位于 `dist/BiliFollowManager.exe`。

## ❓ 常见问题

- 遇到 412/-352 风控：等待 10～30 分钟再继续，工具会自动冷却重试。
- 二维码未弹出：查看日志里的链接，在浏览器中打开扫码。
- 报告/计划打不开：先点“生成报告”生成对应文件。
- 想清空关注：把“未更新天数”设为 0 天，但互相关注和特别关注仍会保留，且取关不可逆。

## 📄 许可证

[MIT](LICENSE)
