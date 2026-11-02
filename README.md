# Telegram 清洁报告机器人

## Docker 启动

1. 安装并启动 Docker（Windows 使用 Docker Desktop 的 Linux containers）。
2. 在项目根目录创建 `token.txt`，只填写从 @BotFather 获取的新 Bot Token。不要提交凭据。
3. 检查 `bot.py` 中 `ADMIN_IDS` 是否为你自己的 Telegram 用户 ID。
4. Windows 和 Linux 都无需手动创建 `data` 或执行 `chown`。
   在项目目录运行：

   ```sh
   docker compose up -d --build
   docker compose logs --tail=100 -f bot
   ```

机器人使用 Telegram polling，不需要开放端口。不要同时运行多个使用同一 Token 的实例。
镜像以 UID/GID `10001:10001` 运行，根文件系统只读，只有数据目录和临时目录可写。
Compose 会先运行一次 `init-storage`：创建照片目录、修复整个数据目录的所有者和写入权限，
并在保留 Token 文件所有者的同时将其组设置为 `10001`、允许组读取并移除其他用户权限。
初始化服务临时以 root 运行，只挂载数据和 Token，不连接网络，成功后退出；
机器人等它成功后才以非 root 用户启动，Token 对机器人仍为只读。
日志中 `init-storage` 显示 `Exited (0)` 是正常状态。
宿主机需要允许 bind mount 的 `chown`/`chmod`；只读磁盘或禁止修改所有者的网络文件系统不适用。

## 拉取代码更新

当前全部应用源码 `bot.py` 以只读方式挂载到容器 `/app/bot.py`。
仅更新 Python 业务代码时不需要重建镜像，但必须重启进程才能加载更新：

```sh
git pull
docker compose restart bot
```

如果 `requirements.txt`、Dockerfile 或 Compose 配置有变化，使用：

```sh
git pull
docker compose up -d --build
```

重启和重建都继续使用相同的宿主机 `data` 目录。

## 持久化数据

Compose 将宿主机 `./data` 直接绑定到容器 `/data`：

| 宿主机路径 | 内容 |
| --- | --- |
| `data/reports.db` | SQLite 任务记录及索引 |
| `data/images/` | 上传的照片 |
| `data/Monthly_Report_*.docx` | Word 导出报告 |
| `data/Task_Log_*.csv` | CSV 导出报告 |
| `data/reports_backup_*.db` | 发送备份期间的临时文件，发送后会删除 |
| `token.txt` | 单独只读挂载的 Telegram Token |
| `bot.py` | 只读挂载的应用源码，更新后重启生效 |

`docker compose build`、`docker compose up -d --build`、`docker compose down`
及重建容器不会删除宿主机 `data`。`down -v` 也不会删除这些 bind mount 文件。
请保留同一宿主机数据目录；移动项目时应一起移动 `data` 和 `token.txt`。
程序内的 `/undo`、`/cleardata` 仍会删除对应数据，挂载不代替备份。
数据、Token、Git 历史均不会被复制进镜像。

备份前先 `docker compose stop bot`，再复制整个 `data` 目录，完成后 `docker compose start bot`。
仅备份数据库不能恢复照片。

## 已有数据迁移

首次切换到 Docker 前停止旧机器人，把原目录的 `reports.db`、`images/` 和需要保留的
Word/CSV 报告复制到 `data/`，保留原始备份。初始化服务会自动设置目录权限。
旧数据库中的照片路径可能是 Windows 或其他机器的绝对路径，复制后不会自动转换。
这些旧记录需要将 `image_path` 和 `pair_image_path` 转换成对应的 `/data/images/...` 路径，
否则旧照片无法用于报告。新容器记录使用 `/data/images/...`，重建后路径保持一致。

## 本地运行

```sh
pip install -r requirements.txt
python bot.py
```

本地默认继续使用 `bot.py` 所在目录存储数据，读取旁边的 `token.txt`。
可以用 `BOT_DATA_DIR`、`BOT_TOKEN_FILE` 环境变量覆盖路径。

## 已知限制

当前群管理员授权逻辑没有限制到指定群组，可能获得全局报告、备份和删除权限。
Docker 配置不修复这一应用权限问题。正式使用前应限制授权范围。
