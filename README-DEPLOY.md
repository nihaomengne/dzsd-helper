# 收单录入辅助小程序 - 部署说明

## 当前状态
此版本已完成基础公网部署加固，适合先部署到 Render 免费版试运行。

## 已做的安全处理
- 密码改为哈希存储
- 默认管理员密码改为环境变量初始化
- SECRET_KEY 改为环境变量
- 默认关闭公开注册
- 增加基础安全响应头
- 增加 Cookie 安全配置

## 推荐部署平台
- Render（免费版）

## 必填环境变量
- SECRET_KEY：随机长字符串
- INIT_ADMIN_PASSWORD：管理员初始密码

## 推荐环境变量
- INIT_ADMIN_USERNAME=admin
- ALLOW_PUBLIC_REGISTRATION=false
- SESSION_COOKIE_SECURE=true
- TRUST_PROXY=true
- DATA_DIR=/opt/render/project/data
- DEFAULT_MAX_COUNT=10
- PORT=10000

## Render 部署步骤
1. 将本项目上传到 GitHub
2. 登录 Render
3. New + → Web Service
4. 选择该 GitHub 仓库
5. 若识别 `render.yaml`，按其默认配置部署
6. 确认已挂载持久磁盘到 `/opt/render/project/data`
7. 填写环境变量并部署

## 登录说明
首次启动时，系统会创建管理员账号：
- 用户名：`INIT_ADMIN_USERNAME`（默认 admin）
- 密码：`INIT_ADMIN_PASSWORD`

## 上线后建议
- 首次登录后立即修改管理员密码
- 默认不要开启公开注册
- 定期备份 Render 磁盘中的数据库文件

## 注意
- 本项目当前使用 SQLite，适合轻量内部使用
- 免费版可能存在休眠 / 冷启动
- 如后续多人长期使用，建议迁移 PostgreSQL
