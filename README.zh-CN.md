# Outlook-OAuth2-Sender

🌍 [English](./README.md) | 简体中文

---

### 📖 项目简介

**Outlook-OAuth2-Sender** 是一个基于 Python 的可靠邮件发送工具，专为无图形界面的服务器环境（如 CentOS / Debian / Ubuntu Server）设计，可使用 Microsoft Outlook / Office 365 账号稳定发送邮件。

随着 Microsoft 正式弃用并关闭 Outlook / Office 365 的 SMTP 基本认证（Basic Authentication），传统基于 SMTP 的邮件发送方案已无法可靠工作。本项目通过 **Microsoft Graph API** 并结合安全的 **OAuth2** 认证机制，提供了一套现代化、长期可用的解决方案。

该工具以稳定性为核心设计目标，适合直接集成到自动化场景中，例如：  
**定时任务（cron）**、**监控告警**、**备份通知**、**运维脚本** 等。

---

### ✨ 核心特性

* **现代且安全**  
  基于 Microsoft Graph API 通过 HTTPS 发送邮件，彻底规避传统 SMTP 的兼容性与认证问题。

* **OAuth2 自动轮转**  
  实现完整 OAuth2 流程，支持 **Refresh Token 自动轮转与持久化**，可长期无人值守运行。

* **可靠的令牌缓存机制**  
  使用原子文件操作保存 Token 状态，有效避免并发或异常情况下的数据损坏。

* **健康检查模式**  
  提供专用的 `--OAuthcheck` 模式，便于监控系统检测邮件服务是否可用。

* **密钥过期预警**  
  内置自检逻辑，在 Azure Client Secret 即将过期前主动发送邮件提醒管理员。

* **HTML 模板支持**  
  支持基于模板的 HTML 邮件发送，并可进行变量替换，适合生成结构化报告邮件。

---

### 🛠️ 项目结构

* `main.py`  
  核心 CLI 程序，负责 Token 管理、自检逻辑以及邮件发送。

* `config.json`  
  存储敏感凭据与路径配置。**必须妥善保护（建议权限 `chmod 600`）。**

* `templates/`  
  HTML 邮件模板目录。

* `.graph_token_cache.json`  
  Graph API Token 缓存文件（自动生成）。

* `.refresh_token_rotated`  
  Refresh Token 轮转状态标记文件（自动生成）。

* `.warning_sent_log`  
  密钥过期提醒发送记录（自动生成）。

---

### 🚀 快速开始

#### 环境要求

* Python 3.6+  
* `requests` 库  
* 已完成 Azure 应用注册，并授予 Microsoft Graph API 权限：
  * `Mail.Send`
  * `Mail.ReadWrite`
  * `User.Read`
* Client ID、Client Secret 以及初始 Refresh Token

#### 安装步骤

1. 将仓库克隆到服务器目录，例如：
   ```bash
   /home/outlookOAuth2
   ```
2.	复制示例配置文件并填写真实凭据：
   ```bash
   cp config.example.json config.json
   ```
3.	加固配置文件权限：
   ```bash
   chmod 600 config.json
   ```
4.	赋予主程序执行权限：
   ```bash
   chmod +x main.py
  ```

#### 使用示例

发送纯文本邮件：
   ```bash
   ./main.py -t recipient@example.com -s "Subject" -b "Body text."
   ```
发送 HTML 模板邮件：
   ```bash
   ./main.py -t recipient@example.com -s "Report" -f report.html
   ```
OAuth / 服务健康检查：
   ```bash
   ./main.py --OAuthcheck
   ```
