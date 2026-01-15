# Outlook-OAuth2-Sender

🇬🇧 English | 🇨🇳 [简体中文](./README.zh-CN.md)

---

### 📖 Introduction

**Outlook-OAuth2-Sender** is a robust Python-based tool designed to reliably send emails from headless servers (like CentOS) using Microsoft Outlook/Office 365 accounts.

As Microsoft has deprecated and disabled SMTP Basic Authentication for Outlook/Office 365, traditional SMTP-based tools no longer work reliably. This project provides a modern solution by leveraging the **Microsoft Graph API** with secure **OAuth2** authentication. It is built for reliability and easy integration into automation workflows (cron jobs, monitoring alerts, backups).

### ✨ Key Features

*   **Modern & Secure:** Uses Graph API via HTTPS, eliminating legacy SMTP issues.
*   **OAuth2 Auto-Rotation:** Implements full OAuth2 flow with **automatic Refresh Token rotation** and persistence for long-term unattended operation.
*   **Robust Caching:** Uses atomic file operations for token caching to prevent data corruption.
*   **Health Check Mode:** Dedicated `--OAuthcheck` mode for monitoring systems to verify service status.
*   **Secret Expiration Alert:** Built-in self-check notifies admins via email before the Azure Client Secret expires.
*   **HTML Templates:** Supports sending rich HTML emails using templates with variable substitution.

### 🛠️ Architecture

*   `main.py`: Core CLI tool for token management, self-checks, and email sending.
*   `config.json`: Stores sensitive credentials and paths. **Must be secured (chmod 600).**
*   `templates/`: Directory for HTML email templates.
*   `.graph_token_cache.json`, `.refresh_token_rotated`, `.warning_sent_log`: Auto-generated state files.

### 🚀 Quick Start

#### Prerequisites

*   Python 3.6+, `requests` library.
*   Azure App Registration with Microsoft Graph API `Mail.Send`, `Mail.ReadWrite`, `User.Read` permissions.
*   Client ID, Client Secret, and initial Refresh Token.

#### Installation

1.  Clone repo to (e.g.) `/home/outlookOAuth2`.
2.  Copy `config.example.json` to `config.json` and edit with your credentials. **Use absolute paths.**
3.  Secure config: `chmod 600 config.json`
4.  Make executable: `chmod +x main.py`

#### Usage

**Send Text Email:**
```bash
./main.py -t recipient@example.com -s "Subject" -b "Body text."
```
#### Send HTML Template Email:
```bash
./main.py -t recipient@example.com -s "Report" -f report.html
```
#### Health Check:
```bash
./main.py --OAuthcheck
```
