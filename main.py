#!/usr/bin/env python3
import requests
import json
import os
import time
import argparse
import sys
from datetime import datetime
# 引入模板引擎，用于替换 HTML 中的变量
from string import Template

# ====================== 基础路径与全局变量 ======================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')
CFG = {}

# ====================== 核心功能函数 ======================

def load_config():
    """加载并验证全局配置"""
    global CFG
    try:
        with open(CONFIG_FILE, 'r') as f:
            CFG = json.load(f)
        # 简单的验证，确保关键 key 存在
        if not all(k in CFG for k in ['azure_app', 'auth', 'security', 'email_settings', 'paths']):
            raise ValueError("配置文件格式不正确，缺少必要字段。")
    except Exception as e:
        print(f"CRITICAL: 无法加载配置文件 {CONFIG_FILE}: {e}")
        sys.exit(1)

def get_valid_token():
    """获取有效的 Access Token (含缓存机制)"""
    token_file = CFG['paths']['token_cache']
    rt_file = CFG['paths'].get('refresh_token')


    # ===== 1 选择 refresh_token 来源 =====
    # refresh_token 优先级：
    # 1. 轮换落盘的 refresh_token 文件
    # 2. config.json 中的初始 refresh_token
    refresh_token = CFG['auth']['refresh_token']
    if rt_file and os.path.exists(rt_file):
        with open(rt_file, 'r') as f:
            refresh_token = f.read().strip()

    # ===== 2 尝试使用 access_token 缓存 =====
    if os.path.exists(token_file):
        try:
            with open(token_file, 'r') as f:
                token_data = json.load(f)
            # 提前 5 分钟认为过期
            if token_data.get('expires_at', 0) > time.time() + 300:
                return token_data['access_token']
        except Exception:
            pass # 缓存无效或读取失败，忽略

    # ===== 3 刷新 Access Token =====
    data = {
        'client_id': CFG['azure_app']['client_id'],
        'client_secret': CFG['azure_app']['client_secret'],
        'refresh_token': refresh_token,
        'grant_type': 'refresh_token',
        'scope': 'https://graph.microsoft.com/.default'
    }
    try:
        resp = requests.post(CFG['auth']['token_url'], data=data, timeout=30)
        resp.raise_for_status()
        new_tokens = resp.json()
        new_tokens['expires_at'] = time.time() + new_tokens['expires_in']

        # ===== 4 refresh_token 轮换落盘（关键）=====
        if 'refresh_token' in new_tokens and rt_file:
            with open(rt_file, 'w') as f:
                f.write(new_tokens['refresh_token'])
            os.chmod(rt_file, 0o600)
            refresh_token = new_tokens['refresh_token']  # ← 更新内存态

        # ===== 5 原子写入 access_token 缓存 =====
        tmp = token_file + '.tmp'
        with open(tmp, 'w') as f:
            json.dump(new_tokens, f)
        os.replace(tmp, token_file)  # 原子替换

        return new_tokens['access_token']
    except Exception as e:
        print(f"ERROR: 刷新令牌失败: {e}")
        print("请检查网络或确认 Refresh Token/Client Secret 是否依然有效。")
        # 如果是自检触发的刷新失败，不能 exit(1)，否则会影响主程序的正常发送
        # 这里抛出异常让调用者决定
        raise

def load_template_content(template_filename, variables=None):
    """
    从 templates 目录加载文件内容，并支持变量替换
    :param template_filename: 模板文件名
    :param variables: 字典，用于替换模板中的 ${key} 占位符
    """
    template_dir = CFG['paths'].get('template_dir')
    if not template_dir:
        print("CRITICAL: paths.template_dir 未在 config.json 中配置")
        sys.exit(1)
    file_path = os.path.join(template_dir, template_filename)
    if not os.path.exists(file_path):
        print(f"ERROR: 模板文件不存在: {file_path}")
        sys.exit(1)
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 如果提供了变量字典，则进行替换
        if variables:
            # safe_substitute 在找不到变量时不会报错，而是保留原样，比较安全
            return Template(content).safe_substitute(variables)
        return content

    except Exception as e:
        print(f"ERROR: 读取或处理模板文件失败: {e}")
        sys.exit(1)

def send_email_core(to_email, subject, body_content, is_html=False):
    """核心发送逻辑，供外部调用或内部自调用"""
    try:
        token = get_valid_token()
    except Exception as e:
        # 如果获取 Token 失败
        print(f"CRITICAL: Token 获取失败: {e}")
        return False


    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    email_data = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body_content
            },
            "toRecipients": [{"emailAddress": {"address": to_email}}]
        },
        "saveToSentItems": "true"
    }
    try:
        resp = requests.post(CFG['email_settings']['graph_send_url'], headers=headers, json=email_data, timeout=30)
        resp.raise_for_status()
        print(f"SUCCESS: 邮件已发送至 {to_email}")
        return True
    except Exception as e:
        print(f"FAIL: 发送失败: {e}")
        if hasattr(e, 'response') and e.response is not None:
            print(f"API Response: {e.response.text}")
        return False # 返回失败状态，供调用者判断

# ====================== 监控自检逻辑 ======================

def run_self_check():
    """检查密钥过期，必要时调用自身功能发送警告"""
    log_file = CFG['paths']['warning_log']
    expire_date_str = CFG['security'].get('secret_expire_date')

    if not expire_date_str or expire_date_str == "202X-XX-XX": return

    try:
        expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
        days_left = (expire_date - datetime.now()).days

        if days_left > CFG['security']['warning_threshold_days']:
            return # 未到预警期

        # 防骚扰：检查今天是否已发送
        today_str = datetime.now().strftime("%Y-%m-%d")
        if os.path.exists(log_file):
            with open(log_file, 'r') as f:
                if f.read().strip() == today_str: return

        print(f"ALERT: 触发密钥过期警告（剩余 {days_left} 天），正在发送通知...")

        # --- 核心复用点 ---
        # 1. 准备模板变量
        template_vars = {
            'days_left': str(days_left),
            'expire_date': expire_date_str,
            'client_id_short': CFG['azure_app']['client_id'][:8]
        }
        # 2. 复用模板加载功能，并进行变量替换
        warning_body = load_template_content(
            CFG['paths']['warning_template_name'],
            variables=template_vars
        )
        # 3. 复用核心发送功能发送给管理员
        admin_email = CFG['email_settings']['admin_notify_email']
        if send_email_core(admin_email, f"【紧急】密钥即将过期 (剩余 {days_left} 天)", warning_body, is_html=True):
            # 发送成功才记录日志
            with open(log_file, 'w') as f: f.write(today_str)

    except Exception as e:
        # 自检错误不应阻断主流程，打印警告即可
        print(f"WARN: 自检模块运行出错: {e}")

# ====================== 主程序入口 ======================

if __name__ == "__main__":
    # 1. 初始化加载配置
    load_config()

    # 2. 在处理任何参数前，先执行自检
    #    这样即使不带参数运行脚本，也能触发检查
    run_self_check()

    # 3. 解析命令行参数 (如果没有参数，到这里就会打印帮助信息并退出)
    parser = argparse.ArgumentParser(description='Outlook OAuth2 Sender Ultimate')

    parser.add_argument('--OAuthcheck', action='store_true',
                        help='仅执行 OAuth 健康检查（不发送邮件）')

    parser.add_argument('-t', '--to', help='收件人')
    parser.add_argument('-s', '--subject', help='主题')

    group = parser.add_mutually_exclusive_group()
    group.add_argument('-b', '--body', help='直接提供文本正文内容')
    group.add_argument('-f', '--file', help='指定 templates 目录下的模板文件名')

    parser.add_argument('--html', action='store_true', help='声明内容为 HTML (使用模板时默认开启)')

    args = parser.parse_args()

    # ===== OAuth 健康检查模式 =====
    if args.OAuthcheck:
        print("== OAuth Health Check ==")
        try:
            # 1. 尝试获取 / 刷新 access_token
            token = get_valid_token()
            print("✔ Access token refresh: OK")

            # 2. 运行 Secret 过期自检（只打印，不影响退出码）
            run_self_check()

            print("✔ OAuth 状态正常")
            sys.exit(0)

        except Exception as e:
            print(f"✖ OAuth 健康检查失败: {e}")
            sys.exit(2)

    # ===== 正常发送模式：参数校验（补回 required=True 行为）=====
    if not args.to or not args.subject:
        parser.error("the following arguments are required: -t/--to, -s/--subject")

    if not (args.body or args.file):
        parser.error("one of the arguments -b/--body -f/--file is required")

    # 4. 准备发送内容
    final_body = ""
    is_html_content = args.html

    if args.file:
        # 加载用户指定的模板 (这里不传 variables，保持原样加载)
        final_body = load_template_content(args.file)
        is_html_content = True
        print(f"INFO: 使用模板文件: {args.file}")
    else:
        final_body = args.body

    # 5. 执行主发送任务
    if not send_email_core(args.to, args.subject, final_body, is_html_content):
        sys.exit(1)