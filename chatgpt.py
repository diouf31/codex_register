import json
import os
import re
import time
import random
import string
import secrets
import hashlib
import base64
import imaplib
import email as email_lib
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
from pathlib import Path
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import Any, Dict, List, Optional
from email.header import decode_header
from email.utils import parsedate_to_datetime
import urllib.parse
import urllib.request
import urllib.error
import threading

from curl_cffi import requests as cffi_requests
from curl_cffi.requests import Session
import requests as std_requests

OUT_DIR = Path(__file__).parent.resolve()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"


# ========== 1. 微软邮箱接码模块 (Graph API + IMAP OAuth2) ==========

class MSMailFetcher:
    """微软邮箱验证码提取器：先用 Graph API，失败则回退到 IMAP OAuth2"""

    def __init__(self, ms_email: str, ms_password: str, client_id: str, refresh_token: str, log_func=None):
        self.ms_email = ms_email
        self.ms_password = ms_password
        self.client_id = client_id
        self.refresh_token = refresh_token
        self.log = log_func or print

    # ---------- Graph API 路径 ----------

    def _graph_get_token(self) -> Optional[str]:
        token_url = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
        data = {
            "client_id": self.client_id,
            "grant_type": "refresh_token",
            "refresh_token": self.refresh_token,
            "scope": "https://graph.microsoft.com/.default",
        }
        try:
            resp = std_requests.post(token_url, data=data, timeout=15)
            resp.raise_for_status()
            return resp.json().get("access_token")
        except Exception as e:
            self.log(f"  [Graph] Token 获取失败: {e}")
            return None

    def _graph_fetch_emails(self, access_token: str, folder: str = "Inbox", top: int = 10) -> List[Dict]:
        url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder}/messages"
        params = {
            "$top": top,
            "$select": "id,subject,body,bodyPreview,receivedDateTime",
            "$orderby": "receivedDateTime desc",
        }
        headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
        try:
            resp = std_requests.get(url, headers=headers, params=params, timeout=15)
            resp.raise_for_status()
            items = (resp.json() or {}).get("value", [])
            results = []
            for item in items:
                text = item.get("bodyPreview") or ""
                html_body = ""
                body = item.get("body") or {}
                if (body.get("contentType") or "").lower() == "html":
                    html_body = body.get("content") or ""
                if not text and html_body:
                    text = re.sub(r"<[^>]+>", " ", html_body)
                results.append({
                    "subject": item.get("subject") or "",
                    "text": text,
                    "date": item.get("receivedDateTime") or "",
                })
            return results
        except Exception as e:
            self.log(f"  [Graph] 邮件获取失败: {e}")
            return []

    def _graph_get_otp(self) -> Optional[str]:
        self.log("  [Graph] 尝试通过 Graph API 获取验证码...")
        token = self._graph_get_token()
        if not token:
            return None
        for attempt in range(40):
            for folder in ("Inbox", "JunkEmail"):
                mails = self._graph_fetch_emails(token, folder=folder, top=5)
                for m in mails:
                    sb = m.get("subject", "")
                    txt = m.get("text", "")
                    if "OpenAI" in sb or "ChatGPT" in sb or "verify" in sb.lower() or "code" in txt.lower():
                        match = re.search(r"(\d{6})", txt) or re.search(r"(\d{6})", sb)
                        if match:
                            self.log(f"  [Graph] 成功提取验证码: {match.group(1)}")
                            return match.group(1)
            time.sleep(8)
        return None

    # ---------- IMAP OAuth2 路径 ----------

    def _imap_get_token(self) -> Optional[str]:
        token_url = "https://login.live.com/oauth20_token.srf"
        data = {
            "client_id": self.client_id,
            "grant_type": "refresh_token",
            "refresh_token": self.refresh_token,
            "redirect_uri": "https://login.live.com/oauth20_desktop.srf",
        }
        try:
            resp = std_requests.post(token_url, data=data, timeout=15)
            resp.raise_for_status()
            return resp.json().get("access_token")
        except Exception as e:
            self.log(f"  [IMAP] Token 获取失败: {e}")
            return None

    def _imap_fetch_emails(self, access_token: str, mailbox: str = "INBOX", max_mails: int = 10) -> List[Dict]:
        auth_string = f"user={self.ms_email}\x01auth=Bearer {access_token}\x01\x01"
        imap = None
        results = []
        try:
            last_err = None
            for host in ("outlook.office365.com", "imap-mail.outlook.com"):
                try:
                    imap = imaplib.IMAP4_SSL(host, 993)
                    imap.authenticate("XOAUTH2", lambda x: auth_string.encode("utf-8"))
                    last_err = None
                    break
                except Exception as e:
                    last_err = e
                    if imap:
                        try: imap.logout()
                        except: pass
                    imap = None
            if last_err:
                raise last_err

            # 尝试选择文件夹
            selected = False
            candidates = [mailbox] if mailbox == "INBOX" else [mailbox, "Junk", "Junk Email", "Spam"]
            for folder in candidates:
                try:
                    status, _ = imap.select(folder, readonly=True)
                    if status == "OK":
                        selected = True
                        break
                except:
                    continue
            if not selected:
                imap.select("INBOX", readonly=True)

            status, messages = imap.search(None, "ALL")
            if status != "OK":
                return []
            email_ids = messages[0].split() if messages and messages[0] else []
            if not email_ids:
                return []
            email_ids = email_ids[-max_mails:]
            for eid in reversed(email_ids):
                try:
                    res, msg_data = imap.fetch(eid, "(RFC822)")
                    if res != "OK" or not msg_data or not msg_data[0]:
                        continue
                    msg = email_lib.message_from_bytes(msg_data[0][1])
                    subject_raw = msg.get("Subject") or ""
                    val, charset = decode_header(subject_raw)[0]
                    if charset:
                        try: subject = val.decode(charset)
                        except: subject = str(val)
                    elif isinstance(val, bytes):
                        subject = val.decode("utf-8", errors="ignore")
                    else:
                        subject = val

                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            ct = part.get_content_type()
                            if "attachment" in str(part.get("Content-Disposition")):
                                continue
                            if ct == "text/plain":
                                try:
                                    body = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="ignore")
                                    break
                                except: pass
                            elif ct == "text/html" and not body:
                                try:
                                    body = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="ignore")
                                except: pass
                    else:
                        try:
                            body = msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8", errors="ignore")
                        except:
                            body = ""
                    results.append({"subject": subject, "text": body.strip()})
                except:
                    continue
            return results
        except Exception as e:
            self.log(f"  [IMAP] 邮件获取失败: {e}")
            return []
        finally:
            if imap:
                try: imap.logout()
                except: pass

    def _imap_get_otp(self) -> Optional[str]:
        self.log("  [IMAP] 尝试通过 IMAP OAuth2 获取验证码...")
        token = self._imap_get_token()
        if not token:
            return None
        for attempt in range(40):
            for mailbox in ("INBOX", "Junk"):
                mails = self._imap_fetch_emails(token, mailbox=mailbox, max_mails=5)
                for m in mails:
                    sb = m.get("subject", "")
                    txt = m.get("text", "")
                    if "OpenAI" in sb or "ChatGPT" in sb or "verify" in sb.lower() or "code" in txt.lower():
                        match = re.search(r"(\d{6})", txt) or re.search(r"(\d{6})", sb)
                        if match:
                            self.log(f"  [IMAP] 成功提取验证码: {match.group(1)}")
                            return match.group(1)
            time.sleep(8)
        return None

    # ---------- 公共入口 ----------

    def get_otp(self) -> Optional[str]:
        """先 Graph API，失败后回退 IMAP OAuth2"""
        code = self._graph_get_otp()
        if code:
            return code
        self.log("  [!] Graph API 未能取到验证码，回退到 IMAP 协议...")
        return self._imap_get_otp()


# ========== 2. OpenAI OAuth2 授权与环境生成模块 ==========

AUTH_URL = "https://auth.openai.com/oauth/authorize"
TOKEN_URL = "https://auth.openai.com/oauth/token"
CLIENT_ID = "app_EMoamEEZ73f0CkXaXp7hrann"
DEFAULT_REDIRECT_URI = "http://localhost:1455/auth/callback"
DEFAULT_SCOPE = "openid email profile offline_access"

def _gen_password() -> str:
    alphabet = string.ascii_letters + string.digits
    special = "!@#$%^&*.-"
    base = [
        random.choice(string.ascii_lowercase),
        random.choice(string.ascii_uppercase),
        random.choice(string.digits),
        random.choice(special),
    ]
    base += [random.choice(alphabet + special) for _ in range(12)]
    random.shuffle(base)
    return "".join(base)

def _random_name() -> str:
    return ''.join(random.choice(string.ascii_lowercase) for _ in range(random.randint(5, 9))).capitalize()

def _random_birthdate() -> str:
    start = datetime(1970,1,1)
    end = datetime(1999,12,31)
    d = start + timedelta(days=random.randrange((end - start).days + 1))
    return d.strftime('%Y-%m-%d')

def _b64url_no_pad(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")

def _sha256_b64url_no_pad(s: str) -> str:
    return _b64url_no_pad(hashlib.sha256(s.encode("ascii")).digest())

def _random_state(nbytes: int = 16) -> str:
    return secrets.token_urlsafe(nbytes)

def _pkce_verifier() -> str:
    return secrets.token_urlsafe(64)

def _parse_callback_url(callback_url: str) -> Dict[str, Any]:
    candidate = callback_url.strip()
    if not candidate:
        return {"code": "","state": "","error": "","error_description": ""}
    if "://" not in candidate:
        if candidate.startswith("?"): candidate = f"http://localhost{candidate}"
        elif any(ch in candidate for ch in "/?#") or ":" in candidate: candidate = f"http://{candidate}"
        elif "=" in candidate: candidate = f"http://localhost/?{candidate}"
    parsed = urllib.parse.urlparse(candidate)
    query = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
    fragment = urllib.parse.parse_qs(parsed.fragment, keep_blank_values=True)
    for key, values in fragment.items():
        if key not in query or not query[key] or not (query[key][0] or "").strip():
            query[key] = values
    def get1(k: str) -> str:
        v = query.get(k, [""])
        return (v[0] or "").strip()
    code = get1("code"); state = get1("state")
    error = get1("error"); error_description = get1("error_description")
    if code and not state and "#" in code:
        code, state = code.split("#",1)
    if not error and error_description:
        error, error_description = error_description, ""
    return {"code": code,"state": state,"error": error,"error_description": error_description}

def _jwt_claims_no_verify(id_token: str) -> Dict[str, Any]:
    if not id_token or id_token.count(".") < 2: return {}
    payload_b64 = id_token.split(".")[1]
    pad = "=" * ((4 - (len(payload_b64) % 4)) % 4)
    try:
        return json.loads(base64.urlsafe_b64decode((payload_b64 + pad).encode("ascii")).decode("utf-8"))
    except: return {}

def _decode_jwt_segment(seg: str) -> Dict[str, Any]:
    raw = (seg or "").strip()
    if not raw: return {}
    pad = "=" * ((4 - (len(raw) % 4)) % 4)
    try: return json.loads(base64.urlsafe_b64decode((raw + pad).encode("ascii")).decode("utf-8"))
    except: return {}

def _to_int(v: Any) -> int:
    try: return int(v)
    except: return 0

def _post_form(url: str, data: Dict[str, str], timeout: int = 30) -> Dict[str, Any]:
    body = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded","Accept": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read()
            if resp.status != 200: raise RuntimeError(f"token exchange failed: {resp.status}")
            return json.loads(raw.decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"token exchange failed: {exc.code}") from exc

@dataclass(frozen=True)
class OAuthStart:
    auth_url: str
    state: str
    code_verifier: str
    redirect_uri: str

def generate_oauth_url(*, redirect_uri: str = DEFAULT_REDIRECT_URI, scope: str = DEFAULT_SCOPE) -> OAuthStart:
    state = _random_state()
    code_verifier = _pkce_verifier()
    code_challenge = _sha256_b64url_no_pad(code_verifier)
    params = {
        "client_id": CLIENT_ID, "response_type": "code", "redirect_uri": redirect_uri,
        "scope": scope, "state": state, "code_challenge": code_challenge,
        "code_challenge_method": "S256", "prompt": "login",
        "id_token_add_organizations": "true", "codex_cli_simplified_flow": "true",
    }
    auth_url = f"{AUTH_URL}?{urllib.parse.urlencode(params)}"
    return OAuthStart(auth_url=auth_url, state=state, code_verifier=code_verifier, redirect_uri=redirect_uri)

def fetch_sentinel_token(*, flow: str, did: str, proxies: Any = None) -> Optional[str]:
    """获取 OpenAI 最新的反爬 Token (Sentinel)"""
    try:
        body = json.dumps({"p": "", "id": did, "flow": flow})
        resp = cffi_requests.post(
            "https://sentinel.openai.com/backend-api/sentinel/req",
            headers={
                "origin": "https://sentinel.openai.com",
                "referer": "https://sentinel.openai.com/backend-api/sentinel/frame.html?sv=20260219f9f6",
                "content-type": "text/plain;charset=UTF-8",
                "user-agent": UA
            },
            data=body, proxies=proxies, impersonate="chrome120", timeout=15,
        )
        if resp.status_code != 200: return None
        return resp.json().get("token")
    except: return None

def submit_callback_url(*, callback_url: str, expected_state: str, code_verifier: str, redirect_uri: str = DEFAULT_REDIRECT_URI) -> str:
    """提取重定向中的 Code 并换取最终的 Access / Refresh Token"""
    cb = _parse_callback_url(callback_url)
    if cb["error"]: raise RuntimeError(f"oauth error: {cb['error']}")
    if not cb["code"] or not cb["state"]: raise ValueError("callback missing code/state")
    if cb["state"] != expected_state: raise ValueError("state mismatch")

    token_resp = _post_form(TOKEN_URL, {
        "grant_type": "authorization_code", "client_id": CLIENT_ID,
        "code": cb["code"], "redirect_uri": redirect_uri, "code_verifier": code_verifier,
    })
    
    access_token = (token_resp.get("access_token") or "").strip()
    refresh_token = (token_resp.get("refresh_token") or "").strip()
    id_token = (token_resp.get("id_token") or "").strip()
    expires_in = _to_int(token_resp.get("expires_in"))

    claims = _jwt_claims_no_verify(id_token)
    email_addr = str(claims.get("email") or "").strip()
    auth_claims = claims.get("https://api.openai.com/auth") or {}
    account_id = str(auth_claims.get("chatgpt_account_id") or "").strip()

    now = int(time.time())
    expired_rfc3339 = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(now + max(expires_in, 0)))
    now_rfc3339 = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(now))

    config = {
        "id_token": id_token, "access_token": access_token, "refresh_token": refresh_token,
        "account_id": account_id, "last_refresh": now_rfc3339, "email": email_addr,
        "type": "codex", "expired": expired_rfc3339,
    }
    return json.dumps(config, ensure_ascii=False, separators=(",", ":"))


# ========== 3. 核心注册与提取流程 ==========

def run(ms_email: str, ms_password: str, client_id: str, refresh_token: str,
        proxy: Optional[str], log_func=None) -> Optional[tuple]:
    log = log_func or print
    proxies = {"http": proxy, "https": proxy} if proxy else None
    s = Session(proxies=proxies, impersonate="chrome120")
    s.headers.update({"user-agent": UA})

    openai_email = ms_email
    openai_password = _gen_password()
    log(f"[*] 使用邮箱: {openai_email}")
    log(f"[*] 生成密码: {openai_password}")

    fetcher = MSMailFetcher(ms_email, ms_password, client_id, refresh_token, log_func=log)

    oauth = generate_oauth_url()
    
    try:
        # 第一步：进入 OAuth
        resp = s.get(oauth.auth_url, timeout=15)
        did = s.cookies.get("oai-did")
        if not did:
            log("[Error] 未能获取到 OpenAI Device ID (oai-did)")
            return None

        # 第二步：获取 Sentinel Token (authorize_continue)
        sen_token = fetch_sentinel_token(flow="authorize_continue", did=did, proxies=proxies)
        sentinel = json.dumps({"p": "", "t": "", "c": sen_token, "id": did, "flow": "authorize_continue"}) if sen_token else None

        # 第三步：获取 Sentinel SO Token (oauth_create_account)
        so_token = fetch_sentinel_token(flow="oauth_create_account", did=did, proxies=proxies)

        # 第四步：提交邮箱授权
        signup_headers = {"referer": "https://auth.openai.com/create-account", "accept": "application/json", "content-type": "application/json"}
        if sentinel: signup_headers["openai-sentinel-token"] = sentinel
        signup_resp = s.post("https://auth.openai.com/api/accounts/authorize/continue", headers=signup_headers, data=json.dumps({"username": {"value": openai_email, "kind": "email"}, "screen_hint": "signup"}))
        if signup_resp.status_code != 200:
            log(f"[Error] 提交邮箱失败: {signup_resp.status_code}")
            return None

        # 第五步：设置密码
        register_headers = {"referer": "https://auth.openai.com/create-account/password", "accept": "application/json", "content-type": "application/json"}
        if sentinel: register_headers["openai-sentinel-token"] = sentinel
        reg_resp = s.post("https://auth.openai.com/api/accounts/user/register", headers=register_headers, data=json.dumps({"password": openai_password, "username": openai_email}))
        if reg_resp.status_code != 200:
            log(f"[Error] 设置密码失败: {reg_resp.status_code}")
            return None

        # 第六步：触发验证码发送
        s.get("https://auth.openai.com/api/accounts/email-otp/send", headers=register_headers, timeout=15)
        log("[*] 已触发验证码发送，开始轮询邮箱...")

        # 第七步：提取验证码 (Graph API -> IMAP OAuth2)
        code = fetcher.get_otp()
        if not code:
            log("[Error] 验证码等待超时或提取失败")
            return None
        log(f"[*] 成功提取验证码: {code}")

        # 第八步：校验验证码
        validate_headers = {"referer": "https://auth.openai.com/email-verification", "accept": "application/json", "content-type": "application/json"}
        if sentinel: validate_headers["openai-sentinel-token"] = sentinel
        code_resp = s.post("https://auth.openai.com/api/accounts/email-otp/validate", headers=validate_headers, data=json.dumps({"code": code}))
        if code_resp.status_code != 200:
            log(f"[Error] 验证码校验失败: {code_resp.status_code}")
            return None

        # 第九步：完成账号注册填写
        create_headers = {"referer": "https://auth.openai.com/about-you", "accept": "application/json", "content-type": "application/json"}
        if so_token: create_headers["openai-sentinel-so-token"] = so_token
        create_resp = s.post("https://auth.openai.com/api/accounts/create_account", headers=create_headers, data=json.dumps({"name": _random_name(), "birthdate": _random_birthdate()}))
        if create_resp.status_code != 200:
            log(f"[Error] 账户信息填写失败: {create_resp.status_code}")
            return None

        # 第十步：选择工作区 Workspace
        auth_cookie = s.cookies.get("oai-client-auth-session")
        if not auth_cookie: return None
        auth_json = _decode_jwt_segment(auth_cookie.split(".")[0])
        workspace_id = str((auth_json.get("workspaces") or [{}])[0].get("id") or "").strip()
        
        select_resp = s.post("https://auth.openai.com/api/accounts/workspace/select", headers={"referer": "https://auth.openai.com/sign-in-with-chatgpt/codex/consent", "content-type": "application/json"}, data=json.dumps({"workspace_id": workspace_id}))
        if select_resp.status_code != 200: return None
        
        continue_url = str((select_resp.json() or {}).get("continue_url") or "").strip()

        # 第十一步：拦截重定向，提取终极 Token
        current_url = continue_url
        for _ in range(6):
            final_resp = s.get(current_url, allow_redirects=False, timeout=15)
            location = final_resp.headers.get("Location") or ""
            if final_resp.status_code not in [301, 302, 303, 307, 308] or not location:
                break
            next_url = urllib.parse.urljoin(current_url, location)
            if "code=" in next_url and "state=" in next_url:
                token_json = submit_callback_url(callback_url=next_url, code_verifier=oauth.code_verifier, redirect_uri=oauth.redirect_uri, expected_state=oauth.state)
                return token_json, openai_email, openai_password
            current_url = next_url

        log("[Error] 未能在重定向链中捕获到最终 Token")
        return None

    except Exception as e:
        log(f"[Error] 运行时异常: {e}")
        return None


# ========== 4. Tkinter GUI (支持拖拽) ==========

def _create_root():
    """创建支持拖拽的主窗口，未安装 tkinterdnd2 时自动降级"""
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
        root._dnd_enabled = True
        return root
    except ImportError:
        root = tk.Tk()
        root._dnd_enabled = False
        return root


class App:
    def __init__(self):
        self.root = _create_root()
        self.root.title("OpenAI Codex 自动注册工具 (微软邮箱接码)")
        self.root.geometry("900x720")
        self.root.resizable(True, True)
        self.root.configure(bg="#f0f0f0")

        self._running = False
        self._stop_event = threading.Event()

        self._build_ui()
        self._setup_drag_drop()

    # ---------- 界面构建 ----------

    def _build_ui(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TButton", font=("Microsoft YaHei UI", 9))
        style.configure("TLabel", font=("Microsoft YaHei UI", 9))

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # ---- 邮箱数据区 ----
        acc_frame = ttk.LabelFrame(main, text="邮箱数据 (每行一个: 账号----密码----clientid----refresh_token，支持粘贴/拖拽文件)", padding=6)
        acc_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        # 邮箱文本框
        self._acc_text = scrolledtext.ScrolledText(acc_frame, wrap=tk.NONE, font=("Consolas", 9),
                                                    height=8, bg="#fff", fg="#333", insertbackground="#333")
        self._acc_text.pack(fill=tk.BOTH, expand=True)

        # 邮箱区按钮行
        acc_btn_row = ttk.Frame(acc_frame)
        acc_btn_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(acc_btn_row, text="导入文件", command=self._import_file).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(acc_btn_row, text="清空", command=self._clear_acc_text).pack(side=tk.LEFT, padx=(0, 6))
        self._acc_count_var = tk.StringVar(value="")
        ttk.Label(acc_btn_row, textvariable=self._acc_count_var, foreground="gray").pack(side=tk.LEFT, padx=(6, 0))
        self._acc_text.bind("<KeyRelease>", lambda e: self._update_acc_count())

        # ---- 代理行 ----
        proxy_frame = ttk.LabelFrame(main, text="代理配置", padding=6)
        proxy_frame.pack(fill=tk.X, pady=(0, 6))
        proxy_row = ttk.Frame(proxy_frame)
        proxy_row.pack(fill=tk.X)
        ttk.Label(proxy_row, text="代理地址:").pack(side=tk.LEFT)
        self._proxy_var = tk.StringVar()
        ttk.Entry(proxy_row, textvariable=self._proxy_var, width=45).pack(side=tk.LEFT, padx=(4, 4))
        ttk.Label(proxy_row, text="(可选, 如 http://127.0.0.1:7890)", foreground="gray").pack(side=tk.LEFT)

        # ---- 控制按钮 ----
        ctrl_frame = ttk.Frame(main)
        ctrl_frame.pack(fill=tk.X, pady=6)
        self._start_btn = ttk.Button(ctrl_frame, text="开始注册", command=self._start)
        self._start_btn.pack(side=tk.LEFT, padx=(0, 6))
        self._stop_btn = ttk.Button(ctrl_frame, text="停止", command=self._stop, state=tk.DISABLED)
        self._stop_btn.pack(side=tk.LEFT, padx=(0, 6))
        self._test_btn = ttk.Button(ctrl_frame, text="测试 (第一行/本地IP)", command=self._test)
        self._test_btn.pack(side=tk.LEFT, padx=(0, 6))
        self._clear_log_btn = ttk.Button(ctrl_frame, text="清空日志", command=self._clear_log)
        self._clear_log_btn.pack(side=tk.LEFT)

        self._status_var = tk.StringVar(value="就绪")
        ttk.Label(ctrl_frame, textvariable=self._status_var, foreground="blue").pack(side=tk.RIGHT)

        # ---- 日志 ----
        log_frame = ttk.LabelFrame(main, text="运行日志", padding=4)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self._log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 9),
                                                    state=tk.DISABLED, bg="#1e1e1e", fg="#d4d4d4",
                                                    insertbackground="#d4d4d4")
        self._log_text.pack(fill=tk.BOTH, expand=True)

    # ---------- 拖拽支持 ----------

    def _setup_drag_drop(self):
        if not getattr(self.root, "_dnd_enabled", False):
            return
        try:
            from tkinterdnd2 import DND_FILES
            self._acc_text.drop_target_register(DND_FILES)
            self._acc_text.dnd_bind("<<Drop>>", self._handle_drop)
        except Exception:
            pass

    def _handle_drop(self, event):
        """拖拽文件到文本框时，读取文件内容追加进去"""
        try:
            files = event.widget.tk.splitlist(event.data)
            if not files:
                return
            file_path = files[0].strip("{}")
            if not os.path.isfile(file_path):
                return
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()
            current = self._acc_text.get("1.0", tk.END).strip()
            if current:
                self._acc_text.insert(tk.END, "\n")
            self._acc_text.insert(tk.END, content.strip())
            self._update_acc_count()
            self._log(f"[*] 已拖入文件: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"读取拖入文件失败: {e}")

    # ---------- 导入 / 清空 ----------

    def _import_file(self):
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
            current = self._acc_text.get("1.0", tk.END).strip()
            if current:
                self._acc_text.insert(tk.END, "\n")
            self._acc_text.insert(tk.END, content.strip())
            self._update_acc_count()
            self._log(f"[*] 已导入文件: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {e}")

    def _clear_acc_text(self):
        self._acc_text.delete("1.0", tk.END)
        self._update_acc_count()

    def _update_acc_count(self):
        accounts = self._parse_accounts()
        self._acc_count_var.set(f"有效行: {len(accounts)}")

    # ---------- 解析文本框内容 ----------

    def _parse_accounts(self) -> List[Dict]:
        text = self._acc_text.get("1.0", tk.END)
        accounts = []
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split("----")
            if len(parts) >= 4:
                accounts.append({
                    "email": parts[0].strip(),
                    "password": parts[1].strip(),
                    "client_id": parts[2].strip(),
                    "refresh_token": parts[3].strip(),
                })
        return accounts

    # ---------- 日志 ----------

    def _log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        def _append():
            self._log_text.configure(state=tk.NORMAL)
            self._log_text.insert(tk.END, f"[{ts}] {msg}\n")
            self._log_text.see(tk.END)
            self._log_text.configure(state=tk.DISABLED)
        if threading.current_thread() is threading.main_thread():
            _append()
        else:
            self.root.after(0, _append)

    def _clear_log(self):
        self._log_text.configure(state=tk.NORMAL)
        self._log_text.delete("1.0", tk.END)
        self._log_text.configure(state=tk.DISABLED)

    # ---------- 测试按钮 ----------

    def _test(self):
        if self._running:
            messagebox.showwarning("提示", "当前有任务正在运行")
            return
        accounts = self._parse_accounts()
        if not accounts:
            messagebox.showwarning("提示", "文本框中没有有效的邮箱数据")
            return
        acc = accounts[0]
        self._running = True
        self._stop_event.clear()
        self._test_btn.configure(state=tk.DISABLED)
        self._start_btn.configure(state=tk.DISABLED)
        self._status_var.set("测试中 (本地IP)...")
        self._log(f"[测试] 使用第一行数据: {acc['email']}，不使用代理")
        threading.Thread(target=self._test_worker, args=(acc,), daemon=True).start()

    def _test_worker(self, acc: Dict):
        tokens_dir = OUT_DIR / "tokens"
        tokens_dir.mkdir(parents=True, exist_ok=True)

        result = run(
            ms_email=acc["email"],
            ms_password=acc["password"],
            client_id=acc["client_id"],
            refresh_token=acc["refresh_token"],
            proxy=None,
            log_func=self._log,
        )

        if result:
            token_json, reg_email, reg_password = result
            fname = reg_email.replace("@", "_")
            file_path = tokens_dir / f"token_{fname}_{int(time.time())}.json"
            file_path.write_text(token_json, encoding="utf-8")
            self._log(f"[测试OK] Token 已保存: {file_path}")
            acc_file = tokens_dir / "accounts.txt"
            with open(acc_file, "a", encoding="utf-8") as f:
                f.write(f"{reg_email}----{reg_password}\n")
            self._log(f"[测试OK] 账号已追加: {acc_file}")
        else:
            self._log("[测试失败] 注册未成功")

        self.root.after(0, self._on_finished)

    # ---------- 开始 / 停止 ----------

    def _start(self):
        accounts = self._parse_accounts()
        if not accounts:
            messagebox.showwarning("提示", "文本框中没有有效的邮箱数据")
            return
        self._running = True
        self._stop_event.clear()
        self._start_btn.configure(state=tk.DISABLED)
        self._test_btn.configure(state=tk.DISABLED)
        self._stop_btn.configure(state=tk.NORMAL)
        self._status_var.set("运行中...")
        threading.Thread(target=self._worker, args=(accounts,), daemon=True).start()

    def _stop(self):
        self._stop_event.set()
        self._status_var.set("正在停止...")
        self._log("[*] 用户请求停止，将在当前任务完成后停止")

    def _on_finished(self):
        self._running = False
        self._start_btn.configure(state=tk.NORMAL)
        self._test_btn.configure(state=tk.NORMAL)
        self._stop_btn.configure(state=tk.DISABLED)
        self._status_var.set("就绪")

    # ---------- 工作线程 ----------

    def _worker(self, accounts: List[Dict]):
        proxy = self._proxy_var.get().strip() or None
        tokens_dir = OUT_DIR / "tokens"
        tokens_dir.mkdir(parents=True, exist_ok=True)

        total = len(accounts)
        success_count = 0
        fail_count = 0

        for idx, acc in enumerate(accounts, 1):
            if self._stop_event.is_set():
                self._log("[*] 已停止")
                break

            self._log(f"\n{'='*50}")
            self._log(f"[{idx}/{total}] 开始注册: {acc['email']}")
            self._log(f"{'='*50}")

            result = run(
                ms_email=acc["email"],
                ms_password=acc["password"],
                client_id=acc["client_id"],
                refresh_token=acc["refresh_token"],
                proxy=proxy,
                log_func=self._log,
            )

            if result:
                token_json, reg_email, reg_password = result
                success_count += 1
                fname = reg_email.replace("@", "_")
                file_path = tokens_dir / f"token_{fname}_{int(time.time())}.json"
                file_path.write_text(token_json, encoding="utf-8")
                self._log(f"[OK] Token 已保存: {file_path}")

                acc_file = tokens_dir / "accounts.txt"
                with open(acc_file, "a", encoding="utf-8") as f:
                    f.write(f"{reg_email}----{reg_password}\n")
                self._log(f"[OK] 账号已追加: {acc_file}")
            else:
                fail_count += 1
                self._log(f"[-] 注册失败: {acc['email']}")

            # 冷却
            if idx < total and not self._stop_event.is_set():
                wait_time = random.randint(5, 15)
                self._log(f"[*] 冷却 {wait_time} 秒...")
                for _ in range(wait_time):
                    if self._stop_event.is_set():
                        break
                    time.sleep(1)

        self._log(f"\n{'='*50}")
        self._log(f"[完成] 成功: {success_count}, 失败: {fail_count}, 总计: {total}")
        self._log(f"{'='*50}")
        self.root.after(0, self._on_finished)


if __name__ == "__main__":
    app = App()
    app.root.mainloop()
