"""离线授权管理模块（够用方案，不对抗专业逆向）。"""

from __future__ import annotations

import getpass
import hashlib
import hmac
import json
import os
import platform
import subprocess
import sys
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Tuple

# 常量定义（按需修改）
APP_NAME = "Oligosaccharide prediction"
DEFAULT_DAYS_VALID = 365
EXPECTED_PW_SHA256 = "a5e8ce9ff4e56d3279072e8c289b1fcdfe3527e387e03b85e9eb28c099da0cdf"  # maplegao
SECRET_HEX = (
    "5fdd9b7d7a3c4d8bb5c6c63102f4a962"
    "79a0781df3aacf2fb4d5f70c6f52880a"
)  # 可按需替换


def _hex_to_bytes(value: str) -> bytes:
    try:
        return bytes.fromhex(value)
    except ValueError as exc:  # pragma: no cover
        raise ValueError("SECRET_HEX 必须为合法的十六进制字符串") from exc


SECRET_BYTES = _hex_to_bytes(SECRET_HEX)


def _get_machine_id() -> Tuple[bool, str]:
    """获取机器 ID，按平台策略逐项尝试。"""
    system_name = platform.system().lower()
    try_getters = []

    if system_name.startswith("win"):
        try_getters.append(_get_machine_guid_windows)
    elif system_name == "linux":
        try_getters.append(_get_machine_id_linux)
    elif system_name == "darwin":
        try_getters.append(_get_machine_uuid_macos)
    else:
        # 未知系统仍尝试通用条目
        try_getters.extend(
            [_get_machine_guid_windows, _get_machine_id_linux, _get_machine_uuid_macos]
        )

    for getter in try_getters:
        ok, value = getter()
        if ok and value:
            return True, value

    # 回退 MAC 地址
    mac = uuid.getnode()
    mac_text = ":".join(f"{(mac >> ele) & 0xFF:02X}" for ele in range(40, -8, -8))
    return True, mac_text


def _get_machine_guid_windows() -> Tuple[bool, str]:
    try:
        import winreg  # type: ignore[attr-defined]

        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Cryptography",
            0,
            winreg.KEY_READ | winreg.KEY_WOW64_64KEY,
        ) as key:
            value, _ = winreg.QueryValueEx(key, "MachineGuid")
            return True, str(value)
    except Exception:
        return False, ""


def _get_machine_id_linux() -> Tuple[bool, str]:
    try:
        machine_id = Path("/etc/machine-id").read_text(encoding="utf-8").strip()
        return bool(machine_id), machine_id
    except Exception:
        return False, ""


def _get_machine_uuid_macos() -> Tuple[bool, str]:
    try:
        result = subprocess.run(
            ["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            return False, ""
        for line in result.stdout.splitlines():
            if "IOPlatformUUID" in line:
                parts = line.split("=")
                if len(parts) == 2:
                    return True, parts[1].strip().strip('"')
    except Exception:
        pass
    return False, ""


def _resolve_license_path() -> Path:
    """根据平台计算 license.json 存放路径。"""
    system_name = platform.system().lower()
    candidate_dirs = []

    if system_name.startswith("win"):
        program_data = os.getenv("PROGRAMDATA")
        if program_data:
            candidate_dirs.append(Path(program_data) / APP_NAME)
    else:
        home = Path.home()
        candidate_dirs.append(home / ".local" / "share" / APP_NAME)

    candidate_dirs.append(Path.cwd())

    for directory in candidate_dirs:
        try:
            directory.mkdir(parents=True, exist_ok=True)
            return directory / "license.json"
        except Exception:
            continue

    # 最后兜底当前目录，不再 mkdir
    return Path("license.json")


LICENSE_PATH = _resolve_license_path()


def _now_utc() -> datetime:
    return datetime.now(timezone.utc)


def _format_time(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_time(value: str) -> datetime:
    return datetime.strptime(value, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)


def _sign_payload(payload: Dict[str, Any]) -> str:
    payload_text = json.dumps(payload, sort_keys=True, separators=(",", ":"))
    digest = hmac.new(SECRET_BYTES, payload_text.encode("utf-8"), hashlib.sha256)
    return digest.hexdigest()


def _read_license(path: Path) -> Tuple[bool, Dict[str, Any], str]:
    try:
        raw = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return False, {}, "license 文件不存在"
    except Exception as exc:
        return False, {}, f"读取 license 失败: {exc}"

    try:
        data = json.loads(raw)
        payload = data.get("payload")
        signature = data.get("signature")
        if not isinstance(payload, dict) or not isinstance(signature, str):
            return False, {}, "license 格式无效"
        return True, {"payload": payload, "signature": signature}, ""
    except Exception as exc:
        return False, {}, f"解析 license 失败: {exc}"


def _write_license(path: Path, payload: Dict[str, Any]) -> Tuple[bool, str]:
    signature = _sign_payload(payload)
    content = {"payload": payload, "signature": signature}
    try:
        path.write_text(json.dumps(content, indent=2, ensure_ascii=False), encoding="utf-8")
        return True, "授权文件已写入"
    except Exception as exc:
        return False, f"写入 license 失败: {exc}"


def verify_license(path: str | Path = LICENSE_PATH) -> Tuple[bool, str]:
    """验证本地 license 有效性，不抛异常。"""
    license_path = Path(path)
    ok, data, message = _read_license(license_path)
    if not ok:
        return False, message

    payload = data["payload"]
    signature = data["signature"]
    expected_signature = _sign_payload(payload)
    if not hmac.compare_digest(signature, expected_signature):
        return False, "license 签名不匹配"

    if payload.get("app") != APP_NAME or payload.get("version") != 1:
        return False, "license 与当前应用不匹配"

    ok_mid, machine_id = _get_machine_id()
    if not ok_mid:
        return False, "无法获取机器信息"

    if payload.get("machine_id") != machine_id:
        return False, "机器 ID 不匹配"

    expires_at_text = payload.get("expires_at")
    try:
        expires_at = _parse_time(expires_at_text)
    except Exception:
        return False, "license 时间格式无效"

    if _now_utc() > expires_at:
        return False, "license 已过期"

    return True, "授权有效"


def _prompt_password() -> str:
    """优先使用命令行读取口令，失败时回退到 Tk 输入框。"""
    try:
        if sys.stdin and sys.stdin.isatty():
            password = getpass.getpass("请输入激活口令: ")
            if password:
                return password
    except Exception:
        pass

    try:
        import tkinter as tk

        result: Dict[str, str] = {"value": ""}

        def submit(_: Any = None) -> None:
            result["value"] = entry.get()
            window.destroy()

        def cancel() -> None:
            result["value"] = ""
            window.destroy()

        window = tk.Tk()
        window.title("License Activation")
        window.resizable(False, False)
        window.attributes("-topmost", True)
        window.after(200, lambda: window.attributes("-topmost", False))

        frame = tk.Frame(window, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        label = tk.Label(frame, text="请输入激活口令")
        label.pack(anchor="w")

        entry = tk.Entry(frame, show="*")
        entry.pack(fill="x", pady=(8, 12))
        entry.focus_set()

        button_row = tk.Frame(frame)
        button_row.pack(fill="x")

        ok_btn = tk.Button(button_row, text="确定", width=10, command=submit)
        ok_btn.pack(side="left", padx=(0, 8))

        cancel_btn = tk.Button(button_row, text="取消", width=10, command=cancel)
        cancel_btn.pack(side="left")

        window.bind("<Return>", submit)
        window.bind("<Escape>", lambda _event: cancel())

        window.mainloop()
        return result["value"]
    except Exception:
        return ""


def _password_valid(password: str) -> bool:
    password_hash = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return hmac.compare_digest(password_hash, EXPECTED_PW_SHA256)


def activate_if_needed(days: int = DEFAULT_DAYS_VALID) -> Tuple[bool, str]:
    """如有必要执行激活流程，否则直接返回已有授权状态。"""
    ok, message = verify_license()
    if ok:
        return True, message

    password = _prompt_password()
    if not password:
        return False, "未输入口令"

    if not _password_valid(password):
        return False, "口令验证失败"

    ok_mid, machine_id = _get_machine_id()
    if not ok_mid:
        return False, "无法获取机器信息"

    issued_at = _now_utc()
    expires_at = issued_at + timedelta(days=days)
    payload = {
        "machine_id": machine_id,
        "issued_at": _format_time(issued_at),
        "expires_at": _format_time(expires_at),
        "app": APP_NAME,
        "version": 1,
    }

    ok_write, msg = _write_license(LICENSE_PATH, payload)
    if not ok_write:
        return False, msg

    return True, "授权生成成功"


def _main() -> int:
    ok, message = activate_if_needed()
    if ok:
        print("授权有效，正常启动。")
        return 0
    print(f"授权失败：{message}")
    return 1


if __name__ == "__main__":  # pragma: no cover
    sys.exit(_main())
