"""示例测试脚本：演示 license_manager 的使用流程。"""

from __future__ import annotations

import json
from pathlib import Path

import license_manager


def main() -> None:
    ok, message = license_manager.verify_license()
    print(f"初次检查: {ok}, {message}")

    if not ok:
        print("尝试激活（请使用测试口令 labpass）...")
        ok, message = license_manager.activate_if_needed()
        print(f"激活结果: {ok}, {message}")

    ok, message = license_manager.verify_license()
    print(f"再次检查: {ok}, {message}")

    if ok:
        data_path = Path(license_manager.LICENSE_PATH)
        if data_path.exists():
            print("license 内容预览:")
            try:
                content = json.loads(data_path.read_text(encoding="utf-8"))
                print(json.dumps(content, indent=2, ensure_ascii=False))
            except Exception as exc:
                print(f"读取 license 失败: {exc}")


if __name__ == "__main__":
    main()
