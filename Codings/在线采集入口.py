"""Interactive Bilibili collector that asks only for a URL and a Cookie."""

from __future__ import annotations

import getpass
from datetime import datetime

from b站评论区爬虫 import BilibiliCommentsSpider
from project_config import ORIGINAL_DATA_DIR


def _read_required(prompt: str, *, hidden: bool = False) -> str:
    reader = getpass.getpass if hidden else input
    while True:
        value = reader(prompt).strip().strip('"').strip("'")
        if value:
            return value
        print("输入不能为空，请重新输入。")


def _normalize_cookie(cookie: str) -> str:
    # Allow pasting either the cookie value or a copied `Cookie: ...` header.
    if cookie.lower().startswith("cookie:"):
        return cookie.split(":", 1)[1].strip()
    return cookie


def main() -> None:
    print("=" * 58)
    print("B站评论快捷采集（Cookie 仅驻留本次进程，不会保存）")
    print("=" * 58)
    video_url = _read_required("1/2 请输入B站视频URL：")
    cookie = _normalize_cookie(_read_required("2/2 请粘贴Cookie（明文输入）："))

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = ORIGINAL_DATA_DIR / f"在线采集评论_{timestamp}.xlsx"

    print("\n开始采集，请保持网络连接……")
    try:
        comments = BilibiliCommentsSpider(
            video_url=video_url,
            save_path=output_path,
            cookie=cookie,
        ).start()
    except KeyboardInterrupt:
        raise SystemExit("\n采集已由用户终止。") from None
    except Exception as exc:
        raise SystemExit(
            "\n采集失败："
            f"{exc}\n可能原因：Cookie失效、URL格式错误、网络异常，或B站接口已调整。"
        ) from exc

    print("\n采集完成。")
    print(f"去重后评论数：{len({item['评论'] for item in comments})}")
    print(f"Excel保存位置：{output_path}")
    print("后续分析可运行：python Codings/main.py offline --input \"上面的Excel路径\"")


if __name__ == "__main__":
    main()
