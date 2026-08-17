"""Best-effort Bilibili top-level comment collector.

The analysis workflow is offline-first. This collector is optional because
Bilibili may change request signing and pagination without notice.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import random
import re
import time
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests

from project_config import DEFAULT_INPUT_FILE, get_bilibili_cookie


class BilibiliCommentsSpider:
    API_URL = "https://api.bilibili.com/x/v2/reply/wbi/main"
    LEGACY_WBI_SUFFIX = "ea1db124af3c7062474693fa704f4ff8"

    def __init__(
        self,
        video_url: str,
        save_path: str | Path = DEFAULT_INPUT_FILE,
        cookie: str | None = None,
        max_pages: int = 1000,
    ) -> None:
        if not video_url:
            raise ValueError("必须提供 B站视频地址")
        self.video_url = video_url
        self.save_path = Path(save_path).expanduser().resolve()
        self.cookie = (cookie if cookie is not None else get_bilibili_cookie()).strip()
        self.max_pages = max_pages
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 Chrome/127.0 Safari/537.36"
                )
            }
        )
        if self.cookie:
            self.session.headers.update({"Cookie": self.cookie})

    def get_aid(self) -> int:#获取aid
        video_url = self.video_url
        if "b23.tv" in video_url:
            video_url = self.session.head(video_url, allow_redirects=True, timeout=10).url
        bvid_match = re.search(r"video/(BV[0-9A-Za-z]+)", video_url)
        avid_match = re.search(r"video/av(\d+)", video_url, flags=re.IGNORECASE)
        if bvid_match:
            endpoint = "https://api.bilibili.com/x/web-interface/view"
            params = {"bvid": bvid_match.group(1)}
        elif avid_match:
            endpoint = "https://api.bilibili.com/x/web-interface/view"
            params = {"aid": avid_match.group(1)}
        else:
            raise ValueError("无法从地址中识别 BV 号或 av 号")
        response = self.session.get(endpoint, params=params, timeout=10)
        response.raise_for_status()
        payload = response.json()
        if payload.get("code") != 0:
            raise RuntimeError(f"视频信息接口返回异常: {payload.get('message')}")
        return int(payload["data"]["aid"])

    @staticmethod
    def _md5(value: str) -> str:
        return hashlib.md5(value.encode("utf-8")).hexdigest()

    def _legacy_w_rid(self, params: dict[str, str]) -> str:
        """Retain the course-project signer as a best-effort compatibility path."""

        order = [
            "mode",
            "oid",
            "pagination_str",
            "plat",
            "seek_rpid",
            "type",
            "web_location",
            "wts",
        ]
        parts = []
        for key in order:
            if key in params:
                value = quote(params[key]) if key == "pagination_str" else params[key]
                parts.append(f"{key}={value}")
        return self._md5("&".join(parts) + self.LEGACY_WBI_SUFFIX)

    @staticmethod
    def _parse_comments(payload: dict) -> list[dict[str, str]]:
        replies = payload.get("data", {}).get("replies") or []
        return [
            {
                "昵称": str(reply.get("member", {}).get("uname", "")),
                "性别": str(reply.get("member", {}).get("sex", "保密")),
                "评论": str(reply.get("content", {}).get("message", "")),
            }
            for reply in replies
            if reply.get("content", {}).get("message")
        ]

    def _save(self, comments: list[dict[str, str]]) -> None:
        self.save_path.parent.mkdir(parents=True, exist_ok=True)
        frame = pd.DataFrame(comments).drop_duplicates(subset=["评论"], keep="first")
        if self.save_path.suffix.lower() == ".csv":
            frame.to_csv(self.save_path, index=False, encoding="utf-8-sig")
        else:
            frame.to_excel(self.save_path, index=False)

    def start(self) -> list[dict[str, str]]:
        aid = self.get_aid()
        offset = ""
        comments: list[dict[str, str]] = []
        for page in range(self.max_pages):
            now = str(int(time.time()))
            params = {
                "oid": str(aid),
                "type": "1",
                "mode": "3",
                "pagination_str": json.dumps({"offset": offset}, ensure_ascii=False),
                "plat": "1",
                "web_location": "1315875",
                "wts": now,
            }
            if page == 0:
                params["seek_rpid"] = ""
            params["w_rid"] = self._legacy_w_rid(params)
            response = self.session.get(self.API_URL, params=params, timeout=10)
            response.raise_for_status()
            payload = response.json()
            if payload.get("code") != 0:
                raise RuntimeError(f"评论接口返回异常: {payload.get('message')}")
            batch = self._parse_comments(payload)
            comments.extend(batch)
            cursor = payload.get("data", {}).get("cursor", {})
            next_offset = cursor.get("pagination_reply", {}).get("next_offset")
            print(f"已采集第 {page + 1} 页，累计 {len(comments)} 条")
            if not next_offset or not batch:
                break
            offset = str(next_offset)
            if (page + 1) % 10 == 0:
                self._save(comments)
            time.sleep(random.uniform(1.0, 2.5))
        self._save(comments)
        return comments


# Backward-compatible name used by the original main.py.
bilibili_comments_spider = BilibiliCommentsSpider


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="可选的 B站顶层评论采集器")
    parser.add_argument("--url", required=True, help="B站视频 URL")
    parser.add_argument("--output", default=str(DEFAULT_INPUT_FILE), help="保存为 xlsx 或 csv")
    parser.add_argument("--max-pages", type=int, default=1000)
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    BilibiliCommentsSpider(args.url, args.output, max_pages=args.max_pages).start()
