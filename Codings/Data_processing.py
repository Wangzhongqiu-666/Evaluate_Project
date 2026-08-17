"""Offline data cleaning, domain sentiment scoring and keyword extraction."""

from __future__ import annotations

import json
import re
from collections import Counter
from pathlib import Path
from typing import Iterable

import jieba
import pandas as pd
from snownlp import SnowNLP, sentiment

from project_config import (
    CUSTOM_MODEL_PREFIX,
    DEFAULT_INPUT_FILE,
    NEGATIVE_THRESHOLD,
    POSITIVE_THRESHOLD,
    STOPWORDS_FILE,
    USER_DICT_FILE,
    ensure_output_dir,
)


REQUIRED_COLUMNS = {"昵称", "性别", "评论"}
TOKEN_PATTERN = re.compile(r"^[\u4e00-\u9fffa-zA-Z0-9]+$")
TEXT_CLEAN_PATTERN = re.compile(r"[^\u4e00-\u9fffa-zA-Z0-9]+")


def classify_score(score: float) -> int:
    if score >= POSITIVE_THRESHOLD:
        return 1
    if score <= NEGATIVE_THRESHOLD:
        return -1
    return 0


class DataProcessing:
    """Process one Excel file or a directory of Excel files without network access."""

    def __init__(
        self,
        input_path: str | Path = DEFAULT_INPUT_FILE,
        output_dir: str | Path | None = None,
        model_prefix: str | Path = CUSTOM_MODEL_PREFIX,
    ) -> None:
        self.input_path = Path(input_path).expanduser().resolve()
        self.output_dir = ensure_output_dir(output_dir)
        self.model_prefix = Path(model_prefix).expanduser().resolve()
        self.data = self._load_input()
        self.raw_record_count = len(self.data)
        self.duplicate_count = int(self.data.duplicated(subset=["评论"]).sum())
        self.data = self.data.drop_duplicates(subset=["评论"], keep="first").copy()
        self.data.insert(0, "原始序号", range(1, len(self.data) + 1))
        self.keywords = pd.DataFrame(columns=["关键词", "频数"])
        self.summary: dict[str, object] = {}

    def _input_files(self) -> list[Path]:
        if self.input_path.is_file():
            return [self.input_path]
        if self.input_path.is_dir():
            return sorted(
                p
                for p in self.input_path.iterdir()
                if p.suffix.lower() in {".xlsx", ".xls"} and not p.name.startswith("~$")
            )
        raise FileNotFoundError(f"找不到输入路径: {self.input_path}")

    def _load_input(self) -> pd.DataFrame:
        frames: list[pd.DataFrame] = []
        for file_path in self._input_files():
            frame = pd.read_excel(file_path)
            missing = REQUIRED_COLUMNS - set(frame.columns)
            if missing:
                raise ValueError(f"{file_path.name} 缺少必要列: {sorted(missing)}")
            frame = frame[["昵称", "性别", "评论"]].copy()
            frame["来源文件"] = file_path.name
            frames.append(frame)
        if not frames:
            raise ValueError("输入目录中没有可处理的 Excel 文件")
        data = pd.concat(frames, ignore_index=True)
        data = data.dropna(subset=["评论"]).copy()
        data["评论"] = data["评论"].astype(str).str.strip()
        return data[data["评论"] != ""].reset_index(drop=True)

    @staticmethod
    def clean_text(text: object) -> str:
        return TEXT_CLEAN_PATTERN.sub("", str(text)).strip()

    @staticmethod
    def _load_stopwords(path: Path = STOPWORDS_FILE) -> set[str]:
        frame = pd.read_excel(path)
        if "停用词" not in frame.columns:
            raise ValueError(f"停用词表缺少“停用词”列: {path}")
        return {str(value).strip() for value in frame["停用词"].dropna() if str(value).strip()}

    @staticmethod
    def _valid_token(token: str, stopwords: set[str]) -> bool:
        token = token.strip().lower()
        if token in stopwords or len(token) < 2 or token.isdigit():
            return False
        return bool(TOKEN_PATTERN.fullmatch(token))

    def score_sentiment(self) -> None:
        actual_model = Path(f"{self.model_prefix}.3")
        if not actual_model.exists():
            raise FileNotFoundError(f"找不到领域情感模型: {actual_model}")
        sentiment.load(str(self.model_prefix))
        scores = [round(float(SnowNLP(text).sentiments), 4) for text in self.data["清洗后评论"]]
        self.data["领域模型分数"] = scores
        self.data["情绪标签"] = [classify_score(score) for score in scores]

    def extract_keywords(self) -> None:
        if USER_DICT_FILE.exists():
            jieba.load_userdict(str(USER_DICT_FILE))
        stopwords = self._load_stopwords()
        counts: Counter[str] = Counter()
        for text in self.data["清洗后评论"]:
            tokens = jieba.lcut(text, cut_all=False)
            counts.update(token for token in tokens if self._valid_token(token, stopwords))
        self.keywords = pd.DataFrame(counts.most_common(), columns=["关键词", "频数"])

    def _build_summary(self) -> dict[str, object]:
        counts = self.data["情绪标签"].value_counts().reindex([1, 0, -1], fill_value=0)
        total = int(len(self.data))
        positive = int(counts.loc[1])
        neutral = int(counts.loc[0])
        negative = int(counts.loc[-1])
        positive_share = positive / total if total else 0.0
        negative_share = negative / total if total else 0.0
        return {
            "raw_records": self.raw_record_count,
            "unique_records": total,
            "duplicates_removed": self.duplicate_count,
            "positive": positive,
            "neutral": neutral,
            "negative": negative,
            "positive_share": round(positive_share, 6),
            "neutral_share": round(neutral / total if total else 0.0, 6),
            "negative_share": round(negative_share, 6),
            "net_sentiment_points": round((positive_share - negative_share) * 100, 2),
            "unique_keywords": int(len(self.keywords)),
            "keyword_occurrences": int(self.keywords["频数"].sum()) if not self.keywords.empty else 0,
            "thresholds": {
                "negative_max": NEGATIVE_THRESHOLD,
                "positive_min": POSITIVE_THRESHOLD,
            },
        }

    def save_outputs(self) -> None:
        self.data.to_csv(
            self.output_dir / "评论分析结果_保研版.csv", index=False, encoding="utf-8-sig"
        )
        self.keywords.to_csv(
            self.output_dir / "关键词统计_保研版.csv", index=False, encoding="utf-8-sig"
        )
        with (self.output_dir / "分析摘要.json").open("w", encoding="utf-8") as stream:
            json.dump(self.summary, stream, ensure_ascii=False, indent=2)

    def start(self) -> dict[str, object]:
        self.data["清洗后评论"] = self.data["评论"].map(self.clean_text)
        self.data = self.data[self.data["清洗后评论"] != ""].reset_index(drop=True)
        self.score_sentiment()
        self.extract_keywords()
        self.summary = self._build_summary()
        self.save_outputs()
        return self.summary


def load_processed_rows(path: str | Path) -> Iterable[dict[str, object]]:
    return pd.read_csv(path, encoding="utf-8-sig").to_dict(orient="records")


if __name__ == "__main__":
    result = DataProcessing().start()
    print(json.dumps(result, ensure_ascii=False, indent=2))
