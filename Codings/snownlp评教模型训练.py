"""Retrain the course-domain SnowNLP model from local positive/negative corpora."""

from __future__ import annotations

import argparse
import time
from pathlib import Path

from snownlp import SnowNLP, sentiment

from project_config import CUSTOM_MODEL_PREFIX, FUNCTIONAL_DIR


DEFAULT_CORPUS_DIR = FUNCTIONAL_DIR / "针对评教的SnowNLP模型"


def train_model(corpus_dir: str | Path, model_prefix: str | Path) -> None:
    corpus = Path(corpus_dir).expanduser().resolve()
    negative = corpus / "neg.txt"
    positive = corpus / "pos.txt"
    output = Path(model_prefix).expanduser().resolve()
    if not negative.exists() or not positive.exists():
        raise FileNotFoundError("训练目录必须同时包含 neg.txt 和 pos.txt")
    output.parent.mkdir(parents=True, exist_ok=True)
    started = time.time()
    sentiment.train(str(negative), str(positive))
    sentiment.save(str(output))
    sentiment.load(str(output))
    elapsed = time.time() - started
    print(f"模型已保存: {output}.3")
    print(f"训练耗时: {elapsed:.2f}s")
    for text in ("老师讲得很清晰", "课程节奏太快，我完全听不懂"):
        print(f"{text}\t{SnowNLP(text).sentiments:.4f}")


def main() -> None:
    parser = argparse.ArgumentParser(description="训练教学评价领域 SnowNLP 模型")
    parser.add_argument("--corpus-dir", default=str(DEFAULT_CORPUS_DIR))
    parser.add_argument("--model-prefix", default=str(CUSTOM_MODEL_PREFIX))
    args = parser.parse_args()
    train_model(args.corpus_dir, args.model_prefix)


if __name__ == "__main__":
    main()
