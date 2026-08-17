"""Project-wide paths and reproducibility constants."""

from __future__ import annotations

import os
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CODINGS_DIR = PROJECT_ROOT / "Codings"
FUNCTIONAL_DIR = PROJECT_ROOT / "Functional_files"
ORIGINAL_DATA_DIR = PROJECT_ROOT / "Original_Datas"
LEGACY_INTEGRATED_DIR = PROJECT_ROOT / "Integrated_data"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs" / "保研版"

DEFAULT_INPUT_FILE = ORIGINAL_DATA_DIR / "评论.xlsx"
USER_DICT_FILE = FUNCTIONAL_DIR / "评教用词典扩充.txt"
STOPWORDS_FILE = FUNCTIONAL_DIR / "学生评教停用词表.xlsx"
# SnowNLP appends ".3" on Python 3. The actual file is therefore *.3.3.
CUSTOM_MODEL_PREFIX = (
    FUNCTIONAL_DIR / "针对评教的SnowNLP模型" / "class_evaluate_model.marshal.3"
)

RANDOM_SEED = 20260811
NEGATIVE_THRESHOLD = 0.3
POSITIVE_THRESHOLD = 0.7


def get_bilibili_cookie() -> str:
    """Read the optional session cookie without storing it in source code."""

    return os.getenv("BILIBILI_COOKIE", "").strip()


def ensure_output_dir(path: str | Path | None = None) -> Path:
    output_dir = Path(path).expanduser().resolve() if path else DEFAULT_OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir

