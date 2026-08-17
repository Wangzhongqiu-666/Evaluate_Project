"""Prepare a blind manual-label set and evaluate default/custom SnowNLP models."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
from sklearn.metrics import (
    accuracy_score,
    classification_report,
    confusion_matrix,
    precision_recall_fscore_support,
)
from snownlp import SnowNLP, sentiment

from Data_processing import classify_score
from project_config import CUSTOM_MODEL_PREFIX, RANDOM_SEED, ensure_output_dir


LABEL_NAMES = {-1: "负面", 0: "中性", 1: "正面"}
QUESTION_MARKERS = (
    "吗",
    "么",
    "怎么",
    "请问",
    "求",
    "哪里",
    "哪一",
    "为什么",
    "不懂",
    "有没有",
    "多少",
    "咋",
)
AD_MARKERS = ("番茄todo", "加入码", "学习群", "微信", "公众号", "广告", "http", "二维码")


def _scores(texts: list[str]) -> list[float]:
    return [round(float(SnowNLP(text).sentiments), 4) for text in texts]


def _score_both_models(texts: list[str]) -> tuple[list[float], list[float]]:
    # Reset to SnowNLP's bundled model before collecting the baseline.
    sentiment.load(sentiment.data_path)
    default_scores = _scores(texts)
    actual_model = Path(f"{CUSTOM_MODEL_PREFIX}.3")
    if not actual_model.exists():
        raise FileNotFoundError(f"找不到领域模型: {actual_model}")
    sentiment.load(str(CUSTOM_MODEL_PREFIX))
    custom_scores = _scores(texts)
    return default_scores, custom_scores


def prepare_evaluation_samples(
    processed_csv: str | Path,
    output_dir: str | Path | None = None,
) -> dict[str, Any]:
    output = ensure_output_dir(output_dir)
    frame = pd.read_csv(processed_csv, encoding="utf-8-sig")
    if len(frame) < 150:
        raise ValueError("去重后评论不足150条，无法按计划构建评估集")
    frame = frame.drop_duplicates(subset=["评论"], keep="first").reset_index(drop=True)

    texts = frame["清洗后评论"].astype(str).tolist()
    default_scores, custom_scores = _score_both_models(texts)
    frame["默认模型分数"] = default_scores
    frame["默认模型预测"] = [classify_score(score) for score in default_scores]
    frame["领域模型分数"] = custom_scores
    frame["领域模型预测"] = [classify_score(score) for score in custom_scores]
    frame["问句请求特征"] = frame["评论"].astype(str).map(
        lambda text: any(marker in text for marker in QUESTION_MARKERS)
    )
    frame["广告特征"] = frame["评论"].astype(str).str.lower().map(
        lambda text: any(marker in text for marker in AD_MARKERS)
    )

    test = frame.sample(n=120, random_state=RANDOM_SEED, replace=False).copy()
    test["样本编号"] = [f"T{i:03d}" for i in range(1, 121)]
    test["样本用途"] = "正式测试"
    used = set(test.index)
    remaining = frame.loc[~frame.index.isin(used)].copy()

    diagnostic_parts: list[pd.DataFrame] = []
    query = remaining[
        (remaining["领域模型预测"] == -1) & remaining["问句请求特征"]
    ].sort_values("领域模型分数").head(12)
    query["诊断类别"] = "低分问句/请求"
    diagnostic_parts.append(query)
    used.update(query.index)

    remaining = frame.loc[~frame.index.isin(used)].copy()
    remaining["阈值距离"] = remaining["领域模型分数"].map(
        lambda value: min(abs(value - 0.3), abs(value - 0.7))
    )
    boundary = remaining.sort_values("阈值距离").head(10)
    boundary["诊断类别"] = "阈值附近"
    diagnostic_parts.append(boundary)
    used.update(boundary.index)

    remaining = frame.loc[~frame.index.isin(used)].copy()
    ads = remaining[remaining["广告特征"]].head(8)
    ads["诊断类别"] = "疑似广告"
    diagnostic_parts.append(ads)
    used.update(ads.index)

    diagnostic = pd.concat(diagnostic_parts, axis=0)
    if len(diagnostic) < 30:
        remaining = frame.loc[~frame.index.isin(used)].copy()
        fill = remaining.assign(
            极端程度=(remaining["领域模型分数"] - 0.5).abs()
        ).sort_values("极端程度", ascending=False).head(30 - len(diagnostic))
        fill["诊断类别"] = "极端分数补充"
        diagnostic = pd.concat([diagnostic, fill], axis=0)

    diagnostic = diagnostic.head(30).copy()
    diagnostic["样本编号"] = [f"D{i:03d}" for i in range(1, 31)]
    diagnostic["样本用途"] = "诊断分析"
    samples = pd.concat([test, diagnostic], ignore_index=True)

    blind_columns = ["样本编号", "样本用途", "评论"]
    blind_records = samples[blind_columns].to_dict(orient="records")
    public_payload = {
        "seed": RANDOM_SEED,
        "formal_test_count": 120,
        "diagnostic_count": 30,
        "label_guide": {
            "1": "明确赞扬、感谢、满意或正向学习结果",
            "0": "提问、求资源、客观陈述、广告或态度不明确",
            "-1": "明确批评、抱怨、不满或负向体验",
        },
        "samples": blind_records,
    }
    with (output / "人工标注样本.json").open("w", encoding="utf-8") as stream:
        json.dump(public_payload, stream, ensure_ascii=False, indent=2)

    cache_columns = [
        "样本编号",
        "样本用途",
        "诊断类别",
        "默认模型分数",
        "默认模型预测",
        "领域模型分数",
        "领域模型预测",
        "问句请求特征",
        "广告特征",
    ]
    samples[cache_columns].to_csv(
        output / "模型预测缓存.csv", index=False, encoding="utf-8-sig"
    )
    status = {
        "seed": RANDOM_SEED,
        "formal_test_count": 120,
        "diagnostic_count": 30,
        "total": 150,
        "label_status": "等待人工填写",
    }
    with (output / "评估状态.json").open("w", encoding="utf-8") as stream:
        json.dump(status, stream, ensure_ascii=False, indent=2)
    return status


def _model_metrics(y_true: list[int], y_pred: list[int]) -> dict[str, Any]:
    precision, recall, f1, _ = precision_recall_fscore_support(
        y_true, y_pred, average="macro", zero_division=0
    )
    report = classification_report(
        y_true,
        y_pred,
        labels=[-1, 0, 1],
        target_names=["负面", "中性", "正面"],
        output_dict=True,
        zero_division=0,
    )
    return {
        "accuracy": round(float(accuracy_score(y_true, y_pred)), 6),
        "macro_precision": round(float(precision), 6),
        "macro_recall": round(float(recall), 6),
        "macro_f1": round(float(f1), 6),
        "per_class": report,
    }


def _save_confusion(
    y_true: list[int], y_pred: list[int], title: str, path: Path
) -> None:
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
    plt.rcParams["axes.unicode_minus"] = False
    sns.set_theme(style="whitegrid", font="Microsoft YaHei")
    matrix = confusion_matrix(y_true, y_pred, labels=[-1, 0, 1])
    plt.figure(figsize=(5.6, 4.6))
    sns.heatmap(
        matrix,
        annot=True,
        fmt="d",
        cmap="Blues",
        xticklabels=["负面", "中性", "正面"],
        yticklabels=["负面", "中性", "正面"],
    )
    plt.title(title)
    plt.xlabel("模型预测")
    plt.ylabel("人工标签")
    plt.tight_layout()
    plt.savefig(path, dpi=180, bbox_inches="tight")
    plt.close()


def evaluate_labels(
    label_workbook: str | Path,
    output_dir: str | Path | None = None,
) -> dict[str, Any]:
    output = ensure_output_dir(output_dir)
    labels = pd.read_excel(label_workbook, sheet_name="人工标注")
    required = {"样本编号", "人工标签"}
    if missing := required - set(labels.columns):
        raise ValueError(f"标注表缺少列: {sorted(missing)}")
    labels["人工标签"] = pd.to_numeric(labels["人工标签"], errors="coerce")
    completed = labels["人工标签"].isin([-1, 0, 1])
    status = {
        "total": int(len(labels)),
        "completed": int(completed.sum()),
        "remaining": int((~completed).sum()),
    }
    if status["remaining"]:
        status["label_status"] = "未完成，暂不生成模型指标"
        with (output / "评估状态.json").open("w", encoding="utf-8") as stream:
            json.dump(status, stream, ensure_ascii=False, indent=2)
        return status

    predictions = pd.read_csv(output / "模型预测缓存.csv", encoding="utf-8-sig")
    merged = labels.merge(predictions, on="样本编号", how="inner", validate="one_to_one")
    if len(merged) != 150:
        raise ValueError(f"标注表与预测缓存未完整匹配: {len(merged)}/150")

    formal = merged[merged["样本用途_y"] == "正式测试"].copy()
    y_true = formal["人工标签"].astype(int).tolist()
    default_pred = formal["默认模型预测"].astype(int).tolist()
    custom_pred = formal["领域模型预测"].astype(int).tolist()
    default_metrics = _model_metrics(y_true, default_pred)
    custom_metrics = _model_metrics(y_true, custom_pred)
    delta = custom_metrics["macro_f1"] - default_metrics["macro_f1"]

    diagnostic = merged[merged["样本用途_y"] == "诊断分析"].copy()
    diagnostic_summary: dict[str, Any] = {}
    for category, group in diagnostic.groupby("诊断类别", dropna=False):
        diagnostic_summary[str(category)] = {
            "count": int(len(group)),
            "default_errors": int(
                (group["默认模型预测"].astype(int) != group["人工标签"].astype(int)).sum()
            ),
            "custom_errors": int(
                (group["领域模型预测"].astype(int) != group["人工标签"].astype(int)).sum()
            ),
        }

    result = {
        "label_status": "已完成",
        "formal_test_count": int(len(formal)),
        "diagnostic_count": int(len(diagnostic)),
        "default_model": default_metrics,
        "custom_model": custom_metrics,
        "macro_f1_delta": round(float(delta), 6),
        "claim_domain_improvement": bool(delta >= 0.02),
        "diagnostic": diagnostic_summary,
    }
    with (output / "模型评估结果.json").open("w", encoding="utf-8") as stream:
        json.dump(result, stream, ensure_ascii=False, indent=2)
    merged.to_csv(output / "模型评估明细.csv", index=False, encoding="utf-8-sig")
    _save_confusion(y_true, default_pred, "默认 SnowNLP 混淆矩阵", output / "默认模型混淆矩阵.png")
    _save_confusion(y_true, custom_pred, "领域 SnowNLP 混淆矩阵", output / "领域模型混淆矩阵.png")
    return result


if __name__ == "__main__":
    raise SystemExit("请通过 main.py 的 prepare-eval 或 evaluate 子命令运行")
