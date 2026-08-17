"""Command-line entry point for the offline-first course-comment NLP project."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from Data_processing import DataProcessing
from evaluation import evaluate_labels, prepare_evaluation_samples
from project_config import DEFAULT_INPUT_FILE, DEFAULT_OUTPUT_DIR, ensure_output_dir


def run_offline(input_path: str | Path, output_dir: str | Path) -> dict:
    result = DataProcessing(input_path=input_path, output_dir=output_dir).start()
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def run_crawl(args: argparse.Namespace) -> None:
    from b站评论区爬虫 import BilibiliCommentsSpider

    BilibiliCommentsSpider(
        video_url=args.url,
        save_path=args.input,
        max_pages=args.max_pages,
    ).start()
    run_offline(args.input, args.output)


def run_report(output_dir: str | Path, labels: str | Path | None = None) -> None:
    from 可视化报告生成 import ReportGenerator

    ReportGenerator(output_dir=output_dir, label_workbook=labels).start()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="B站教学评论采集、情感分析、评估与自动报告"
    )
    parser.add_argument(
        "command",
        nargs="?",
        default="offline",
        choices=["offline", "crawl", "prepare-eval", "evaluate", "report", "all"],
    )
    parser.add_argument("--input", default=str(DEFAULT_INPUT_FILE), help="原始评论 Excel")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_DIR), help="保研版输出目录")
    parser.add_argument("--url", help="crawl 命令所需的 B站视频 URL")
    parser.add_argument("--max-pages", type=int, default=1000)
    parser.add_argument("--labels", help="已填写的人工标注 Excel")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    output = ensure_output_dir(args.output)
    processed = output / "评论分析结果_保研版.csv"

    if args.command == "crawl":
        if not args.url:
            raise SystemExit("crawl 命令必须提供 --url")
        run_crawl(args)
        return
    if args.command == "offline":
        run_offline(args.input, output)
        return
    if args.command == "prepare-eval":
        if not processed.exists():
            run_offline(args.input, output)
        print(
            json.dumps(
                prepare_evaluation_samples(processed, output), ensure_ascii=False, indent=2
            )
        )
        return
    if args.command == "evaluate":
        if not args.labels:
            raise SystemExit("evaluate 命令必须提供 --labels")
        print(json.dumps(evaluate_labels(args.labels, output), ensure_ascii=False, indent=2))
        return
    if args.command == "report":
        run_report(output, args.labels)
        return
    if args.command == "all":
        run_offline(args.input, output)
        prepare_evaluation_samples(processed, output)
        if args.labels:
            evaluate_labels(args.labels, output)
        run_report(output, args.labels)


if __name__ == "__main__":
    main()
