"""Command-line entry point."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Sequence

from .client import VmallClient
from .exporter import VmallExporter


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vmall-export",
        description="Export VMALL product specifications grouped by category.",
    )
    parser.add_argument(
        "--category",
        help="Built-in category key to export, such as phones or tablets.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Export all built-in categories.",
    )
    parser.add_argument(
        "--output-dir",
        default="./output",
        help="Directory used to store generated JSON files.",
    )
    parser.add_argument(
        "--diff-dir",
        default="./diff",
        help="Directory used to store new-product diff JSON files.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Optional maximum number of unique products per category.",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=10,
        help="Maximum number of search result pages per category.",
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=20,
        help="Search page size per request.",
    )
    parser.add_argument(
        "--list-categories",
        action="store_true",
        help="Print built-in categories and exit.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging level.",
    )
    parser.add_argument(
        "--log-file",
        help="Optional log file path.",
    )
    parser.add_argument(
        "--request-interval",
        type=float,
        default=0.8,
        help="Minimum interval in seconds between HTTP requests.",
    )
    parser.add_argument(
        "--max-retries",
        type=int,
        default=4,
        help="Maximum retry attempts for HTTP requests.",
    )
    parser.add_argument(
        "--backoff-factor",
        type=float,
        default=1.5,
        help="Retry backoff factor in seconds.",
    )
    return parser


def configure_logging(level_name: str, log_file: str | None) -> None:
    level = getattr(logging, level_name.upper(), logging.INFO)
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    if log_file:
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))

    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=handlers,
        force=True,
    )


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    configure_logging(args.log_level, args.log_file)
    logger = logging.getLogger(__name__)
    client = VmallClient(
        min_request_interval=args.request_interval,
        max_retries=args.max_retries,
        backoff_factor=args.backoff_factor,
        logger=logging.getLogger("vmall_export_new.client"),
    )
    exporter = VmallExporter(client=client)

    if args.list_categories:
        for category in exporter.list_categories():
            print(
                f"{category.key}\t{category.name}\tkeyword={category.keyword}\tcategory_id={category.category_id}\tqueries={len(category.queries)}"
            )
        return 0

    if args.all and args.category:
        parser.error("--all and --category cannot be used together.")
    if not args.all and not args.category:
        parser.error("Either --category or --all is required unless --list-categories is used.")

    categories = (
        exporter.list_categories()
        if args.all
        else [exporter.get_category(args.category)]
    )

    logger.info("Resolved %s categories for export", len(categories))
    for category in categories:
        file_path, count = exporter.export_category(
            category,
            output_dir=Path(args.output_dir),
            diff_dir=Path(args.diff_dir),
            limit=args.limit,
            max_pages=args.max_pages,
            page_size=args.page_size,
        )
        print(f"{category.key}\t{count}\t{file_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
