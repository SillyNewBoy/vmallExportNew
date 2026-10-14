"""Command-line entry point."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Sequence

from .client import VmallClient
from .consumer_categories import CONSUMER_CATEGORY_PRESETS
from .consumer_client import ConsumerClient
from .consumer_exporter import ConsumerHuaweiExporter
from .excel_export import write_csv, write_excel
from .exporter import VmallExporter
from .hima_exporter import HIMA_CATEGORY_KEY, HimaCarExporter


DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parents[1] / "out"
DEFAULT_DIFF_DIR = Path(__file__).resolve().parents[1] / "diff"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vmall-export",
        description="Export VMALL product specifications grouped by category.",
    )
    parser.add_argument(
        "--site",
        choices=["vmall", "consumer"],
        default="vmall",
        help=(
            "Data source site. 'vmall' uses the VMALL mall APIs (default); "
            "'consumer' exports full specifications from the HUAWEI consumer "
            "site (consumer.huawei.com)."
        ),
    )
    parser.add_argument(
        "--region",
        default="my",
        help="Consumer site region code used with --site consumer, e.g. my, cn, en.",
    )
    parser.add_argument(
        "--excel",
        action="store_true",
        help="Also write an Excel (.xlsx) export per category.",
    )
    parser.add_argument(
        "--csv",
        action="store_true",
        help="Also write a CSV export per category.",
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
        default=None,
        help=(
            "Directory used to store generated JSON files. "
            "Defaults to src/out (VMALL) or src/out_{region} (consumer)."
        ),
    )
    parser.add_argument(
        "--diff-dir",
        default=None,
        help=(
            "Directory used to store new-product diff JSON files. "
            "Defaults to src/diff (VMALL) or src/diff_{region} (consumer)."
        ),
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
    if args.site == "consumer":
        return run_consumer_export(args, parser)
    return run_vmall_export(args, parser)


def run_consumer_export(args: argparse.Namespace, parser: argparse.ArgumentParser) -> int:
    """Export full specifications from consumer.huawei.com."""
    logger = logging.getLogger(__name__)
    consumer_client = ConsumerClient(
        region=args.region,
        min_request_interval=args.request_interval,
        max_retries=args.max_retries,
        backoff_factor=args.backoff_factor,
        logger=logging.getLogger("vmall_export_new.consumer_client"),
    )
    exporter = ConsumerHuaweiExporter(client=consumer_client)

    if args.list_categories:
        for category in exporter.list_categories():
            print(f"{category['key']}\t{category['name']}\tpath={category['path']}")
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

    # The consumer workflow defaults to also writing Excel + CSV exports.
    write_excel_enabled = args.excel or not args.csv
    write_csv_enabled = args.csv or not args.excel

    output_dir = Path(args.output_dir) if args.output_dir else DEFAULT_OUTPUT_DIR.parent / f"out_{args.region}"
    diff_dir = Path(args.diff_dir) if args.diff_dir else DEFAULT_DIFF_DIR.parent / f"diff_{args.region}"

    logger.info(
        "Resolved %s consumer categories for export (region=%s)",
        len(categories),
        args.region,
    )
    for category in categories:
        file_path, count, payload = exporter.export_category(
            category,
            output_dir=output_dir,
            diff_dir=diff_dir,
            limit=args.limit,
        )
        print(f"{category['key']}\t{count}\t{file_path}")
        if write_excel_enabled:
            print(f"{category['key']}\t{count}\t{write_excel(payload, output_dir)}")
        if write_csv_enabled:
            print(f"{category['key']}\t{count}\t{write_csv(payload, output_dir)}")

    return 0


def run_vmall_export(args: argparse.Namespace, parser: argparse.ArgumentParser) -> int:
    """Original VMALL / HIMA export flow."""
    logger = logging.getLogger(__name__)
    output_dir = Path(args.output_dir) if args.output_dir else DEFAULT_OUTPUT_DIR
    diff_dir = Path(args.diff_dir) if args.diff_dir else DEFAULT_DIFF_DIR
    vmall_client = VmallClient(
        min_request_interval=args.request_interval,
        max_retries=args.max_retries,
        backoff_factor=args.backoff_factor,
        logger=logging.getLogger("vmall_export_new.client"),
    )
    exporter = VmallExporter(client=vmall_client)
    hima_exporter = HimaCarExporter(
        request_interval=args.request_interval,
        logger=logging.getLogger("vmall_export_new.hima_exporter"),
    )

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
        if category.key == HIMA_CATEGORY_KEY:
            file_path, count = hima_exporter.export_category(
                output_dir=output_dir,
                diff_dir=diff_dir,
                limit=args.limit,
            )
        else:
            file_path, count = exporter.export_category(
                category,
                output_dir=output_dir,
                diff_dir=diff_dir,
                limit=args.limit,
                max_pages=args.max_pages,
                page_size=args.page_size,
            )
        print(f"{category.key}\t{count}\t{file_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
