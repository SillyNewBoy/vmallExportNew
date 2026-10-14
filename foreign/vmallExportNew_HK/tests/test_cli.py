from vmall_export_new.cli import DEFAULT_DIFF_DIR, DEFAULT_OUTPUT_DIR, build_parser


def test_default_output_and_diff_dirs_use_src_folders():
    args = build_parser().parse_args(["--category", "phones"])

    # defaults are None on the CLI and resolved per site inside cli.main()
    assert args.output_dir is None
    assert args.diff_dir is None
    assert DEFAULT_OUTPUT_DIR.name == "out"
    assert DEFAULT_DIFF_DIR.name == "diff"
    assert DEFAULT_OUTPUT_DIR.parent.name == "src"
    assert DEFAULT_DIFF_DIR.parent.name == "src"


def test_consumer_region_defaults_to_my():
    args = build_parser().parse_args(["--site", "consumer", "--category", "phones"])

    assert args.site == "consumer"
    assert args.region == "my"
