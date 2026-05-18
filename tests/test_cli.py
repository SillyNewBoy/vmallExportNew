from vmall_export_new.cli import DEFAULT_DIFF_DIR, DEFAULT_OUTPUT_DIR, build_parser


def test_default_output_and_diff_dirs_use_src_folders():
    args = build_parser().parse_args(["--category", "phones"])

    assert args.output_dir == str(DEFAULT_OUTPUT_DIR)
    assert args.diff_dir == str(DEFAULT_DIFF_DIR)
    assert DEFAULT_OUTPUT_DIR.name == "out"
    assert DEFAULT_DIFF_DIR.name == "diff"
    assert DEFAULT_OUTPUT_DIR.parent.name == "src"
    assert DEFAULT_DIFF_DIR.parent.name == "src"
