import argparse
import gc
import glob
import os
import sys
import time
from typing import Iterable, Optional

import slack


def _cleanup_intermediate_reports(output_dir: str, version: str, keep_files: Iterable[str]) -> None:
    keep_set = set(keep_files)
    generated_reports = [
        f"PrepPac_Comparison_Report_{version}.xlsx",
        f"final_accessibility_status_only_{version}.xlsx",
        f"version_summary_report_{version}.xlsx",
        f"final_summary_{version}.xlsx",
        f"splitup_report_{version}.xlsx",
    ]

    for file_name in generated_reports:
        if file_name in keep_set:
            continue

        file_path = os.path.join(output_dir, file_name)
        if not os.path.exists(file_path):
            continue

        deleted = False
        for _ in range(8):
            try:
                os.remove(file_path)
                print(f"🗑 Deleted intermediate report: {file_name}")
                deleted = True
                break
            except PermissionError:
                gc.collect()
                time.sleep(1)

        if not deleted and os.path.exists(file_path):
            print(f"⚠️ Could not delete (locked): {file_name}")


def _resolve_input_excel(base_dir: str, version: str, provided_path: str, result_folder: str) -> str:
    provided_path = provided_path.strip()

    # If caller provided a real path (absolute or with folders), use it directly.
    has_folder = (os.path.sep in provided_path) or (os.path.altsep and os.path.altsep in provided_path)
    if os.path.isabs(provided_path) or has_folder:
        return os.path.abspath(provided_path)

    # 1) Check command version folder first for convenience.
    default_candidate = os.path.join(base_dir, version, result_folder, provided_path)
    if os.path.exists(default_candidate):
        return default_candidate

    # 2) Search all version folders; version arg is for output location only.
    pattern = os.path.join(base_dir, "v*", result_folder, provided_path)
    matches = glob.glob(pattern)
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise FileNotFoundError(
            f"Multiple matches found for {provided_path} in */{result_folder}/. "
            f"Please pass an explicit path. Matches: {matches}"
        )

    # Return expected path so caller gets a clear not-found message.
    return default_candidate


def _auto_pick_input_excel(
    base_dir: str,
    version: str,
    result_folder: str,
    preferred_names: Iterable[str],
) -> str:
    target_dir = os.path.join(base_dir, version, result_folder)
    if not os.path.isdir(target_dir):
        return os.path.join(target_dir, next(iter(preferred_names), ""))

    for name in preferred_names:
        candidate = os.path.join(target_dir, name)
        if os.path.exists(candidate):
            return candidate

    xlsx_candidates = sorted(
        glob.glob(os.path.join(target_dir, "*.xlsx")),
        key=os.path.getmtime,
        reverse=True,
    )
    if xlsx_candidates:
        return xlsx_candidates[0]

    return os.path.join(target_dir, next(iter(preferred_names), ""))



def run_slack_stage_only(
    version: str,
    pac_file: Optional[str],
    prep_file: Optional[str],
    output_dir: str,
) -> None:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pac_input = pac_file or "<auto>"
    prep_input = prep_file or "<auto>"

    if pac_file:
        pac_file = _resolve_input_excel(base_dir, version, pac_file, "pac_results")
    else:
        pac_file = _auto_pick_input_excel(
            base_dir,
            version,
            "pac_results",
            preferred_names=[f"PAC_Final_Summary_{version}.xlsx"],
        )

    if prep_file:
        prep_file = _resolve_input_excel(base_dir, version, prep_file, "prep_results")
    else:
        prep_file = _auto_pick_input_excel(
            base_dir,
            version,
            "prep_results",
            preferred_names=[f"prep_final_summary_{version}.xlsx"],
        )

    output_dir = os.path.abspath(output_dir)

    print("\n" + "=" * 70)
    print("Slack Stage-Only Run")
    print("=" * 70)
    print(f"Version           : {version}")
    print(f"PAC input arg     : {pac_input}")
    print(f"PREP input arg    : {prep_input}")
    print(f"Resolved PAC path : {pac_file}")
    print(f"Resolved PREP path: {prep_file}")
    print(f"Output directory  : {output_dir}")

    missing_items = []
    if not os.path.exists(pac_file):
        missing_items.append(("PAC", pac_file, "pac_results"))
    if not os.path.exists(prep_file):
        missing_items.append(("PREP", prep_file, "prep_results"))

    if missing_items:
        print("\n:x: Input validation failed.")
        for label, path, folder_hint in missing_items:
            print(f"  - Missing {label} report: {path}")
            print(
                f"    Tip: pass full path OR pass file name that exists in <version>/{folder_hint}."
            )
        raise FileNotFoundError("One or more required input files were not found.")

    os.makedirs(output_dir, exist_ok=True)

    # Update module globals used by existing slack.py functions.
    slack.version = version
    slack.CURRENT_VERSION = version

    final_report_file = os.path.join(output_dir, f"PrepPac_Comparison_Report_{version}.xlsx")
    output_file = os.path.join(output_dir, f"PAC_PREP_Final_{version}.xlsx")
    status_file = os.path.join(output_dir, f"final_accessibility_status_only_{version}.xlsx")
    version_summary_file = os.path.join(output_dir, f"version_summary_report_{version}.xlsx")
    splitup_report = os.path.join(output_dir, f"splitup_report_{version}.xlsx")

    final_report = slack.generate_accessibility_report(
        pac_file,
        prep_file,
        final_report_file,
    )

    result = slack.generate_comparison_report(
        pac_file,
        prep_file,
        output_file,
    )

    status_output = slack.generate_status_file(final_report, status_file)
    version_summary_output = slack.generate_version_summary(status_output, version_summary_file)
    slack.generate_final_column_summary(final_report)
    slack.generate_pac_prep_report(pac_file, prep_file, splitup_report)

    # Consolidate report files into one workbook.
    slack.combine_reports(output_dir, version)

    final_combined = f"Final_Accessibility_Report_{version}.xlsx"
    pac_prep_final = f"PAC_PREP_Final_{version}.xlsx"
    _cleanup_intermediate_reports(output_dir, version, keep_files=[final_combined, pac_prep_final])

    # Uses existing .env variables: SLACK_TOKEN and CHANNEL_ID.
    # Slack failures should not block report generation/cleanup outputs.
    try:
        slack.send_slack_summary(version_summary_output, result)
    except Exception as exc:
        print(f"⚠️ Slack summary skipped: {exc}")

    print("\n✅ Slack stage-only workflow completed.")



def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Run only the slack/comparison stage using PAC and PREP Excel files "
            "while keeping the same report outputs as slack.py. "
            "You can pass full paths or just file names."
        )
    )
    parser.add_argument("version", help="Version label (example: v1.3)")
    parser.add_argument(
        "pac_excel",
        nargs="?",
        default=None,
        help=(
            "Optional PAC Excel path or file name. If omitted, auto-picks from "
            "<version>/pac_results."
        ),
    )
    parser.add_argument(
        "prep_excel",
        nargs="?",
        default=None,
        help=(
            "Optional PREP Excel path or file name. If omitted, auto-picks from "
            "<version>/prep_results."
        ),
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Output folder for generated reports (default: <script_dir>/<version>)",
    )
    return parser



def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = args.output_dir or os.path.join(base_dir, args.version)

    try:
        run_slack_stage_only(args.version, args.pac_excel, args.prep_excel, output_dir)
        return 0
    except Exception as exc:
        print(f"❌ Stage-only slack run failed: {exc}")
        return 1



if __name__ == "__main__":
    sys.exit(main())
