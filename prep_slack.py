import subprocess
import sys
import os
import shutil


def normalize_path(path):
    normalized = path.strip().replace("\\", "/")
    if len(normalized) > 2 and normalized[1] == ";" and normalized[2] == "/":
        normalized = f"{normalized[0]}:{normalized[2:]}"
    return normalized


def resolve_pac_source_paths(source_path=None):
    base_dir = os.path.dirname(os.path.abspath(__file__))

    if source_path:
        normalized = normalize_path(source_path)
        source_summary = normalized if os.path.isabs(normalized) else os.path.join(base_dir, normalized)
        source_version_dir = os.path.dirname(os.path.dirname(source_summary))
    else:
        source_version_dir = os.path.join(base_dir, "v1")
        source_summary = os.path.join(source_version_dir, "pac_results", "PAC_Final_Summary_v1.xlsx")

    return source_version_dir, source_summary


def resolve_jar_name(version, jar_name=None):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    version_dir = os.path.join(base_dir, version)

    if jar_name:
        jar_path = os.path.join(version_dir, jar_name)
        if not os.path.isfile(jar_path):
            print(f"JAR not found in selected version folder: {jar_path}")
            print("Provide a valid jar file name present inside the version folder.")
            sys.exit(1)
        return jar_name

    if not os.path.isdir(version_dir):
        print(f"Version folder not found: {version_dir}")
        sys.exit(1)

    jar_files = sorted([f for f in os.listdir(version_dir) if f.lower().endswith(".jar")])
    if not jar_files:
        print(f"No JAR file found in: {version_dir}")
        print("Place a JAR file in this version folder or pass jar name explicitly.")
        sys.exit(1)

    selected = jar_files[0]
    print(f"Using JAR from version folder: {selected}")
    if len(jar_files) > 1:
        print("Multiple JARs found; selected the first one alphabetically.")

    return selected


def is_placeholder_arg(value):
    if value is None:
        return True
    return value.strip().lower() in {"na", "none", "null", "-", ""}


# -------------------------------------------------
# Copy PAC processed and skipped folders
# -------------------------------------------------
def copy_pac_folders(version, source_version_dir):

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    dst_version = os.path.join(BASE_DIR, version)

    folders = ["pac_processed", "pac_skipped"]

    for folder in folders:

        src = os.path.join(source_version_dir, folder)
        dst = os.path.join(dst_version, folder)

        if not os.path.exists(src):
            print(f"Required PAC folder not found: {src}")
            print("Provide a valid source PAC summary path from a version folder that has pac_processed and pac_skipped.")
            sys.exit(1)

        if os.path.exists(dst):
            shutil.rmtree(dst)

        shutil.copytree(src, dst)

    print(f"PAC processed and skipped folders copied from: {source_version_dir}")


# -------------------------------------------------
# Copy PAC summary file
# -------------------------------------------------
def copy_pac_summary(version, source_path=None):

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    _, source = resolve_pac_source_paths(source_path)

    if not os.path.exists(source):
        if source_path:
            print(f"PAC summary file not found: {source}")
            print("Pass a valid PAC summary file path as the 3rd argument.")
        else:
            print("PAC baseline data not found in v1.")
            print("Please run the baseline pipeline first:")
            print("python pipeline.py v1")
        sys.exit(1)

    target_folder = os.path.join(BASE_DIR, version, "pac_results")
    os.makedirs(target_folder, exist_ok=True)

    target = os.path.join(target_folder, f"PAC_Final_Summary_{version}.xlsx")

    shutil.copy(source, target)

    print(f"PAC summary copied from: {source}")


# -------------------------------------------------
# Run PREP
# -------------------------------------------------
def run_prep(version, jar_name):

    print("Running PREP...")

    subprocess.run(["python", "prep.py", version, jar_name], check=True)


def ensure_prep_summary_exists(version):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    prep_summary = os.path.join(base_dir, version, "prep_results", f"prep_final_summary_{version}.xlsx")
    if not os.path.isfile(prep_summary):
        print(f"PREP summary not found: {prep_summary}")
        print("Stopping before Slack stage because PREP did not produce output.")
        sys.exit(1)


# -------------------------------------------------
# Run Slack comparison
# -------------------------------------------------
def run_slack(version):

    print("Running Slack comparison...")

    subprocess.run(["python", "slack.py", version], check=True)


# -------------------------------------------------
# Pipeline
# -------------------------------------------------
def run_pipeline(version, jar_name, pac_summary_source=None):

    resolved_jar_name = resolve_jar_name(version, jar_name)

    source_version_dir, _ = resolve_pac_source_paths(pac_summary_source)

    copy_pac_folders(version, source_version_dir)

    copy_pac_summary(version, pac_summary_source)

    run_prep(version, resolved_jar_name)

    ensure_prep_summary_exists(version)

    run_slack(version)

    print("Pipeline completed.")


# -------------------------------------------------
# Main
# -------------------------------------------------
if __name__ == "__main__":

    if len(sys.argv) < 2:
        print("Usage:")
        print("python prep_slack.py <version> [jar_name] [pac_summary_source]")
        print("Examples:")
        print("python prep_slack.py v3")
        print("python prep_slack.py v3 my.jar")
        print("python prep_slack.py v3 my.jar C:/project/v2/pac_results/PAC_Final_Summary_v2.xlsx")
        print("python prep_slack.py v3 C:/project/v2/pac_results/PAC_Final_Summary_v2.xlsx")
        sys.exit(1)

    version = sys.argv[1]

    jar_name = None
    pac_summary_source = None

    if len(sys.argv) >= 3:
        third_arg = sys.argv[2]
        if is_placeholder_arg(third_arg):
            jar_name = None
            pac_summary_source = None
        elif third_arg.lower().endswith(".jar"):
            jar_name = third_arg
            if len(sys.argv) > 3 and not is_placeholder_arg(sys.argv[3]):
                pac_summary_source = sys.argv[3]
        else:
            pac_summary_source = third_arg
            if len(sys.argv) > 3 and not is_placeholder_arg(sys.argv[3]) and sys.argv[3].lower().endswith(".jar"):
                jar_name = sys.argv[3]

    if is_placeholder_arg(jar_name):
        jar_name = None
    if is_placeholder_arg(pac_summary_source):
        pac_summary_source = None

    try:
        run_pipeline(version, jar_name, pac_summary_source)
    except subprocess.CalledProcessError as e:
        print(f"Pipeline stopped because step failed: {' '.join(e.cmd)}")
        sys.exit(1)
    except SystemExit:
        raise
    except Exception as e:
        print(f"Pipeline stopped due to unexpected error: {e}")
        sys.exit(1)

