
import subprocess
import sys
import os


DEFAULT_JAR_NAME = "pdfremediation-0.0.1-SNAPSHOT.jar"


# =====================================
# RUN PAC SCRIPT
# =====================================
def run_pac(version):
    print("🔹 Running PAC script...")

    result = subprocess.run(["python", "pac.py", version])

    if result.returncode != 0:
        print("❌ PAC script failed. Stopping pipeline.")
        sys.exit(1)

    print("✅ PAC script completed.\n")


# =====================================
# RUN PREP SCRIPT
# =====================================
def run_prep(version, jar_name):
    print("🔹 Running PREP script...")

    result = subprocess.run(["python", "prep.py", version, jar_name])

    if result.returncode != 0:
        print("❌ PREP script failed. Stopping pipeline.")
        sys.exit(1)

    print("✅ PREP script completed.\n")


# =====================================
# RUN SLACK SCRIPT
# =====================================
def run_slack(version):
    print("🔹 Running SLACK + COMPARISON script...")

    result = subprocess.run(["python", "slack.py", version])

    if result.returncode != 0:
        print("❌ Slack script failed.")
        sys.exit(1)

    print("✅ Slack script completed.\n")


def _ensure_input_pdfs(version):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    version_dir = os.path.join(base_dir, version)

    if not os.path.isdir(version_dir):
        print(f"❌ Version folder not found: {version_dir}")
        print("Add input PDF files to process in this folder and run again.")
        sys.exit(0)

    pdfs = [f for f in os.listdir(version_dir) if f.lower().endswith(".pdf")]
    if not pdfs:
        print(f"⚠️ No input PDF files found in: {version_dir}")
        print("Add input PDF files to process in this folder and run again.")
        sys.exit(0)


def _ensure_stage_outputs_for_slack(version):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pac_summary = os.path.join(base_dir, version, "pac_results", f"PAC_Final_Summary_{version}.xlsx")
    prep_summary = os.path.join(base_dir, version, "prep_results", f"prep_final_summary_{version}.xlsx")

    missing = []
    if not os.path.exists(pac_summary):
        missing.append(pac_summary)
    if not os.path.exists(prep_summary):
        missing.append(prep_summary)

    if missing:
        print("⚠️ Required summary files are missing. Skipping Slack/comparison stage.")
        for path in missing:
            print(f"   - Missing: {path}")
        print("Add input PDF files to process in the version folder and run again.")
        sys.exit(0)


def _ensure_prep_jar(version):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    version_dir = os.path.join(base_dir, version)

    jar_files = sorted([f for f in os.listdir(version_dir) if f.lower().endswith(".jar")])
    if jar_files:
        if DEFAULT_JAR_NAME in jar_files:
            print(f"🧩 Using JAR: {DEFAULT_JAR_NAME}")
            return DEFAULT_JAR_NAME

        selected_jar = jar_files[0]
        print(f"🧩 Using JAR from folder: {selected_jar}")
        if len(jar_files) > 1:
            print("ℹ️ Multiple JARs found; selected the first one alphabetically.")
        return selected_jar

    print(f"⚠️ No JAR file found in: {version_dir}")
    print("Add the PREP JAR file and run again.")
    print(f"Expected default name: {DEFAULT_JAR_NAME}")
    sys.exit(0)


# =====================================
# MAIN PIPELINE FUNCTION
# =====================================
def run_pipeline(version):

    print(f"\n🚀 Starting Automation Pipeline for version: {version}\n")

    _ensure_input_pdfs(version)

    run_pac(version)
    print("Waiting for PAC to release file locks...")
    import time
    time.sleep(5)

    jar_name = _ensure_prep_jar(version)
    run_prep(version, jar_name)

    _ensure_stage_outputs_for_slack(version)
    run_slack(version)

    print("🎉 PIPELINE COMPLETED SUCCESSFULLY!")


# =====================================
# ENTRY POINT
# =====================================
if __name__ == "__main__":

    version = sys.argv[1] if len(sys.argv) > 1 else "v1"

    run_pipeline(version)