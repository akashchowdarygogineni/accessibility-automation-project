# 🚀 Accessibility Automation Pipeline

> End-to-End PDF Accessibility Validation System using **PAC (UI Tool) + PREP (JAR API) + Automated Reporting + Slack Integration**

---
```bash
git clone https://github.com/akashchowdarygogineni/accessibility-automation-project.git
cd accessibility-automation-project
pip install -r requirements.txt
python pipeline.py v1
```

## 📌 Executive Summary

This project automates accessibility validation of PDF files through a **three-stage pipeline**:

1. **PAC Stage** → Desktop-based validation (UI automation)
2. **PREP Stage** → API/JAR-based validation
3. **Comparison Stage** → Report consolidation + Slack notifications

Each execution is **version-controlled (v1, v1.3, v1.5, etc.)**, ensuring traceability of inputs, outputs, and reports.

---

## 🎯 Problem Statement

* Manual PDF accessibility validation is **time-consuming and repetitive**
* PAC requires **manual UI interaction**
* No unified system exists to **combine PAC and PREP results**
* Difficult to scale validation for **large volumes of PDFs**

---

## 💡 Solution

A scalable automation pipeline that:

* Processes PDFs in batch
* Automates PAC using UI scripting
* Integrates PREP via JAR API
* Compares PAC vs PREP results
* Generates multiple Excel reports
* Sends Slack notifications

---

## ⚙️ System Architecture

<img width="1677" height="498" alt="workflow" src="https://github.com/user-attachments/assets/fa13442a-6a4c-46a6-a4d1-39198425056e" />

---

## 🔄 End-to-End Workflow

1. Read PDFs from version folder
2. Run PAC validation (`pac.py`)
3. Run PREP validation (`prep.py`)
4. Compare results (`slack.py`)
5. Generate multiple reports
6. Consolidate into final report
7. Cleanup intermediate files

---

## 📂 Project Structure

```id="5q0p8m"
pipeline.py            → Full pipeline orchestration
pac.py                 → PAC UI automation & extraction
prep.py                → PREP JAR API processing
slack.py               → Comparison & report generation

prep_slack.py          → PREP + Slack (skip PAC)
slack_only_runner.py   → Slack-only execution
```

---

## ▶️ Execution Commands (IMPORTANT)

### 🔹 Full Pipeline (PAC → PREP → Slack)

```bash id="7c2t2j"
python pipeline.py <version>
```

Example:

```bash id="g3x0qk"
python pipeline.py v1.5
```

---

### 🔹 PAC Stage Only

```bash id="zv7u3y"
python pac.py <version>
```

Example:

```bash id="n2xk5r"
python pac.py v1
```

---

### 🔹 PREP Stage Only

#### Without JAR (auto-detect)

```bash id="gpk7p6"
python prep.py <version>
```

#### With JAR (Recommended)

```bash id="x6k9z2"
python prep.py <version> pdfremediation-0.0.1-SNAPSHOT.jar
```

Example:

```bash id="1j9m0c"
python prep.py v1 pdfremediation-0.0.1-SNAPSHOT.jar
```

---

### 🔹 Slack / Comparison Stage Only

```bash id="0l2k7f"
python slack.py <version>
```

Example:

```bash id="v9c3d1"
python slack.py v1
```

---

### 🔹 PREP + Slack (Skip PAC)

```bash id="y8x6r2"
python prep_slack.py <version> <jar_path>
```

Example:

```bash id="m4n8p1"
python prep_slack.py v1.5 pdfremediation-0.0.1-SNAPSHOT.jar
```

---

### 🔹 Slack Only Runner (Using Existing Files)

#### Auto-detect files

```bash id="a5k8z9"
python slack_only_runner.py <version>
```

#### With file names

```bash id="b7m2q4"
python slack_only_runner.py <version> PAC.xlsx PREP.xlsx
```

#### With full paths (Recommended)

```bash id="c9r5x1"
python slack_only_runner.py <version> <pac_excel_path> <prep_excel_path>
```

Example:

```bash id="d2t6y8"
python slack_only_runner.py v1 C:/project/v1/pac_results/PAC_Final_Summary_v1.xlsx C:/project/v1/prep_results/prep_final_summary_v1.xlsx
```

---

## 📊 Input & Output Design

### 📥 Input

* `BASE_DIR/<version>/` → PDF files

---

### 📤 Output

#### PAC Stage

* `pac_results/PAC_Final_Summary_<version>.xlsx`
* Per-file PAC outputs
* `pac_processed/`, `pac_skipped/`

#### PREP Stage

* `prep_results/prep_final_summary_<version>.xlsx`
* JSON outputs
* `prep_processed/`, `prep_skipped/`

#### Comparison Stage

* `PrepPac_Comparison_Report_<version>.xlsx`
* `PAC_PREP_Final_<version>.xlsx`
* `final_accessibility_status_only_<version>.xlsx`
* `version_summary_report_<version>.xlsx`
* `Final_Accessibility_Report_<version>.xlsx`

---

## 🏆 Final Deliverable

```id="e3u8k1"
Final_Accessibility_Report_<version>.xlsx
```

---

## 🛠️ Technology Stack

* Python 3.x
* pandas, openpyxl
* pywinauto (PAC UI automation)
* requests (API communication)
* python-dotenv (.env handling)
* slack_sdk (notifications)
* Java JAR (PREP tool)

---

## 🔔 Slack Integration

### 🔹 Configuration (.env)

```id="t8k3p6"
SLACK_TOKEN=xoxb-your-token
CHANNEL_ID=your-channel-id
```

### 🔹 Features

* Sends summary after execution
* Includes:

  * Total files processed
  * Pass/Fail/Skip counts
  * Comparison breakdown
  * Report references

---

## ⚠️ Known Constraints

* PAC requires Windows UI (no CLI support)
* UI automation depends on timing
* PREP requires local JAR server
* Excel files may get locked temporarily

---

## 🧪 Test Coverage

* Empty input handling
* Invalid folders
* Corrupted PDFs
* API failures
* Timeout handling
* Slack failure resilience
* Re-run overwrite handling

📌 Based on detailed test cases (*Documentation Pages 6–16*).

---



---

## ⚠️ Design Considerations
PAC is a desktop UI-based tool and does not support CLI/API execution
UI automation is implemented using pywinauto
PREP runs via a local JAR service
Pipeline is designed to handle failures, retries, and partial execution

---

## 📈 Results & Impact
✅ Automated end-to-end accessibility validation
✅ Reduced manual effort significantly (~60–70%)
✅ Enabled batch processing of multiple PDFs
✅ Improved reporting accuracy and consistency
✅ Integrated multiple tools into a unified pipeline

---

## 🙌 Acknowledgements
PAC (PDF Accessibility Checker)
PREP (PDF Remediation Tool - JAR API)
Slack API for notifications

---

## 🔐 Security

* Secrets stored in `.env`
* `.env` excluded via `.gitignore`
* Secure handling of API tokens

---

## 🌟 Future Enhancements

* Full PAC automation improvements
* Web dashboard for reporting
* Cloud deployment support
* Advanced analytics

---



## Author
- Name: akash gogineni
- Role: Accessibility Automation Engineer
- Contact: www.linkedin.com/in/akash-gogineni-b68102300

