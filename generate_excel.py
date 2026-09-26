import os
import pandas as pd

EXCEL_FILE = "QMS_Data.xlsx"

# 1. Protection Check: Abort if file exists to prevent overwriting user entries
if os.path.exists(EXCEL_FILE):
  print(
      f"⚠️  ABORTED: '{EXCEL_FILE}' already exists! Script will not overwrite"
      " your custom rows."
  )
  print(
      "If you intentionally want to reset to baseline data, rename or delete"
      " the existing Excel file first."
  )
else:
  # 2. Seed Data Definition (Executed ONLY if file is missing)
  kpi_data = {
      "Site": ["OLD BLD", "OLD BLD", "NEW BLD", "NEW BLD"],
      "Department": [
          "Certification scheme and standard mark adm",
          "Certification scheme and standard mark adm",
          "Testing Laboratory",
          "Testing Laboratory",
      ],
      "Period": ["26-Apr", "26-May", "26-Apr", "26-May"],
      "Total_Deviations": [14, 10, 8, 5],
      "Closed_On_Time": [11, 9, 7, 5],
      "Overdue": [3, 1, 1, 0],
      "CAPA_Closed": [10, 8, 6, 5],
      "CAPA_Failed": [1, 1, 1, 0],
      "Training_Compliance": [92, 95, 88, 90],
  }

  rc_data = {
      "Site": ["OLD BLD", "OLD BLD", "NEW BLD", "NEW BLD"],
      "Department": [
          "Certification scheme and standard mark adm",
          "Certification scheme and standard mark adm",
          "Testing Laboratory",
          "Testing Laboratory",
      ],
      "Period": ["26-Apr", "26-May", "26-Apr", "26-May"],
      "RC_People": [4, 2, 2, 1],
      "RC_Process": [5, 4, 3, 2],
      "RC_System": [3, 2, 2, 1],
      "RC_Equipment": [1, 1, 1, 1],
      "RC_Other": [1, 1, 0, 0],
  }

  defn_data = {
      "Category": [
          "1. RC_People (Human Factors & Personnel)",
          "2. RC_Process (Operational & Methodological Procedures)",
          "3. RC_System (Management Systems & Information Technology)",
          "4. RC_Equipment (Hardware, Instruments & Infrastructure)",
          "5. RC_Other (External & Unclassified Factors)",
      ],
      "Definition": [
          (
              "Deviations originating from human error, procedural"
              " non-compliance, training deficits, oversight, fatigue, or gaps"
              " in technical competency."
          ),
          (
              "Issues stemming from flawed, ambiguous, outdated, or overly"
              " complex workflows, execution protocols, and standard"
              " operational procedures."
          ),
          (
              "Failures within broader governance structures, document"
              " control networks, digital traceability architectures (e.g., QR"
              " code serialization/GTIN databases), software glitches, or"
              " cross-departmental communication gaps."
          ),
          (
              "Deviations caused by physical machinery, testing laboratory"
              " equipment calibration drift, tool degradation, unexpected"
              " hardware breakdowns, or utility interruptions."
          ),
          (
              "Root causes that do not fall under traditional operational"
              " categories, including raw material/supplier anomalies,"
              " environmental variations (e.g., temperature/humidity spikes),"
              " or external third-party disruptions."
          ),
      ],
      "QMS Context & Executive Application": [
          (
              "Identifies areas requiring refresher training, standard"
              " operating procedure (SOP) reassessment, or improved operational"
              " supervision."
          ),
          (
              "Highlights standard operating procedures that need revision,"
              " simplification, or structural alignment to meet quality"
              " standards."
          ),
          (
              "Focuses on upgrading digital tools, IT workflows, quality"
              " records systems, and management control systems."
          ),
          (
              "Points to required preventative maintenance schedules,"
              " recalibrations, equipment upgrades, or repair protocols."
          ),
          (
              "Triggers supplier quality audits, environmental control reviews,"
              " or external risk mitigation strategies."
          ),
      ],
  }

  audits_data = {
      "Site": ["OLD BLD", "OLD BLD", "NEW BLD", "NEW BLD"],
      "Department": [
          "Certification scheme and standard mark adm",
          "Certification scheme and standard mark adm",
          "Testing Laboratory",
          "Testing Laboratory",
      ],
      "Period": ["26-Apr", "26-May", "26-Apr", "26-May"],
      "Audit_Internal_Conducted": [2, 1, 1, 1],
      "Audit_Internal_Obs": [5, 3, 2, 1],
      "Audit_Supplier_Conducted": [1, 0, 1, 0],
      "Audit_Supplier_Obs": [2, 0, 1, 0],
      "Audit_Regulatory_Conducted": [0, 1, 0, 0],
      "Audit_Regulatory_Obs": [0, 1, 0, 0],
  }

  actions_data = {
      "Site": ["OLD BLD", "OLD BLD", "NEW BLD", "NEW BLD"],
      "Department": [
          "Certification scheme and standard mark adm",
          "Certification scheme and standard mark adm",
          "Testing Laboratory",
          "Testing Laboratory",
      ],
      "Period": ["26-Apr", "26-May", "26-Apr", "26-May"],
      "Actions_Taken": [
          "Initiated staff training on standard operating procedure revision.",
          "Conducted internal calibration audit for testing apparatus.",
          "Updated digital record archives and QR code traceability tags.",
          "Finalized corrective action plan for non-conformities.",
      ],
      "Future_Plans": [
          "Implement automated dashboard tracking across all site branches.",
          "Schedule external ISO/IEC 17025 surveillance assessment.",
          "Upgrade laboratory document management system infrastructure.",
          "Expand quality mark verification coverage to regional centers.",
      ],
      "Executive_Notes": [
          "Compliance metrics within expected target thresholds.",
          "Requires management review for resource allocation.",
          "System integration proceeding according to implementation roadmap.",
          "Audit findings closed ahead of target deadline.",
      ],
  }

  sec_data = {
      "Site": ["OLD BLD", "OLD BLD", "NEW BLD", "NEW BLD"],
      "Department": [
          "Certification scheme and standard mark adm",
          "Certification scheme and standard mark adm",
          "Testing Laboratory",
          "Testing Laboratory",
      ],
      "Period": ["26-Apr", "26-May", "26-Apr", "26-May"],
      "Investigation_Effectiveness": [94, 96, 90, 92],
      "Total_CAPAs": [11, 9, 7, 5],
      "Open_Changes": [2, 1, 1, 0],
      "Doc_Up_To_Date": [98, 99, 95, 97],
      "Complaints_Received": [1, 0, 0, 0],
  }

  # Write out Excel sheets safely
  with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
    pd.DataFrame(kpi_data).to_excel(
        writer, sheet_name="Overview_KPIs", index=False
    )
    pd.DataFrame(rc_data).to_excel(
        writer, sheet_name="Deviations_RootCauses", index=False
    )
    pd.DataFrame(defn_data).to_excel(
        writer, sheet_name="Defn of ACs", index=False
    )
    pd.DataFrame(audits_data).to_excel(
        writer, sheet_name="Audits_Compliance", index=False
    )
    pd.DataFrame(actions_data).to_excel(
        writer, sheet_name="Operational_Actions", index=False
    )
    pd.DataFrame(sec_data).to_excel(
        writer, sheet_name="Secondary_Metrics", index=False
    )

  print(
      f"✅ Successfully initialized baseline workbook '{EXCEL_FILE}' with 6"
      " sheets."
  )