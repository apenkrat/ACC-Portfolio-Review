#!/bin/bash
export PATH="/opt/homebrew/bin:$PATH"

# ── CONFIG ──────────────────────────────────────────────
ALIAS="apenkrat"
OUTPUT_DIR="$HOME/Library/CloudStorage/GoogleDrive-apenkrat@salesforce.com/My Drive/TMT Reports"
FILENAME="TMT_Project_Pulses_$(date +%Y-%m-%d).csv"
LATEST="TMT_Project_Pulses_LATEST.csv"
TMPDIR="$HOME/tmt-reports/tmp"
# ────────────────────────────────────────────────────────

echo "🔄 Starting TMT Project Pulses export - $(date)"
mkdir -p "$OUTPUT_DIR"
mkdir -p "$TMPDIR"

# ── QUERY 1: Status Indicators (PRIMARY — drives row count) ──
echo "📋 Query 1: Status Indicators..."
sf data query --target-org $ALIAS --result-format csv \
  --query "SELECT Id, Name, Project_Pulse__c, Project_Pulse__r.Project__c, Project_Pulse__r.Name, Project_Pulse__r.High_Watch_Visibility__c, Project_Pulse__r.Overall_Health_Color__c, Project_Pulse__r.Leadership_Notes__c, Project_Pulse__r.LRBC_Stage__c, Project_Pulse__r.Overall_Summary__c, Project_Pulse__r.Red_Account_Reason__c, Project_Pulse__r.Go_Live__c, Project_Pulse__r.Not_Primary_Pulse_Record__c, Project_Pulse__r.Project_on_Hold__c, Project_Pulse__r.Pulse_Update_Frequency_Required__c, Project_Pulse__r.NextSteps__c, Project_Pulse__r.Control__c, Project_Pulse__r.Overall_Health__c, Project_Pulse__r.Action_Needed_from_Leadership__c, Project_Pulse__r.SWE_or_CO_anticipated__c, Project_Pulse__r.Project_Methodology_Stage__c, Project_Pulse__r.Strategic_Account__c, Project_Pulse__r.Project_Description__c, Project_Pulse__r.Summary_Status__c, Project_Pulse__r.Trend_new__c, Project_Pulse__r.Overall_Pulse_Status__c, Project_Pulse__r.Overall_Status__c, Project_Pulse__r.Governance_Status__c, Summary_Project_Status__c, Project_Scope__c, Reason_for_RY_Path_to_Green_Scope__c, Project_Schedule__c, Reason_for_RY_Path_to_Green_Schedule__c, Project_Budget__c, Reason_for_RY_Path_to_Green_Budget__c, Resource_Status__c, Reason_for_RY_Path_to_Green_Resource__c, Customer_Status__c, Reason_for_RY_Path_to_Green_Customer__c, Methodology_Comments__c, Project_Risk__c FROM Status_Indicator__c WHERE Project_Pulse__r.Project__r.Subregion_new__c IN ('AMER TMT - 1', 'AMER TMT - 2') AND Project_Pulse__r.Project__r.pse__Stage__c IN ('Draft', 'On Hold', 'In Progress', 'In Progress - SWE') AND Project_Pulse__r.Project__r.pse__Practice__r.Name != 'FDE' AND Project_Pulse__r.Project__r.pse__Account__r.Name != 'Salesforce' ORDER BY Project_Pulse__r.Project__r.Name LIMIT 3000" \
  > "$TMPDIR/indicators.csv"
echo "✅ $(wc -l < "$TMPDIR/indicators.csv") indicator rows"

# ── QUERY 2: PSA Projects ────────────────────────────────
echo "📋 Query 2: PSA Projects..."
sf data query --target-org $ALIAS --result-format csv \
  --query "SELECT Id, Name, pse__Region__r.Name, pse__Account__r.Name, pse__Account__r.AOV_Band__c, pse__Account__r.CSG_Region__c, ProjectManager2Contact__r.Name, pse__Parent_Project__r.Name, pse__Project_Phase__c, pse__Project_Status__c, pse__Start_Date__c, pse__End_Date__c, OpportunityStage__c, pse__Opportunity__r.Name, pse__Opportunity__r.CloseDate, pse__Opportunity__r.Sub_region__c, pse__Practice__r.Name, Revenue_Treatment__c, pse__Pre_Billed__c, Timecard_Milestones_Required__c, pse__Bookings__c, pse__Billings__c, Total_Amount_Remaining__c, Revenue_Recognized_Comments__c, pse__Project_Status_Notes__c, pse__Project_Manager__r.Name, ActualProjectMargin__c, Margin_at_Close__c, Margin_at_Close_Percent__c, Overall_Bid_Margin_new__c, Operating_Model__c, Program_Size__c, Work_Complete__c, pse__Project_ID__c, pse__Stage__c, pse__Opportunity__c FROM pse__Proj__c WHERE Subregion_new__c IN ('AMER TMT - 1', 'AMER TMT - 2') AND pse__Stage__c IN ('Draft', 'On Hold', 'In Progress', 'In Progress - SWE') AND pse__Practice__r.Name != 'FDE' AND pse__Account__r.Name != 'Salesforce' ORDER BY Name LIMIT 3000" \
  > "$TMPDIR/projects.csv"
echo "✅ $(wc -l < "$TMPDIR/projects.csv") project rows"

# ── MERGE ────────────────────────────────────────────────
echo "🔗 Merging tables..."
python3 - << 'PYEOF'
import csv, os

tmpdir   = os.path.expanduser("~/tmt-reports/tmp")
out_dir  = os.environ.get("OUTPUT_DIR", tmpdir)
filename = os.environ.get("FILENAME", "test_output.csv")

def load_csv(path):
    with open(path, newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))

indicators = load_csv(f"{tmpdir}/indicators.csv")
projects   = load_csv(f"{tmpdir}/projects.csv")

print(f"Loaded: {len(indicators)} indicators, {len(projects)} projects")

projects_by_id = {}
for r in projects:
    pid = r.get("Id","")
    if pid:
        projects_by_id[pid] = r

OUTPUT_COLUMNS = [
    "Project Pulse ID","Region: Region Name","High Watch Visibility Selection",
    "Overall Health Color","Account: AOV Band","Project Manager 2 (Contact): Full Name",
    "Project Name","Project: Project Name","Parent Project: Project Name",
    "Project Phase","Project Status","Start Date","End Date","Go Live",
    "Opportunity Stage","Opportunity: Opportunity Name","Practice: Practice Name",
    "Revenue Treatment","Leadership Notes","LRBC Stage","Overall Summary",
    "Status Indicator Name","Path to Green","Pre-Billed (converted) Currency",
    "Pre-Billed (converted)","Pre-Billed Currency","Pre-Billed","Close Date",
    "Project Pulse Name","Red Account Reason","Timecard Milestones Required",
    "Bookings (converted) Currency","Bookings (converted)",
    "Billings (converted) Currency","Billings (converted)",
    "Forecasted Amount Remaining Currency","Forecasted Amount Remaining",
    "Forecasted Amount Remaining Reason","Project Status Notes",
    "Opportunity: Subregion","Project Manager: Full Name","Actual Project Margin",
    "Margin at Close Amount Currency","Margin at Close Amount","Margin at Close %",
    "Project Methodology Stage","Strategic Account","Project Description",
    "Project Progress Summary","Overall Bid Margin","Trend","Not Primary Pulse",
    "On Hold","Pulse Update Frequency Required","Next Steps","Current Risk",
    "Operating Model","Account: CSG Territory","Program Size","Work % Complete",
    "PSA Project ID","Overall Health","Project ID","Action Needed from Leadership",
    "SWE or CO anticipated?","Scope Status","RY Reason & Path to Green (Scope)",
    "Schedule Status","RY Reason & Path to Green (Schedule)","Budget Status",
    "RY Reason & Path to Green (Budget)","Resource Status",
    "RY Reason & Path to Green (Resource)","Customer Status",
    "RY Reason & Path to Green (Customer)","Overall Pulse Status","Overall Status",
    "Governance Status","Specify Details - Methodology Used","Risk Synopsis","Stage"
]

out_path = f"{out_dir}/{filename}"
written = 0

with open(out_path, "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=OUTPUT_COLUMNS, extrasaction="ignore")
    writer.writeheader()

    for ind in indicators:
        proj_id = ind.get("Project_Pulse__r.Project__c","")
        proj    = projects_by_id.get(proj_id, {})

        row = {
            "Project Pulse ID":                       ind.get("Project_Pulse__c",""),
            "Region: Region Name":                    proj.get("pse__Region__r.Name",""),
            "High Watch Visibility Selection":         ind.get("Project_Pulse__r.High_Watch_Visibility__c",""),
            "Overall Health Color":                    ind.get("Project_Pulse__r.Overall_Health_Color__c",""),
            "Account: AOV Band":                       proj.get("pse__Account__r.AOV_Band__c",""),
            "Project Manager 2 (Contact): Full Name":  proj.get("ProjectManager2Contact__r.Name",""),
            "Project Name":                            proj.get("Name",""),
            "Project: Project Name":                   proj.get("Name",""),
            "Parent Project: Project Name":            proj.get("pse__Parent_Project__r.Name",""),
            "Project Phase":                           proj.get("pse__Project_Phase__c",""),
            "Project Status":                          proj.get("pse__Project_Status__c",""),
            "Start Date":                              proj.get("pse__Start_Date__c",""),
            "End Date":                                proj.get("pse__End_Date__c",""),
            "Go Live":                                 ind.get("Project_Pulse__r.Go_Live__c",""),
            "Opportunity Stage":                       proj.get("OpportunityStage__c",""),
            "Opportunity: Opportunity Name":           proj.get("pse__Opportunity__r.Name",""),
            "Practice: Practice Name":                 proj.get("pse__Practice__r.Name",""),
            "Revenue Treatment":                       proj.get("Revenue_Treatment__c",""),
            "Leadership Notes":                        ind.get("Project_Pulse__r.Leadership_Notes__c",""),
            "LRBC Stage":                              ind.get("Project_Pulse__r.LRBC_Stage__c",""),
            "Overall Summary":                         ind.get("Project_Pulse__r.Overall_Summary__c",""),
            "Status Indicator Name":                   ind.get("Name",""),
            "Path to Green":                           ind.get("Summary_Project_Status__c",""),
            "Pre-Billed (converted) Currency":         "USD",
            "Pre-Billed (converted)":                  proj.get("pse__Pre_Billed__c",""),
            "Pre-Billed Currency":                     "USD",
            "Pre-Billed":                              proj.get("pse__Pre_Billed__c",""),
            "Close Date":                              proj.get("pse__Opportunity__r.CloseDate",""),
            "Project Pulse Name":                      ind.get("Project_Pulse__r.Name",""),
            "Red Account Reason":                      ind.get("Project_Pulse__r.Red_Account_Reason__c",""),
            "Timecard Milestones Required":            proj.get("Timecard_Milestones_Required__c",""),
            "Bookings (converted) Currency":           "USD",
            "Bookings (converted)":                    proj.get("pse__Bookings__c",""),
            "Billings (converted) Currency":           "USD",
            "Billings (converted)":                    proj.get("pse__Billings__c",""),
            "Forecasted Amount Remaining Currency":    "USD",
            "Forecasted Amount Remaining":             proj.get("Total_Amount_Remaining__c",""),
            "Forecasted Amount Remaining Reason":      proj.get("Revenue_Recognized_Comments__c",""),
            "Project Status Notes":                    proj.get("pse__Project_Status_Notes__c",""),
            "Opportunity: Subregion":                  proj.get("pse__Opportunity__r.Sub_region__c",""),
            "Project Manager: Full Name":              proj.get("pse__Project_Manager__r.Name",""),
            "Actual Project Margin":                   proj.get("ActualProjectMargin__c",""),
            "Margin at Close Amount Currency":         "USD",
            "Margin at Close Amount":                  proj.get("Margin_at_Close__c",""),
            "Margin at Close %":                       proj.get("Margin_at_Close_Percent__c",""),
            "Project Methodology Stage":               ind.get("Project_Pulse__r.Project_Methodology_Stage__c",""),
            "Strategic Account":                       ind.get("Project_Pulse__r.Strategic_Account__c",""),
            "Project Description":                     ind.get("Project_Pulse__r.Project_Description__c",""),
            "Project Progress Summary":                ind.get("Project_Pulse__r.Summary_Status__c",""),
            "Overall Bid Margin":                      proj.get("Overall_Bid_Margin_new__c",""),
            "Trend":                                   ind.get("Project_Pulse__r.Trend_new__c",""),
            "Not Primary Pulse":                       ind.get("Project_Pulse__r.Not_Primary_Pulse_Record__c",""),
            "On Hold":                                 ind.get("Project_Pulse__r.Project_on_Hold__c",""),
            "Pulse Update Frequency Required":         ind.get("Project_Pulse__r.Pulse_Update_Frequency_Required__c",""),
            "Next Steps":                              ind.get("Project_Pulse__r.NextSteps__c",""),
            "Current Risk":                            ind.get("Project_Pulse__r.Control__c",""),
            "Operating Model":                         proj.get("Operating_Model__c",""),
            "Account: CSG Territory":                  proj.get("pse__Account__r.CSG_Region__c",""),
            "Program Size":                            proj.get("Program_Size__c",""),
            "Work % Complete":                         proj.get("Work_Complete__c",""),
            "PSA Project ID":                          ind.get("Project_Pulse__r.Project__c",""),
            "Overall Health":                          ind.get("Project_Pulse__r.Overall_Health__c",""),
            "Project ID":                              proj.get("pse__Project_ID__c",""),
            "Action Needed from Leadership":           ind.get("Project_Pulse__r.Action_Needed_from_Leadership__c",""),
            "SWE or CO anticipated?":                  ind.get("Project_Pulse__r.SWE_or_CO_anticipated__c",""),
            "Scope Status":                            ind.get("Project_Scope__c",""),
            "RY Reason & Path to Green (Scope)":       ind.get("Reason_for_RY_Path_to_Green_Scope__c",""),
            "Schedule Status":                         ind.get("Project_Schedule__c",""),
            "RY Reason & Path to Green (Schedule)":    ind.get("Reason_for_RY_Path_to_Green_Schedule__c",""),
            "Budget Status":                           ind.get("Project_Budget__c",""),
            "RY Reason & Path to Green (Budget)":      ind.get("Reason_for_RY_Path_to_Green_Budget__c",""),
            "Resource Status":                         ind.get("Resource_Status__c",""),
            "RY Reason & Path to Green (Resource)":    ind.get("Reason_for_RY_Path_to_Green_Resource__c",""),
            "Customer Status":                         ind.get("Customer_Status__c",""),
            "RY Reason & Path to Green (Customer)":    ind.get("Reason_for_RY_Path_to_Green_Customer__c",""),
            "Overall Pulse Status":                    ind.get("Project_Pulse__r.Overall_Pulse_Status__c",""),
            "Overall Status":                          ind.get("Project_Pulse__r.Overall_Status__c",""),
            "Governance Status":                       ind.get("Project_Pulse__r.Governance_Status__c",""),
            "Specify Details - Methodology Used":      ind.get("Methodology_Comments__c",""),
            "Risk Synopsis":                           ind.get("Project_Risk__c",""),
            "Stage":                                   proj.get("pse__Stage__c",""),
        }
        writer.writerow(row)
        written += 1

print(f"Done! Wrote {written} rows to {out_path}")