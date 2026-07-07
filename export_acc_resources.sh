#!/bin/bash
export PATH="/opt/homebrew/bin:$PATH"

# ── CONFIG ──────────────────────────────────────────────
ALIAS="apenkrat"
OUTPUT_DIR="$HOME/Library/CloudStorage/GoogleDrive-apenkrat@salesforce.com/My Drive/TMT Reports"
FILENAME="TMT_Project_Resources_$(date +%Y-%m-%d).csv"
LATEST="TMT_Project_Resources_LATEST.csv"
# ────────────────────────────────────────────────────────

echo "🔄 Starting TMT Project Resources export - $(date)"

# Create output dir if it doesn't exist
mkdir -p "$OUTPUT_DIR"

# Run the SOQL query and export to CSV
sf data query \
  --target-org $ALIAS \
  --query "SELECT Name, pse__Project__r.Name, pse__Project__r.Subregion_new__c, pse__Project__r.pse__Account__r.Name, pse__Project__r.pse__Stage__c, pse__Resource__r.Name, pse__Role__c, pse__Start_Date__c, pse__End_Date__c, pse__Status__c, pse__Is_Billable__c, pse__Planned_Hours__c, pse__Scheduled_Hours__c, Forecasted_Hours_Remaining__c, Total_Billable_and_Credited_Hours__c, pse__Bill_Rate__c, Planned_Amount__c, pse__Projected_Revenue__c, Project_Region__c, Resource_Region__c, Resource_Practice__c FROM pse__Assignment__c WHERE pse__Status__c = 'Scheduled' AND pse__End_Date__c >= TODAY AND pse__Role__c NOT IN ('Advisory - Technical Account Manager','Advisory - Technical Account Manager - 0 rate','Nonbillable Role') AND pse__Project__r.Subregion_new__c IN ('AMER TMT - 1','AMER TMT - 2') ORDER BY pse__Project__r.Name, pse__Resource__r.Name LIMIT 1000" \
  --result-format csv > "$OUTPUT_DIR/$FILENAME"

# Also keep a LATEST version that always overwrites
cp "$OUTPUT_DIR/$FILENAME" "$OUTPUT_DIR/$LATEST"

echo "✅ Export complete - $(wc -l < "$OUTPUT_DIR/$FILENAME") rows written to $OUTPUT_DIR/$FILENAME"
