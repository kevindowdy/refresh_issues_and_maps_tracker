# Refresh Issues and MAPs Tracker Per Business Unit

## Overview
The Refresh Issues & MAPs Tracker Script is designed to generate standardized, business-unit-specific trackers for Issue and MAP oversight using data refreshed daily from Archer (via the FIG Issue Management Model).

## Setup Instructions
1. Clone repository
2. Update `USERNAME` 
3. Update `USERNAME`
4. Update `DFS_INPUT_FILE`, `DPS_EPS_INPUT_FILE`, `CARDS_INPUT_FILE`, `CAPS_INPUT_FILE`, `MR_INPUT_FILE`
5. Update `BUSINESS_UNITS.business_leader_names`
6. Update `BUSINESS_UNITS.col_to_search`

## End-to-End Workflow
1. Download the updated FIG model file 
2. Download the previous date's tracker
3. Run the Script: `python generate_tracker.py`
4. Upload generated trackers to the designated SharePoint team folder

## Data Processing and Transformation Details
1. `Source Data Ingestion` - The script loads two sheets from the FIG model:
   Issues dataset (Issue metadata, status, risk rating, hierarchy - dump issues sheet)
   MAPs dataset (MAP ownership, status, due dates, update history - dump maps sheet)
2. `MAP Deduplication` - MAP data is deduplicated to retain only the most recent record per MAP ID based on the Last Updated field.
3. `Issue–MAP Merging` - Issues and MAPs are merged into a single dataset using Issue ID relationships.
4. `Hash Generation` - A unique Hash is generated using the Issue ID and MAP ID.
5. `Discussion-Driven Filtering` - The script isolates only records requiring active governance review including items where:
   Issue status indicates "open", "past due", "past due - pending map owner approval", "past due - pending aso approval" AND
   MAP status is active and not "cancelled", "draft", "completed", "draft", "map cancelled", "draft - pending approvals", "map cancellation pending aso approval"
6. `Business Unit Segmentation` - Data is segmented into business-unit-specific datasets using configurable filters:
   Filtering is applied using a specified column (e.g., MC-3 Name)
   Business leader mappings are fully configurable
   Supports organizational changes without code restructuring
7. `Comment and ETA Carry-Forward` - If a previous tracker is provided, the script pulls these columns to the new data using the Hash key:
   Comments, Past Due ETA updates
8. `Data Enrichment` - Summary updates prefixed with last update date. Comments automatically stamped with current date.
9. `MAP Compliance Calculation` - The script evaluates whether MAP updates meet governance-defined cadence requirements.

   * 7 days → Past Due MAPs
   * 14 days → MAPs due within 60 days
   * 30 days → MAPs due beyond 60 days

## Output Structure
For each business unit, the script generates an Excel workbook with the following sheets:

1. `Issues and MAPs Sheet` - Detailed Issue and MAP records requiring discussion
2. `MAPs Compliance Sheet` - Open and past-due MAPs only. Includes compliance calculations and update metrics and color-coded compliance indicators (green/red)
3. `Summary Sheet` - Executive-level overview including: _Total MAP counts_, _Open/Past Due volume_, _Compliance counts and percentages_
