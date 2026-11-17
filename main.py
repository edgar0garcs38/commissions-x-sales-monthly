import pandas as pd
import re

# -----------------------------
# Helpers – Monetary parsing
# -----------------------------

def parse_value(val):
    """Convert Excel string with currency/parenthesis into float, robust to EU/US formats."""
    if pd.isna(val) or str(val).strip() == "":
        return 0.0

    val = str(val)
    val = val.replace("$", "").replace(" ", "").replace("´", "").strip()

    # Remove thousand separators like 1.234,56 → 1234,56
    val = re.sub(r"\.(?=\d{3}(?:,|$))", "", val)
    val = val.replace(",", ".")

    if val.startswith("(") and val.endswith(")"):
        val = "-" + val.strip("()")

    try:
        return float(val)
    except ValueError:
        return 0.0

# -----------------------------
# Helpers – Classification
# -----------------------------

def classify(val):
    if val > 0:
        return "Paid"
    elif val == 0:
        return "Neutral"
    else:
        return "Anomaly"

# -----------------------------
# Helpers – Policy normalization
# -----------------------------

def normalize_policy(pol):
    if pd.isna(pol):
        return ""
    # Force to string, strip whitespace, remove all non-alphanumeric, lowercase
    cleaned = re.sub(r'[^A-Za-z0-9]', '', str(pol)).lower()
    return cleaned


# -----------------------------
# Load data
# -----------------------------

# Main dataset
df_main = pd.read_excel("main.xlsx")

# Cancel report
df_cancel = pd.read_excel("cancelreport.xlsx")

# EmiReno dataset
df_emireno = pd.read_excel("Renovaciones.xlsx")

# -----------------------------
# Step 1: Paid Status Report
# -----------------------------

# Parse monetary values (adjust column name if needed)
df_main["PrimaCom_val"] = df_main["Prima Com."].apply(parse_value)

# Classify each row
df_main["Status"] = df_main["PrimaCom_val"].apply(classify)

paid_status_df = df_main[["Poliza", "Cliente", "Agente", "PrimaCom_val", "Status"]]

import pandas as pd
import re

# -----------------------------
# Helpers – Monetary parsing
# -----------------------------

def parse_value(val):
    """Convert Excel string with currency/parenthesis into float, robust to EU/US formats."""
    if pd.isna(val) or str(val).strip() == "":
        return 0.0

    val = str(val)
    val = val.replace("$", "").replace(" ", "").replace("´", "").strip()

    # Remove thousand separators like 1.234,56 → 1234,56
    val = re.sub(r"\.(?=\d{3}(?:,|$))", "", val)
    val = val.replace(",", ".")

    if val.startswith("(") and val.endswith(")"):
        val = "-" + val.strip("()")

    try:
        return float(val)
    except ValueError:
        return 0.0

# -----------------------------
# Helpers – Classification
# -----------------------------

def classify(val):
    if val > 0:
        return "Paid"
    elif val == 0:
        return "Neutral"
    else:
        return "Anomaly"

# -----------------------------
# Helpers – Policy normalization
# -----------------------------

def normalize_policy(pol):
    if pd.isna(pol):
        return ""
    # Remove all non-alphanumeric characters and lowercase
    return re.sub(r'[^a-zA-Z0-9]', '', str(pol)).lower()

# -----------------------------
# Load data
# -----------------------------

# Main dataset
df_main = pd.read_excel("main.xlsx")

# Cancel report
df_cancel = pd.read_excel("cancelreport.xlsx")

# EmiReno dataset
df_emireno = pd.read_excel("Renovaciones.xlsx")

# -----------------------------
# Step 1: Paid Status Report
# -----------------------------

# Parse monetary values (adjust column name if needed)
df_main["PrimaCom_val"] = df_main["Prima Com."].apply(parse_value)

# Classify each row
df_main["Status"] = df_main["PrimaCom_val"].apply(classify)

paid_status_df = df_main[["Poliza", "Cliente", "Agente", "PrimaCom_val", "Status"]]

# -----------------------------
# Step 2: Cancel Report Matching
# -----------------------------

# Normalize cancel policies
cancel_policies = df_cancel["POLIZA"].apply(normalize_policy).tolist()

# Normalize both datasets
df_main["Poliza_norm"] = df_main["Poliza"].apply(normalize_policy)
df_emireno["Poliza_norm"] = df_emireno["Poliza"].apply(normalize_policy)

# Drop canceled policies from BOTH datasets
df_main = df_main[~df_main["Poliza_norm"].isin(cancel_policies)]
df_emireno = df_emireno[~df_emireno["Poliza_norm"].isin(cancel_policies)]

# -----------------------------
# Step 3: EmiReno Matching (EmiReno → Main)
# -----------------------------

# Build lookup from main.xlsx
main_lookup = df_main.set_index("Poliza_norm")["Status"].to_dict()

def match_to_main(row):
    pol = row["Poliza_norm"]
    if pol in main_lookup:
        return pd.Series({
            "MatchedPolicyInMain": row["Poliza"],   # original format from EmiReno
            "FinalStatus": main_lookup[pol]
        })
    else:
        return pd.Series({
            "MatchedPolicyInMain": "",
            "FinalStatus": "Unpaid"
        })

match_results = df_emireno.copy()
match_results[["MatchedPolicyInMain", "FinalStatus"]] = match_results.apply(match_to_main, axis=1)
match_results = match_results[["Poliza", "MatchedPolicyInMain", "AGENTE", "FinalStatus"]]


# -----------------------------
# Step 4: Payment Global Summary
# -----------------------------

payments_global = (
    df_main["Status"]
    .value_counts()
    .rename_axis("Status")
    .reset_index(name="Count")
)

# -----------------------------
# Final Excel Output
# -----------------------------

with pd.ExcelWriter("final_report.xlsx") as writer:
    paid_status_df.to_excel(writer, sheet_name="Paid Status", index=False)
    match_results.to_excel(writer, sheet_name="Match Results", index=False)
    payments_global.to_excel(writer, sheet_name="Payment Global Summary", index=False)

print("Report generated: final_report.xlsx")


# -----------------------------
# Step 3: EmiReno Matching (EmiReno → Main)
# -----------------------------

# Normalize policy numbers for matching
df_main["Poliza_norm"] = df_main["Poliza"].apply(normalize_policy)
df_emireno["Poliza_norm"] = df_emireno["Poliza"].apply(normalize_policy)

main_policies = set(df_main["Poliza_norm"].tolist())

def match_to_main(row):
    pol = row["Poliza_norm"]
    if pol in main_policies:
        # If found in main, keep its paid status
        status = df_main.loc[df_main["Poliza_norm"] == pol, "Status"].iloc[0]
        return pd.Series({
            "MatchedPolicyInMain": row["Poliza"],   # original format from EmiReno
            "FinalStatus": status
        })
    else:
        return pd.Series({
            "MatchedPolicyInMain": "",
            "FinalStatus": "Unpaid"
        })

match_results = df_emireno.copy()
match_results[["MatchedPolicyInMain", "FinalStatus"]] = match_results.apply(match_to_main, axis=1)

match_results = match_results[["Poliza", "MatchedPolicyInMain", "AGENTE", "COMPANY", "ASEGURADO", "FinalStatus"]]


# -----------------------------
# Step 4: Payment Global Summary
# -----------------------------

payments_global = (
    df_main["Status"]
    .value_counts()
    .rename_axis("Status")
    .reset_index(name="Count")
)

# -----------------------------
# Final Excel Output
# -----------------------------

with pd.ExcelWriter("final_report.xlsx") as writer:
    paid_status_df.to_excel(writer, sheet_name="Paid Status", index=False)
    match_results.to_excel(writer, sheet_name="Match Results", index=False)
    payments_global.to_excel(writer, sheet_name="Payment Global Summary", index=False)

print("Report generated: final_report.xlsx")
