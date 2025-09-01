import re
import pandas as pd
import streamlit as st
import altair as alt
from io import BytesIO

# -----------------------------
# Streamlit App UI
# -----------------------------
st.set_page_config(page_title="Weekly Validation Comparison", layout="wide")
st.title("Weekly Validation Comparison Report")

st.header("Upload Files")

meta_file = st.file_uploader("Upload Meta Data File (Excel)", type=["xlsx"])
week_files = st.file_uploader("Upload TWO Weekly CSV Files (e.g., week_34.csv and week_35.csv)", type=["csv"], accept_multiple_files=True)

if not meta_file or len(week_files) != 2:
    st.warning("Please upload the Meta Data File and TWO weekly CSV files.")
    st.stop()

# -----------------------------
# Helper Functions
# -----------------------------
def extract_week_number(filename):
    match = re.search(r"week[_ ](\d+)", filename, re.IGNORECASE)
    return int(match.group(1)) if match else None

def status_change(before, after):
    if pd.isna(before):
        return f"Missing → {after}"
    if pd.isna(after):
        return f"{before} → Missing"
    if before == after:
        return "No Change"
    return f"{before} → {after}"

def color_status(val):
    if str(val).lower() == "green":
        return "background-color: lightgreen"
    if str(val).lower() == "red":
        return "background-color: salmon"
    if str(val).lower() == "yellow":
        return "background-color: khaki"
    return ""

def color_changes(val):
    if "→" in str(val):
        last_status = str(val).split("→")[-1].strip()
        return color_status(last_status)
    if str(val) == "No Change":
        return "background-color: lightgrey"
    return ""

def color_meta(val):
    if str(val) == "Green":
        return "background-color: lightgreen"
    if str(val) == "Yellow":
        return "background-color: khaki"
    if str(val) == "Red":
        return "background-color: salmon"
    return ""

def meta_style(df):
    return df.style.applymap(color_meta, subset=df.columns[1:])

def color_status_change(val):
    if "→" in str(val):
        last_status = str(val).split("→")[-1].strip()
        return color_status(last_status)
    return ""

# -----------------------------
# Process Uploaded Files
# -----------------------------
weeks = {}
for file in week_files:
    week_num = extract_week_number(file.name)
    if week_num:
        weeks[week_num] = pd.read_csv(file)

sorted_weeks = sorted(weeks.keys())
week_before, week_after = sorted_weeks[0], sorted_weeks[1]

week_before_df = weeks[week_before].add_suffix(f"_W{week_before}")
week_after_df = weeks[week_after].add_suffix(f"_W{week_after}")

week_before_df = week_before_df.rename(columns={f"OMNI_MODULE_W{week_before}": "OMNI_MODULE"})
week_after_df = week_after_df.rename(columns={f"OMNI_MODULE_W{week_after}": "OMNI_MODULE"})

df = pd.merge(week_before_df, week_after_df, on="OMNI_MODULE", how="left")

df["CHANGES"] = df.apply(lambda row: status_change(row[f"STATUS_W{week_before}"], row[f"STATUS_W{week_after}"]), axis=1)

pct_cols = [
    f"VAL_PCT_W{week_before}",
    f"VAL_PCT_W{week_after}",
    f"VAL_SALES_W{week_before}",
    f"VAL_SALES_W{week_after}",
    f"NBL_SHARE_W{week_before}",
    f"NBL_SHARE_W{week_after}",
]

for col in pct_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce") * 100
    df[col] = df[col].round(0)
    df[col] = df[col].apply(lambda x: f"{int(x)}%" if pd.notna(x) else "N/A")

final_df = df[
    [
        "OMNI_MODULE",
        f"STATUS_W{week_before}",
        f"STATUS_W{week_after}",
        "CHANGES",
        f"VAL_PCT_W{week_before}",
        f"VAL_PCT_W{week_after}",
        f"VAL_SALES_W{week_before}",
        f"VAL_SALES_W{week_after}",
        f"NBL_SHARE_W{week_before}",
        f"NBL_SHARE_W{week_after}",
    ]
]

final_df = final_df.rename(
    columns={
        f"STATUS_W{week_before}": f"STATUS (Week {week_before})",
        f"STATUS_W{week_after}": f"STATUS (Week {week_after})",
        f"VAL_PCT_W{week_before}": f"VAL_PCT (W{week_before})",
        f"VAL_PCT_W{week_after}": f"VAL_PCT (W{week_after})",
        f"VAL_SALES_W{week_before}": f"VAL_SALES (W{week_before})",
        f"VAL_SALES_W{week_after}": f"VAL_SALES (W{week_after})",
        f"NBL_SHARE_W{week_before}": f"NBL_SHARE (W{week_before})",
        f"NBL_SHARE_W{week_after}": f"NBL_SHARE (W{week_after})",
    }
)

# ✅ Remove rows with "No Change"
comparison_df = final_df[final_df["CHANGES"] != "No Change"].reset_index(drop=True)

nbl_share_before = pd.to_numeric(final_df[f"NBL_SHARE (W{week_before})"].str.replace("%", ""), errors="coerce")
nbl_share_after = pd.to_numeric(final_df[f"NBL_SHARE (W{week_after})"].str.replace("%", ""), errors="coerce")

diff = ((nbl_share_after - nbl_share_before) / nbl_share_before.replace(0, pd.NA)) * 100
final_df["NBL_Share_Pct"] = diff.round(0).apply(lambda x: f"{int(x)}%" if pd.notna(x) else "N/A")

# -----------------------------
# Meta Data Updates
# -----------------------------
meta_df = pd.read_excel(meta_file)
status_counts = week_after_df[f"STATUS_W{week_after}"].value_counts().to_dict()
meta_df[f"Week {week_after}"] = meta_df["OMNI_MODULE"].map(status_counts).fillna(0).astype(int)

week_cols = [c for c in meta_df.columns if c.startswith("Week ")]
last6weeks_df = meta_df.loc[:, ["OMNI_MODULE"] + week_cols[-6:]]

# -----------------------------
# Summary & Recommendations
# -----------------------------
summary_lines = []
success_modules = []
attention_modules = []

for _, row in final_df.iterrows():
    module = row["OMNI_MODULE"]
    change = row["CHANGES"]
    val_pct_wafter = row[f"VAL_PCT (W{week_after})"]
    val_sales_wafter = row[f"VAL_SALES (W{week_after})"]
    nbl_share_wafter = row[f"NBL_SHARE (W{week_after})"]

    if "Red → Green" in change or "Yellow → Green" in change:
        success_modules.append(f"{module} (status improved, Validation ↑, NBL {nbl_share_wafter})")
    elif "Green → Red" in change or "Green → Yellow" in change or "Yellow → Red" in change:
        attention_modules.append(f"{module} (status regressed, NBL {nbl_share_wafter}, Validation {val_pct_wafter})")

summary_lines.append("Summary & Recommendations")
summary_lines.append("• Success Highlights:")
if success_modules:
    summary_lines.append("Modules showing strong improvements: " + ", ".join(success_modules))
else:
    summary_lines.append("No major success shifts this week.")

summary_lines.append("")
summary_lines.append("• Focus Needed:")
if attention_modules:
    summary_lines.append("Modules that need immediate attention: " + ", ".join(attention_modules))
else:
    summary_lines.append("No critical regressions this week.")

summary_text = "\n".join(summary_lines)
summary_df = pd.DataFrame({"Summary": summary_lines})

# -----------------------------
# Key Changes Insights
# -----------------------------
key_changes_rows = []
for _, row in final_df.iterrows():
    if row["CHANGES"] != "No Change":
        val_pct_diff = (
            pd.to_numeric(row[f"VAL_PCT (W{week_after})"].replace("%",""), errors="coerce") -
            pd.to_numeric(row[f"VAL_PCT (W{week_before})"].replace("%",""), errors="coerce")
        )
        val_sales_diff = (
            pd.to_numeric(row[f"VAL_SALES (W{week_after})"].replace("%",""), errors="coerce") -
            pd.to_numeric(row[f"VAL_SALES (W{week_before})"].replace("%",""), errors="coerce")
        )
        nbl_share_diff = (
            pd.to_numeric(row[f"NBL_SHARE (W{week_after})"].replace("%",""), errors="coerce") -
            pd.to_numeric(row[f"NBL_SHARE (W{week_before})"].replace("%",""), errors="coerce")
        )

        insight = ""
        if "Red → Green" in row["CHANGES"]:
            insight = "Full validation achieved; NBL completely eliminated" if nbl_share_diff == -100 else "Strong improvement"
        elif "Red → Yellow" in row["CHANGES"]:
            insight = "Validation and sales percentage improved; moderate NBL drop"
        elif "Yellow → Green" in row["CHANGES"]:
            insight = "Improved validation coverage and sales volume; better brand identification"
        elif "→ Red" in row["CHANGES"]:
            insight = "Regression observed; focus needed"
        else:
            insight = "Status shift observed"

        key_changes_rows.append({
            "OMNI_MODULE": row["OMNI_MODULE"],
            "STATUS CHANGE": row["CHANGES"],
            "VAL_PCT": f"{int(val_pct_diff)}%" if pd.notna(val_pct_diff) else "N/A",
            "VAL_SALES": f"{int(val_sales_diff)}%" if pd.notna(val_sales_diff) else "N/A",
            "NBL_SHARE": f"{int(nbl_share_diff)}%" if pd.notna(nbl_share_diff) else "N/A",
            "KEY INSIGHTS": insight
        })

key_changes_df = pd.DataFrame(key_changes_rows)

# -----------------------------
# Save Final Excel Report
# -----------------------------
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    (
        comparison_df.style
        .applymap(color_status, subset=[f"STATUS (Week {week_before})", f"STATUS (Week {week_after})"])
        .applymap(color_changes, subset=["CHANGES"])
        .to_excel(writer, sheet_name="Comparison", index=False)
    )
    meta_style(last6weeks_df).to_excel(writer, sheet_name="Last_6_Weeks_Status", index=False)
    summary_df.to_excel(writer, sheet_name="Summary & Recommendations", index=False)
    (
        key_changes_df.style
        .applymap(color_status_change, subset=["STATUS CHANGE"])
        .to_excel(writer, sheet_name="Key_Changes", index=False)
    )

# -----------------------------
# Streamlit Outputs
# -----------------------------
st.header("Download Report")
st.download_button(
    label="Download Final Report Excel",
    data=output.getvalue(),
    file_name="Final_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

meta_output = BytesIO()
meta_df.to_excel(meta_output, index=False, engine="openpyxl")
st.download_button(
    label="Download Updated Meta Data File",
    data=meta_output.getvalue(),
    file_name="Meta Data File.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.header("Preview Data")
st.subheader("Comparison Table (Only Changes)")
st.dataframe(comparison_df)

st.subheader("Last 6 Weeks Status")
st.dataframe(last6weeks_df)

st.subheader("Summary & Recommendations")
st.text(summary_text)

st.subheader("Key Changes")
st.dataframe(key_changes_df)

# -----------------------------
# 📊 Visual Insights with Altair
# -----------------------------
st.header("📊 Visual Insights")

# 1. Status Distribution
status_counts_df = pd.DataFrame({
    "Week": ["Week "+str(week_before)] * len(week_before_df[f"STATUS_W{week_before}"].value_counts()) +
            ["Week "+str(week_after)] * len(week_after_df[f"STATUS_W{week_after}"].value_counts()),
    "Status": list(week_before_df[f"STATUS_W{week_before}"].value_counts().index) +
              list(week_after_df[f"STATUS_W{week_after}"].value_counts().index),
    "Count": list(week_before_df[f"STATUS_W{week_before}"].value_counts().values) +
             list(week_after_df[f"STATUS_W{week_after}"].value_counts().values)
})

status_chart = alt.Chart(status_counts_df).mark_bar().encode(
    x="Week:N",
    y="Count:Q",
    color="Status:N",
    column="Status:N"
).properties(title="Status Distribution (Before vs After)")

st.altair_chart(status_chart, use_container_width=True)

# 2. Validation % Movement Scatter
val_chart = alt.Chart(final_df).mark_circle(size=80).encode(
    x=f"VAL_PCT (W{week_before}):Q",
    y=f"VAL_PCT (W{week_after}):Q",
    tooltip=["OMNI_MODULE", f"VAL_PCT (W{week_before})", f"VAL_PCT (W{week_after})", "CHANGES"],
    color="CHANGES:N"
).properties(title="Validation % Movement per Module")

line = alt.Chart(pd.DataFrame({'x':[0,100], 'y':[0,100]})).mark_line(strokeDash=[5,5]).encode(x="x", y="y")
st.altair_chart(val_chart + line, use_container_width=True)

# 3. NBL Share Change for Changed Modules
key_changes_df["NBL_CHANGE"] = pd.to_numeric(key_changes_df["NBL_SHARE"].str.replace("%",""), errors="coerce")
changes_chart = alt.Chart(key_changes_df).mark_bar().encode(
    x="OMNI_MODULE:N",
    y="NBL_CHANGE:Q",
    color="STATUS CHANGE:N",
    tooltip=["OMNI_MODULE", "STATUS CHANGE", "VAL_PCT", "VAL_SALES", "NBL_SHARE", "KEY INSIGHTS"]
).properties(title="Key Module Changes (NBL Share Δ)")

st.altair_chart(changes_chart, use_container_width=True)
