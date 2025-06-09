import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dealer Group Rebate Calculator", layout="centered")
st.title("Dealer Group Rebate Calculator Dashboard")

st.markdown("""
This dashboard mirrors the SUMMARY calculator in your Excel, but uses your accurate, updated data.\
- **Upload your formatted Excel** (DealerGroup_TAM_SAM_DUM_NUM_FORMATTED_MAY.xlsx) below.\
- **Select a dealer group** to view all metrics.\
- **Edit rebate tiers** and see live calculations.
""")

# Upload the corrected Excel file (now Rebate_data.xlsx)
data_file = st.file_uploader(
    "Upload Rebate_data.xlsx",
    type=["xlsx"],
    help="This should be your corrected output, not the original Dealer Group Rebate Summary.xlsx."
)

if data_file:
    df = pd.read_excel(data_file)
    st.write("Columns in your file:", list(df.columns))  # Show columns for debugging and mapping
    group_names = df["Group Name | Website"].dropna().unique()
    selected_group = st.selectbox("Select Dealer Group", group_names)
    group_row = df[df["Group Name | Website"] == selected_group].iloc[0]

    # --- Calculator Box ---
    st.header(f"Summary for {selected_group}")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("DUM (Customer Stores)", int(group_row["DUM"]))
        st.metric("Prospects in SAM", int(group_row["DUM"])) # If you want a different mapping, adjust here
        st.metric("SAM", int(group_row["SAM"]))
    with col2:
        st.metric("TAM", int(group_row["TAM"]))
        st.metric("Current SAM Penetration %", f"{group_row['SAM PEN']*100:.2f}%")

        # Try to find the Net New ARR column dynamically
        net_new_arr_col = None
        for col in df.columns:
            if 'arr' in col.lower() and 'net' in col.lower():
                net_new_arr_col = col
                break
        if net_new_arr_col:
            st.metric("Net New ARR", f"${group_row[net_new_arr_col]:,.0f}")
        else:
            st.warning("Net New ARR column not found. Check your file's column names.")

    st.divider()
    st.subheader("Rebate Tier Table (Editable)")
    # Editable rebate tiers
    default_tiers = {
        "Below 75%": {"min": 0.00, "max": 0.74, "rebate": 0.00},
        "Tier 1": {"min": 0.75, "max": 0.89, "rebate": 0.10},
        "Tier 2": {"min": 0.90, "max": 1.00, "rebate": 0.15}
    }
    tiers = {}
    for tier, vals in default_tiers.items():
        st.markdown(f"**{tier}**")
        col_min, col_max, col_rebate = st.columns(3)
        min_val = col_min.number_input(f"{tier} Min", value=vals["min"], key=f"{tier}_min")
        max_val = col_max.number_input(f"{tier} Max", value=vals["max"], key=f"{tier}_max")
        rebate_val = col_rebate.number_input(f"{tier} Rebate %", value=vals["rebate"]*100, step=0.01, key=f"{tier}_rebate")
        tiers[tier] = {"min": min_val, "max": max_val, "rebate": rebate_val/100}

    # --- Live Rebate Calculation ---
    sam_pen = group_row['SAM PEN']
    rebate = 0.0
    current_tier = None
    for tier, vals in tiers.items():
        if vals['min'] <= sam_pen <= vals['max']:
            rebate = vals['rebate']
            current_tier = tier
            break
    st.success(f"**Calculated Rebate for {selected_group}: {rebate*100:.2f}% (Tier: {current_tier if current_tier else 'N/A'})**")

    st.divider()
    st.subheader("Rebate Calculator Table (All Tiers)")
    # --- Actuals ---
    dum = int(group_row["DUM"]) if "DUM" in group_row else 0
    vum = int(group_row["TAM"]) if "TAM" in group_row else 0  # If you want a different mapping for VUM, adjust here
    # Use columns named 'MRR' and ARR ('ARR', 'Net New ARR', or 'Net ARR', case-insensitive)
    mrr_col = None
    arr_col = None
    arr_candidates = ['arr', 'net new arr', 'net arr']
    for col in df.columns:
        if col.strip().lower() == 'mrr':
            mrr_col = col
        if col.strip().lower() in arr_candidates:
            arr_col = col
    if not arr_col:
        st.warning(f"Could not find an ARR column ('ARR', 'Net New ARR', or 'Net ARR') in your file. Please check your data. Columns found: {list(df.columns)}")
        arr = 0
    else:
        arr = float(group_row[arr_col]) if pd.notnull(group_row[arr_col]) else 0
    if mrr_col:
        mrr = float(group_row[mrr_col]) if pd.notnull(group_row[mrr_col]) else 0
    else:
        mrr = arr / 12 if arr else 0

    # --- Calculate per tier ---
    table_data = []
    below_75_net_mrr = None
    below_75_net_arr = None
    for tier, vals in tiers.items():
        rebate_pct = vals['rebate']
        rebate_mo = mrr * rebate_pct
        rebate_yr = arr * rebate_pct
        mrr_net = mrr - rebate_mo
        arr_net = arr - rebate_yr
        # Save Below 75% net for gain/loss calculation
        if tier == "Below 75%":
            below_75_net_mrr = mrr_net
            below_75_net_arr = arr_net
        mrr_gain_loss = mrr_net - (below_75_net_mrr if below_75_net_mrr is not None else 0)
        arr_gain_loss = arr_net - (below_75_net_arr if below_75_net_arr is not None else 0)
        table_data.append({
            "Tier": tier,
            "DUM": dum,
            "VUM (EoM)": vum,
            "MRR": f"${mrr:,.0f}",
            "ARR": f"${arr:,.0f}",
            "Rebate $ (Mo.)": f"${rebate_mo:,.0f}",
            "Rebate $ (Yr.)": f"${rebate_yr:,.0f}",
            "MRR Net of Rebate": f"${mrr_net:,.0f}",
            "ARR Net of Rebate": f"${arr_net:,.0f}",
            "MRR Net Gain/Loss": f"${mrr_gain_loss:,.0f}",
            "ARR Net Gain/Loss": f"${arr_gain_loss:,.0f}"
        })

    # Display as styled DataFrame
    import pandas as pd
    table_df = pd.DataFrame(table_data)
    def highlight_current(s):
        return ["background-color: #d1ffd6; font-weight: bold" if v == current_tier else "" for v in s]
    st.dataframe(table_df.style.apply(highlight_current, subset=["Tier"]))

    st.divider()
    st.caption("All values are calculated from your uploaded, corrected data. Edit tiers above to see live results.")
else:
    st.info("Upload your corrected Excel file to begin.")
