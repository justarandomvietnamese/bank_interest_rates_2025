import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
from typing import List, Dict

# =========================
# Config
# =========================
st.set_page_config(page_title="BANK DEPOSIT INTEREST RATES", page_icon="🏦", layout="wide")

EXCEL_PATH = "dashboard data.xlsx"  # Updated Excel file path
# SHEET_NAME = "Oct" # Updated Sheet name to "Oct" - This will be determined by the filter now

TERM_COLS_EN = ['1M(%)', '3M(%)', '6M(%)', '9M(%)', '12M(%)']
DIM_COLS_EN = ['BANK', 'CHANNEL', 'PRODUCT', 'CUSTOMERTYPE'] # Removed 'NOTE'

BRAND_COLORS = {
    "VIETCOMBANK": "#2E7D32",
    "VIETTINBANK": "#1565C0",
    "BIDV": "#D32F2F",
    "TECHCOMBANK": "#C62828",
    "SACOMBANK": "#1976D2",
    "TPBANK": "#8E24AA",
    "VPBANK": "#008A37",
    "MB": "#283593",
    "ACB": "#0277BD",
    "VIB": "#EF6C00",
    "HDBANK": "#C2185B",
}

# =========================
# Helper Functions
# =========================
def clean_str(s):
    if pd.isna(s):
        return s
    s = str(s).strip()
    s = re.sub(r"\\s+", " ", s)
    return s

def bank_alias(x: str) -> str:
    if x is None or pd.isna(x):
        return x
    s = str(x).strip().upper()
    alias = {
        "STB": "SACOMBANK",
        "TPB": "TPBANK",
        "VPB": "VPBANK",
        "HDB": "HDBANK",
        "SACOMBANK ": "SACOMBANK",
    }
    return alias.get(s, s)

def channel_normalize(x: str) -> str:
    if x is None or pd.isna(x):
        return x
    s = str(x).strip().upper()
    s = re.sub(r"\\s+", " ", s)
    if s in ["QUẦY"]:
        return "COUNTER"
    if s == "ONLINE":
        return "ONLINE"
    if s == "ONLINE ":
        return "ONLINE"
    return s

def product_normalize(x: str) -> str:
    if x is None or pd.isna(x):
        return x
    s = str(x).upper().strip()
    s = re.sub(r"\\s+", " ", s)
    if "ONLINE" in s and "SAVING" in s:
        return "ONLINE SAVINGS"
    if "ONLINE" in s and "DEPOSIT" in s:
        return "ONLINE DEPOSIT"
    if ("SAVING" in s or "SAVINGS" in s) and ("COUNTER" in s or "DEPOSIT" in s):
        return "SAVINGS AT COUNTER"
    if "TERM" in s and "DEPOSIT" in s:
        return "TERM DEPOSIT"
    return s

def customer_normalize(x: str) -> str:
    if x is None or pd.isna(x):
        return x
    s = str(x).strip().upper()
    s = re.sub(r"\\s+", " ", s)
    if s in ["SME", "SMES"]:
        return "SMEs"
    if s == "CORPORATE ":
        return "CORPORATE"
    return s

def build_plotly_color_map(categories: List[str], palette_name: str, color_by: str) -> Dict[str, str]:
    plotly_palette = None
    try:
        if hasattr(px.colors, 'qualitative') and hasattr(px.colors.qualitative, palette_name):
             palette_object = getattr(px.colors.qualitative, palette_name)
             if isinstance(palette_object, list):
                  plotly_palette = palette_object
             else:
                  st.warning(f"Attribute '{palette_name}' found in px.colors.qualitative but is not a list.")

        else:
             st.warning(f"Palette '{palette_name}' not found in px.colors.qualitative.")

    except Exception as e:
        st.warning(f"An error occurred accessing palette '{palette_name}': {e}")

    if plotly_palette is None or not isinstance(plotly_palette, list) or not plotly_palette:
        st.info(f"Using fallback 'Plotly' palette.")
        if hasattr(px.colors, 'qualitative') and hasattr(px.colors.qualitative, 'Plotly') and isinstance(px.colors.qualitative.Plotly, list):
            plotly_palette = px.colors.qualitative.Plotly
        else:
             st.error("Could not retrieve any color palette, including fallback. Using a hardcoded default.")
             plotly_palette = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']


    if color_by == "BANK":
        cmap = {}
        palette_index = 0
        for cat in categories:
            if cat in BRAND_COLORS:
                cmap[cat] = BRAND_COLORS[cat]
            else:
                if palette_index < len(plotly_palette):
                     cmap[cat] = plotly_palette[palette_index]
                     palette_index += 1
                else:
                     cmap[cat] = plotly_palette[0]


        return cmap
    else:
        return {cat: plotly_palette[i % len(plotly_palette)] for i, cat in enumerate(categories)}


@st.cache_data
def load_and_process_data(path: str, sheet_name="Oct") -> pd.DataFrame: # Updated sheet_name to be a parameter
    df = pd.DataFrame()
    try:
        # Read the first few rows to detect the header
        temp_df = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=10, engine="openpyxl")

        # Convert potential header rows to string and strip whitespace for matching
        temp_df_str = temp_df.applymap(lambda x: str(x).strip() if pd.notna(x) else '')

        # Define target column names based on user's input (exact names without extra spaces)
        TARGET_HEADER_NAMES_EXACT = ['BANK'] + TERM_COLS_EN # Look for 'BANK' or any of the term columns

        # Try to find the header row index by looking for the very first row that contains 'BANK' or any term column
        header_row_index = None
        for index, row in temp_df_str.iterrows():
             # Check if 'BANK' is in the row values or if any of the term columns are in the row values
             if 'BANK' in row.values or any(term in row.values for term in TERM_COLS_EN):
                  header_row_index = index
                  break

        # If header is not found in the first few rows, default to row 0 (index 0)
        if header_row_index is None:
            st.warning("Could not automatically detect header row containing 'BANK' or any Term columns in the first few rows. Assuming header is at row 1 (index 0).")
            header_row_index = 0 # Assume header at row 1 (index 0)


        # Now read the full data using the detected or assumed header row
        skip_rows = list(range(header_row_index)) # Skip rows before the header
        df = pd.read_excel(path, sheet_name=sheet_name, header=header_row_index, skiprows=skip_rows, engine="openpyxl")


        # Strip whitespace from column names immediately after reading (keep original casing from Excel)
        df.columns = [str(c).strip() for c in df.columns]


        # Define expected column names for selection and processing
        EXPECTED_COLS = DIM_COLS_EN + TERM_COLS_EN

        # Check if essential columns are present after reading with detected header and stripping whitespace
        if 'BANK' not in df.columns or not any(c in df.columns for c in TERM_COLS_EN):
             st.error("Error: 'BANK' column or any Term columns are missing after final read and stripping whitespace. Please check your Excel file structure and column names.")
             return pd.DataFrame()


        # Keep only relevant columns defined in DIM_COLS_EN and TERM_COLS_EN
        # Ensure columns actually exist in the dataframe before selecting
        valid_cols_to_keep = [c for c in EXPECTED_COLS if c in df.columns]

        # Check if essential dimension columns are in the valid_cols_to_keep list
        essential_dim_cols = ['BANK', 'CHANNEL'] # Define essential dimension columns
        if not set(essential_dim_cols).issubset(valid_cols_to_keep):
             missing_essential = set(essential_dim_cols) - set(valid_cols_to_keep)
             st.error(f"Error: Essential dimension columns are missing after selecting valid columns: {list(missing_essential)}. Please check your Excel file structure.")
             return pd.DataFrame()


        df = df.loc[:, valid_cols_to_keep]


        # Drop rows where all values are NaN after column selection
        df = df.dropna(how="all")

        # Drop rows where 'BANK' is NaN (assuming BANK is a key identifier)
        if 'BANK' in df.columns:
            df = df[df["BANK"].notna()].reset_index(drop=True)
        else:
             st.error("'BANK' column not found after cleaning steps.")
             return pd.DataFrame()


        # Clean and normalize string columns (using original casing)
        for c in set(DIM_COLS_EN).intersection(df.columns):
            if c in df.columns:
                df[c] = df[c].apply(clean_str)

        # Apply specific normalization functions (using original casing)
        if "BANK" in df.columns: df["BANK"] = df["BANK"].apply(bank_alias)
        if "CHANNEL" in df.columns: df["CHANNEL"] = df["CHANNEL"].apply(channel_normalize)
        if "PRODUCT" in df.columns: df["PRODUCT"] = df["PRODUCT"].apply(product_normalize)
        if "CUSTOMERTYPE" in df.columns: df["CUSTOMERTYPE"] = df["CUSTOMERTYPE"].apply(customer_normalize)


        # Convert term columns to numeric, coercing errors (using original casing)
        for c in TERM_COLS_EN:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")

        # Drop duplicates
        df = df.drop_duplicates().reset_index(drop=True)

        return df

    except FileNotFoundError:
        st.error(f"Error: File not found at {path}")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"An error occurred during data loading and initial processing: {e}")
        return pd.DataFrame()


@st.cache_data
def melt_data(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    term_cols = [c for c in TERM_COLS_EN if c in df.columns]
    id_cols = [c for c in DIM_COLS_EN if c in df.columns]

    if 'BANK' not in id_cols:
        st.error("Error: 'BANK' column is missing from dimension columns for melting.")
        return pd.DataFrame()

    if not id_cols or not term_cols:
         st.error("Error: Missing dimension or term columns for melting.")
         return pd.DataFrame()

    try:
        long = df.melt(
            id_vars=id_cols,
            value_vars=term_cols,
            var_name="Term",
            value_name="Interest Rate (%)"
        )
        if "Term" in long.columns and term_cols:
            long["Term"] = pd.Categorical(long["Term"], categories=term_cols, ordered=True)
        return long
    except Exception as e:
        st.error(f"An error occurred during data melting: {e}")
        return pd.DataFrame()

def build_plotly_color_map(categories: List[str], palette_name: str, color_by: str) -> Dict[str, str]:
    plotly_palette = None
    try:
        if hasattr(px.colors, 'qualitative') and hasattr(px.colors.qualitative, palette_name):
             palette_object = getattr(px.colors.qualitative, palette_name)
             if isinstance(palette_object, list):
                  plotly_palette = palette_object
             else:
                  st.warning(f"Attribute '{palette_name}' found in px.colors.qualitative but is not a list.")

        else:
             st.warning(f"Palette '{palette_name}' not found in px.colors.qualitative.")

    except Exception as e:
        st.warning(f"An error occurred accessing palette '{palette_name}': {e}")

    if plotly_palette is None or not isinstance(plotly_palette, list) or not plotly_palette:
        st.info(f"Using fallback 'Plotly' palette.")
        if hasattr(px.colors, 'qualitative') and hasattr(px.colors.qualitative, 'Plotly') and isinstance(px.colors.qualitative.Plotly, list):
            plotly_palette = px.colors.qualitative.Plotly
        else:
             st.error("Could not retrieve any color palette, including fallback. Using a hardcoded default.")
             plotly_palette = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']


    if color_by == "BANK":
        cmap = {}
        palette_index = 0
        for cat in categories:
            if cat in BRAND_COLORS:
                cmap[cat] = BRAND_COLORS[cat]
            else:
                if palette_index < len(plotly_palette):
                     cmap[cat] = plotly_palette[palette_index]
                     palette_index += 1
                else:
                     cmap[cat] = plotly_palette[0]


        return cmap
    else:
        return {cat: plotly_palette[i % len(plotly_palette)] for i, cat in enumerate(categories)}


# =========================
# Load & Prepare
# =========================
# Use a try-except block around the data loading and melting process
# raw_df = load_and_process_data(EXCEL_PATH, SHEET_NAME) # This will be called after month selection

# =========================
# Layout
# =========================
st.set_page_config(layout="wide")

# Sidebar for filters
with st.sidebar:
    st.title("BANK DEPOSIT INTEREST RATES")
    st.header("Filter Options")

    st.subheader("Select Filters")

    # Add month selection filter
    selected_month = st.selectbox("Select Month", ["Oct", "Sep", "Nov"]) # Added month selectbox


    # Load and process data based on selected month
    try:
        raw_df = load_and_process_data(EXCEL_PATH, selected_month)

        # Check if raw_df is empty after loading and processing
        if raw_df.empty:
             # Error message is already handled inside load_and_process_data
             st.stop()

        # Ensure raw_df has the 'BANK' column before melting (redundant check for safety)
        if "BANK" not in raw_df.columns:
             st.error("Internal Error: 'BANK' column is missing in the data after loading and initial processing (post-function check).")
             st.stop()


        df_long = melt_data(raw_df)

        # Check if df_long is empty after melting
        if df_long.empty:
            # Error message is already handled inside melt_data
            st.stop()

        # Ensure df_long has the 'BANK' column before proceeding to layout and filters (redundant check for safety)
        if "BANK" not in df_long.columns:
            st.error("Internal Error: 'BANK' column is still missing in the processed data after melting (post-function check).")
            st.stop()

    except Exception as e:
        # Catch potential exceptions during the overall load and melt process
        st.error(f"An unexpected error occurred during data loading and preparation: {e}")
        st.stop()


    def checkbox_multi_expander(title: str, series: pd.Series, key_prefix: str) -> List[str]:
        with st.expander(title, expanded=True if title == "BANK" else False):
            # Ensure series is not empty before getting unique values
            if not series.empty:
                values = series.dropna().unique().tolist()
                values = sorted(values, key=lambda x: str(x))
                selected = []
                # Check all by default
                for v in values:
                    checked = st.checkbox(str(v), value=True, key=f"sidebar_{key_prefix}_{v}") # Added sidebar_ prefix
                    if checked:
                        selected.append(v)
                return selected
            else:
                st.warning(f"No data available for {title} filter.")
                return []


    # Ensure columns exist in df_long before creating filters
    # Pass pd.Series(dtype='object') if column is missing to prevent error in checkbox_multi_expander
    selected_banks = checkbox_multi_expander("BANK", df_long["BANK"] if "BANK" in df_long.columns else pd.Series(dtype='object'), "bank")
    selected_products = checkbox_multi_expander("PRODUCT", df_long["PRODUCT"] if "PRODUCT" in df_long.columns else pd.Series(dtype='object'), "product")
    selected_customer_types = checkbox_multi_expander("CUSTOMER TYPE", df_long["CUSTOMERTYPE"] if "CUSTOMERTYPE" in df_long.columns else pd.Series(dtype='object'), "cust_type")
    selected_channels = checkbox_multi_expander("CHANNEL", df_long["CHANNEL"] if "CHANNEL" in df_long.columns else pd.Series(dtype='object'), "channel")


    # Color & Style Expander
    with st.expander("COLOR & STYLE", expanded=True):
        # Get available Plotly palettes
        # Dynamically get palette names from px.colors.qualitative
        plotly_palette_names = []
        if hasattr(px.colors, 'qualitative'):
            plotly_palette_names = sorted([attr for attr in dir(px.colors.qualitative) if isinstance(getattr(px.colors.qualitative, attr), list)])


        palette_name = st.selectbox(
            "Palette",
            plotly_palette_names if plotly_palette_names else ["Plotly"], # Provide a default if no palettes found
            index=plotly_palette_names.index("Plotly") if "Plotly" in plotly_palette_names else (0 if plotly_palette_names else 0)
        )
        color_by_options = [c for c in DIM_COLS_EN if c in df_long.columns] # Exclude NOTE already in DIM_COLS_EN
        # Set a default color_by if options exist, otherwise None or a default string
        color_by = st.selectbox("Color by", color_by_options if color_by_options else [],
                                 index=color_by_options.index("BANK") if "BANK" in color_by_options else (0 if color_by_options else None))


        sort_desc = st.checkbox("Sort by interest desc", value=True)
        # Removed Top N slider as per user request
        # top_n = st.slider("Top N by selected term", 3, 30, 15)
        show_text = st.checkbox("Show values on bars", value=True)


# Main content area
# Create 3 columns for the main content area: one empty, one for bar chart, one for pie chart
col1_main, col2_main, col3_main = st.columns([1, 2, 1])

# Apply filters based on selections from sidebar
filter_cols_to_check = ['BANK', 'CHANNEL', 'CUSTOMERTYPE', 'PRODUCT']
current_filters = {col: selected_banks if col == 'BANK' else selected_channels if col == 'CHANNEL' else selected_customer_types if col == 'CUSTOMERTYPE' else selected_products
                   for col in filter_cols_to_check if col in df_long.columns}

filter_condition = pd.Series(True, index=df_long.index)
for col, values in current_filters.items():
    if col in df_long.columns and values:
         filter_condition &= (df_long[col].isin(values))
    elif col in df_long.columns and not values:
         filter_condition &= False

filtered = df_long[filter_condition].copy()


term_options_filtered = filtered["Term"].dropna().unique().tolist() if "Term" in filtered.columns and not filtered.empty else []
# Add 'Overall Average' option and ensure correct order
term_options_ordered = [term for term in TERM_COLS_EN if term in term_options_filtered] + ["Overall Average"]


# Term Slider - Moved to the main content area, above the charts
if not term_options_ordered:
     st.warning("No terms available for selection with current filters.")
     selected_term = None
else:
    default_term_value = '12M(%)' if '12M(%)' in term_options_ordered else (term_options_ordered[-1] if len(term_options_ordered) > 1 else term_options_ordered[0])
    # Check if 'Overall Average' is the only option or if '12M(%)' is not available
    if default_term_value not in term_options_ordered and "Overall Average" in term_options_ordered:
        default_term_value = "Overall Average"
    elif default_term_value not in term_options_ordered and term_options_ordered:
         default_term_value = term_options_ordered[0] # Default to the first available term if 12M(%) and Overall Average are not options


    selected_term = st.select_slider("Select Term", options=term_options_ordered, value=default_term_value)


# Further filter data based on the selected term or use all data if 'Overall Average' is selected
term_df = pd.DataFrame()
if selected_term:
    if selected_term == "Overall Average":
        # For overall average, we need to calculate the average rate across all terms *per row* first,
        # then use this average for subsequent aggregations (like by bank or customer type).
        # Let's create a temporary column for the average rate across terms for each original row.
        # We need to go back to the wide format (raw_df filtered by sidebar) to calculate row-wise average.
        # However, the charts below (col2_main bar, col3_main pies) expect the 'Interest Rate (%)' column in the melted format.
        # A simpler approach for 'Overall Average' in these charts is to aggregate the already melted and filtered data ('filtered')
        # directly by the relevant dimensions (BANK for bar chart, CUSTOMERTYPE for pie charts)
        # taking the mean of 'Interest Rate (%)' across ALL terms for the selected filters.
        term_df = filtered.copy() # Use the fully filtered data for 'Overall Average' aggregations


    elif not filtered.empty and "Term" in filtered.columns:
        # For specific terms, filter the data as before
        term_df = filtered[filtered["Term"] == selected_term].copy()

    if not term_df.empty:
        term_df = term_df.dropna(subset=["Interest Rate (%)"])


# Overview Metrics - Moved to the main content area, below the term slider
st.subheader("Overview")
# Calculate max and min interest rates based on term_df (filtered by selected term)
# Ensure term_df is not empty before calculating max/min
max_rate = np.nanmax(term_df["Interest Rate (%)"]) if not term_df.empty and "Interest Rate (%)" in term_df.columns and term_df["Interest Rate (%)"].notna().any() else "N/A"
min_rate = np.nanmin(term_df["Interest Rate (%)"]) if not term_df.empty and "Interest Rate (%)" in term_df.columns and term_df["Interest Rate (%)"].notna().any() else "N/A"

# Use filtered for overall avg rate and record count
overall_avg_rate = np.nanmean(filtered['Interest Rate (%)']) if not filtered.empty and "Interest Rate (%)" in filtered.columns and filtered['Interest Rate (%)'].notna().any() else "N/A"
record_count = len(filtered)


# Adjusted columns for metrics and swapped positions
k1, k2, k3, k4, k5, k6 = st.columns(6)
with k1:
    st.metric("Records", f"{record_count:,}")
with k2:
    if not filtered.empty and "BANK" in filtered.columns:
         st.metric("Banks", filtered["BANK"].nunique()) # Swapped position
    else:
         st.metric("Banks", 0)
with k3:
    if not filtered.empty and "Interest Rate (%)" in filtered.columns and filtered['Interest Rate (%)'].notna().any():
        st.metric("Avg Rate (all terms)", f"{overall_avg_rate:.2f}%" if isinstance(overall_avg_rate, (int, float)) else overall_avg_rate) # Swapped position
    else:
        st.metric("Avg Rate (all terms)", "N/A")
with k4:
    if not term_df.empty and "Interest Rate (%)" in term_df.columns and term_df["Interest Rate (%)"].notna().any():
        # Calculate median based on the data used for plotting (term_df)
        st.metric(f"Median Rate ({selected_term or ''})", f"{np.nanmedian(term_df['Interest Rate (%)']):.2f}%")
    else:
        st.metric(f"Median Rate ({selected_term or ''})", "N/A")
with k5:
    st.metric(f"Highest Rate ({selected_term or ''})", f"{max_rate:.2f}%" if isinstance(max_rate, (int, float)) else max_rate)
with k6:
    st.metric(f"Lowest Rate ({selected_term or ''})", f"{min_rate:.2f}%" if isinstance(min_rate, (int, float)) else min_rate)


# Comparison Chart in col1_main
with col1_main:
    st.subheader("Avg Rate by Channel")
    # Use term_df for the channel comparison chart as it reflects the selected term or 'Overall Average'
    if not term_df.empty and "CHANNEL" in term_df.columns and "Interest Rate (%)" in term_df.columns:
        # Calculate average rate by channel for the selected term or all terms
        # Ensure there's data after grouping before plotting
        avg_rate_by_channel = term_df.groupby("CHANNEL")["Interest Rate (%)"].mean().reset_index()

        if not avg_rate_by_channel.empty and "Interest Rate (%)" in avg_rate_by_channel.columns:
            fig_channel_avg = px.bar(
                avg_rate_by_channel,
                x="CHANNEL",
                y="Interest Rate (%)",
                title=f"Average Interest Rate by Channel for {'All Terms' if selected_term == 'Overall Average' else selected_term}",
                text="Interest Rate (%)", # Add text labels
                labels={"CHANNEL": "Channel", "Interest Rate (%)": "Average Rate (%)"},
                hover_data={"CHANNEL": False, "Interest Rate (%)": ':.2f'} # Customize hover to show only formatted rate
            )
            # Reverted texttemplate to its previous state
            fig_channel_avg.update_traces(texttemplate='%{y:.2f}%', textposition='outside') # Format text labels and position, removed text color
            fig_channel_avg.update_layout(
                height=480, # Increased height to match main bar chart
                margin=dict(t=50, r=10, b=10, l=10),
                bargap=0.2,
                template="plotly_white",
                yaxis_title="Average Rate (%)",
                hoverlabel=dict(bgcolor="white", font_size=12, font_family="Lato") # Optional: Customize hover label appearance
            )
            st.plotly_chart(fig_channel_avg, use_container_width=True)
        else:
            st.info("No data available to calculate average rate by channel with current filters and term.")
    else:
        st.info("Required data for Channel Comparison chart is missing with current filters and term.")

    # Add a horizontal line separator
    st.markdown("---")

    # Add the new stacked bar chart here, inside col1_main
    st.subheader("Customer Distribution by Product Type")
    if not filtered.empty and "PRODUCT" in filtered.columns and "CUSTOMERTYPE" in filtered.columns:
        # Group by PRODUCT and CUSTOMERTYPE and count occurrences
        customer_product_counts = filtered.groupby(["PRODUCT", "CUSTOMERTYPE"]).size().reset_index(name="count")

        if not customer_product_counts.empty:
            # Build color map for CUSTOMERTYPE
            cust_type_categories = customer_product_counts["CUSTOMERTYPE"].dropna().unique().tolist()
            cust_type_color_map = build_plotly_color_map(cust_type_categories, palette_name, "CUSTOMERTYPE")

            fig_stacked_bar = px.bar(
                customer_product_counts,
                x="PRODUCT",
                y="count",
                color="CUSTOMERTYPE",
                color_discrete_map=cust_type_color_map,
                title="Customer Type Distribution by Product",
                labels={"PRODUCT": "Product Type", "count": "Number of Records", "CUSTOMERTYPE": "Customer Type"},
                hover_data=["PRODUCT", "CUSTOMERTYPE", "count"],
                text='count' # Add text labels for counts
            )

            fig_stacked_bar.update_layout(
                height=400, # Adjusted height
                margin=dict(t=50, r=10, b=10, l=10),
                xaxis_title="Product Type",
                yaxis_title="Number of Records",
                legend_title="Customer Type",
                bargap=0.25,
                template="plotly_white",
                hovermode="x unified"
            )
            fig_stacked_bar.update_traces(textposition='inside') # Position text labels inside bars, removed text color
            fig_stacked_bar.update_xaxes(tickangle=-45)
            st.plotly_chart(fig_stacked_bar, use_container_width=True)
        else:
            st.info("No data to display for Customer Distribution by Product Type.")
    else:
        st.warning("Required data for Customer Distribution by Product Type is missing with current filters.")


# Main Bar Chart - Using Plotly
with col2_main: # Place in col2_main
     st.subheader(f"Interest Rates by Bank for {'All Terms (Average)' if selected_term == 'Overall Average' else selected_term}")
     # Use term_df which is already filtered by selected_term or is the full 'filtered' data for Overall Average
     if not term_df.empty and "BANK" in term_df.columns and "Interest Rate (%)" in term_df.columns:
        plot_df = term_df.copy()

        # If 'Overall Average' is selected, group by BANK and calculate the mean interest rate across terms
        if selected_term == "Overall Average":
            # Group by BANK and calculate the mean interest rate across ALL terms for the selected sidebar filters
            # The data for this aggregation is already in term_df (which is 'filtered' when Overall Average is selected)
            plot_df = plot_df.groupby("BANK")["Interest Rate (%)"].mean().reset_index()
            current_color_by = "BANK" if "BANK" in plot_df.columns else None # Force color by BANK when Overall Average is selected
            legend_title = "Bank"
            agg_cols = ["BANK"] # Define columns used for aggregation

        else:
            # For specific terms, use the data filtered by the specific term (term_df)
            # Aggregate by BANK and the color_by dimension if applicable.
            current_color_by = color_by if color_by in plot_df.columns else None
            legend_title = current_color_by if current_color_by else None

            # Explicitly group by BANK and the color_by dimension (if exists and not BANK) to ensure aggregation.
            agg_cols = ["BANK"]
            if current_color_by and current_color_by != "BANK" and current_color_by in plot_df.columns:
                 agg_cols.append(current_color_by)

            # Decide aggregation function - using mean for general case
            agg_func = "mean"
            if not plot_df.empty and agg_cols and "Interest Rate (%)" in plot_df.columns:
                 # Ensure aggregation columns exist in the dataframe
                 valid_agg_cols = [col for col in agg_cols if col in plot_df.columns]
                 if valid_agg_cols:
                      plot_df = plot_df.groupby(valid_agg_cols)["Interest Rate (%)"].agg(agg_func).reset_index()
                 else:
                      st.warning("Invalid aggregation columns for main bar chart with specific term.")
                      plot_df = pd.DataFrame() # Clear plot_df if aggregation is not possible
            else:
                 # plot_df is already empty, or missing essential columns
                 plot_df = pd.DataFrame()


        if not plot_df.empty and "Interest Rate (%)" in plot_df.columns:
            if sort_desc:
                plot_df = plot_df.sort_values("Interest Rate (%)", ascending=False)

            bank_order = plot_df["BANK"].astype(str).tolist() if "BANK" in plot_df.columns else []

            # Re-build color map after aggregation if color_by changed
            color_categories = plot_df[current_color_by].astype(str).fillna("N/A").unique().tolist() if current_color_by and current_color_by in plot_df.columns else []
            color_map = build_plotly_color_map(color_categories, palette_name, current_color_by) if current_color_by and current_color_by in plot_df.columns else {}

            fig_bar = px.bar(
                plot_df,
                x="BANK",
                y="Interest Rate (%)",
                color=current_color_by,
                color_discrete_map=color_map,
                category_orders={"BANK": bank_order} if bank_order else None,
                hover_data=[col for col in agg_cols if col in plot_df.columns] + ["Interest Rate (%)"] if not plot_df.empty else [],
                text=plot_df["Interest Rate (%)"].round(2).astype(str) if "Interest Rate (%)" in plot_df.columns else None
            )
            if show_text and fig_bar.data and "text" in fig_bar.data[0]: # Check if fig_bar.data is not empty
                 fig_bar.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0) # Removed text color
            else:
                 fig_bar.update_traces(textposition="none", marker_line_width=0)

            y_min_plot = plot_df["Interest Rate (%)"].min() if not plot_df.empty and "Interest Rate (%)" in plot_df.columns and plot_df["Interest Rate (%)"].notna().any() else 0
            y_max_plot = plot_df["Interest Rate (%)"].max() if not plot_df.empty and "Interest Rate (%)" in plot_df.columns and plot_df["Interest Rate (%)"].notna().any() else 10
            y_range = [max(0, y_min_plot * 0.95), y_max_plot * 1.05] if y_min_plot is not None and y_max_plot is not None else [0, 10]


            fig_bar.update_layout(
                height=480,
                margin=dict(t=50, r=10, b=10, l=10),
                xaxis_title="Bank",
                yaxis_title="Interest Rate (%)",
                legend_title=legend_title,
                bargap=0.25,
                template="plotly_white",
                hovermode="x unified",
                showlegend=bool(current_color_by),
                yaxis=dict(range=y_range)
            )
            fig_bar.update_xaxes(tickangle=-45)
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
             st.info("No data available to display the bar chart after aggregation.")


     # Add a horizontal line separator
     st.markdown("---")

     # Add the new line chart code Block OR Scatter Plot code Block
     if selected_term != "Overall Average":
         st.subheader(f"Interest Rate vs. Product Type for {selected_term or ''}, Grouped by Channel")
         if not filtered.empty and "Term" in filtered.columns and "Interest Rate (%)" in filtered.columns and "PRODUCT" in filtered.columns and "CHANNEL" in filtered.columns:

             # Filter data for the selected term (using filtered, not term_df, as term_df might be aggregated for Overall Avg)
             line_chart_df = filtered[filtered["Term"] == selected_term].copy()

             if not line_chart_df.empty:
                 # Group by Product and Channel and calculate the mean interest rate
                 avg_rate_by_product_channel = line_chart_df.groupby(["PRODUCT", "CHANNEL"])["Interest Rate (%)"].mean().reset_index()

                 if not avg_rate_by_product_channel.empty and "Interest Rate (%)" in avg_rate_by_product_channel.columns:
                      # Build color map for CHANNEL
                      channel_categories = avg_rate_by_product_channel["CHANNEL"].dropna().unique().tolist()
                      channel_color_map = build_plotly_color_map(channel_categories, palette_name, "CHANNEL")

                      fig_line = px.line(
                          avg_rate_by_product_channel,
                          x="PRODUCT",
                          y="Interest Rate (%)",
                          color="CHANNEL", # Color by Channel
                          color_discrete_map=channel_color_map,
                          title=f"Average Interest Rate vs. Product Type for {selected_term or ''}, Grouped by Channel",
                          markers=True, # Add markers to the line
                          labels={"PRODUCT": "Product Type", "Interest Rate (%)": "Average Rate (%)", "CHANNEL": "Channel"},
                           hover_data=["PRODUCT", "CHANNEL", "Interest Rate (%)"]
                      )
                      fig_line.update_layout(
                          height=400,
                          margin=dict(t=50, r=10, b=10, l=10),
                          yaxis_title="Average Rate (%)",
                          xaxis_title="Product Type",
                          legend_title="Channel",
                          template="plotly_white",
                          hovermode="x unified"
                      )
                      fig_line.update_xaxes(tickangle=-45) # Rotate x-axis labels for readability
                      st.plotly_chart(fig_line, use_container_width=True)
                 else:
                     st.info(f"No data available to plot interest rate by product type and channel for {selected_term}.")
             else:
                 st.info(f"No data available for the selected term ({selected_term}) to plot interest rate vs. product type by channel.")
         else:
             st.info("Required columns for the line chart (Term, Interest Rate (%), PRODUCT, CHANNEL) are missing with current filters.")
     else:
          # Scatter Plot for Overall Average
          st.subheader("Average Interest Rate by Bank (Overall Average)")
          # For the scatter plot with Overall Average, we need the average rate per bank across all terms
          if not filtered.empty and "BANK" in filtered.columns and "Interest Rate (%)" in filtered.columns:
               avg_rate_by_bank_overall = filtered.groupby("BANK")["Interest Rate (%)"].mean().reset_index()

               if not avg_rate_by_bank_overall.empty and "Interest Rate (%)" in avg_rate_by_bank_overall.columns:
                   # Build color map for BANK
                   bank_categories_scatter = avg_rate_by_bank_overall["BANK"].dropna().unique().tolist()
                   bank_color_map_scatter = build_plotly_color_map(bank_categories_scatter, palette_name, "BANK")

                   fig_scatter = px.scatter(
                       avg_rate_by_bank_overall,
                       x="BANK",
                       y="Interest Rate (%)",
                       color="BANK", # Color by Bank
                       color_discrete_map=bank_color_map_scatter,
                       title="Average Interest Rate by Bank (Overall Average)",
                       size="Interest Rate (%)", # Size points by interest rate
                       hover_name="BANK", # Show Bank name on hover
                       hover_data=["BANK", "Interest Rate (%)"], # Show Bank and Avg Rate on hover
                       labels={"BANK": "Bank", "Interest Rate (%)": "Average Interest Rate (%)"}
                   )
                   fig_scatter.update_layout(
                       height=480, # Increased height to match main bar chart
                       margin=dict(t=50, r=10, b=10, l=10),
                       xaxis_title="Bank",
                       yaxis_title="Average Interest Rate (%)",
                       legend_title="Bank",
                       template="plotly_white",
                       hovermode="closest"
                   )
                   fig_scatter.update_xaxes(tickangle=-45) # Rotate x-axis labels for readability
                   st.plotly_chart(fig_scatter, use_container_width=True)
               else:
                    st.info("No data available to display the scatter plot for Overall Average.")
          else:
               st.warning("Required columns for the scatter plot (BANK, Interest Rate (%)) are missing with current filters.")


with col3_main: # Place in col3_main
    st.subheader(f"Customer Type Distribution (Count) for {'All Terms' if selected_term == 'Overall Average' else selected_term}")
    # Use filtered for the count plot when selected_term is Overall Average, otherwise use term_df
    data_for_pie_count = filtered.copy() if selected_term == "Overall Average" else term_df.copy()

    if not data_for_pie_count.empty and "CUSTOMERTYPE" in data_for_pie_count.columns:
        counts = data_for_pie_count["CUSTOMERTYPE"].value_counts(dropna=False)
        if not counts.empty:
            pie_df = counts.reset_index()
            pie_df.columns = ["CUSTOMERTYPE", "count"]
            pie_color_by = "CUSTOMERTYPE" if "CUSTOMERTYPE" in pie_df.columns else None
            pie_color_map = build_plotly_color_map(pie_df["CUSTOMERTYPE"].astype(str).tolist(), palette_name, "CUSTOMERTYPE") if pie_color_by else {}

            fig_pie = px.pie(
                pie_df,
                names="CUSTOMERTYPE",
                values="count",
                color=pie_color_by,
                color_discrete_map=pie_color_map,
                hole=0.3,
                 # Define hover data for pie chart as a list of column names
                hover_data=["CUSTOMERTYPE", "count"]
            )
            # Update traces to show percentage and value on hover and percentage on slices
            fig_pie.update_traces(textinfo='percent', hovertemplate='%{label}: %{value}<br>Percentage: %{percent}<extra></extra>') # Changed textinfo to 'percent', removed text color


            fig_pie.update_layout(height=447, margin=dict(t=40, r=10, b=10, l=10), legend_title="CUSTOMERTYPE", showlegend=True) # Adjusted height
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
             st.info("No data to display for Customer Type Distribution.")
    else:
        st.warning("No data available to display Customer Type Distribution with current filters and term.")

    # Add a horizontal line separator
    st.markdown("---")

    # Add a new pie chart for distribution of interest rates by customer type
    st.subheader(f"Interest Rate Distribution by Customer Type for {'All Terms (Average)' if selected_term == 'Overall Average' else selected_term}")
    # Use filtered for aggregation when selected_term is Overall Average, otherwise use term_df
    data_for_pie_interest = filtered.copy() if selected_term == "Overall Average" else term_df.copy()

    if not data_for_pie_interest.empty and "CUSTOMERTYPE" in data_for_pie_interest.columns and "Interest Rate (%)" in data_for_pie_interest.columns:
        # Calculate the average interest rate per customer type
        # Grouping by CUSTOMERTYPE and taking the mean of 'Interest Rate (%)' across all terms is the correct behavior here for Overall Average.
        # For specific terms, term_df already contains only the data for that term.
        interest_distribution_df = data_for_pie_interest.groupby("CUSTOMERTYPE")["Interest Rate (%)"].mean().reset_index()

        if not interest_distribution_df.empty and "Interest Rate (%)" in interest_distribution_df.columns:
            pie_interest_color_by = "CUSTOMERTYPE" if "CUSTOMERTYPE" in interest_distribution_df.columns else None
            pie_interest_color_map = build_plotly_color_map(interest_distribution_df["CUSTOMERTYPE"].astype(str).tolist(), palette_name, "CUSTOMERTYPE") if pie_interest_color_by else {}

            fig_pie_interest = px.pie(
                interest_distribution_df,
                names="CUSTOMERTYPE",
                values="Interest Rate (%)", # Use average interest rate as values
                color=pie_interest_color_by,
                color_discrete_map=pie_interest_color_map,
                hole=0.3,
                 # Define hover data to show customer type and average interest rate
                hover_data=["CUSTOMERTYPE", "Interest Rate (%)"]
            )
            # Update traces to show percentage and value on hover and percentage on slices
            fig_pie_interest.update_traces(textinfo='percent', hovertemplate='%{label}: %{value:.2f}%<br>Percentage: %{percent}<extra></extra>') # Changed textinfo to 'percent', removed text color


            fig_pie_interest.update_layout(height=360, margin=dict(t=40, r=10, b=10, l=10), legend_title="CUSTOMERTYPE", showlegend=True) # Show legend for pie chart
            st.plotly_chart(fig_pie_interest, use_container_width=True)
        else:
            st.info("No data available to display Interest Rate Distribution by Customer Type.")
    else:
        st.warning("Required data for Interest Rate Distribution by Customer Type is missing with current filters and term.")


# Heatmap - Below the columns
st.subheader("Heatmap: interest by term (max per bank)")

# Heatmap should always show individual terms, not an aggregate 'Overall Average' view in this format
# Removed the conditional check for selected_term != "Overall Average"

with st.expander("Heatmap filters", expanded=False):
    # These filters should apply to the heatmap data source (which is 'filtered' now)
    hm_channel_options = ["(All)"] + sorted(filtered["CHANNEL"].dropna().unique().tolist()) if not filtered.empty and "CHANNEL" in filtered.columns else ["(All)"]
    hm_channel = st.selectbox("Channel", hm_channel_options, key="hm_channel_filter", index=0)

    hm_cust_options = ["(All)"] + sorted(filtered["CUSTOMERTYPE"].dropna().unique().tolist()) if not filtered.empty and "CUSTOMERTYPE" in filtered["CUSTOMERTYPE"].dropna().unique().tolist() else ["(All)"]
    hm_cust = st.selectbox("Customer Type", hm_cust_options, key="hm_cust_filter", index=0)

    hm_product_options = ["(All)"] + sorted(filtered["PRODUCT"].dropna().unique().tolist()) if not filtered.empty and "PRODUCT" in filtered.columns else ["(All)"]
    hm_product = st.selectbox("Product", hm_product_options, key="hm_product_filter", index=0)


heat_df = filtered.copy()
if not heat_df.empty:
    # Apply heatmap-specific filters
    if hm_channel != "(All)" and "CHANNEL" in heat_df.columns:
        heat_df = heat_df[heat_df["CHANNEL"] == hm_channel]
    if hm_cust != "(All)" and "CUSTOMERTYPE" in heat_df.columns:
        heat_df = heat_df[heat_df["CUSTOMERTYPE"] == hm_cust]
    if hm_product != "(All)" and "PRODUCT" in heat_df.columns:
        heat_df = heat_df[heat_df["PRODUCT"] == hm_product]

    # The heatmap should always show data for all terms based on the *filtered* data, not term_df
    # Check if necessary columns for pivot_table are in heat_df
    required_heatmap_cols = {"BANK", "Term", "Interest Rate (%)"}
    if not heat_df.empty and required_heatmap_cols.issubset(heat_df.columns):
        # Ensure there is data with non-null Interest Rate for pivoting
        if heat_df["Interest Rate (%)"].notna().any():
            pvt = heat_df.pivot_table(
                index="BANK",
                columns="Term",
                values="Interest Rate (%)",
                aggfunc="max" # Using max as aggregation function for the heatmap as in previous versions
            )
            if not pvt.empty:
                # Ensure columns (terms) are in the correct order
                ordered_cols = [c for c in TERM_COLS_EN if c in pvt.columns]
                # Only keep columns that exist in the pivot table
                ordered_cols_existing = [col for col in ordered_cols if col in pvt.columns]
                pvt = pvt[ordered_cols_existing]

                fig_hm = px.imshow(
                    pvt,
                    text_auto=".2f",
                    aspect="auto",
                    color_continuous_scale="Blues",
                    labels=dict(x="Term", y="Bank", color="Rate (%)")
                    # Removed hover_data as it's not supported by px.imshow
                )
                fig_hm.update_layout(height=520, margin=dict(t=40, r=10, b=10, l=10), coloraxis_colorbar=dict(title="Rate (%)"))
                 # Update x-axis to use ordered categories for correct display
                fig_hm.update_xaxes(categoryorder="array", categoryarray=ordered_cols_existing)
                st.plotly_chart(fig_hm, use_container_width=True)
            else:
                 st.info("No data to display heatmap after filtering and pivoting.")
        else:
             st.info("No valid interest rate data found for heatmap after filtering.")
    else:
        st.info("Required columns for heatmap are missing after filtering, or heatmap data is empty.")

# Filtered Data Table - Below the heatmap
st.subheader("Filtered Data")
# Display filtered data based on sidebar filters, not term selection
if not filtered.empty:
    st.dataframe(filtered, use_container_width=True, height=380)
    if not filtered.empty:
         csv = filtered.to_csv(index=False).encode("utf-8-sig")
         st.download_button("Download CSV", data=csv, file_name="interest_filtered.csv", mime="text/csv")
else:
    st.warning("No data nào phù hợp với bộ lọc hiện tại.")
