import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# 1. Import your pipeline function
from remas_pipeline import run_pipeline

st.set_page_config(page_title="REMAS Live Calculator", layout="wide")
st.title("🐄 REMAS Farm Ammonia Emission Live Dashboard")

st.sidebar.header("1. Upload Input Data")
uploaded_file = st.sidebar.file_uploader("Upload REMAS Input Excel File", type=["xlsx"])

st.sidebar.header("2. Parameter Settings")
method = st.sidebar.radio("Select Calculation Method", ["VCRE", "MUN"])
ref_line = st.sidebar.slider("Set Total Emission Reference (kg NH3 / GVE / yr)", min_value=5.0, max_value=20.0, value=11.3, step=0.1)

if uploaded_file is not None:
    with st.spinner("Running REMAS model... Please wait."):
        try:
            df = run_pipeline(uploaded_file, sheet_name='Main input')
            st.sidebar.success("Model calculation successful!")
        except Exception as e:
            st.error(f"An error occurred during calculation: {e}")
            st.stop()
    
    # ==========================================
    # CORE FIX 1: Filter out ghost empty rows from Excel
    # We use 'Nr_koe' (Number of cows) as the validation metric.
    # ==========================================
    if 'Nr_koe' in df.columns:
        df = df.dropna(subset=['Nr_koe'])  # Remove NaN rows
        df = df[df['Nr_koe'] > 0]          # Remove rows where cow count is 0

    # Fallback/mocking to ensure Farm_ID and Round exist
    if 'Farm_ID' not in df.columns:
        df['Farm_ID'] = [f"Farm_{i+1}" for i in range(len(df))]
    if 'Round' not in df.columns:
        df['Round'] = ["Round_1" if i % 2 == 0 else "Round_2" for i in range(len(df))]

    # Dynamic column names based on selected method
    col_total_emission = f"NH3_Emission_per_GVE_{method}"
    col_stable_emission = f"Stable_Emission_per_GVE_{method}"
    
    # ==========================================
    # CORE FIX 2: Normalize Field Application Emission
    # Sum up Manure and Fertiliser app emissions, then divide by GVE
    # ==========================================
    safe_gve = np.where(df['Total_GVE_Farm'] > 0, df['Total_GVE_Farm'], 1.0)
    df['App_Emission_per_GVE'] = (df['Emission_ManureApp_Total'] + df['Emission_FertiliserApp']) / safe_gve

    # Flag farms that exceed the total threshold
    df['Status'] = df[col_total_emission].apply(lambda x: 'Above Threshold' if x > ref_line else 'Below Threshold')

    st.header("📊 Individual Farm Emission Comparison (Normalized per GVE)")
    
    # ==========================================
    # CORE FIX 3: Split charts into Stable and Application
    # ==========================================
    
    # ---- Chart 1: Stable & Storage Emission ----
    st.subheader("🏠 1. Stable & Storage Emission")
    fig_stable = px.bar(
        df, 
        x="Farm_ID", 
        y=col_stable_emission, 
        color_discrete_sequence=["#1f77b4"], # Uniform Blue
        hover_data=['Total_GVE_Farm'],
        labels={col_stable_emission: f"Stable Emission ({method}) [kg/GVE/yr]"}
    )
    # Set reference line for stable emission at 70% of total reference line
    stable_ref = ref_line * 0.70
    fig_stable.add_hline(y=stable_ref, line_dash="dash", line_color="orange", annotation_text=f"Stable Ref (~{stable_ref:.1f})")
    st.plotly_chart(fig_stable, use_container_width=True)


    # ---- Chart 2: Application Emission ----
    st.subheader("🚜 2. Field Application Emission")
    fig_app = px.bar(
        df, 
        x="Farm_ID", 
        y="App_Emission_per_GVE", 
        color_discrete_sequence=["#2ca02c"], # Uniform Green
        hover_data=['Emission_ManureApp_Total', 'Emission_FertiliserApp'],
        labels={"App_Emission_per_GVE": "Application Emission [kg/GVE/yr]"}
    )
    # Set reference line for application emission at 30% of total reference line
    app_ref = ref_line * 0.30
    fig_app.add_hline(y=app_ref, line_dash="dash", line_color="orange", annotation_text=f"Application Ref (~{app_ref:.1f})")
    st.plotly_chart(fig_app, use_container_width=True)

    
    # Display statistics
    above_count = len(df[df['Status'] == 'Above Threshold'])
    st.warning(f"💡 Note: Combining Stable and Application, a total of **{above_count}** farm(s) exceed the overall reference line of {ref_line} kg/GVE/yr.")

    st.markdown("---")

    # ==========================
    # MODULE B: Round Aggregation (Absolute Total Emission)
    # ==========================
    st.header("🌍 Round Aggregation (Absolute Total Emission)")
    df['Absolute_Total_Emission'] = df[col_total_emission] * df['Total_GVE_Farm']
    df_round = df.groupby('Round')['Absolute_Total_Emission'].sum().reset_index()
    
    fig_round = px.bar(
        df_round, 
        x="Round", 
        y="Absolute_Total_Emission", 
        text_auto='.2s',
        color="Round",
        title=f"Absolute Total Emission Aggregated by Round ({method} Method)",
        labels={"Absolute_Total_Emission": "Total Emission (kg NH3 / year)"}
    )
    st.plotly_chart(fig_round, use_container_width=True)

    # ==========================
    # MODULE C: Data Details Table
    # ==========================
    st.header("📋 Calculated Data Details")
    display_cols = ['Farm_ID', 'Round', 'Total_GVE_Farm', col_stable_emission, 'App_Emission_per_GVE', col_total_emission, 'Status']
    existing_cols = [c for c in display_cols if c in df.columns]
    st.dataframe(df[existing_cols])

else:
    st.info("👈 Please upload your REMAS Input Excel file in the left sidebar to run the model and view the visualization.")