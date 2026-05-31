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

# --- NEW: Independent Reference Sliders ---
st.sidebar.subheader("Reference Thresholds")
ref_stable = st.sidebar.slider("Stable Emission Reference (kg NH3 / GVE / yr)", min_value=3.0, max_value=15.0, value=7.9, step=0.1)
ref_app = st.sidebar.slider("Application Emission Reference (kg NH3 / GVE / yr)", min_value=1.0, max_value=10.0, value=3.4, step=0.1)

# Calculate total reference dynamically
ref_total = ref_stable + ref_app

if uploaded_file is not None:
    with st.spinner("Running REMAS model... Please wait."):
        try:
            df = run_pipeline(uploaded_file, sheet_name='Main input')
            st.sidebar.success("Model calculation successful!")
        except Exception as e:
            st.error(f"An error occurred during calculation: {e}")
            st.stop()
    
    # Filter out empty rows
    if 'Nr_koe' in df.columns:
        df = df.dropna(subset=['Nr_koe'])
        df = df[df['Nr_koe'] > 0]

    # Fallback/mocking to ensure Farm_ID and Round exist
    if 'Farm_ID' not in df.columns:
        df['Farm_ID'] = [f"Farm_{i+1}" for i in range(len(df))]
    if 'Round' not in df.columns:
        df['Round'] = ["Round_1" if i % 2 == 0 else "Round_2" for i in range(len(df))]

    col_total_emission = f"NH3_Emission_per_GVE_{method}"
    col_stable_emission = f"Stable_Emission_per_GVE_{method}"
    
    # Normalize Field Application Emission
    safe_gve = np.where(df['Total_GVE_Farm'] > 0, df['Total_GVE_Farm'], 1.0)
    df['App_Emission_per_GVE'] = (df['Emission_ManureApp_Total'] + df['Emission_FertiliserApp']) / safe_gve

    # Overall Status based on the SUM of the two references
    df['Status'] = df[col_total_emission].apply(lambda x: 'Above Total Threshold' if x > ref_total else 'Below Total Threshold')
    
    # Independent statuses for color coding the separate charts
    df['Stable_Status'] = df[col_stable_emission].apply(lambda x: 'Above' if x > ref_stable else 'Below')
    df['App_Status'] = df['App_Emission_per_GVE'].apply(lambda x: 'Above' if x > ref_app else 'Below')

    st.header("📊 Individual Farm Emission Comparison (Normalized per GVE)")
    
    # ---- Chart 1: Stable & Storage Emission ----
    st.subheader(f"🏠 1. Stable & Storage Emission (Threshold: {ref_stable} kg/GVE)")
    fig_stable = px.bar(
        df, 
        x="Farm_ID", 
        y=col_stable_emission, 
        color="Stable_Status",
        color_discrete_map={"Above": "#d62728", "Below": "#1f77b4"}, # Red if above, Blue if below
        hover_data=['Total_GVE_Farm'],
        labels={col_stable_emission: f"Stable Emission ({method}) [kg/GVE/yr]", "Stable_Status": "Status"}
    )
    fig_stable.add_hline(y=ref_stable, line_dash="dash", line_color="black", annotation_text=f"Ref: {ref_stable}")
    st.plotly_chart(fig_stable, use_container_width=True)

    # ---- Chart 2: Application Emission ----
    st.subheader(f"🚜 2. Field Application Emission (Threshold: {ref_app} kg/GVE)")
    fig_app = px.bar(
        df, 
        x="Farm_ID", 
        y="App_Emission_per_GVE", 
        color="App_Status",
        color_discrete_map={"Above": "#d62728", "Below": "#2ca02c"}, # Red if above, Green if below
        hover_data=['Emission_ManureApp_Total', 'Emission_FertiliserApp'],
        labels={"App_Emission_per_GVE": "Application Emission [kg/GVE/yr]", "App_Status": "Status"}
    )
    fig_app.add_hline(y=ref_app, line_dash="dash", line_color="black", annotation_text=f"Ref: {ref_app}")
    st.plotly_chart(fig_app, use_container_width=True)
    
    # Display statistics
    above_count = len(df[df['Status'] == 'Above Total Threshold'])
    st.warning(f"💡 Note: Combining Stable and Application, a total of **{above_count}** farm(s) exceed the overall reference line of **{ref_total:.1f} kg/GVE/yr**.")

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