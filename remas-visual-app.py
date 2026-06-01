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

# 🌟 ADJUSTMENT: Application slider unit is now kg/ha/yr with adjusted default/range
st.sidebar.subheader("Emission Reference Lines")
ref_stable = st.sidebar.slider("Stable Emission Reference (kg NH3 / GVE / yr)", min_value=3.0, max_value=15.0, value=7.9, step=0.1)
ref_app = st.sidebar.slider("Application Emission Reference (kg NH3 / ha / yr)", min_value=1.0, max_value=50.0, value=15.0, step=0.5)

if uploaded_file is not None:
    with st.spinner("Running REMAS model... Please wait."):
        try:
            df = run_pipeline(uploaded_file, sheet_name='Main input')
            st.sidebar.success("Model calculation successful!")
        except Exception as e:
            st.error(f"An error occurred during calculation: {e}")
            st.stop()
    
    # Filter out empty rows (ghost rows from Excel)
    if 'Nr_koe' in df.columns:
        df = df.dropna(subset=['Nr_koe'])
        df = df[df['Nr_koe'] > 0]

    # Fallback to ensure Farm_ID and Round exist, and force Round to string to prevent decimal years on axes
    if 'Farm_ID' not in df.columns:
        df['Farm_ID'] = [f"Farm_{i+1}" for i in range(len(df))]
    if 'Round' not in df.columns:
        df['Round'] = ["2023" if i % 2 == 0 else "2024" for i in range(len(df))]
    df['Round'] = df['Round'].astype(str)

    col_stable_emission = f"Stable_Emission_per_GVE_{method}"
    
    # ==========================================
    # 🌟 CORE ADJUSTMENT: Calculate total area and normalize by Hectare (Ha)
    # ==========================================
    # Safely get land areas (in case some columns are missing in Excel)
    ha_grass = df.get('Ha_Grass', pd.Series(0, index=df.index)).fillna(0)
    ha_mais = df.get('Ha_Mais', pd.Series(0, index=df.index)).fillna(0)
    ha_crop = df.get('Ha_Crop', pd.Series(0, index=df.index)).fillna(0)
    
    df['Total_Ha_Farm'] = ha_grass + ha_mais + ha_crop
    
    # Prevent division by zero if area or GVE is 0
    safe_ha = np.where(df['Total_Ha_Farm'] > 0, df['Total_Ha_Farm'], 1.0)
    safe_gve = np.where(df['Total_GVE_Farm'] > 0, df['Total_GVE_Farm'], 1.0)

    # Calculate application emission based on hectares
    df['App_Emission_per_Ha'] = (df['Emission_ManureApp_Total'] + df['Emission_FertiliserApp']) / safe_ha

    # Independently evaluate if stable and application emissions exceed thresholds
    df['Stable_Status'] = df[col_stable_emission].apply(lambda x: 'Above Threshold' if x > ref_stable else 'Below Threshold')
    df['App_Status'] = df['App_Emission_per_Ha'].apply(lambda x: 'Above Threshold' if x > ref_app else 'Below Threshold')
    
    # If either threshold is exceeded, the farm gets a warning flag
    df['Is_Warning'] = (df['Stable_Status'] == 'Above Threshold') | (df['App_Status'] == 'Above Threshold')

    st.header("📊 Individual Farm Emission Comparison (Normalized)")
    
    # ---- Chart 1: Stable & Storage Emission (kg/GVE) ----
    st.subheader(f"🏠 1. Stable & Storage Emission (Reference: {ref_stable} kg/GVE/yr)")
    fig_stable = px.bar(
        df, 
        x="Farm_ID", 
        y=col_stable_emission, 
        color="Stable_Status",
        color_discrete_map={"Above Threshold": "#d62728", "Below Threshold": "#1f77b4"}, 
        hover_data=['Total_GVE_Farm'],
        labels={col_stable_emission: f"Stable Emission ({method}) [kg/GVE/yr]", "Stable_Status": "Status"}
    )
    fig_stable.add_hline(y=ref_stable, line_dash="dash", line_color="black", annotation_text=f"Stable Ref: {ref_stable}")
    st.plotly_chart(fig_stable, use_container_width=True)

    # ---- Chart 2: Application Emission (kg/Ha) ----
    st.subheader(f"🚜 2. Field Application Emission (Reference: {ref_app} kg/ha/yr)")
    fig_app = px.bar(
        df, 
        x="Farm_ID", 
        y="App_Emission_per_Ha", 
        color="App_Status",
        color_discrete_map={"Above Threshold": "#d62728", "Below Threshold": "#2ca02c"}, 
        hover_data=['Emission_ManureApp_Total', 'Emission_FertiliserApp', 'Total_Ha_Farm'],
        labels={"App_Emission_per_Ha": "Application Emission [kg/ha/yr]", "App_Status": "Status"}
    )
    fig_app.add_hline(y=ref_app, line_dash="dash", line_color="black", annotation_text=f"Application Ref: {ref_app}")
    st.plotly_chart(fig_app, use_container_width=True)
    
    # Display warning statistics
    warning_count = df['Is_Warning'].sum()
    st.warning(f"💡 Overall, there are **{warning_count}** farm(s) that exceeded either the stable or the field application emission threshold.")

    st.markdown("---")

    # ==========================
    # MODULE B: Yearly Aggregation (Absolute Total Emission)
    # ==========================
    st.header("🌍 Yearly Aggregation (Absolute Total Emission)")
    
    # Calculate Absolute Total = (Stable per GVE * Total GVE) + (App per Ha * Total Ha)
    df['Absolute_Total_Emission'] = (df[col_stable_emission] * df['Total_GVE_Farm']) + (df['App_Emission_per_Ha'] * df['Total_Ha_Farm'])
    
    df_round = df.groupby('Round')['Absolute_Total_Emission'].sum().reset_index()
    
    fig_round = px.bar(
        df_round, 
        x="Round", 
        y="Absolute_Total_Emission", 
        text_auto='.2s',
        color="Round",
        title=f"Absolute Total Emission Aggregated by Year ({method} Method)",
        labels={"Absolute_Total_Emission": "Total Emission (kg NH3 / year)", "Round": "Year"}
    )
    st.plotly_chart(fig_round, use_container_width=True)

    # ==========================
    # MODULE C: Data Details Table
    # ==========================
    st.header("📋 Calculated Data Details")
    display_cols = ['Farm_ID', 'Round', 'Total_GVE_Farm', 'Total_Ha_Farm', col_stable_emission, 'Stable_Status', 'App_Emission_per_Ha', 'App_Status']
    existing_cols = [c for c in display_cols if c in df.columns]
    st.dataframe(df[existing_cols])

else:
    st.info("👈 Please upload your REMAS Input Excel file in the left sidebar to run the model and view the visualization.")