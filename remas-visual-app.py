# app.py
import streamlit as st
import pandas as pd
import plotly.express as px
from remas_pipeline import run_pipeline

# Set page configuration
st.set_page_config(page_title="REMAS Live Calculator", layout="wide")
st.title("🐄 REMAS Farm Ammonia Emission Live Calculator")

# Sidebar: File upload and parameter settings
st.sidebar.header("1. Upload Input Data")
uploaded_file = st.sidebar.file_uploader("Upload REMAS Input Excel File", type=["xlsx"])

st.sidebar.header("2. Parameter Settings")
method = st.sidebar.radio("Select Calculation Method", ["VCRE", "MUN"])
ref_line = st.sidebar.slider("Set Emission Reference Line (kg NH3 / cow / year)", min_value=5.0, max_value=20.0, value=11.3, step=0.1)

if uploaded_file is not None:
    # Run the pipeline with a loading spinner
    with st.spinner("Calculating REMAS model... Please wait."):
        try:
            # Pass the uploaded file directly to your pipeline
            df = run_pipeline(uploaded_file, sheet_name='Main input')
            st.sidebar.success("Calculation Successful!")
        except Exception as e:
            st.error(f"An error occurred during calculation: {e}")
            st.stop()
    
    # Ensure necessary grouping columns exist for visualization
    if 'Farm_ID' not in df.columns:
        df['Farm_ID'] = [f"Farm_{i+1}" for i in range(len(df))]
    if 'Round' not in df.columns:
        df['Round'] = ["Round_1" if i % 2 == 0 else "Round_2" for i in range(len(df))]

    # Dynamically get the column names based on the selected method
    col_total_emission = f"NH3_Emission_per_GVE_{method}"
    col_stable_emission = f"Stable_Emission_per_GVE_{method}"
    col_tan = f"Total_Corrected_TAN_{method}"

    # Differentiate between farms above and below the reference line
    df['Status'] = df[col_total_emission].apply(lambda x: 'Above Threshold' if x > ref_line else 'Below Threshold')

    # ==========================
    # MODULE A: Individual Farm Comparison (Normalized)
    # ==========================
    st.header("📊 Individual Farm Emission Comparison (Normalized per GVE)")
    
    fig = px.bar(
        df, 
        x="Farm_ID", 
        y=col_total_emission, 
        color="Status",
        color_discrete_map={"Above Threshold": "red", "Below Threshold": "green"},
        hover_data=[col_tan, col_stable_emission, "Emission_FertiliserApp"],
        labels={col_total_emission: f"Total Emission ({method}) [kg/GVE/yr]"}
    )
    
    fig.add_hline(y=ref_line, line_dash="dash", line_color="black", annotation_text=f"Reference: {ref_line}")
    st.plotly_chart(fig, use_container_width=True)

    above_count = len(df[df['Status'] == 'Above Threshold'])
    st.warning(f"💡 There are **{above_count}** farm(s) exceeding the {ref_line} kg/GVE/yr reference line.")

    # ==========================
    # MODULE B: Round Aggregation (Absolute Total Emission)
    # ==========================
    st.header("🌍 Round Aggregation (Absolute Total Emission)")
    
    # Calculate absolute total emission = Normalized Emission * Cow Number (Nr_koe)
    df['Absolute_Total_Emission'] = df[col_total_emission] * df['Nr_koe']
    
    # Group by Round and sum
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
    display_cols = ['Farm_ID', 'Round', 'Nr_koe', col_tan, col_stable_emission, 'Emission_FertiliserApp', col_total_emission, 'Status']
    
    # Only display columns that actually exist in the dataframe to prevent errors
    existing_cols = [c for c in display_cols if c in df.columns]
    st.dataframe(df[existing_cols])

else:
    st.info("👈 Please upload your REMAS **Input** Excel file in the left sidebar to run the model.")