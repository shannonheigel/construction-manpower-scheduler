import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference

st.set_page_config(page_title="Construction Manpower Scheduler", layout="wide")

# --- Generate manpower curve ---
def generate_manpower_curve(weeks, peak_manpower):
    x = np.linspace(-2, 2, weeks)
    y = np.exp(-x**2)
    y = y / y.max() * peak_manpower
    return y

# --- Build schedule ---
def build_schedule(total_weeks, departments, total_scope_units):
    df = pd.DataFrame({'Week': range(1, total_weeks + 1)})
    for dept, scope_ratio in departments.items():
        dept_scope = total_scope_units * scope_ratio
        peak_manpower = dept_scope / total_weeks * 2
        df[dept] = generate_manpower_curve(total_weeks, peak_manpower).round(2)
    df['Total Manpower'] = df[list(departments.keys())].sum(axis=1)
    return df

# --- Excel Export ---
def to_excel(df):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Manpower Schedule"

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    chart = LineChart()
    chart.title = "Construction Manpower Schedule"
    chart.style = 10
    chart.y_axis.title = "Manpower"
    chart.x_axis.title = "Week"
    chart.width = 20
    chart.height = 12

    data = Reference(ws, min_col=2, max_col=1+len(df.columns)-1, min_row=1, max_row=len(df)+1)
    cats = Reference(ws, min_col=1, min_row=2, max_row=len(df)+1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "H2")

    wb.save(output)
    return output.getvalue()

# --- UI ---
st.title("🏗️ Construction Manpower Scheduler")

st.sidebar.header("📅 Project Settings")
total_weeks = st.sidebar.slider("Total Project Weeks", 4, 52, 20)
total_scope_units = st.sidebar.slider("Total Scope Units", 100, 5000, 1000, step=100)

st.sidebar.header("🛠️ Department Allocation")
engineering = st.sidebar.slider("Engineering", 0.0, 1.0, 0.25)
superintendents = st.sidebar.slider("Superintendents", 0.0, 1.0, 0.15)
foreman = st.sidebar.slider("Foreman", 0.0, 1.0, 0.2)
electricians = st.sidebar.slider("Electricians", 0.0, 1.0, 0.4)

total_ratio = engineering + superintendents + foreman + electricians

if total_ratio > 1.0:
    st.sidebar.error("Total allocation exceeds 100%. Adjust the sliders.")
    st.stop()

departments = {
    'Engineering': engineering,
    'Superintendents': superintendents,
    'Foreman': foreman,
    'Electricians': electricians
}

df = build_schedule(total_weeks, departments, total_scope_units)

st.subheader("📊 Manpower Schedule Preview")
st.line_chart(df.set_index("Week")[list(departments.keys()) + ["Total Manpower"]])
st.dataframe(df, use_container_width=True)

excel_data = to_excel(df)

st.download_button(
    label="📥 Download Excel Schedule",
    data=excel_data,
    file_name="Construction_Schedule.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
