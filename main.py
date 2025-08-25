import streamlit as st
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt

# Load Excel workbook and sheet using openpyxl
workbook = openpyxl.load_workbook("/Users/paul/Desktop/Python_Projects/Portfolio_Project/Sample_Investment_Portfolio.xlsx")
sheet = workbook.active  # First sheet

# Extract data from sheet into a list of rows
data = []
for row in sheet.iter_rows(values_only=True):
    data.append(row)

# Convert to pandas DataFrame
headers = data[0]  # first row as header
rows = data[1:]    # remaining rows as data
df = pd.DataFrame(rows, columns=headers)

# Streamlit App
st.set_page_config(page_title="Investment Portfolio Dashboard", layout="wide")

st.title("📊 Investment Portfolio Dashboard")

# Show data
st.subheader("Portfolio Allocation Table")
st.dataframe(df)

# Pie Chart for Allocation %
st.subheader("Portfolio Allocation (%)")
fig1, ax1 = plt.subplots()
ax1.pie(df["Allocation (%)"], labels=df["Asset Class"], autopct="%1.1f%%", startangle=90)
ax1.axis("equal")
st.pyplot(fig1)

# Bar Chart for Value
st.subheader("Portfolio Value Distribution (INR)")
fig2, ax2 = plt.subplots()
ax2.bar(df["Asset Class"], df["Value (INR)"])
ax2.set_ylabel("Value (INR)")
ax2.set_xlabel("Asset Class")
st.pyplot(fig2)
