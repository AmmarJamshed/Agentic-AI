#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import xlwings as xw
import os
import tkinter as tk
from tkinter import filedialog

# 🔹 Ensure 'openpyxl' is up-to-date
try:
    import openpyxl
    if openpyxl.__version__ < "3.1.0":
        print("⚠️ Updating openpyxl...")
        os.system("pip install --upgrade openpyxl")
except ImportError:
    print("⚠️ Installing openpyxl...")
    os.system("pip install openpyxl")

# 🔹 File Selection Dialog
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title="Select an Excel File", filetypes=[("Excel Files", "*.xlsx;*.xls")])

if not file_path:
    print("❌ No file selected. Exiting...")
    exit()

# 🔹 Load the Selected Excel File
try:
    df = pd.read_excel(file_path, engine="openpyxl")

    # ✅ Identify Numeric and Date Columns
    numeric_cols = df.select_dtypes(include=["number"]).columns
    date_cols = df.select_dtypes(include=["datetime"]).columns

    # ✅ Data Summary - Compute Statistics
    summary = df.describe(datetime_is_numeric=True) if not df.empty else "No numerical data available"
    print("\n✅ Data Summary:")
    print(summary)

    # ✅ Save Summary Statistics to Excel
    app = xw.App(visible=False)  # Run Excel in background
    wb = app.books.open(file_path)
    sheet = wb.sheets[0]

    sheet.range("K1").value = "Summary Statistics"
    sheet.range("K2").value = summary

    wb.save()
    wb.close()
    app.quit()

    print("✅ Summary Statistics Saved in Excel.")

except Exception as e:
    print(f"🚨 Excel Error: {e}")

# ✅ Create Dashboard-Style Visualizations
try:
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))  # 2x2 Grid

    if not numeric_cols.empty:
        # 🔹 Histogram for First Numeric Column
        sns.histplot(df[numeric_cols[0]], bins=20, kde=True, color="blue", ax=axes[0, 0])
        axes[0, 0].set_title(f"Distribution of {numeric_cols[0]}")

    if "Sales_Volume" in df.columns:
        # 🔹 Boxplot for Sales Volume
        sns.boxplot(x=df["Sales_Volume"], color="green", ax=axes[0, 1])
        axes[0, 1].set_title("Sales Volume Boxplot")

    if not date_cols.empty:
        # 🔹 Line Chart for Stock Over Time
        df[date_cols[0]] = pd.to_datetime(df[date_cols[0]])  # Convert to datetime
        df.sort_values(date_cols[0], inplace=True)  # Ensure time order
        sns.lineplot(x=df[date_cols[0]], y=df[numeric_cols[1]], color="red", ax=axes[1, 0])
        axes[1, 0].set_title(f"{numeric_cols[1]} Over Time")

    if len(numeric_cols) > 2:
        # 🔹 Histogram for Another Numeric Column
        sns.histplot(df[numeric_cols[2]], bins=15, color="purple", kde=True, ax=axes[1, 1])
        axes[1, 1].set_title(f"Distribution of {numeric_cols[2]}")

    plt.tight_layout()
    plt.savefig("dashboard_visualizations.png")  # Save the figure
    plt.show()

    print("✅ Dashboard Visualizations Completed.")

except Exception as e:
    print(f"🚨 Visualization Error: {e}")

