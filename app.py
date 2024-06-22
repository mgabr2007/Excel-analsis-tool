import streamlit as st
import pandas as pd
import tempfile
import os
import plotly.express as px
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Utility Functions
def handle_file_upload(upload_type, file_types):
    uploaded_file = st.file_uploader(f"Choose a {upload_type} file", type=file_types, key=upload_type)
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{file_types[0]}') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        return tmp_file_path, uploaded_file.name
    return None, None

def read_excel(file):
    try:
        return pd.read_excel(file, engine='openpyxl')
    except Exception as e:
        error_message = f"Failed to read Excel file: {e}"
        logging.error(error_message)
        st.error(error_message)
        return pd.DataFrame()

# Visualization Functions
def visualize_data(df, columns):
    figs = []
    for column in columns:
        if pd.api.types.is_numeric_dtype(df[column]):
            fig = px.histogram(df, x=column)
            fig.update_layout(paper_bgcolor='white', plot_bgcolor='white', font_color='black')
            st.plotly_chart(fig)
            figs.append(fig)
        else:
            fig = px.bar(df, x=column, title=f"Bar chart of {column}")
            fig.update_layout(paper_bgcolor='white', plot_bgcolor='white', font_color='black')
            st.plotly_chart(fig)
            figs.append(fig)
    return figs

def generate_insights(df):
    if not df.empty:
        st.write("Descriptive Statistics:", df.describe())

# Excel File Analysis Function
def excel_file_analysis():
    st.write("""
    ### Instructions for Analyzing Excel Files:

    1. **Upload an Excel File:** Click on the "Choose an Excel file" button to upload an Excel spreadsheet.

    2. **Select Columns to Display:** Choose the columns you want to display from the uploaded Excel file.

    3. **Visualize Data:** Click on "Visualize Data" to generate charts for the selected columns.

    4. **Generate Insights:** Click on "Generate Insights" to view descriptive statistics and other insights from the data.
    """)

    file_path, _ = handle_file_upload("Excel", ['xlsx'])
    if file_path:
        st.write(f"File uploaded: {file_path}")
        df = read_excel(file_path)
        if not df.empty:
            st.write("File read successfully!")
            selected_columns = st.multiselect("Select columns to display", df.columns.tolist(), default=df.columns.tolist(), key="columns")
            if selected_columns:
                st.dataframe(df[selected_columns])
                figs = []
                if st.button("Visualize Data", key="visualize"):
                    figs = visualize_data(df, selected_columns)
                if st.button("Generate Insights", key="insights"):
                    generate_insights(df)
            os.remove(file_path)

# Main Function
def main():
    st.title("Excel File Analysis Tool")
    st.sidebar.title("Navigation")
    if st.sidebar.button("Analyze Excel File"):
        excel_file_analysis()

if __name__ == "__main__":
    main()
