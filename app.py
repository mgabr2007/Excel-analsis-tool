import streamlit as st
import pandas as pd
import tempfile
import os
import plotly.express as px
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message=s')

# Utility Functions
def handle_file_upload(upload_type, file_types):
    uploaded_file = st.file_uploader(f"Choose a {upload_type} file", type=file_types, key=upload_type)
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{file_types[0]}') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        logging.info(f"File uploaded: {uploaded_file.name}")
        return tmp_file_path, uploaded_file.name
    return None, None

def read_excel(file):
    try:
        logging.info("Reading Excel file...")
        df = pd.read_excel(file, engine='openpyxl')
        logging.info("Excel file read successfully!")
        return df
    except Exception as e:
        error_message = f"Failed to read Excel file: {e}"
        logging.error(error_message)
        st.error(error_message)
        return pd.DataFrame()

# Visualization Functions
def visualize_data(df, x_column, y_column, z_column, chart_type):
    if chart_type == "3D Scatter Plot" and z_column:
        fig = px.scatter_3d(df, x=x_column, y=y_column, z=z_column, title=f"3D Scatter Plot of {x_column} vs {y_column} vs {z_column}")
    elif chart_type == "3D Line Chart" and z_column:
        fig = px.line_3d(df, x=x_column, y=y_column, z=z_column, title=f"3D Line Chart of {x_column} vs {y_column} vs {z_column}")
    elif chart_type == "3D Surface Plot" and z_column:
        fig = px.surface(df, x=x_column, y=y_column, z=z_column, title=f"3D Surface Plot of {x_column} vs {y_column} vs {z_column}")
    elif chart_type == "Scatter Plot":
        fig = px.scatter(df, x=x_column, y=y_column, title=f"Scatter Plot of {x_column} vs {y_column}")
    elif chart_type == "Line Chart":
        fig = px.line(df, x=x_column, y=y_column, title=f"Line Chart of {x_column} vs {y_column}")
    elif chart_type == "Bar Chart":
        fig = px.bar(df, x=x_column, y=y_column, title=f"Bar Chart of {x_column} vs {y_column}")
    elif chart_type == "Histogram":
        fig = px.histogram(df, x=x_column, title=f"Histogram of {x_column}")
    elif chart_type == "Box Plot":
        fig = px.box(df, y=y_column, x=x_column, title=f"Box Plot of {x_column} vs {y_column}")
    else:
        st.error("Unsupported chart type selected or missing Z column for 3D chart.")
        return None
    
    fig.update_layout(paper_bgcolor='white', plot_bgcolor='white', font_color='black')
    st.plotly_chart(fig)

def generate_insights(df):
    if not df.empty:
        st.write("Descriptive Statistics:", df.describe())
        numeric_df = df.select_dtypes(include=['number'])
        if not numeric_df.empty:
            st.write("Correlation Matrix:")
            corr_matrix = numeric_df.corr()
            st.dataframe(corr_matrix)
        else:
            st.write("No numeric columns available for correlation analysis.")
    else:
        st.write("No data available to generate insights.")

# Excel File Analysis Function
def excel_file_analysis():
    st.write("""
    ### Instructions for Analyzing Excel Files:

    1. **Upload an Excel File:** Click on the "Choose an Excel file" button to upload an Excel spreadsheet.

    2. **Select Columns for Analysis:** Choose the columns you want to use for analysis from the uploaded Excel file.

    3. **Select Chart Type:** Choose the type of chart to visualize the relationship between the selected columns.

    4. **Visualize Data:** Click on "Visualize Data" to generate the chart for the selected columns and chart type.

    5. **Generate Insights:** Click on "Generate Insights" to view descriptive statistics and other insights from the data.
    """)

    file_path, file_name = handle_file_upload("Excel", ['xlsx'])
    if file_path:
        st.write(f"File uploaded: {file_name}")
        df = read_excel(file_path)
        if not df.empty:
            st.write("File read successfully! Here is a preview of the data:")
            st.dataframe(df.head())
            
            columns = df.columns.tolist()
            x_column = st.selectbox("Select X-axis column", columns)
            y_column = st.selectbox("Select Y-axis column", columns)
            z_column = st.selectbox("Select Z-axis column (optional, for 3D charts)", [None] + columns)
            chart_type = st.selectbox("Select chart type", ["Scatter Plot", "Line Chart", "Bar Chart", "Histogram", "Box Plot", "3D Scatter Plot", "3D Line Chart", "3D Surface Plot"])
            
            if x_column and y_column and chart_type:
                if st.button("Visualize Data"):
                    st.write(f"Visualizing {chart_type} for {x_column} vs {y_column}" + (f" vs {z_column}" if z_column else "") + "...")
                    visualize_data(df, x_column, y_column, z_column, chart_type)
                if st.button("Generate Insights"):
                    st.write("Generating insights...")
                    generate_insights(df)
            else:
                st.warning("Please select columns and chart type for visualization.")
        else:
            st.error("The uploaded file is empty or could not be read.")
        os.remove(file_path)
    else:
        st.info("Please upload an Excel file to proceed.")

# Main Function
def main():
    st.title("Excel File Analysis Tool")
    excel_file_analysis()

if __name__ == "__main__":
    main()
