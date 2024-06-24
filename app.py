import streamlit as st
import pandas as pd
import tempfile
import os
import logging
import pygwalker as pyg

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Utility Functions
def handle_file_upload(upload_type, file_types):import streamlit as st
import pandas as pd
import tempfile
import os
import logging
import pygwalker as pyg

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

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

    3. **Generate Insights:** Click on "Generate Insights" to view descriptive statistics and other insights from the data.

    4. **Visualize Data:** Use Pygwalker below to create interactive visualizations.
    """)

    file_path, file_name = handle_file_upload("Excel", ['xlsx'])
    if file_path:
        st.write(f"File uploaded: {file_name}")
        df = read_excel(file_path)
        if not df.empty:
            st.write("File read successfully! Here is a preview of the data:")
            st.dataframe(df.head())

            columns = df.columns.tolist()
            selected_columns = st.multiselect("Select columns for analysis", columns, default=columns)
            
            if selected_columns:
                df_selected = df[selected_columns]
                if st.button("Generate Insights"):
                    st.write("Generating insights...")
                    generate_insights(df_selected)

                st.write("### Interactive Visualization")
                # Initialize Pygwalker interface and render as HTML in Streamlit
                walker_html = pyg.walk(df_selected)
                st.components.v1.html(walker_html.to_html(), height=800, scrolling=True)
            else:
                st.warning("Please select columns for analysis.")
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

    3. **Generate Insights:** Click on "Generate Insights" to view descriptive statistics and other insights from the data.

    4. **Visualize Data:** Use Pygwalker below to create interactive visualizations.
    """)

    file_path, file_name = handle_file_upload("Excel", ['xlsx'])
    if file_path:
        st.write(f"File uploaded: {file_name}")
        df = read_excel(file_path)
        if not df.empty:
            st.write("File read successfully! Here is a preview of the data:")
            st.dataframe(df.head())

            columns = df.columns.tolist()
            selected_columns = st.multiselect("Select columns for analysis", columns, default=columns)
            
            if selected_columns:
                df_selected = df[selected_columns]
                if st.button("Generate Insights"):
                    st.write("Generating insights...")
                    generate_insights(df_selected)

                st.write("### Interactive Visualization")
                # Initialize Pygwalker interface and render as HTML in Streamlit
                walker_html = pyg.walk(df_selected)
                st.components.v1.html(walker_html.to_html(), height=600, scrolling=True)
            else:
                st.warning("Please select columns for analysis.")
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
