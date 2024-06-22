import pandas as pd
import streamlit as st
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
        # Placeholder for more sophisticated analysis or predictive modeling

# PDF Export Function
def export_analysis_to_pdf(ifc_metadata, component_count, figs, author, subject, cover_text):
    buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    doc = SimpleDocTemplate(buffer.name, pagesize=letter)
    styles = getSampleStyleSheet()
    flowables = []

    # Cover Page
    flowables.append(Spacer(1, 1 * inch))
    flowables.append(Paragraph(subject, styles['Title']))
    flowables.append(Spacer(1, 0.5 * inch))
    flowables.append(Paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}", styles['Normal']))
    flowables.append(Paragraph(f"Author: {author}", styles['Normal']))
    flowables.append(Spacer(1, 1 * inch))
    flowables.append(Paragraph(cover_text, styles['Normal']))
    flowables.append(Spacer(1, 2 * inch))

    # Adding Images
    for idx, fig in enumerate(figs):
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_file:
            try:
                fig.update_layout(paper_bgcolor='white', plot_bgcolor='white', font_color='black')
                fig.write_image(tmp_file.name, format='png', engine='kaleido')
                flowables.append(Spacer(1, 0.5 * inch))
                flowables.append(Paragraph(f"Chart {idx + 1}", styles['Heading2']))
                flowables.append(Image(tmp_file.name))
            except Exception as e:
                logging.error(f"Error exporting chart to image: {e}")
                st.error(f"Error exporting chart to image: {e}")

    doc.build(flowables)
    return buffer.name

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
        df = read_excel(file_path)
        if not df.empty:
            selected_columns = st.multiselect("Select columns to display", df.columns.tolist(), default=df.columns.tolist(), key="columns")
            if selected_columns:
                st.dataframe(df[selected_columns])
                figs = []
                if st.button("Visualize Data", key="visualize"):
                    figs = visualize_data(df, selected_columns)
                if st.button("Generate Insights", key="insights"):
                    generate_insights(df)
                if figs and st.button("Export Analysis as PDF"):
                    pdf_file_path = export_analysis_to_pdf({"Name": "Excel Data Analysis"}, {}, figs, "Author Name", "Excel Data Analysis Report", "This report contains the analysis of Excel data.")
                    with open(pdf_file_path, 'rb') as f:
                        st.download_button('Download PDF Report', f, 'excel_analysis.pdf')
            os.remove(file_path)
