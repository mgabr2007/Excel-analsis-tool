import streamlit as st
import pandas as pd
import tempfile
import os
import logging
import pygwalker as pyg

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Set Streamlit page configuration to wide layout
st.set_page_config(layout="wide")

# Translations dictionary
translations = {
    "en": {
        "title": "Excel File Analysis Tool",
        "instructions_title": "Instructions for Analyzing Excel Files:",
        "instruction_1": "1. **Upload an Excel File**: Click on the \"Choose an Excel file\" button to upload an Excel spreadsheet in `.xlsx` format.",
        "instruction_2": "2. **Preview Data**: After uploading, a preview of the first few rows of the file will be displayed. This helps you confirm that the correct file has been uploaded.",
        "instruction_3": "3. **Select Columns for Analysis**: Choose the columns you want to use for analysis from the uploaded Excel file. Use the multiselect dropdown to select multiple columns.",
        "instruction_4": "4. **Generate Insights**: Click on the \"Generate Insights\" button to view descriptive statistics and other insights from the data. This includes basic statistics and a correlation matrix for numeric columns.",
        "instruction_5": "5. **Visualize Data**: Below the insights, use Pygwalker to create interactive visualizations. These visualizations are highly customizable and allow you to explore the data in depth.",
        "choose_file": "Choose a file",
        "file_uploaded": "File uploaded:",
        "file_read_success": "File read successfully! Here is a preview of the data:",
        "select_columns": "Select columns for analysis",
        "generate_insights": "Generate Insights",
        "interactive_visualization": "Interactive Visualization",
        "select_columns_warning": "Please select columns for analysis.",
        "file_empty_error": "The uploaded file is empty or could not be read.",
        "upload_prompt": "Please upload an Excel file to proceed.",
        "descriptive_statistics": "Descriptive Statistics:",
        "correlation_matrix": "Correlation Matrix:",
        "no_numeric_columns": "No numeric columns available for correlation analysis.",
        "no_data_available": "No data available to generate insights."
    },
    "ar": {
        "title": "أداة تحليل ملفات Excel",
        "instructions_title": "إرشادات لتحليل ملفات Excel:",
        "instruction_1": "1. **تحميل ملف Excel**: انقر فوق الزر \"اختر ملف Excel\" لتحميل جدول بيانات Excel بتنسيق `.xlsx`.",
        "instruction_2": "2. **معاينة البيانات**: بعد التحميل، سيتم عرض معاينة لأول بضعة صفوف من الملف. يساعدك هذا في تأكيد تحميل الملف الصحيح.",
        "instruction_3": "3. **اختر الأعمدة للتحليل**: اختر الأعمدة التي تريد استخدامها للتحليل من ملف Excel الذي تم تحميله. استخدم القائمة المنسدلة المتعددة لتحديد أعمدة متعددة.",
        "instruction_4": "4. **توليد الإحصاءات**: انقر فوق الزر \"توليد الإحصاءات\" لعرض الإحصاءات الوصفية والرؤى الأخرى من البيانات. يتضمن ذلك الإحصاءات الأساسية ومصفوفة الارتباط للأعمدة الرقمية.",
        "instruction_5": "5. **تصور البيانات**: أسفل الإحصاءات، استخدم Pygwalker لإنشاء تصورات تفاعلية. هذه التصورات قابلة للتخصيص بدرجة كبيرة وتتيح لك استكشاف البيانات بعمق.",
        "choose_file": "اختر ملفًا",
        "file_uploaded": "تم تحميل الملف:",
        "file_read_success": "تم قراءة الملف بنجاح! فيما يلي معاينة للبيانات:",
        "select_columns": "اختر الأعمدة للتحليل",
        "generate_insights": "توليد الإحصاءات",
        "interactive_visualization": "التصور التفاعلي",
        "select_columns_warning": "يرجى اختيار الأعمدة للتحليل.",
        "file_empty_error": "الملف الذي تم تحميله فارغ أو لا يمكن قراءته.",
        "upload_prompt": "يرجى تحميل ملف Excel للمتابعة.",
        "descriptive_statistics": "الإحصاءات الوصفية:",
        "correlation_matrix": "مصفوفة الارتباط:",
        "no_numeric_columns": "لا توجد أعمدة رقمية متاحة لتحليل الارتباط.",
        "no_data_available": "لا توجد بيانات متاحة لتوليد الإحصاءات."
    },
    "fr": {
        "title": "Outil d'Analyse de Fichier Excel",
        "instructions_title": "Instructions pour Analyser les Fichiers Excel:",
        "instruction_1": "1. **Téléchargez un Fichier Excel**: Cliquez sur le bouton \"Choisir un fichier Excel\" pour télécharger une feuille de calcul Excel au format `.xlsx`.",
        "instruction_2": "2. **Aperçu des Données**: Après le téléchargement, un aperçu des premières lignes du fichier sera affiché. Cela vous aide à confirmer que le bon fichier a été téléchargé.",
        "instruction_3": "3. **Sélectionner les Colonnes pour l'Analyse**: Choisissez les colonnes que vous souhaitez utiliser pour l'analyse à partir du fichier Excel téléchargé. Utilisez la liste déroulante multisélection pour sélectionner plusieurs colonnes.",
        "instruction_4": "4. **Générer des Informations**: Cliquez sur le bouton \"Générer des Informations\" pour afficher les statistiques descriptives et autres informations sur les données. Cela inclut les statistiques de base et une matrice de corrélation pour les colonnes numériques.",
        "instruction_5": "5. **Visualiser les Données**: Sous les informations, utilisez Pygwalker pour créer des visualisations interactives. Ces visualisations sont hautement personnalisables et vous permettent d'explorer les données en profondeur.",
        "choose_file": "Choisissez un fichier",
        "file_uploaded": "Fichier téléchargé:",
        "file_read_success": "Fichier lu avec succès! Voici un aperçu des données:",
        "select_columns": "Sélectionnez les colonnes pour l'analyse",
        "generate_insights": "Générer des Informations",
        "interactive_visualization": "Visualisation Interactive",
        "select_columns_warning": "Veuillez sélectionner les colonnes pour l'analyse.",
        "file_empty_error": "Le fichier téléchargé est vide ou ne peut pas être lu.",
        "upload_prompt": "Veuillez télécharger un fichier Excel pour continuer.",
        "descriptive_statistics": "Statistiques Descriptives:",
        "correlation_matrix": "Matrice de Corrélation:",
        "no_numeric_columns": "Aucune colonne numérique disponible pour l'analyse de corrélation.",
        "no_data_available": "Aucune donnée disponible pour générer des informations."
    },
    "de": {
        "title": "Excel-Dateianalysetool",
        "instructions_title": "Anleitung zur Analyse von Excel-Dateien:",
        "instruction_1": "1. **Laden Sie eine Excel-Datei hoch**: Klicken Sie auf die Schaltfläche \"Wählen Sie eine Excel-Datei aus\", um eine Excel-Tabelle im `.xlsx`-Format hochzuladen.",
        "instruction_2": "2. **Datenvorschau**: Nach dem Hochladen wird eine Vorschau der ersten Zeilen der Datei angezeigt. Dies hilft Ihnen zu bestätigen, dass die richtige Datei hochgeladen wurde.",
        "instruction_3": "3. **Wählen Sie Spalten zur Analyse aus**: Wählen Sie die Spalten aus, die Sie aus der hochgeladenen Excel-Datei zur Analyse verwenden möchten. Verwenden Sie das Dropdown-Menü zur Mehrfachauswahl, um mehrere Spalten auszuwählen.",
        "instruction_4": "4. **Erzeugen Sie Erkenntnisse**: Klicken Sie auf die Schaltfläche \"Erkenntnisse generieren\", um beschreibende Statistiken und andere Erkenntnisse aus den Daten anzuzeigen. Dies umfasst grundlegende Statistiken und eine Korrelationsmatrix für numerische Spalten.",
        "instruction_5": "5. **Daten visualisieren**: Unterhalb der Erkenntnisse verwenden Sie Pygwalker, um interaktive Visualisierungen zu erstellen. Diese Visualisierungen sind hochgradig anpassbar und ermöglichen es Ihnen, die Daten im Detail zu erkunden.",
        "choose_file": "Wählen Sie eine Datei",
        "file_uploaded": "Datei hochgeladen:",
        "file_read_success": "Datei erfolgreich gelesen! Hier ist eine Vorschau der Daten:",
        "select_columns": "Wählen Sie Spalten zur Analyse aus",
        "generate_insights": "Erkenntnisse generieren",
        "interactive_visualization": "Interaktive Visualisierung",
        "select_columns_warning": "Bitte wählen Sie Spalten zur Analyse aus.",
        "file_empty_error": "Die hochgeladene Datei ist leer oder konnte nicht gelesen werden.",
        "upload_prompt": "Bitte laden Sie eine Excel-Datei hoch, um fortzufahren.",
        "descriptive_statistics": "Beschreibende Statistiken:",
        "correlation_matrix": "Korrelationsmatrix:",
        "no_numeric_columns": "Keine numerischen Spalten zur Korrelationsanalyse verfügbar.",
        "no_data_available": "Keine Daten verfügbar, um Erkenntnisse zu generieren."
    }
}

# Utility Functions
def translate_text(language, key):
    return translations[language].get(key, key)

def handle_file_upload(upload_type, file_types, language):
    uploaded_file = st.file_uploader(translate_text(language, "choose_file"), type=file_types, key=upload_type)
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{file_types[0]}') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        logging.info(f"File uploaded: {uploaded_file.name}")
        return tmp_file_path, uploaded_file.name
    return None, None

def read_excel(file, language):
    try:
        logging.info("Reading Excel file...")
        df = pd.read_excel(file, engine='openpyxl')
        logging.info("Excel file read successfully!")
        return df
    except Exception as e:
        error_message = translate_text(language, "file_empty_error") + f": {e}"
        logging.error(error_message)
        st.error(error_message)
        return pd.DataFrame()

def generate_insights(df, language):
    if not df.empty:
        st.write(translate_text(language, "descriptive_statistics"), df.describe())
        numeric_df = df.select_dtypes(include=['number'])
        if not numeric_df.empty:
            st.write(translate_text(language, "correlation_matrix"))
            corr_matrix = numeric_df.corr()
            st.dataframe(corr_matrix)
        else:
            st.write(translate_text(language, "no_numeric_columns"))
    else:
        st.write(translate_text(language, "no_data_available"))

# Excel File Analysis Function
def excel_file_analysis(language):
    st.write(f"""
    ### {translate_text(language, "instructions_title")}

    {translate_text(language, "instruction_1")}
    {translate_text(language, "instruction_2")}
    {translate_text(language, "instruction_3")}
    {translate_text(language, "instruction_4")}
    {translate_text(language, "instruction_5")}
    """)

    file_path, file_name = handle_file_upload("Excel", ['xlsx'], language)
    if file_path:
        st.write(f"### {translate_text(language, 'file_uploaded')} {file_name}")
        df = read_excel(file_path, language)
        if not df.empty:
            st.write(f"#### {translate_text(language, 'file_read_success')}")
            st.dataframe(df.head())

            columns = df.columns.tolist()
            selected_columns = st.multiselect(translate_text(language, "select_columns"), columns, default=columns)
            
            if selected_columns:
                df_selected = df[selected_columns]
                if st.button(translate_text(language, "generate_insights")):
                    st.write("Generating insights...")
                    generate_insights(df_selected, language)

                st.write(f"### {translate_text(language, 'interactive_visualization')}")
                # Initialize Pygwalker interface and render as HTML in Streamlit
                walker_html = pyg.walk(df_selected)
                st.components.v1.html(walker_html.to_html(), height=800, scrolling=True)
            else:
                st.warning(translate_text(language, "select_columns_warning"))
        else:
            st.error(translate_text(language, "file_empty_error"))
        os.remove(file_path)
    else:
        st.info(translate_text(language, "upload_prompt"))

# Main Function
def main():
    # Language selection
    language = st.sidebar.selectbox("Select Language", ["en", "ar", "fr", "de"])
    
    st.title(translate_text(language, "title"))
    excel_file_analysis(language)

if __name__ == "__main__":
    main()
