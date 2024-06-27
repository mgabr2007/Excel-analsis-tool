import streamlit as st
import pandas as pd
import tempfile
import os
import logging
import pygwalker as pyg
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.tree import DecisionTreeRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.metrics import mean_squared_error, r2_score

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
        "instruction_6": "6. **Train a Machine Learning Model**: Select features and a target column to train a simple linear regression or decision tree regression model.",
        "ml_instruction": """
        ### What does the Machine Learning model do?

        The machine learning model implemented in this tool is a simple linear regression or decision tree regression model. Here’s what it does:

        1. **Feature Selection**: Choose one or more columns from your dataset to use as features (independent variables) for the model.
        2. **Target Selection**: Choose one column from your dataset to use as the target (dependent variable) for the model.
        3. **Train the Model**: The tool splits the data into training and testing sets, trains a linear regression or decision tree regression model on the training data, and evaluates it on the testing data.
        4. **Model Performance**: The tool provides the Mean Squared Error (MSE) and the R² Score to evaluate the model's performance.
        """,
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
        "no_data_available": "No data available to generate insights.",
        "sidebar_instructions": "### Instructions:\n1. Select your preferred language.\n2. Follow the instructions on the main page to upload and analyze your Excel file.",
        "ml_section_title": "Train a Machine Learning Model",
        "ml_select_features": "Select feature columns",
        "ml_select_target": "Select target column",
        "ml_train_button": "Train Model",
        "ml_model_performance": "Model Performance",
        "ml_mse": "Mean Squared Error",
        "ml_r2": "R² Score",
        "ml_model_choice": "Choose a Machine Learning Model",
        "ml_performance_explanation": """
        ### Model Performance Explanation

        **Mean Squared Error (MSE)**: This is the average of the squared differences between the actual and predicted values. A lower MSE indicates a better fit.

        **R² Score**: This score represents the proportion of the variance in the dependent variable that is predictable from the independent variables. An R² score close to 1 indicates a good fit.
        """
    },
    "ar": {
        "title": "أداة تحليل ملفات Excel",
        "instructions_title": "إرشادات لتحليل ملفات Excel:",
        "instruction_1": "1. **تحميل ملف Excel**: انقر فوق الزر \"اختر ملف Excel\" لتحميل جدول بيانات Excel بتنسيق `.xlsx`.",
        "instruction_2": "2. **معاينة البيانات**: بعد التحميل، سيتم عرض معاينة لأول بضعة صفوف من الملف. يساعدك هذا في تأكيد تحميل الملف الصحيح.",
        "instruction_3": "3. **اختر الأعمدة للتحليل**: اختر الأعمدة التي تريد استخدامها للتحليل من ملف Excel الذي تم تحميله. استخدم القائمة المنسدلة المتعددة لتحديد أعمدة متعددة.",
        "instruction_4": "4. **توليد الإحصاءات**: انقر فوق الزر \"توليد الإحصاءات\" لعرض الإحصاءات الوصفية والرؤى الأخرى من البيانات. يتضمن ذلك الإحصاءات الأساسية ومصفوفة الارتباط للأعمدة الرقمية.",
        "instruction_5": "5. **تصور البيانات**: أسفل الإحصاءات، استخدم Pygwalker لإنشاء تصورات تفاعلية. هذه التصورات قابلة للتخصيص بدرجة كبيرة وتتيح لك استكشاف البيانات بعمق.",
        "instruction_6": "6. **تدريب نموذج التعلم الآلي**: اختر الميزات وعمود الهدف لتدريب نموذج الانحدار الخطي البسيط أو نموذج شجرة القرار.",
        "ml_instruction": """
        ### ماذا يفعل نموذج التعلم الآلي؟

        النموذج المطبق في هذه الأداة هو نموذج انحدار خطي بسيط أو نموذج شجرة القرار. إليك ما يفعله:

        1. **اختيار الميزات**: اختر عمودًا أو أكثر من بياناتك لاستخدامها كميزات (متغيرات مستقلة) للنموذج.
        2. **اختيار الهدف**: اختر عمودًا واحدًا من بياناتك لاستخدامه كهدف (متغير تابع) للنموذج.
        3. **تدريب النموذج**: تقوم الأداة بتقسيم البيانات إلى مجموعات تدريب واختبار، وتدريب نموذج الانحدار الخطي أو نموذج شجرة القرار على بيانات التدريب، وتقييمه على بيانات الاختبار.
        4. **أداء النموذج**: توفر الأداة متوسط ​​الخطأ التربيعي (MSE) ودرجة R² لتقييم أداء النموذج.
        """,
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
        "no_data_available": "لا توجد بيانات متاحة لتوليد الإحصاءات.",
        "sidebar_instructions": "### تعليمات:\n1. اختر لغتك المفضلة.\n2. اتبع التعليمات في الصفحة الرئيسية لتحميل وتحليل ملف Excel الخاص بك.",
        "ml_section_title": "تدريب نموذج التعلم الآلي",
        "ml_select_features": "اختر أعمدة الميزات",
        "ml_select_target": "اختر عمود الهدف",
        "ml_train_button": "تدريب النموذج",
        "ml_model_performance": "أداء النموذج",
        "ml_mse": "متوسط ​​الخطأ التربيعي",
        "ml_r2": "درجة R²",
        "ml_model_choice": "اختر نموذج التعلم الآلي",
        "ml_performance_explanation": """
        ### شرح أداء النموذج

        **متوسط ​​الخطأ التربيعي (MSE)**: هذا هو متوسط ​​الفروق المربعة بين القيم الفعلية والقيم المتوقعة. يشير انخفاض MSE إلى مطابقة أفضل.

        **درجة R²**: يمثل هذا الدرجة نسبة التباين في المتغير التابع التي يمكن التنبؤ بها من المتغيرات المستقلة. يشير اقتراب درجة R² من 1 إلى مطابقة جيدة.
        """
    },
    "fr": {
        "title": "Outil d'Analyse de Fichier Excel",
        "instructions_title": "Instructions pour Analyser les Fichiers Excel:",
        "instruction_1": "1. **Téléchargez un Fichier Excel**: Cliquez sur le bouton \"Choisir un fichier Excel\" pour télécharger une feuille de calcul Excel au format `.xlsx`.",
        "instruction_2": "2. **Aperçu des Données**: Après le téléchargement, un aperçu des premières lignes du fichier sera affiché. Cela vous aide à confirmer que le bon fichier a été téléchargé.",
        "instruction_3": "3. **Sélectionner les Colonnes pour l'Analyse**: Choisissez les colonnes que vous souhaitez utiliser pour l'analyse à partir du fichier Excel téléchargé. Utilisez la liste déroulante multisélection pour sélectionner plusieurs colonnes.",
        "instruction_4": "4. **Générer des Informations**: Cliquez sur le bouton \"Générer des Informations\" pour afficher les statistiques descriptives et autres informations sur les données. Cela inclut les statistiques de base et une matrice de corrélation pour les colonnes numériques.",
        "instruction_5": "5. **Visualiser les Données**: Sous les informations, utilisez Pygwalker pour créer des visualisations interactives. Ces visualisations sont hautement personnalisables et vous permettent d'explorer les données en profondeur.",
        "instruction_6": "6. **Former un Modèle de Machine Learning**: Sélectionnez les caractéristiques et une colonne cible pour former un modèle de régression linéaire simple ou un modèle d'arbre de décision.",
        "ml_instruction": """
        ### Que fait le modèle de Machine Learning ?

        Le modèle de machine learning implémenté dans cet outil est un simple modèle de régression linéaire ou un modèle d'arbre de décision. Voici ce qu'il fait :

        1. **Sélection des caractéristiques**: Choisissez une ou plusieurs colonnes de votre jeu de données à utiliser comme caractéristiques (variables indépendantes) pour le modèle.
        2. **Sélection de la cible**: Choisissez une colonne de votre jeu de données à utiliser comme cible (variable dépendante) pour le modèle.
        3. **Entraîner le modèle**: L'outil divise les données en ensembles d'entraînement et de test, entraîne un modèle de régression linéaire ou un modèle d'arbre de décision sur les données d'entraînement et l'évalue sur les données de test.
        4. **Performance du modèle**: L'outil fournit l'erreur quadratique moyenne (MSE) et le score R² pour évaluer les performances du modèle.
        """,
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
        "no_data_available": "Aucune donnée disponible pour générer des informations.",
        "sidebar_instructions": "### Instructions:\n1. Sélectionnez votre langue préférée.\n2. Suivez les instructions sur la page principale pour télécharger et analyser votre fichier Excel.",
        "ml_section_title": "Former un Modèle de Machine Learning",
        "ml_select_features": "Sélectionnez les colonnes de caractéristiques",
        "ml_select_target": "Sélectionnez la colonne cible",
        "ml_train_button": "Former le Modèle",
        "ml_model_performance": "Performance du Modèle",
        "ml_mse": "Erreur Quadratique Moyenne",
        "ml_r2": "Score R²",
        "ml_model_choice": "Choisissez un modèle de Machine Learning",
        "ml_performance_explanation": """
        ### Explication de la Performance du Modèle

        **Erreur Quadratique Moyenne (MSE)**: Il s'agit de la moyenne des différences quadratiques entre les valeurs réelles et prévues. Une MSE plus faible indique un meilleur ajustement.

        **Score R²**: Ce score représente la proportion de la variance dans la variable dépendante qui est prévisible à partir des variables indépendantes. Un score R² proche de 1 indique un bon ajustement.
        """
    },
    "de": {
        "title": "Excel-Dateianalysetool",
        "instructions_title": "Anleitung zur Analyse von Excel-Dateien:",
        "instruction_1": "1. **Laden Sie eine Excel-Datei hoch**: Klicken Sie auf die Schaltfläche \"Wählen Sie eine Excel-Datei aus\", um eine Excel-Tabelle im `.xlsx`-Format hochzuladen.",
        "instruction_2": "2. **Datenvorschau**: Nach dem Hochladen wird eine Vorschau der ersten Zeilen der Datei angezeigt. Dies hilft Ihnen zu bestätigen, dass die richtige Datei hochgeladen wurde.",
        "instruction_3": "3. **Wählen Sie Spalten zur Analyse aus**: Wählen Sie die Spalten aus, die Sie aus der hochgeladenen Excel-Datei zur Analyse verwenden möchten. Verwenden Sie das Dropdown-Menü zur Mehrfachauswahl, um mehrere Spalten auszuwählen.",
        "instruction_4": "4. **Erzeugen Sie Erkenntnisse**: Klicken Sie auf die Schaltfläche \"Erkenntnisse generieren\", um beschreibende Statistiken und andere Erkenntnisse aus den Daten anzuzeigen. Dies umfasst grundlegende Statistiken und eine Korrelationsmatrix für numerische Spalten.",
        "instruction_5": "5. **Daten visualisieren**: Unterhalb der Erkenntnisse verwenden Sie Pygwalker, um interaktive Visualisierungen zu erstellen. Diese Visualisierungen sind hochgradig anpassbar und ermöglichen es Ihnen, die Daten im Detail zu erkunden.",
        "instruction_6": "6. **Trainieren Sie ein Machine Learning Modell**: Wählen Sie Funktionen und eine Zielspalte, um ein einfaches lineares Regressionsmodell oder ein Entscheidungsbaum-Regressionsmodell zu trainieren.",
        "ml_instruction": """
        ### Was macht das Machine Learning Modell?

        Das Machine Learning Modell, das in diesem Tool implementiert ist, ist ein einfaches lineares Regressionsmodell oder ein Entscheidungsbaum-Regressionsmodell. Hier ist, was es tut:

        1. **Merkmalsauswahl**: Wählen Sie eine oder mehrere Spalten aus Ihrem Datensatz aus, die als Merkmale (unabhängige Variablen) für das Modell verwendet werden sollen.
        2. **Zielauswahl**: Wählen Sie eine Spalte aus Ihrem Datensatz aus, die als Ziel (abhängige Variable) für das Modell verwendet werden soll.
        3. **Modell trainieren**: Das Tool teilt die Daten in Trainings- und Testmengen auf, trainiert ein lineares Regressionsmodell oder ein Entscheidungsbaum-Regressionsmodell mit den Trainingsdaten und bewertet es mit den Testdaten.
        4. **Modellleistung**: Das Tool liefert den mittleren quadratischen Fehler (MSE) und den R²-Score zur Bewertung der Modellleistung.
        """,
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
        "no_data_available": "Keine Daten verfügbar, um Erkenntnisse zu generieren.",
        "sidebar_instructions": "### Anweisungen:\n1. Wählen Sie Ihre bevorzugte Sprache.\n2. Befolgen Sie die Anweisungen auf der Hauptseite, um Ihre Excel-Datei hochzuladen und zu analysieren.",
        "ml_section_title": "Trainieren Sie ein Machine Learning Modell",
        "ml_select_features": "Wählen Sie Feature-Spalten",
        "ml_select_target": "Wählen Sie die Zielspalte",
        "ml_train_button": "Modell trainieren",
        "ml_model_performance": "Modellleistung",
        "ml_mse": "Mittlerer quadratischer Fehler",
        "ml_r2": "R²-Score",
        "ml_model_choice": "Wählen Sie ein Machine Learning Modell",
        "ml_performance_explanation": """
        ### Erklärung der Modellleistung

        **Mittlerer quadratischer Fehler (MSE)**: Dies ist der Durchschnitt der quadrierten Unterschiede zwischen den tatsächlichen und vorhergesagten Werten. Ein niedrigerer MSE weist auf eine bessere Übereinstimmung hin.

        **R²-Score**: Dieser Score gibt den Anteil der Varianz in der abhängigen Variable an, der durch die unabhängigen Variablen vorhergesagt werden kann. Ein R²-Score nahe 1 weist auf eine gute Übereinstimmung hin.
        """
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

# Machine Learning Model Training Function
def train_ml_model(df, language):
    st.write(f"### {translate_text(language, 'ml_section_title')}")
    
    model_choice = st.selectbox(translate_text(language, "ml_model_choice"), ["Linear Regression", "Decision Tree Regressor"])

    columns = df.columns.tolist()
    feature_columns = st.multiselect(translate_text(language, "ml_select_features"), columns)
    target_column = st.selectbox(translate_text(language, "ml_select_target"), columns)
    
    if feature_columns and target_column:
        X = df[feature_columns]
        y = df[target_column]
        
        if not pd.api.types.is_numeric_dtype(y):
            st.error("Target column must be numeric.")
            return

        for col in feature_columns:
            if not pd.api.types.is_numeric_dtype(df[col]):
                st.error(f"Feature column '{col}' must be numeric.")
                return

        # Preprocessing pipeline
        preprocessor = ColumnTransformer(
            transformers=[
                ('num', StandardScaler(), feature_columns)
            ])

        if model_choice == "Linear Regression":
            model = Pipeline(steps=[('preprocessor', preprocessor),
                                    ('regressor', LinearRegression())])
        elif model_choice == "Decision Tree Regressor":
            model = Pipeline(steps=[('preprocessor', preprocessor),
                                    ('regressor', DecisionTreeRegressor(random_state=42))])

        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)

        mse = mean_squared_error(y_test, y_pred)
        r2 = r2_score(y_test, y_pred)
        
        st.write(f"### {translate_text(language, 'ml_model_performance')}")
        st.write(f"{translate_text(language, 'ml_mse')}: {mse}")
        st.write(f"{translate_text(language, 'ml_r2')}: {r2}")

        st.write(translate_text(language, "ml_performance_explanation"))

# Excel File Analysis Function
def excel_file_analysis(language):
    st.write(f"""
    ### {translate_text(language, "instructions_title")}

    {translate_text(language, "instruction_1")}
    {translate_text(language, "instruction_2")}
    {translate_text(language, "instruction_3")}
    {translate_text(language, "instruction_4")}
    {translate_text(language, "instruction_5")}
    {translate_text(language, "instruction_6")}
    """)

    st.write(translate_text(language, "ml_instruction"))

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
                
                # Train Machine Learning Model
                train_ml_model(df_selected, language)
            else:
                st.warning(translate_text(language, "select_columns_warning"))
        else:
            st.error(translate_text(language, "file_empty_error"))
        os.remove(file_path)
    else:
        st.info(translate_text(language, "upload_prompt"))

# Main Function
def main():
    # Language selection with flags
    language = st.sidebar.radio(
        "🌐 Select Language",
        options=["en", "ar", "fr", "de"],
        format_func=lambda lang: {
            "en": "English 🇺🇸",
            "ar": "Arabic 🇸🇦",
            "fr": "French 🇫🇷",
            "de": "German 🇩🇪"
        }[lang]
    )

    if language == "ar":
        # Inject CSS for RTL layout
        st.markdown(
            """
            <style>
            .css-1outpf7 {
                direction: rtl;
            }
            .css-1v3fvcr {
                direction: rtl;
            }
            </style>
            """,
            unsafe_allow_html=True
        )

    # Sidebar instructions
    st.sidebar.markdown(translate_text(language, "sidebar_instructions"))

    st.title(translate_text(language, "title"))
    excel_file_analysis(language)

if __name__ == "__main__":
    main()
