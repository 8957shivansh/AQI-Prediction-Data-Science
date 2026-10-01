import os
try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
except ImportError:
    print("Please install python-docx first by running: pip install python-docx")
    exit()

def create_doc(filename, title, content_sections):
    doc = Document()
    
    # Title
    title_run = doc.add_heading(title, 0).runs[0]
    title_run.font.size = Pt(24)
    doc.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    doc.add_paragraph("Data Science Apprenticeship Internship - Yuva Intern\n").alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    for section_title, section_text in content_sections.items():
        doc.add_heading(section_title, level=1)
        for paragraph in section_text.split('\n\n'):
            if paragraph.strip():
                p = doc.add_paragraph(paragraph.strip())
                p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                
    doc.save(filename)
    print(f"Successfully created: {filename}")

# ==========================================
# WEEK 1 CONTENT
# ==========================================
week1_content = {
    "1. Project Summary": "This document outlines the foundational strategy for a data science project focused on predicting the Air Quality Index (AQI) using Machine Learning. The project will leverage publicly available environmental and meteorological data to forecast air quality, providing actionable insights for public health and environmental monitoring.",
    
    "2. Problem Statement": "Poor air quality is a significant global health hazard, contributing to respiratory diseases and environmental degradation. City administrations and citizens lack precise, localized forecasting of air quality. This project aims to address this gap by developing a robust predictive model for AQI based on historical pollution levels, weather conditions, and traffic patterns.",
    
    "3. Project Objectives": "- To collect and consolidate public air quality and weather data.\n- To identify key meteorological and environmental factors influencing AQI.\n- To build a machine learning model capable of predicting AQI for the next 24-48 hours.\n- To evaluate model performance using standard regression metrics.\n- To present findings through an interactive and visually appealing dashboard.",
    
    "4. Research Questions": "1. Which meteorological variables (temperature, humidity, wind speed) have the highest correlation with AQI levels?\n2. How do seasonal variations impact PM2.5 and PM10 concentrations?\n3. Can machine learning models outperform traditional statistical forecasting for AQI?\n4. What is the lag effect of weather changes on air pollution levels?",
    
    "5. Initial Hypotheses": "- Hypothesis 1: Wind speed is negatively correlated with AQI (higher wind disperses pollutants, lowering AQI).\n- Hypothesis 2: Temperature and humidity have a non-linear relationship with particulate matter concentrations.\n- Hypothesis 3: Ensemble models like Random Forest will yield higher predictive accuracy compared to simple Linear Regression.",
    
    "6. Scope of Analysis": "The analysis will focus on a specific metropolitan area using 3-5 years of historical daily data. It will cover data cleaning, exploratory data analysis, feature engineering, and the training of regression models. Deployment into a live application is outside the scope of this theoretical plan.",
    
    "7. Proposed Methodology (CRISP-DM)": "The project will follow the Cross-Industry Standard Process for Data Mining (CRISP-DM):\n1. Business Understanding: Defining objectives and success criteria.\n2. Data Understanding: Initial collection and exploration of datasets.\n3. Data Preparation: Cleaning, handling missing values, and formatting.\n4. Modeling: Applying ML algorithms.\n5. Evaluation: Testing models against benchmarks.\n6. Deployment: Structuring the final report and visualizations.",
    
    "8. Potential Challenges & Solutions": "Challenge: Missing or inconsistent data from sensor failures.\nSolution: Apply advanced imputation techniques (e.g., K-Nearest Neighbors imputation) and drop days with excessive missing data.\n\nChallenge: Non-stationary time series data.\nSolution: Incorporate time-based features (month, day of week) and lag variables.",
    
    "9. Tools and Technologies": "- Language: Python\n- Libraries: Pandas, NumPy (Data Manipulation), Scikit-Learn (Modeling), Matplotlib, Seaborn (Visualization)\n- Environment: Jupyter Notebook"
}

# ==========================================
# WEEK 2 CONTENT
# ==========================================
week2_content = {
    "1. Objective & Dataset Selection": "This document outlines the data preprocessing workflow and Exploratory Data Analysis (EDA) plan for the AQI Prediction Project. The hypothetical dataset selected is the publicly available 'Air Quality Data' from government meteorological portals or open-source datasets like OpenAQ. It includes features like PM2.5, PM10, NO2, CO, SO2, O3, Temperature, Humidity, and Wind Speed.",
    
    "2. Data Cleaning Strategy": "1. Handling Missing Values:\n- Numerical features with <5% missing data will be imputed using the median to avoid outlier skewness.\n- Time-series gaps will be handled using forward-fill (ffill) or linear interpolation to maintain temporal continuity.\n\n2. Data Type Conversion:\n- Ensure the 'Date' column is converted to a DateTime object and set as the dataframe index.\n- Convert categorical text data (e.g., 'City' or 'Season') into suitable categorical formats.",
    
    "3. Outlier Detection and Treatment": "Outliers in pollution data can represent true anomalies (like crop burning or fireworks) rather than errors.\n- Detection: We will use Boxplots and the Interquartile Range (IQR) method to identify statistical outliers in PM2.5 and PM10 columns.\n- Treatment: Extreme values resulting from sensor errors will be capped at the 99th percentile (Winsorization) rather than deleted, preserving the integrity of high-pollution events.",
    
    "4. Data Normalization and Feature Scaling": "Machine learning models like Support Vector Machines and Gradient Descent-based algorithms are sensitive to feature scales.\n- Technique: We will apply Standard Scaling (Z-score normalization) using Scikit-Learn's StandardScaler. This transforms data to have a mean of 0 and standard deviation of 1.\n- Application: Applied to all continuous independent variables (e.g., Temperature, Wind Speed, NO2 levels) before model training.",
    
    "5. Exploratory Data Analysis (EDA) Plan": "The EDA will systematically uncover patterns, trends, and relationships within the data.\n\n1. Univariate Analysis:\n- Histograms & KDE Plots: To understand the distribution of the target variable (AQI). Expected to be right-skewed.\n- Rationale: Helps decide if logarithmic transformation is needed for the target variable.\n\n2. Bivariate Analysis:\n- Scatter Plots: Plotting AQI against Temperature and Wind Speed.\n- Rationale: To visually confirm hypotheses (e.g., higher wind speed correlates with lower AQI).\n\n3. Multivariate Analysis & Correlation:\n- Heatmap: A Seaborn correlation matrix of all numerical features.\n- Rationale: To identify multicollinearity (e.g., PM2.5 and PM10 might be highly correlated) and select the most relevant features for modeling.\n\n4. Time-Series Analysis:\n- Line Charts with Rolling Averages: Plotting AQI over time (monthly and yearly).\n- Rationale: To detect seasonal trends (e.g., higher pollution in winter months).",
    
    "6. Summary": "This preprocessing and EDA phase ensures data quality and provides the statistical foundation necessary for informed feature engineering and robust model selection in the subsequent stages of the project."
}

# ==========================================
# WEEK 3 CONTENT
# ==========================================
week3_content = {
    "1. Objective": "This document outlines the strategic approach to feature engineering and model selection for the AQI prediction project. The goal is to transform raw environmental data into predictive signals and select appropriate machine learning algorithms.",
    
    "2. Feature Engineering Strategy": "To enhance the predictive power of our models, we will create new features derived from the existing dataset.\n\n1. Temporal Feature Extraction:\n- Extract 'Month', 'Day of Week', and 'Is_Weekend' from the DateTime index. \n- Rationale: AQI often shows strong weekly (lower on weekends due to less traffic) and seasonal patterns.\n\n2. Lag Features for Time Series:\n- Create 'AQI_Lag_1' (AQI of the previous day) and 'AQI_Lag_7' (AQI of the previous week).\n- Rationale: Air pollution is highly autocorrelated; yesterday's AQI is a strong predictor of today's AQI.\n\n3. Rolling Window Statistics:\n- Calculate 3-day and 7-day rolling averages of temperature and wind speed.\n- Rationale: Prolonged periods of stagnant air contribute more to high AQI than single-day anomalies.\n\n4. Handling Categorical Variables:\n- Apply One-Hot Encoding to the 'Season' variable (Winter, Summer, Monsoon, Post-Monsoon) to allow models to interpret seasonal effects mathematically.",
    
    "3. Dimensionality Reduction": "If the feature space becomes too large after engineering (e.g., including data from multiple neighboring cities), we will consider Principal Component Analysis (PCA).\n- Approach: Check explained variance ratio. If 90% of variance can be captured by fewer components, PCA will be applied to reduce noise and computation time. However, for tree-based models, raw features are preferred to maintain interpretability.",
    
    "4. Model Selection Matrix": "We will evaluate three diverse machine learning algorithms:\n\n1. Multiple Linear Regression (Baseline)\n- Strengths: Highly interpretable, fast to train, provides clear feature coefficients.\n- Weaknesses: Assumes linear relationships; performs poorly with complex, non-linear weather interactions.\n\n2. Random Forest Regressor\n- Strengths: Handles non-linear data well, robust to outliers, provides feature importance scores.\n- Weaknesses: Prone to overfitting on noisy data, acts as a 'black box' compared to linear models.\n\n3. XGBoost (Extreme Gradient Boosting)\n- Strengths: State-of-the-art performance for structured tabular data, handles missing values internally, highly customizable hyperparameters.\n- Weaknesses: Computationally intensive, sensitive to hyperparameter tuning.",
    
    "5. Decision Process & Summary": "The decision tree for model selection is as follows:\nStep 1: Train Linear Regression to establish a baseline performance metric (RMSE).\nStep 2: Train Random Forest to capture non-linear relationships.\nStep 3: Apply XGBoost for maximum predictive accuracy.\nStep 4: The final model will be chosen based on the best balance of low RMSE, high R-squared, and reasonable training time. Given the complex interactions in environmental data, XGBoost is theoretically expected to be the optimal choice."
}

# ==========================================
# WEEK 4 CONTENT
# ==========================================
week4_content = {
    "1. Objective": "This document details the strategy for model training, hyperparameter tuning, and validation for predicting the Air Quality Index (AQI). It articulates the step-by-step machine learning workflow.",
    
    "2. Data Splitting Strategy": "Because the dataset is time-series in nature (daily AQI readings), a random train-test split would cause data leakage (predicting the past using future data).\n\nMethodology: Time-Series Split\n- Training Set: First 70% of the chronological data (e.g., Jan 2018 - Dec 2021).\n- Validation Set: Next 15% of the data (e.g., Jan 2022 - Sep 2022) for hyperparameter tuning.\n- Test Set: Final 15% of the data (e.g., Oct 2022 - Dec 2023) for final model evaluation.",
    
    "3. Model Training and Tuning Plan": "Algorithm Chosen for Optimization: XGBoost Regressor.\n\nStep-by-Step Tuning Process (using GridSearchCV or RandomizedSearchCV):\n1. Learning Rate (eta): Start with testing [0.01, 0.05, 0.1]. A lower rate makes the model more robust but requires more trees.\n2. Max Depth: Test [3, 5, 7]. Controls the depth of the trees to prevent overfitting. Environmental interactions typically require moderate depth (around 5).\n3. N_Estimators: Test [100, 300, 500]. The number of boosting rounds.\n\nValidation approach: Instead of standard K-Fold cross-validation, we will use TimeSeriesSplit from Scikit-Learn to ensure the model is always validated on future data relative to the training folds.",
    
    "4. Evaluation Metrics": "Since AQI prediction is a regression problem, classification metrics (accuracy, precision, recall, F1) are not applicable. We will use standard regression metrics:\n\n1. Root Mean Squared Error (RMSE):\n- Definition: The square root of the average squared differences between predicted and actual values.\n- Justification: Heavily penalizes large errors. In AQI forecasting, failing to predict a severe pollution day (large error) is highly detrimental, making RMSE the most critical metric.\n\n2. Mean Absolute Error (MAE):\n- Definition: The average of absolute errors.\n- Justification: Provides a straightforward interpretation (e.g., 'The model predictions are off by an average of 15 AQI points').\n\n3. R-squared (R2) Score:\n- Definition: The proportion of variance in the dependent variable explained by the model.\n- Justification: Helps assess the overall goodness-of-fit. A theoretical benchmark target is an R2 score > 0.80.",
    
    "5. Workflow Summary": "1. Initialize XGBoost with default parameters.\n2. Perform TimeSeriesSplit cross-validation.\n3. Apply RandomizedSearchCV to find optimal hyperparameters.\n4. Train the final model on the combined Train + Validation sets using optimal parameters.\n5. Evaluate on the unseen Test set using RMSE and MAE.\n6. Compare performance against the baseline Linear Regression model to quantify improvement."
}

# ==========================================
# WEEK 5 CONTENT
# ==========================================
week5_content = {
    "1. Objective": "This document outlines the data visualization, reporting, and storytelling strategy for the AQI prediction project. The focus is on translating technical model outputs into a compelling, business-oriented narrative for decision-makers and the public.",
    
    "2. Narrative Arc & Storyboard": "The presentation will follow a logical progression to engage stakeholders:\n\n1. The Hook (The Problem): Start with the real-world impact. Visual: A dual-image slide showing the city skyline on a clear day vs. a severe smog day, accompanied by a headline statistic on health impacts.\n\n2. The Context (EDA Insights): Explain what drives the problem. Visual: Seasonal trends showing when AQI is at its worst, guiding policymakers on when to enforce restrictions.\n\n3. The Solution (The Model): Briefly explain how the machine learning model works without overly technical jargon. Emphasize predictive capability.\n\n4. The Impact (Results): Show the model's success. Visual: A line chart overlapping Actual AQI vs. Predicted AQI, demonstrating accuracy.\n\n5. The Call to Action (Recommendations): What should be done with this tool? (e.g., Early warning systems for schools, dynamic traffic rerouting).",
    
    "3. Selection of Visualization Tools": "1. Matplotlib & Seaborn:\n- Use Case: Static, high-resolution plots for the printed technical report.\n- Rationale: Best for deep statistical exploration (heatmaps, KDE plots) and ensuring consistency in document formatting.\n\n2. Plotly (or Streamlit for Dashboards):\n- Use Case: Interactive charts for the final presentation.\n- Rationale: Allows stakeholders to hover over data points (e.g., seeing the exact AQI on a specific date) and zoom into specific high-pollution events, making the data more accessible.",
    
    "4. Key Proposed Visualizations": "1. Feature Importance Bar Chart:\n- Description: A horizontal bar chart generated from the XGBoost model showing the top 5 most influential factors (e.g., Lag_1 AQI, Wind Speed, Winter Season).\n- Narrative Value: Explains 'WHY' the model makes its predictions. Empowers policy makers to target the biggest factors (like traffic-induced emissions during low wind).\n\n2. Seasonal Heatmap:\n- Description: Months on the X-axis, Years on the Y-axis, colored by average AQI (Green to Dark Red).\n- Narrative Value: Instantly highlights the 'pollution season', validating the need for targeted, time-specific interventions.\n\n3. Actual vs. Predicted Time-Series Plot:\n- Description: A line graph covering a 2-month testing period showing how closely the model's forecast line follows the actual AQI line.\n- Narrative Value: Builds trust in the model's reliability for future deployment.",
    
    "5. Interpreting Data for Business/Research": "The transition from data to narrative involves answering the 'So What?' question. For instance, discovering a high correlation between low wind speeds and AQI spikes isn't just a statistical fact; it translates to a recommendation: 'The city should issue public health advisories 48 hours in advance when meteorological forecasts show dropping wind speeds.' This storytelling approach ensures the data science effort translates into tangible public value."
}

# ==========================================
# GENERATE ALL DOCUMENTS
# ==========================================
if __name__ == "__main__":
    print("Generating Yuva Intern Task Reports...")
    
    # Create an output directory
    output_dir = "Yuva_Intern_Submissions"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    create_doc(f"{output_dir}/Week_1_Project_Planning.docx", "Week 1: Project Planning and Strategic Blueprint", week1_content)
    create_doc(f"{output_dir}/Week_2_Data_Preprocessing_EDA.docx", "Week 2: Data Preprocessing and EDA", week2_content)
    create_doc(f"{output_dir}/Week_3_Feature_Engineering.docx", "Week 3: Feature Engineering and Model Selection", week3_content)
    create_doc(f"{output_dir}/Week_4_Model_Development.docx", "Week 4: Machine Learning Model Development", week4_content)
    create_doc(f"{output_dir}/Week_5_Data_Storytelling.docx", "Week 5: Data Visualization and Storytelling", week5_content)
    
    print(f"\nAll 5 documents have been successfully created in the '{output_dir}' folder!")
    print("You can now upload these .docx files to the Yuva Intern portal.")
