import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = 'C://Users//Administration//OneDrive//Desktop//Hotel score.xlsx'
excel_data = pd.ExcelFile(file_path)

# Load the first sheet into a dataframe
df = excel_data.parse(excel_data.sheet_names[0])

# Display the first few rows of the dataframe to understand its structure
print("First few rows of the dataframe:")
print(df.head())

# Descriptive statistics
descriptive_stats = df.describe()

# Count missing values
missing_values = df.isnull().sum().reset_index()
missing_values.columns = ['Column', 'Missing Values']

# Average scores per hotel
average_scores_per_hotel = df.groupby('Hotel Name').mean()

# Select only numeric columns for correlation analysis
numeric_cols = df.select_dtypes(include=['float64', 'int64'])
correlation_matrix = numeric_cols.corr()

# 1. Identify Top-Performing Hotels
top_performing_hotels = average_scores_per_hotel.mean(axis=1).sort_values(ascending=False).reset_index()
top_performing_hotels.columns = ['Hotel Name', 'Average Score']

# 2. Category-Specific Top Performers
category_top_performers = df.groupby('Hotel Name').mean().idxmax().reset_index()
category_top_performers.columns = ['Category', 'Top Performer']

# 3. Improvement Areas
improvement_areas = df.groupby('Hotel Name').mean().idxmin().reset_index()
improvement_areas.columns = ['Category', 'Needs Improvement']

# Save all outputs to a new Excel file
output_file_path = 'C://Users//Administration//OneDrive//Desktop//Hotel_score_analysis.xlsx'

with pd.ExcelWriter(output_file_path) as writer:
    df.to_excel(writer, sheet_name='Original Data', index=False)
    descriptive_stats.to_excel(writer, sheet_name='Descriptive Stats')
    missing_values.to_excel(writer, sheet_name='Missing Values', index=False)
    average_scores_per_hotel.to_excel(writer, sheet_name='Average Scores')
    correlation_matrix.to_excel(writer, sheet_name='Correlation Matrix')
    top_performing_hotels.to_excel(writer, sheet_name='Top Performers', index=False)
    category_top_performers.to_excel(writer, sheet_name='Category Top Performers', index=False)
    improvement_areas.to_excel(writer, sheet_name='Improvement Areas', index=False)

print(f"Analysis saved to {output_file_path}")

# 4. Distribution Analysis
fig, axes = plt.subplots(3, 2, figsize=(14, 14))

# Distribution of Presentability scores
sns.histplot(df['Presentability'], kde=True, ax=axes[0, 0])
axes[0, 0].set_title('Distribution of Presentability Scores')

# Distribution of Conference Facilities scores
sns.histplot(df['Conference Facilities'], kde=True, ax=axes[0, 1])
axes[0, 1].set_title('Distribution of Conference Facilities Scores')

# Distribution of Food scores
sns.histplot(df['Food'], kde=True, ax=axes[1, 0])
axes[1, 0].set_title('Distribution of Food Scores')

# Distribution of Customer Service scores
sns.histplot(df['Customer Service'], kde=True, ax=axes[1, 1])
axes[1, 1].set_title('Distribution of Customer Service Scores')

# Distribution of Accommodation Facilities scores
sns.histplot(df['Accommodation Facilities'].dropna(), kde=True, ax=axes[2, 0])
axes[2, 0].set_title('Distribution of Accommodation Facilities Scores')

# Distribution of WIFI Availability scores
sns.histplot(df['WIFI Availability'], kde=True, ax=axes[2, 1])
axes[2, 1].set_title('Distribution of WIFI Availability Scores')

plt.tight_layout()
plt.show()
