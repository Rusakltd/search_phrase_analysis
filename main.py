import pandas as pd
from openpyxl import load_workbook
from collections import defaultdict

# Load the workbook and select the active sheet
wb = load_workbook('export.xlsx')
sheet = wb.active

# Delete the first four rows (assuming you want to skip them)
sheet.delete_rows(0, 4)

# Create the DataFrame with the first row as column names
df = pd.DataFrame(sheet.values, columns=[cell.value for cell in next(sheet.rows)])

# Delete the first row
df.drop(0, inplace=True)

# Subset the DataFrame
df = df[['Поисковый запрос', 'Показы', 'Клики', 'Расход (руб.)', 'Конверсии']]

# Rename the columns, replace '-' with '0' and convert data types
df.columns = ['search phrase', 'impressions', 'clicks', 'spent', 'conversions']
df['conversions'] = df['conversions'].replace('-','0')
df['conversions'] = pd.to_numeric(df['conversions'])

# Initialize defaultdict to store aggregated metrics
word_metrics = defaultdict(lambda: {'impressions': 0, 'clicks': 0, 'spent':0, 'conversions': 0})

# Iterate through each row in DataFrame
for index, row in df.iterrows():
    search_phrase = row['search phrase']
    impressions = row['impressions']
    clicks = row['clicks']
    spent = row['spent']
    conversions = row['conversions']
    
    # Split search phrase into words
    words = search_phrase.split()
    
    # Update aggregated metrics for each word
    for word in words:
        word_metrics[word]['impressions'] += impressions
        word_metrics[word]['clicks'] += clicks
        word_metrics[word]['spent'] += spent
        word_metrics[word]['conversions'] += conversions

# Create DataFrame from aggregated metrics
word_df = pd.DataFrame.from_dict(word_metrics, orient='index').reset_index()
word_df.columns = ['word', 'impressions', 'clicks', 'spent', 'conversions']

# Save DataFrame to Excel
word_df.to_excel('output.xlsx', index=False)