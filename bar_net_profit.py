import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px

# Replace 'your_excel_file.xlsx' with the actual path to your Excel file
excel_file_path = r'C:\Users\marke\Desktop\wip\IQRAinvest\offer_export.xlsx'

# Read the Excel file into a Pandas DataFrame
df = pd.read_excel(excel_file_path)

# Define a function to round numerical columns to 2 decimal places
def round_numeric_columns(column):
    if pd.api.types.is_numeric_dtype(column):
        return column.round(2)
    return column

# Apply the rounding function to all columns in the DataFrame
df = df.apply(round_numeric_columns)

# Display the first few rows of the DataFrame to verify the import
print(df.head())



# Assuming your DataFrame is named 'df'
# Extract relevant columns
products = df['Product Title']
net_profit = df['Net Profit - Leadtime']

# Create an interactive bar chart with Plotly
fig = px.bar(x=products, y=net_profit, color=net_profit,
             labels={'x': 'Product Title', 'y': 'Net Profit - Leadtime'},
             title='Net Profit for Each Product',
             text=net_profit)

# Rotate x-axis labels for better visibility
fig.update_layout(xaxis=dict(tickangle=45))

# Show the interactive plot
fig.show()
