import pandas as pd


data = {    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'City': ['New York', 'Los Angeles', 'Chicago']}
df = pd.DataFrame(data)
df.to_excel('output.xlsx', index=False)
# This code creates a DataFrame with sample data and saves it to an Excel file named 'output.xlsx'.