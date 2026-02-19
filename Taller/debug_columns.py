import pandas as pd

# Load the CSV and check columns
df = pd.read_csv('/Users/javierfernandez/Desktop/Taller/Libro1.csv', sep=';')
print("Original columns:")
print(df.columns.tolist())
print("\nFirst few rows:")
print(df.head(2))
print("\nColumn data types:")
print(df.dtypes)
