import pandas as pd

def calculate_statistics_for_column(df, column_name):
    # Drop NaN values in the column
    df = df.dropna(subset=[column_name])

    # Ensure the column is numeric
    df[column_name] = pd.to_numeric(df[column_name], errors='coerce')

    # Remove NaN values resulting from the conversion and zero values
    df = df[df[column_name] > 0]

    # Calculate statistics
    total = df[column_name].sum()
    average = df[column_name].mean()
    median = df[column_name].median()
    variance = df[column_name].var()
    stdev = df[column_name].std()
    max_value = df[column_name].max()
    min_value = df[column_name].min()


    # Display statistical results
    print(f"\nColumn: {column_name}")
    print(f"Total: {total}")
    print(f"Average: {average}")
    print(f"Median: {median}")
    print(f"Variance: {variance}")
    print(f"Standard Deviation: {stdev}")
    print(f"Max: {max_value}")
    print(f"Min: {min_value}")


# Load the Excel files
df1 = pd.read_excel('Suspected.xlsx')
df2 = pd.read_excel('Diagnosed.xlsx')
df3 = pd.read_excel('normalCounted11.xlsx')
df4 = pd.read_excel('normalCounted22.xlsx')


# Column names to analyze
columns_to_analyze = ['num of character', 'num of word']

print("\nStatistics for suspectCounted.xlsx")
for column in columns_to_analyze:
    calculate_statistics_for_column(df1, column)

print("\n\nStatistics for sureCounted.xlsx")
for column in columns_to_analyze:
    calculate_statistics_for_column(df2, column)

print("\n\nStatistics for normalCounted11.xlsx")
for column in columns_to_analyze:
    calculate_statistics_for_column(df3, column)

print("\n\nStatistics for normalCounted22.xlsx")
for column in columns_to_analyze:
    calculate_statistics_for_column(df4, column)
