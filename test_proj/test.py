import pandas as pd

# Create a DataFrame from a dictionary
data = {
    'Name': ['Alice', 'Bob', 'Charlie', 'David'],
    'Age': [25, 30, 35, 40],
    'City': ['New York', 'Los Angeles', 'Chicago', 'Houston']
}

df = pd.DataFrame(data)

# Display the whole DataFrame
print("=== Original DataFrame ===")
print(df)

# Filter: People older than 30
print("\n=== Age > 30 ===")
print(df[df['Age'] > 30])

# Add a new column (age in months)
df['AgeMonths'] = df['Age'] * 12
print("\n=== With Age in Months ===")
print(df)

# Basic statistics
print("\n=== Summary Statistics ===")
print(df.describe())

# Sort by Age descending
print("\n=== Sorted by Age (desc) ===")
print(df.sort_values('Age', ascending=False))