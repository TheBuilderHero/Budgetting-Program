import pandas as pd

# Create a DataFrame with a date column
df = pd.DataFrame({'date': ['01-01-2023', '02-15-2023', '03-31-2023']})

# Convert the 'date' column to datetime format
df['date'] = pd.to_datetime(df['date'])

# Extract the month name
df['month_name'] = df['date'].dt.month_name()

print(df)