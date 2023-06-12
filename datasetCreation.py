import pandas as pd

# Create a dictionary with the sample data
# data = {
#     'Product Name': ['Product A', 'Product B', 'Product C', 'Product D', 'Product E', 'Product F', 'Product G', 'Product H', 'Product I', 'Product J'],
#     'Date': ['2023-04-01', '2023-04-05', '2023-04-10', '2023-04-15', '2023-04-20', '2023-04-25', '2023-04-30', '2023-05-05', '2023-05-10', '2023-05-15'],
#     'Quantity': [100, 150, 200, 120, 180, 90, 250, 160, 140, 220]
# }

# # Create a DataFrame from the dictionary
# df = pd.DataFrame(data)

# Step 1: Data Collection
# Create a sample dataset
data = {
    'Date': ['2023-04-01', '2023-04-02', '2023-04-03', '2023-04-04', '2023-04-05', '2023-04-06', '2023-04-07', '2023-04-08', '2023-04-09', '2023-04-10', '2023-04-11', '2023-04-12', '2023-04-13', '2023-04-14', '2023-04-15', '2023-04-16', '2023-04-17', '2023-04-18', '2023-04-19', '2023-04-20'],
    'Product Name': ['Product A', 'Product B', 'Product C', 'Product D', 'Product E', 'Product A', 'Product B', 'Product C', 'Product D', 'Product E', 'Product A', 'Product B', 'Product C', 'Product D', 'Product E', 'Product A', 'Product B', 'Product C', 'Product D', 'Product E'],
    'Quantity': [10, 12, 8, 5, 15, 9, 11, 7, 4, 13, 8, 10, 6, 3, 14, 7, 12, 9, 6, 11]
}

# Create a DataFrame from the 

# Convert the 'Date' column to datetime type
df['Date'] = pd.to_datetime(df['Date'])

# Save the DataFrame as a CSV file
df.to_csv('product_sales.csv', index=False)
