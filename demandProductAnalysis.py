import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Step 1: Data Collection
# Assume you have sales data for the last one month stored in a CSV file called 'sales_data.csv'
sales_data = pd.read_csv('sales_data.csv')

# Step 2: Analyze Daily Sales Data
# Group sales data by product name and date, and calculate the total quantity sold per day
daily_sales = sales_data.groupby(['Product Name', 'Date']).sum()['Quantity'].reset_index()

# Step 3: Find the Most Demanded Products
# Calculate the total quantity sold for each product across all days
product_demand = daily_sales.groupby('Product Name').sum()['Quantity'].reset_index()
# Sort products based on total quantity sold in descending order
most_demanded_products = product_demand.sort_values(by='Quantity', ascending=False)

# Step 4: Visualize Sales Data for Each Product
# for product in most_demanded_products['Product Name']:
#     product_sales = daily_sales[daily_sales['Product Name'] == product]
    
#     # Prepare data for plotting
#     dates = pd.to_datetime(product_sales['Date'])
#     quantities = product_sales['Quantity']
    
#     # Create a line plot for the sales data
#     plt.figure()
#     plt.plot(dates, quantities)
#     plt.title(f'Sales Data for {product}')
#     plt.xlabel('Date')
#     plt.ylabel('Quantity Sold')
#     plt.xticks(rotation=45)
#     plt.show()



sns.set(style='darkgrid')

for product in most_demanded_products['Product Name']:
    product_sales = daily_sales[daily_sales['Product Name'] == product]
    
    # Prepare data for plotting
    dates = pd.to_datetime(product_sales['Date'])
    quantities = product_sales['Quantity']
    
    # Create a line plot for the sales data
    plt.figure()
    plt.plot(dates, quantities)
    plt.title(f'Sales Data for {product}')
    plt.xlabel('Date')
    plt.ylabel('Quantity Sold')
    plt.xticks(rotation=45)
    plt.show()

# Step 5: Output the Most Demanded Products List
print("Most Demanded Products:")
print(most_demanded_products)
