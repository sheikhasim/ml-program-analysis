import pandas as pd
import matplotlib.pyplot as plt

# Step 1: Data Collection
# Assume you have sales data for the last one month stored in a CSV file called 'sales_data.csv'
sales_data = pd.read_csv('product_sales.csv')

# Step 2: Analyze Daily Sales Data
# Group sales data by product name and date, and calculate the total quantity sold per day
daily_sales = sales_data.groupby(['Product Name', 'Date']).sum()['Quantity'].reset_index()

# Step 3: Find the Most Demanded Products
# Calculate the total quantity sold for each product across all days
product_demand = daily_sales.groupby('Product Name').sum()['Quantity'].reset_index()
# Sort products based on total quantity sold in descending order
most_demanded_products = product_demand.sort_values(by='Quantity', ascending=False)

# Step 4: Visualize Sales Data for Each Product
# Create a line plot for each product
for product in daily_sales['Product Name'].unique():
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

# Step 5: Generate Pie Chart for Sales Distribution
# Prepare data for pie chart
product_names = product_demand['Product Name']
quantities_sold = product_demand['Quantity']

# Create a pie chart
plt.figure()
plt.pie(quantities_sold, labels=product_names, autopct='%1.1f%%')
plt.title('Sales Distribution of Products')
plt.axis('equal')
plt.show()

# Step 6: Output the Most Demanded Products List
print("Most Demanded Products:")
print(most_demanded_products)
