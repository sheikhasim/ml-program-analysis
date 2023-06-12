import pandas as pd
import matplotlib.pyplot as plt

# Step 1: Data Collection
# Assume you have sales data for the last two months stored in a CSV file called 'sales_data.csv'
sales_data = pd.read_csv('sales_data1.csv')

# Step 2: Analyze Weekly Sales Data
# Convert the 'Date' column to datetime type
sales_data['Date'] = pd.to_datetime(sales_data['Date'])
# Extract the week number from the 'Date' column
sales_data['Week'] = sales_data['Date'].dt.isocalendar().week

# Group sales data by product name, week, and date, and calculate the total quantity sold per day
weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# Step 3: Find the Demanded Products for the Last Two Weeks
# Get the last two weeks' data
last_two_weeks = weekly_sales['Week'].unique()[-2:]

# Calculate the total quantity sold for each product in the last two weeks
product_demand = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
product_demand = product_demand[product_demand['Week'].isin(last_two_weeks)]
# Sort products based on total quantity sold in descending order
most_demanded_products = product_demand.groupby('Product Name')['Quantity'].sum().reset_index().sort_values(by='Quantity', ascending=False)

# Step 4: Visualize Sales Data for Each Product
# Create a line plot for each product
for product in weekly_sales['Product Name'].unique():
    product_sales = weekly_sales[weekly_sales['Product Name'] == product]
    
    # Prepare data for plotting
    weeks = product_sales['Week'].astype(int)  # Convert week numbers to integers
    quantities = product_sales['Quantity']
    
    # Create a line plot for the sales data
    plt.figure()
    plt.plot(weeks, quantities, marker='o')
    plt.title(f'Sales Data for {product}')
    plt.xlabel('Week')
    plt.ylabel('Quantity Sold')
    plt.xticks(weeks.unique())
    plt.show()

# Step 5: Output the Demanded Products for the Last Two Weeks
print("Demanded Products for the Last Two Weeks:")
print(most_demanded_products)

# Step 6: Create a Pie Chart for Total Sales of All Products
# Calculate the total sales for each product
total_sales = weekly_sales.groupby('Product Name')['Quantity'].sum()

# Plot the pie chart
plt.figure()
plt.pie(total_sales, labels=total_sales.index, autopct='%1.1f%%')
plt.title('Total Sales for All Products')
plt.show()
