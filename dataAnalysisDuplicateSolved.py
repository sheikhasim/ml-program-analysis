import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date

# Step 1: Data Collection
# Assume you have sales data for the last two months stored in a CSV file called 'sales_data.csv'
sales_data = pd.read_csv('./deveopedData/createdData.csv')

# Step 2: Analyze Weekly Sales Data
# Convert the 'Date' column to datetime type
sales_data['Date'] = pd.to_datetime(sales_data['Date'])

# Consolidate the quantity sold for the same date and product
sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()

# Step 3: Analyze Weekly Sales Data
# Extract the week number from the 'Date' column
sales_data['Week'] = sales_data['Date'].dt.isocalendar().week

# Group sales data by product name, week, and date, and calculate the total quantity sold per day
weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# Step 4: Find the Demanded Products for the Last Two Weeks
# Get the last two weeks' data
last_two_weeks = weekly_sales['Week'].unique()[-2:]

# Calculate the total quantity sold for each product in the last two weeks
product_demand = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
product_demand = product_demand[product_demand['Week'].isin(last_two_weeks)]
# Sort products based on total quantity sold in descending order
most_demanded_products = product_demand.groupby('Product Name')['Quantity'].sum().reset_index().sort_values(
    by='Quantity', ascending=False)

# Step 5: Create a Pandas Excel Writer
excel_folder = 'demanded_products'
excel_filename = f"{excel_folder}/demanded_products_{date.today()}.xlsx"
excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# Step 6: Export Most Demanded Products to Excel
most_demanded_products.to_excel(excel_writer, sheet_name='Demanded Products', index=False)

# Step 7: Create Line Plots for Each Product and Save them as Images
os.makedirs('line_plots', exist_ok=True)  # Create the line_plots directory

for product in weekly_sales['Product Name'].unique():
    product_sales = weekly_sales[weekly_sales['Product Name'] == product]

    # Prepare data for plotting
    dates = product_sales['Date']
    quantities = product_sales['Quantity']

    # Create a line plot for the sales data
    plt.figure()
    plt.plot(dates, quantities, marker='o')
    plt.title(f'Sales Data for {product}')
    plt.xlabel('Date')
    plt.ylabel('Quantity Sold')
    plt.xticks(rotation=45)

    # Save the line plot as an image
    image_filename = f"line_plots/{product}_line_plot.png"
    plt.savefig(image_filename)
    plt.close()

    # Add the line plot image to the Excel file
    worksheet_name = f"{product} Line Plot"
    worksheet = excel_writer.book.add_worksheet(worksheet_name)
    worksheet.insert_image('A1', image_filename)

# Step 8: Create a Pie Chart for Total Sales Distribution and Save it as an Image
total_sales = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
product_sales_total = total_sales.groupby('Product Name')['Quantity'].sum()
plt.figure()
plt.pie(product_sales_total, labels=product_sales_total.index, autopct='%1.1f%%')
plt.title("Total Sales Distribution")

# Save the pie chart as an image
pie_chart_filename = "total_sales_pie_chart.png"
plt.savefig(pie_chart_filename)
plt.close()

# Add the pie chart image to the Excel file
worksheet_name = "Total Sales Pie Chart"
worksheet = excel_writer.book.add_worksheet(worksheet_name)
worksheet.insert_image('A1', pie_chart_filename)

# Step 9: Save and Close the Excel File
excel_writer._save()
excel_writer.close()

# Delete the line plot images
for product in weekly_sales['Product Name'].unique():
    image_filename = f"line_plots/{product}_line_plot.png"
    os.remove(image_filename)

# Delete the pie chart image
os.remove(pie_chart_filename)

print(f"Exported most demanded products and charts to {excel_filename}")






# import os
# import pandas as pd
# import matplotlib.pyplot as plt
# from datetime import date

# # Step 1: Data Collection
# # Assume you have sales data for the last two months stored in a CSV file called 'sales_data.csv'
# sales_data = pd.read_csv('sales_data1.csv')

# # Step 2: Analyze Weekly Sales Data
# # Convert the 'Date' column to datetime type
# sales_data['Date'] = pd.to_datetime(sales_data['Date'])

# # Consolidate the quantity sold for the same date and product
# sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()

# # Step 3: Analyze Weekly Sales Data
# # Extract the week number from the 'Date' column
# sales_data['Week'] = sales_data['Date'].dt.isocalendar().week

# # Group sales data by product name, week, and date, and calculate the total quantity sold per day
# weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# # Step 4: Find the Demanded Products for the Last Two Weeks
# # Get the last two weeks' data
# last_two_weeks = weekly_sales['Week'].unique()[-2:]

# # Calculate the total quantity sold for each product in the last two weeks
# product_demand = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
# product_demand = product_demand[product_demand['Week'].isin(last_two_weeks)]
# # Sort products based on total quantity sold in descending order
# most_demanded_products = product_demand.groupby('Product Name')['Quantity'].sum().reset_index().sort_values(
#     by='Quantity', ascending=False)

# # Step 5: Create a Pandas Excel Writer
# excel_filename = f"demanded_products_{date.today()}.xlsx"
# excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# # Step 6: Export Most Demanded Products to Excel
# most_demanded_products.to_excel(excel_writer, sheet_name='Demanded Products', index=False)

# # Step 7: Create Line Plots for Each Product and Export to Excel
# for product in weekly_sales['Product Name'].unique():
#     product_sales = weekly_sales[weekly_sales['Product Name'] == product]

#     # Prepare data for plotting
#     dates = product_sales['Date']
#     quantities = product_sales['Quantity']

#     # Create a line plot for the sales data
#     plt.figure()
#     plt.plot(dates, quantities, marker='o')
#     plt.title(f'Sales Data for {product}')
#     plt.xlabel('Date')
#     plt.ylabel('Quantity Sold')
#     plt.xticks(rotation=45)

#     # Save the line plot as an image and add it to the Excel file
#     image_filename = f"{product}_line_plot.png"
#     plt.savefig(image_filename)
#     plt.close()

#     # Add the line plot image to the Excel file
#     worksheet_name = f"{product} Line Plot"
#     worksheet = excel_writer.book.add_worksheet(worksheet_name)
#     worksheet.insert_image('A1', image_filename)

# # Step 8: Create a Pie Chart for all Products and Add it to Excel
# product_sales_total = sales_data.groupby('Product Name')['Quantity'].sum()

# # Create a pie chart
# plt.figure()
# plt.pie(product_sales_total, labels=product_sales_total.index, autopct='%1.1f%%')
# plt.title("Total Sales Distribution")

# # Save the pie chart as an image and add it to the Excel file
# pie_chart_filename = "total_sales_pie_chart.png"
# plt.savefig(pie_chart_filename)
# plt.close()

# # Add the pie chart image to the Excel file
# worksheet = excel_writer.book.add_worksheet("Total Sales Pie Chart")
# worksheet.insert_image('A1', pie_chart_filename)

# # Save and close the Excel file
# excel_writer._save()

# print(f"Exported most demanded products and charts to {excel_filename}")


# import os
# import pandas as pd
# import matplotlib.pyplot as plt
# from datetime import date

# # Step 1: Data Collection
# # Assume you have sales data for the last two months stored in a CSV file called 'sales_data.csv'
# sales_data = pd.read_csv('sales_data1.csv')

# # Step 2: Analyze Weekly Sales Data
# # Convert the 'Date' column to datetime type
# sales_data['Date'] = pd.to_datetime(sales_data['Date'])

# # Consolidate the quantity sold for the same date and product
# sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()

# # Step 3: Analyze Weekly Sales Data
# # Extract the week number from the 'Date' column
# sales_data['Week'] = sales_data['Date'].dt.isocalendar().week

# # Group sales data by product name, week, and date, and calculate the total quantity sold per day
# weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# # Step 4: Find the Demanded Products for the Last Two Weeks
# # Get the last two weeks' data
# last_two_weeks = weekly_sales['Week'].unique()[-2:]

# # Calculate the total quantity sold for each product in the last two weeks
# product_demand = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
# product_demand = product_demand[product_demand['Week'].isin(last_two_weeks)]
# # Sort products based on total quantity sold in descending order
# most_demanded_products = product_demand.groupby('Product Name')['Quantity'].sum().reset_index().sort_values(
#     by='Quantity', ascending=False)

# # Step 5: Visualize Sales Data for Each Product
# # Create a line plot for each product
# for product in weekly_sales['Product Name'].unique():
#     product_sales = weekly_sales[weekly_sales['Product Name'] == product]

#     # Prepare data for plotting
#     dates = product_sales['Date']
#     quantities = product_sales['Quantity']

#     # Create a line plot for the sales data
#     plt.figure()
#     plt.plot(dates, quantities, marker='o')
#     plt.title(f'Sales Data for {product}')
#     plt.xlabel('Date')
#     plt.ylabel('Quantity Sold')
#     plt.xticks(rotation=45)
#     plt.show()

# # Step 6: Output the Demanded Products for the Last Two Weeks
# print("Demanded Products for the Last Two Weeks:")
# print(most_demanded_products)

# # Step 7: Create a folder and Export Most Demanded Products to Excel
# folder_name = 'demanded_products'
# os.makedirs(folder_name, exist_ok=True)
# filename = f"{folder_name}/{date.today()}_demanded_products.csv"
# most_demanded_products.to_csv(filename, index=False)
# print(f"Exported most demanded products to {filename}")

# # Step 8: Create a Pie Chart for all Products
# # Calculate the total quantity sold for each product
# product_sales_total = sales_data.groupby('Product Name')['Quantity'].sum()

# # Create a pie chart
# plt.figure()
# plt.pie(product_sales_total, labels=product_sales_total.index, autopct='%1.1f%%')
# plt.title("Total Sales Distribution")
# plt.show()