import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date
from sklearn.cluster import KMeans
from sklearn.mixture import GaussianMixture
from sklearn.cluster import DBSCAN
from sklearn.cluster import AgglomerativeClustering
from sklearn.cluster import Birch

# Step 1: Data Collection
sales_data = pd.read_csv('deveopedData/createdData.csv')

# Step 2: Analyze Weekly Sales Data
sales_data['Date'] = pd.to_datetime(sales_data['Date'])
sales_data = sales_data.groupby(['Date', 'Region', 'Product Name'])['Quantity'].sum().reset_index()

# Step 3: Analyze Weekly Sales Data
sales_data['Week'] = sales_data['Date'].dt.isocalendar().week
weekly_sales = sales_data.groupby(['Region', 'Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# Step 4: Find the Demanded Products for the Last Two Weeks in each region
last_two_weeks = weekly_sales['Week'].unique()[-2:]
region_product_demand = weekly_sales.groupby(['Region', 'Product Name', 'Week'])['Quantity'].sum().reset_index()
region_product_demand = region_product_demand[region_product_demand['Week'].isin(last_two_weeks)]
most_demanded_products_by_region = region_product_demand.groupby(['Region', 'Product Name'])['Quantity'].sum().reset_index()
most_demanded_products_by_region = most_demanded_products_by_region.sort_values(by=['Region', 'Quantity'], ascending=False)


# Step 5: Calculate Increase and Decrease in Demand by Region
last_week_demand_by_region = region_product_demand[region_product_demand['Week'] == last_two_weeks[0]]
current_week_demand_by_region = region_product_demand[region_product_demand['Week'] == last_two_weeks[1]]
demand_change_by_region = pd.merge(current_week_demand_by_region, last_week_demand_by_region,
                                   on=['Region', 'Product Name'], how='left')
demand_change_by_region['Change'] = demand_change_by_region['Quantity_x'] - demand_change_by_region['Quantity_y']
demand_change_by_region['Change(%)'] = (demand_change_by_region['Change'] / demand_change_by_region['Quantity_y']) * 100

# Separate increase and decrease in demand by region
increase_demand_by_region = demand_change_by_region[demand_change_by_region['Change'] > 0]
decrease_demand_by_region = demand_change_by_region[demand_change_by_region['Change'] < 0]

# Print Increase and Decrease in Demand by Region
print("\nIncrease in Demand by Region:")
print(increase_demand_by_region[['Region', 'Product Name', 'Week_x', 'Quantity_x', 'Week_y', 'Quantity_y', 'Change', 'Change(%)']].to_string(index=False))
print("\nDecrease in Demand by Region:")
print(decrease_demand_by_region[['Region', 'Product Name', 'Week_x', 'Quantity_x', 'Week_y', 'Quantity_y', 'Change', 'Change(%)']].to_string(index=False))


# Step 8: Create a Pandas Excel Writer
excel_folder = 'demanded_products'
os.makedirs(excel_folder, exist_ok=True)  # Create the demanded_products directory
excel_filename = f"{excel_folder}/demanded_products_based_on_regions{date.today()}.xlsx"
excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# Step 9: Export Increase and Decrease in Demand by Region to Excel
increase_demand_by_region.to_excel(excel_writer, sheet_name='Increase in Demand by Region', index=False)
decrease_demand_by_region.to_excel(excel_writer, sheet_name='Decrease in Demand by Region', index=False)

# Step 10: Export Most Demanded Products by Region to Excel
most_demanded_products_by_region.to_excel(excel_writer, sheet_name='Demanded Products by Region', index=False)

# Step 11: Create Bar Charts for Most Demanded Products by Region and Save them as Images
os.makedirs('bar_charts', exist_ok=True)  # Create the bar_charts directory

for region in most_demanded_products_by_region['Region'].unique():
    region_products = most_demanded_products_by_region[most_demanded_products_by_region['Region'] == region]

    # Prepare data for plotting
    products = region_products['Product Name']
    quantities = region_products['Quantity']

    # Create a bar chart for the demanded products in the region
    plt.figure()
    plt.bar(products, quantities)
    plt.title(f'Demanded Products in {region}')
    plt.xlabel('Product Name')
    plt.ylabel('Quantity Demanded')
    plt.xticks(rotation=45)

    # Save the bar chart as an image
    image_filename = f"bar_charts/{region}_bar_chart.png"
    plt.savefig(image_filename)
    plt.close()

    # Shorten the worksheet name if it exceeds the limit
    worksheet_name = f"{region[:30]} Bar Chart" if len(region) > 30 else f"{region} Bar Chart"

    # Add the bar chart image to the Excel file
    worksheet = excel_writer.book.add_worksheet(worksheet_name)
    worksheet.insert_image('A1', image_filename)

# Step 12: Save and Close the Excel File
excel_writer._save()
excel_writer.close()

# Delete the bar chart images
for region in most_demanded_products_by_region['Region'].unique():
    image_filename = f"bar_charts/{region}_bar_chart.png"
    os.remove(image_filename)

print(f"Exported most demanded products, demand changes, and bar charts by region to {excel_filename}")
