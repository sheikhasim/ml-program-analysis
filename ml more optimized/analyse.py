import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date
import math

# Step 1: Data Collection
sales_data = pd.read_csv('developed_data/createdData.csv')


# In this step, we start by reading the sales data from a CSV file using the Pandas library.
# Step 2: Analyze Weekly Sales Data
sales_data['Date'] = pd.to_datetime(sales_data['Date'])
sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()


# Here, we convert the 'Date' column to a datetime format and group the data by 
# 'Date' and 'Product Name', aggregating the 'Quantity' column to get the total sales quantity for each product on each date.
# Step 3: Analyze Weekly Sales Data
sales_data['Week'] = sales_data['Date'].dt.isocalendar().week
weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()


# Next, we extract the week number from the 'Date' column and add it as a new 'Week' column. 
# We then group the data by 'Product Name', 'Week', and 'Date', and calculate the total sales quantity for each product in each week.
# Step 4: Calculate Demand and EOQ
demand_per_week = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
demand_per_week = demand_per_week.rename(columns={'Quantity': 'Demand'})

# Calculate EOQ for each row
demand_per_week['EOQ'] = demand_per_week.apply(lambda row: round(math.sqrt((2 * 500 * row['Demand']) / 100), 2), axis=1)  # Assuming carrying cost is $500 and ordering cost is $100


# In this step, we calculate the demand per week by summing the weekly sales quantity for each product. We rename the 'Quantity' column to 'Demand' for clarity. 
# Then, we calculate the Economic Order Quantity (EOQ) for each row using the EOQ formula and assuming carrying cost as $500 and ordering cost as $100.
# Step 5: Export Demand and EOQ to Excel
excel_folder = 'inventory_analysis'
os.makedirs(excel_folder, exist_ok=True)  # Create the inventory_analysis directory
excel_filename = f"{excel_folder}/inventory_analysis_{date.today().strftime('%Y-%m-%d')}.xlsx"
excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
demand_per_week.to_excel(excel_writer, sheet_name='Demand and EOQ', index=False)


# Here, we create a folder to store the inventory analysis and define the filename for the Excel file.
#  We then create an Excel writer object and save the demand and EOQ data to a new worksheet named 'Demand and EOQ'.
# Step 6: Create Line Plots for Each Product and Save them as Images
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



# In this step, we iterate over each unique product in the weekly sales data. For each product, we filter the data and extract the dates and quantities. 
# Then, we create a line plot of the sales data, customize the plot, save it as an image, and close the plot. Finally, we add the line plot image to the Excel file as a new worksheet.
# Step 7: Create a Pie Chart for Total Sales Distribution and Save it as an Image
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



# In this step, we calculate the total sales quantity for each product and create a pie chart to visualize the distribution of total sales.
#  We save the pie chart as an image, close the plot, and add the image to the Excel file as a new worksheet.
# Step 8: Create a Line Plot for Demand and EOQ
demand_eoq_plot = demand_per_week.groupby('Week')[['Demand', 'EOQ']].sum().reset_index()
plt.figure()
plt.plot(demand_eoq_plot['Week'], demand_eoq_plot['Demand'], label='Demand', marker='o')
plt.plot(demand_eoq_plot['Week'], demand_eoq_plot['EOQ'], label='EOQ', marker='o')
plt.title("Demand vs. EOQ")
plt.xlabel("Week")
plt.ylabel("Quantity")
plt.legend()

# Save the line plot as an image
demand_eoq_plot_filename = "demand_eoq_plot.png"
plt.savefig(demand_eoq_plot_filename)
plt.close()

# Add the demand vs. EOQ line plot image to the Excel file
worksheet_name = "Demand vs. EOQ Plot"
worksheet = excel_writer.book.add_worksheet(worksheet_name)
worksheet.insert_image('A1', demand_eoq_plot_filename)



# This step involves creating a line plot to compare the demand and EOQ values over the weeks. 
# We group the demand and EOQ data by week, plot the lines, add labels and titles, and save the plot as an image. 
# The image is then added to the Excel file as a new worksheet.
# Step 9: Save and Close the Excel File
excel_writer._save()
excel_writer.close()

print(f"Exported demand analysis, line plots, charts, and EOQ graph to {excel_filename}")
