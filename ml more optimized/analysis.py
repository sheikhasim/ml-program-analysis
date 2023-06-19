# In this updated code, we calculate the demand per week for
# each product and then use the EOQ formula to determine the 
# optimal order quantity. We assume carrying cost is $500 and
# ordering cost is $100. The demand and EOQ values are then 
# exported to an Excel file. Additionally, line plots for 
# each product's sales data and a pie chart for total sales
# distribution are created and saved as images, which are then added to the Excel file.


# iterate over the rows of the DataFrame and calculate the EOQ value for each row
#  use the dt.isocalendar().week method to extract the week number.




import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date
import math

# Step 1: Data Collection
sales_data = pd.read_csv('developed_data/createdData.csv')

# Step 2: Analyze Weekly Sales Data
sales_data['Date'] = pd.to_datetime(sales_data['Date'])
sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()

# Step 3: Analyze Weekly Sales Data
sales_data['Week'] = sales_data['Date'].dt.isocalendar().week
weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# Step 4: Calculate Demand and EOQ
demand_per_week = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
demand_per_week = demand_per_week.rename(columns={'Quantity': 'Demand'})

# Calculate EOQ for each row
demand_per_week['EOQ'] = demand_per_week.apply(lambda row: round(math.sqrt((2 * 500 * row['Demand']) / 100), 2), axis=1)  # Assuming carrying cost is $500 and ordering cost is $100

# Step 5: Export Demand and EOQ to Excel
excel_folder = 'inventory_analysis'
os.makedirs(excel_folder, exist_ok=True)  # Create the inventory_analysis directory
excel_filename = f"{excel_folder}/inventory_analysis_{date.today().strftime('%Y-%m-%d')}.xlsx"
excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
demand_per_week.to_excel(excel_writer, sheet_name='Demand and EOQ', index=False)

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

# # Step 8: Compare Current Week's Demand with Previous Week's Demand
# increase_decrease_sheet = excel_writer.book.add_worksheet('Increase in Demand')
# decrease_in_demand_sheet = excel_writer.book.add_worksheet('Decrease in Demand')

# # Set column headers for Increase sheet
# increase_decrease_sheet.write(0, 0, 'Product')
# increase_decrease_sheet.write(0, 1, 'Current Week')
# increase_decrease_sheet.write(0, 2, 'Previous Week')
# increase_decrease_sheet.write(0, 3, 'Increase')
# increase_decrease_sheet.write(0, 4, 'Percentage Change')

# # Set column headers for Decrease sheet
# decrease_in_demand_sheet.write(0, 0, 'Product')
# decrease_in_demand_sheet.write(0, 1, 'Current Week')
# decrease_in_demand_sheet.write(0, 2, 'Previous Week')
# decrease_in_demand_sheet.write(0, 3, 'Decrease')
# decrease_in_demand_sheet.write(0, 4, 'Percentage Change')

# # Get the maximum week number
# max_week = demand_per_week['Week'].max()

# # Compare the demand for each product between the current and previous week
# for product in demand_per_week['Product Name'].unique():
#     product_demand = demand_per_week[demand_per_week['Product Name'] == product]

#     # Get the demand for the current week
#     current_week = max_week
#     current_week_demand = product_demand[product_demand['Week'] == current_week]['Demand'].values

#     if len(current_week_demand) == 0:
#         continue

#     current_week_demand = current_week_demand[0]

#     # Get the demand for the previous week
#     previous_week = current_week - 1
#     previous_week_demand = product_demand[product_demand['Week'] == previous_week]['Demand'].values

#     if len(previous_week_demand) == 0:
#         continue

#     previous_week_demand = previous_week_demand[0]

#     # Calculate the percentage change
#     if previous_week_demand != 0:
#         percentage_change = ((current_week_demand - previous_week_demand) / previous_week_demand) * 100
#     else:
#         percentage_change = 0

#     # Compare and store the results in the appropriate worksheet
#     if current_week_demand > previous_week_demand:
#         increase_decrease_sheet.write('A2', product)
#         increase_decrease_sheet.write('B2', current_week)
#         increase_decrease_sheet.write('C2', previous_week)
#         increase_decrease_sheet.write('D2', current_week_demand - previous_week_demand)
#         increase_decrease_sheet.write('E2', percentage_change)
#     else:
#         decrease_in_demand_sheet.write('A2', product)
#         decrease_in_demand_sheet.write('B2', current_week)
#         decrease_in_demand_sheet.write('C2', previous_week)
#         decrease_in_demand_sheet.write('D2', previous_week_demand - current_week_demand)
#         decrease_in_demand_sheet.write('E2', percentage_change)












# Step 2: Analyze Weekly Sales Data
sales_data['Date'] = pd.to_datetime(sales_data['Date'])
sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()

# Step 3: Analyze Weekly Sales Data
sales_data['Week'] = sales_data['Date'].dt.isocalendar().week
weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# Step 4: Find the Demanded Products for the Last Two Weeks
last_two_weeks = weekly_sales['Week'].unique()[-2:]
product_demand = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
product_demand = product_demand[product_demand['Week'].isin(last_two_weeks)]
most_demanded_products = product_demand.groupby('Product Name')['Quantity'].sum().reset_index().sort_values(
    by='Quantity', ascending=False)

# Step 5: Calculate Increase and Decrease in Demand
last_week_demand = product_demand[product_demand['Week'] == last_two_weeks[0]]
current_week_demand = product_demand[product_demand['Week'] == last_two_weeks[1]]
demand_change = pd.merge(current_week_demand, last_week_demand, on='Product Name', how='left')
demand_change['Change'] = demand_change['Quantity_x'] - demand_change['Quantity_y']
demand_change['Change(%)'] = (demand_change['Change'] / demand_change['Quantity_y']) * 100

# Separate increase and decrease in demand
increase_demand = demand_change[demand_change['Change'] > 0]
decrease_demand = demand_change[demand_change['Change'] < 0]

# Print Increase and Decrease in Demand
print("\nIncrease in Demand:")
print(increase_demand[['Product Name', 'Week_x', 'Quantity_x', 'Week_y', 'Quantity_y', 'Change', 'Change(%)']].to_string(index=False))
print("\nDecrease in Demand:")
print(decrease_demand[['Product Name', 'Week_x', 'Quantity_x', 'Week_y', 'Quantity_y', 'Change', 'Change(%)']].to_string(index=False))

# # Step 8: Create a Pandas Excel Writer
# excel_folder = 'demanded_products'
# os.makedirs(excel_folder, exist_ok=True)  # Create the demanded_products directory
# excel_filename = f"{excel_folder}/demanded_products_based_on_products{date.today()}.xlsx"
# excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# Step 9: Export Increase and Decrease in Demand to Excel
increase_demand.to_excel(excel_writer, sheet_name='Increase in Demand', index=False)
decrease_demand.to_excel(excel_writer, sheet_name='Decrease in Demand', index=False)



# Step 9: Save and Close the Excel File
excel_writer._save()
excel_writer.close()

# Delete the line plot images
for product in weekly_sales['Product Name'].unique():
    image_filename = f"line_plots/{product}_line_plot.png"
    os.remove(image_filename)

# Delete the pie chart image
os.remove(pie_chart_filename)

print(f"Exported demand analysis, line plots, and charts to {excel_filename}")














# import os
# import pandas as pd
# import matplotlib.pyplot as plt
# from datetime import date
# import math

# # Step 1: Data Collection
# sales_data = pd.read_csv('developed_data/createdData.csv')

# # Step 2: Analyze Weekly Sales Data
# sales_data['Date'] = pd.to_datetime(sales_data['Date'])
# sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()

# # Step 3: Analyze Weekly Sales Data
# sales_data['Week'] = sales_data['Date'].dt.isocalendar().week
# weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# # Step 4: Calculate Demand and EOQ
# demand_per_week = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
# demand_per_week = demand_per_week.rename(columns={'Quantity': 'Demand'})

# # Calculate EOQ for each row
# demand_per_week['EOQ'] = demand_per_week.apply(lambda row: round(math.sqrt((2 * 500 * row['Demand']) / 100), 2), axis=1)  # Assuming carrying cost is $500 and ordering cost is $100

# # Step 5: Export Demand and EOQ to Excel
# excel_folder = 'inventory_analysis'
# os.makedirs(excel_folder, exist_ok=True)  # Create the inventory_analysis directory
# excel_filename = f"{excel_folder}/inventory_analysis_{date.today().strftime('%Y-%m-%d')}.xlsx"
# excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
# demand_per_week.to_excel(excel_writer, sheet_name='Demand and EOQ', index=False)

# # Step 6: Create Line Plots for Each Product and Save them as Images
# os.makedirs('line_plots', exist_ok=True)  # Create the line_plots directory

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

#     # Save the line plot as an image
#     image_filename = f"line_plots/{product}_line_plot.png"
#     plt.savefig(image_filename)
#     plt.close()

#     # Add the line plot image to the Excel file
#     worksheet_name = f"{product} Line Plot"
#     worksheet = excel_writer.book.add_worksheet(worksheet_name)
#     worksheet.insert_image('A1', image_filename)

# # Step 7: Create a Pie Chart for Total Sales Distribution and Save it as an Image
# total_sales = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
# product_sales_total = total_sales.groupby('Product Name')['Quantity'].sum()
# plt.figure()
# plt.pie(product_sales_total, labels=product_sales_total.index, autopct='%1.1f%%')
# plt.title("Total Sales Distribution")

# # Save the pie chart as an image
# pie_chart_filename = "total_sales_pie_chart.png"
# plt.savefig(pie_chart_filename)
# plt.close()

# # Add the pie chart image to the Excel file
# worksheet_name = "Total Sales Pie Chart"
# worksheet = excel_writer.book.add_worksheet(worksheet_name)
# worksheet.insert_image('A1', pie_chart_filename)

# # Step 8: Calculate Increase and Decrease in Demand and Save in Separate Worksheets
# increase_sheet = excel_writer.book.add_worksheet('Increase in Demand')
# increase_sheet.write('A1', 'Product')
# increase_sheet.write('B1', 'Week')
# increase_sheet.write('C1', 'Comparison Week')
# increase_sheet.write('D1', 'Increase')

# decrease_sheet = excel_writer.book.add_worksheet('Decrease in Demand')
# decrease_sheet.write('A1', 'Product')
# decrease_sheet.write('B1', 'Week')
# decrease_sheet.write('C1', 'Comparison Week')
# decrease_sheet.write('D1', 'Decrease')

# demand_per_week = demand_per_week.sort_values(['Product Name', 'Week'])  # Sort the DataFrame by 'Product Name' and 'Week'

# for product in demand_per_week['Product Name'].unique():
#     product_demand = demand_per_week[demand_per_week['Product Name'] == product]

#     for i in range(1, len(product_demand)):
#         current_week_demand = product_demand.iloc[i]['Demand']
#         previous_week_demand = product_demand.iloc[i-1]['Demand']
#         percentage_change = (current_week_demand - previous_week_demand) / previous_week_demand * 100

#         if current_week_demand > previous_week_demand:
#             increase_sheet.write(i, 0, product)
#             increase_sheet.write(i, 1, product_demand.iloc[i]['Week'])
#             increase_sheet.write(i, 2, product_demand.iloc[i-1]['Week'])
#             increase_sheet.write(i, 3, current_week_demand - previous_week_demand)
#         elif current_week_demand < previous_week_demand:
#             decrease_sheet.write(i, 0, product)
#             decrease_sheet.write(i, 1, product_demand.iloc[i]['Week'])
#             decrease_sheet.write(i, 2, product_demand.iloc[i-1]['Week'])
#             decrease_sheet.write(i, 3, previous_week_demand - current_week_demand)

# # Step 9: Save and Close the Excel File
# excel_writer._save()
# excel_writer.close()

# # Delete the line plot images
# for product in weekly_sales['Product Name'].unique():
#     image_filename = f"line_plots/{product}_line_plot.png"
#     os.remove(image_filename)

# # Delete the pie chart image
# os.remove(pie_chart_filename)

# print(f"Exported demand analysis, line plots, and charts to {excel_filename}")







# import os
# import pandas as pd
# import matplotlib.pyplot as plt
# from datetime import date
# import math

# # Step 1: Data Collection
# sales_data = pd.read_csv('developed_data/createdData.csv')

# # Step 2: Analyze Weekly Sales Data
# sales_data['Date'] = pd.to_datetime(sales_data['Date'])
# sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()

# # Step 3: Analyze Weekly Sales Data
# weekly_sales = sales_data.groupby(['Product Name', pd.Grouper(key='Date', freq='W-MON')])['Quantity'].sum().reset_index()

# # Step 4: Calculate Demand and EOQ
# demand_per_week = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
# demand_per_week = demand_per_week.rename(columns={'Quantity': 'Demand'})

# # Calculate increase/decrease in demand and percentage change
# current_week = weekly_sales['Date'].dt.isocalendar().week.max()
# last_week_demand = demand_per_week.copy()
# last_week_demand['Week'] = current_week - 1
# demand_per_week = demand_per_week.merge(last_week_demand[['Product Name', 'Demand', 'Week']], on='Product Name', suffixes=('', '_last_week'))
# demand_per_week['Increase/Decrease'] = demand_per_week['Demand'] - demand_per_week['Demand_last_week']
# demand_per_week['Percentage Change'] = (demand_per_week['Increase/Decrease'] / demand_per_week['Demand_last_week']) * 100

# # Step 5: Export Demand and EOQ to Excel
# excel_folder = 'inventory_analysis'
# os.makedirs(excel_folder, exist_ok=True)  # Create the inventory_analysis directory
# excel_filename = f"{excel_folder}/inventory_analysis_{date.today()}.xlsx"
# excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
# demand_per_week.to_excel(excel_writer, sheet_name='Demand and EOQ', index=False)

# # Step 6: Create Line Plots for Each Product and Save them as Images
# os.makedirs('line_plots', exist_ok=True)  # Create the line_plots directory

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

#     # Save the line plot as an image
#     image_filename = f"line_plots/{product}_line_plot.png"
#     plt.savefig(image_filename)
#     plt.close()

#     # Add the line plot image to the Excel file
#     worksheet_name = f"{product} Line Plot"
#     worksheet = excel_writer.book.add_worksheet(worksheet_name)
#     worksheet.insert_image('A1', image_filename)

# # Step 7: Create a Pie Chart for Total Sales Distribution and Save it as an Image
# total_sales = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
# product_sales_total = total_sales.groupby('Product Name')['Quantity'].sum()
# plt.figure()
# plt.pie(product_sales_total, labels=product_sales_total.index, autopct='%1.1f%%')
# plt.title("Total Sales Distribution")
# plt.axis('equal')

# # Save the pie chart as an image
# pie_chart_filename = "total_sales_pie_chart.png"
# plt.savefig(pie_chart_filename)
# plt.close()

# # Add the pie chart image to the Excel file
# worksheet_name = "Total Sales Pie Chart"
# worksheet = excel_writer.book.add_worksheet(worksheet_name)
# worksheet.insert_image('A1', pie_chart_filename)

# # Step 8: Save and Close the Excel File
# excel_writer._save()
# excel_writer.close()

# # Delete the line plot images
# for product in weekly_sales['Product Name'].unique():
#     image_filename = f"line_plots/{product}_line_plot.png"
#     os.remove(image_filename)

# # Delete the pie chart image
# os.remove(pie_chart_filename)

# print(f"Exported demand analysis, line plots, and charts to {excel_filename}")








# import os
# import pandas as pd
# import matplotlib.pyplot as plt
# from datetime import date
# import math

# # Step 1: Data Collection
# sales_data = pd.read_csv('developed_data/createdData.csv')

# # Step 2: Analyze Weekly Sales Data
# sales_data['Date'] = pd.to_datetime(sales_data['Date'])
# sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()

# # Step 3: Analyze Weekly Sales Data
# sales_data['Week'] = sales_data['Date'].dt.isocalendar().week
# weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# # Step 4: Calculate Demand and EOQ
# demand_per_week = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
# demand_per_week = demand_per_week.rename(columns={'Quantity': 'Demand'})

# # Calculate EOQ for each row
# demand_per_week['EOQ'] = demand_per_week.apply(lambda row: round(math.sqrt((2 * 500 * row['Demand']) / 100), 2), axis=1)  # Assuming carrying cost is $500 and ordering cost is $100

# # Step 5: Calculate Increase and Decrease in Demand
# last_week_demand = demand_per_week.iloc[-2]
# current_week_demand = demand_per_week.iloc[-1]
# demand_change = pd.DataFrame({
#     'Product Name': [current_week_demand['Product Name']],
#     'Week_x': [current_week_demand['Week']],
#     'Quantity_x': [current_week_demand['Demand']],
#     'Week_y': [last_week_demand['Week']],
#     'Quantity_y': [last_week_demand['Demand']]
# })
# demand_change['Change'] = demand_change['Quantity_x'] - demand_change['Quantity_y']
# demand_change['Change(%)'] = (demand_change['Change'] / demand_change['Quantity_y']) * 100

# # Separate increase and decrease in demand
# increase_demand = demand_change[demand_change['Change'] > 0]
# decrease_demand = demand_change[demand_change['Change'] < 0]

# # Step 6: Export Increase and Decrease in Demand to Excel
# excel_folder = 'inventory_analysis'
# os.makedirs(excel_folder, exist_ok=True)  # Create the inventory_analysis directory
# excel_filename = f"{excel_folder}/inventory_analysis_{date.today()}.xlsx"
# excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
# increase_demand.to_excel(excel_writer, sheet_name='Increase in Demand', index=False)
# decrease_demand.to_excel(excel_writer, sheet_name='Decrease in Demand', index=False)

# # Step 7: Export Demand and EOQ to Excel
# demand_per_week.to_excel(excel_writer, sheet_name='Demand and EOQ', index=False)


# # Step 8: Create Line Plots for Each Product and Save them as Images
# os.makedirs('line_plots', exist_ok=True)  # Create the line_plots directory

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

#     # Save the line plot as an image
#     image_filename = f"line_plots/{product}_line_plot.png"
#     plt.savefig(image_filename)
#     plt.close()

#     # Add the line plot image to the Excel file
#     worksheet_name = f"{product} Line Plot"
#     worksheet = excel_writer.book.add_worksheet(worksheet_name)
#     worksheet.insert_image('A1', image_filename)

# # Step 9: Create a Pie Chart for Total Sales Distribution and Save it as an Image
# total_sales = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
# product_sales_total = total_sales.groupby('Product Name')['Quantity'].sum()
# plt.figure()
# plt.pie(product_sales_total, labels=product_sales_total.index, autopct='%1.1f%%')
# plt.title("Total Sales Distribution")

# # Save the pie chart as an image
# pie_chart_filename = "total_sales_pie_chart.png"
# plt.savefig(pie_chart_filename)
# plt.close()

# # Add the pie chart image to the Excel file
# worksheet_name = "Total Sales Pie Chart"
# worksheet = excel_writer.book.add_worksheet(worksheet_name)
# worksheet.insert_image('A1', pie_chart_filename)

# # Step 8: Save and Close the Excel File
# excel_writer._save()
# excel_writer.close()

# # Delete the line plot images
# for product in weekly_sales['Product Name'].unique():
#     image_filename = f"line_plots/{product}_line_plot.png"
#     os.remove(image_filename)

# # Delete the pie chart image
# os.remove(pie_chart_filename)

# print(f"Exported demand analysis, line plots, and charts to {excel_filename}")















# import os
# import pandas as pd
# import matplotlib.pyplot as plt
# from datetime import date
# import math

# # Step 1: Data Collection
# sales_data = pd.read_csv('developed_data/createdData.csv')

# # Step 2: Analyze Weekly Sales Data
# sales_data['Date'] = pd.to_datetime(sales_data['Date'])
# sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()

# # Step 3: Analyze Weekly Sales Data
# sales_data['Week'] = sales_data['Date'].dt.isocalendar().week
# weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# # Step 4: Calculate Demand and EOQ
# demand_per_week = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
# demand_per_week = demand_per_week.rename(columns={'Quantity': 'Demand'})

# # Calculate EOQ for each row
# demand_per_week['EOQ'] = demand_per_week.apply(lambda row: round(math.sqrt((2 * 500 * row['Demand']) / 100), 2), axis=1)  # Assuming carrying cost is $500 and ordering cost is $100

# # Step 5: Export Demand and EOQ to Excel
# excel_folder = 'inventory_analysis'
# os.makedirs(excel_folder, exist_ok=True)  # Create the inventory_analysis directory
# excel_filename = f"{excel_folder}/inventory_analysis_{date.today()}.xlsx"
# excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
# demand_per_week.to_excel(excel_writer, sheet_name='Demand and EOQ', index=False)


# # Step 6: Create Line Plots for Each Product and Save them as Images
# os.makedirs('line_plots', exist_ok=True)  # Create the line_plots directory

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

#     # Save the line plot as an image
#     image_filename = f"line_plots/{product}_line_plot.png"
#     plt.savefig(image_filename)
#     plt.close()

#     # Add the line plot image to the Excel file
#     worksheet_name = f"{product} Line Plot"
#     worksheet = excel_writer.book.add_worksheet(worksheet_name)
#     worksheet.insert_image('A1', image_filename)

# # Step 7: Create a Pie Chart for Total Sales Distribution and Save it as an Image
# total_sales = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
# product_sales_total = total_sales.groupby('Product Name')['Quantity'].sum()
# plt.figure()
# plt.pie(product_sales_total, labels=product_sales_total.index, autopct='%1.1f%%')
# plt.title("Total Sales Distribution")

# # Save the pie chart as an image
# pie_chart_filename = "total_sales_pie_chart.png"
# plt.savefig(pie_chart_filename)
# plt.close()

# # Add the pie chart image to the Excel file
# worksheet_name = "Total Sales Pie Chart"
# worksheet = excel_writer.book.add_worksheet(worksheet_name)
# worksheet.insert_image('A1', pie_chart_filename)

# # Step 8: Save and Close the Excel File
# excel_writer._save()
# excel_writer.close()

# # Delete the line plot images
# for product in weekly_sales['Product Name'].unique():
#     image_filename = f"line_plots/{product}_line_plot.png"
#     os.remove(image_filename)

# # Delete the pie chart image
# os.remove(pie_chart_filename)

# print(f"Exported demand analysis, line plots, and charts to {excel_filename}")
