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

# Step 5: Calculate Increase in Demand
last_week_demand = product_demand[product_demand['Week'] == last_two_weeks[0]]
current_week_demand = product_demand[product_demand['Week'] == last_two_weeks[1]]
demand_increase = pd.merge(current_week_demand, last_week_demand, on='Product Name', how='left')
demand_increase['Increase'] = demand_increase['Quantity_x'] - demand_increase['Quantity_y']
demand_increase['Increase(%)'] = (demand_increase['Increase'] / demand_increase['Quantity_y']) * 100

# Print Increase in Demand
# print("\nIncrease in Demand:")
# Print Increase in Demand
print("\nIncrease in Demand:")
print(demand_increase[['Product Name', 'Week_x', 'Quantity_x', 'Week_y', 'Quantity_y', 'Increase', 'Increase(%)']].to_string(index=False))

# Step 8: Create a Pandas Excel Writer
excel_folder = 'demanded_products'
os.makedirs(excel_folder, exist_ok=True)  # Create the demanded_products directory
excel_filename = f"{excel_folder}/demanded_products_{date.today()}.xlsx"
excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# Step 9: Export Increase in Demand to Excel
demand_increase.to_excel(excel_writer, sheet_name='Increase in Demand', index=False)

# Step 10: Define the list of algorithms
algorithms = [
    KMeans(n_clusters=3, random_state=42),
    GaussianMixture(n_components=3, random_state=42),
    DBSCAN(eps=3, min_samples=2),
    AgglomerativeClustering(n_clusters=3),
    Birch(n_clusters=3)
]

# Step 11: Export Most Demanded Products to Excel for each algorithm
algorithm_names = {
    0: 'KMeans',
    1: 'GaussianMix',
    2: 'DBSCAN',
    3: 'AggClustering',
    4: 'Birch'
}

# Step 12: Calculate product_sales_total
product_sales_total = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
product_sales_total = product_sales_total.rename(columns={'Quantity': 'Total Quantity'})
product_sales_total['Demand Rank'] = product_sales_total['Total Quantity'].rank(ascending=False)

for algorithm_index, algorithm in enumerate(algorithms):
    algorithm.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
    algorithm_name = algorithm_names[algorithm_index]
    if algorithm_name == 'GaussianMix':
        # For Gaussian Mixture, we can calculate the cluster based on the highest probability
        product_sales_total['Cluster'] = algorithm.predict_proba(product_sales_total[['Total Quantity', 'Demand Rank']].values).argmax(axis=1)
    else:
        product_sales_total['Cluster'] = algorithm.labels_
    most_demanded_products = product_sales_total.groupby('Product Name')['Cluster'].sum().reset_index().sort_values(
        by='Cluster', ascending=False)
    sheet_name = f"Algorithm_{algorithm_index}_{algorithm_name}"
    most_demanded_products.to_excel(excel_writer, sheet_name=sheet_name, index=False)


# Step 11: Create Line Plots for Each Product and Save them as Images
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

print(f"Exported most demanded products, line plots, and charts to {excel_filename}")



# # Rest of the code...


# # Step 8: Create a Pandas Excel Writer
# excel_folder = 'demanded_products'
# excel_filename = f"{excel_folder}/demanded_products_{date.today()}.xlsx"
# excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# # Step 6: Export Most Demanded Products to Excel
# most_demanded_products.to_excel(excel_writer, sheet_name='Demanded Products', index=False)

# # Step 6: Apply Machine Learning Optimization Algorithms
# product_sales_total = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
# product_sales_total = product_sales_total.rename(columns={'Quantity': 'Total Quantity'})
# product_sales_total['Demand Rank'] = product_sales_total['Total Quantity'].rank(ascending=False)

# algorithms = [
#     KMeans(n_clusters=3, random_state=42),
#     GaussianMixture(n_components=3, random_state=42),
#     DBSCAN(eps=3, min_samples=2),
#     AgglomerativeClustering(n_clusters=3),
#     Birch(n_clusters=3)
# ]

# # Create a dictionary to map algorithm indices to shorter names
# algorithm_names = {
#     0: 'KMeans',
#     1: 'GaussianMix',
#     2: 'DBSCAN',
#     3: 'AggClustering',
#     4: 'Birch'
# }

# print("Optimization Algorithms and Most Demanded Products:")
# for algorithm_index, algorithm in enumerate(algorithms):
#     algorithm.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
#     algorithm_name = algorithm_names[algorithm_index]
#     if algorithm_name == 'GaussianMix':
#         # For Gaussian Mixture, we can calculate the cluster based on the highest probability
#         product_sales_total['Cluster'] = algorithm.predict_proba(product_sales_total[['Total Quantity', 'Demand Rank']].values).argmax(axis=1)
#     else:
#         product_sales_total['Cluster'] = algorithm.labels_
#     most_demanded_products = product_sales_total.groupby('Product Name')['Cluster'].sum().reset_index().sort_values(
#         by='Cluster', ascending=False)
#     # print(f"\nOptimization Algorithm: {algorithm_name}\n")
#     # print("Most Demanded Products:")
#     # print(most_demanded_products.to_string(index=False))

#     # Create a truncated worksheet name
#     sheet_name = f"Algorithm_{algorithm_index}_{algorithm_name[:20]}"

#     # Save the most demanded products to the Excel file
#     most_demanded_products.to_excel(excel_writer, sheet_name=sheet_name, index=False)




# # Step 7: Create Line Plots for Each Product and Save them as Images
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

# # Step 8: Create a Pie Chart for Total Sales Distribution and Save it as an Image
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

# # Step 9: Save and Close the Excel File
# excel_writer._save()
# excel_writer.close()

# # Delete the line plot images
# for product in weekly_sales['Product Name'].unique():
#     image_filename = f"line_plots/{product}_line_plot.png"
#     os.remove(image_filename)

# # Delete the pie chart image
# os.remove(pie_chart_filename)

# print(f"Exported most demanded products and charts to {excel_filename}")













# import os
# import pandas as pd
# import matplotlib.pyplot as plt
# from datetime import date
# from sklearn.cluster import KMeans
# from sklearn.mixture import GaussianMixture
# from sklearn.cluster import DBSCAN
# from sklearn.cluster import AgglomerativeClustering
# from sklearn.cluster import Birch

# # Step 1: Data Collection
# sales_data = pd.read_csv('deveopedData/chatgptData.csv')

# # Step 2: Analyze Weekly Sales Data
# sales_data['Date'] = pd.to_datetime(sales_data['Date'])
# sales_data = sales_data.groupby(['Date', 'Product Name'])['Quantity'].sum().reset_index()

# # Step 3: Analyze Weekly Sales Data
# sales_data['Week'] = sales_data['Date'].dt.isocalendar().week
# weekly_sales = sales_data.groupby(['Product Name', 'Week', 'Date'])['Quantity'].sum().reset_index()

# # Step 4: Find the Demanded Products for the Last Two Weeks
# last_two_weeks = weekly_sales['Week'].unique()[-2:]
# product_demand = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
# product_demand = product_demand[product_demand['Week'].isin(last_two_weeks)]
# most_demanded_products = product_demand.groupby('Product Name')['Quantity'].sum().reset_index().sort_values(
#     by='Quantity', ascending=False)

# # Step 5: Calculate Increase in Demand
# last_week_demand = product_demand[product_demand['Week'] == last_two_weeks[0]]
# current_week_demand = product_demand[product_demand['Week'] == last_two_weeks[1]]
# demand_increase = pd.merge(current_week_demand, last_week_demand, on='Product Name', how='left')
# demand_increase['Increase'] = demand_increase['Quantity_x'] - demand_increase['Quantity_y']
# demand_increase['Increase(%)'] = (demand_increase['Increase'] / demand_increase['Quantity_y']) * 100

# # Step 6: Apply Machine Learning Optimization Algorithms
# product_sales_total = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
# product_sales_total = product_sales_total.rename(columns={'Quantity': 'Total Quantity'})
# product_sales_total['Demand Rank'] = product_sales_total['Total Quantity'].rank(ascending=False)

# algorithms = [
#     KMeans(n_clusters=3, random_state=42),
#     GaussianMixture(n_components=3, random_state=42),
#     DBSCAN(eps=3, min_samples=2),
#     AgglomerativeClustering(n_clusters=3),
#     Birch(n_clusters=3)
# ]

# print("Optimization Algorithms and Most Demanded Products:")
# for algorithm in algorithms:
#     algorithm.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
#     algorithm_name = algorithm.__class__.__name__
#     if algorithm_name == 'GaussianMixture':
#         # For Gaussian Mixture, we can calculate the cluster based on the highest probability
#         product_sales_total['Cluster'] = algorithm.predict_proba(product_sales_total[['Total Quantity', 'Demand Rank']].values).argmax(axis=1)
#     else:
#         product_sales_total['Cluster'] = algorithm.labels_
#     most_demanded_products = product_sales_total.groupby('Product Name')['Cluster'].sum().reset_index().sort_values(
#         by='Cluster', ascending=False)
#     print(f"\nOptimization Algorithm: {algorithm_name}\n")
#     print("Most Demanded Products:")
#     print(most_demanded_products.to_string(index=False))

# # Step 7: Output Increase in Demand
# print("\nIncrease in Demand:")
# print(demand_increase[['Product Name', 'Increase', 'Increase(%)']].to_string(index=False))

# # Step 8: Create a Pandas Excel Writer
# excel_folder = 'demanded_products'
# excel_filename = f"{excel_folder}/demanded_products_{date.today()}.xlsx"
# excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# # Step 6: Export Most Demanded Products to Excel
# most_demanded_products.to_excel(excel_writer, sheet_name='Demanded Products', index=False)

# # Step 7: Create Line Plots for Each Product and Save them as Images
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

# # Step 8: Create a Pie Chart for Total Sales Distribution and Save it as an Image
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

# # Step 9: Save and Close the Excel File
# excel_writer._save()
# excel_writer.close()

# # Delete the line plot images
# for product in weekly_sales['Product Name'].unique():
#     image_filename = f"line_plots/{product}_line_plot.png"
#     os.remove(image_filename)

# # Delete the pie chart image
# os.remove(pie_chart_filename)

# print(f"Exported most demanded products and charts to {excel_filename}")
