import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date
from sklearn.cluster import KMeans, DBSCAN, AgglomerativeClustering, Birch
from sklearn.mixture import GaussianMixture

# Step 1: Data Collection
sales_data = pd.read_csv('./deveopedData/createdData.csv')

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

# Step 5: Apply Machine Learning Optimization Algorithms
product_sales_total = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
product_sales_total = product_sales_total.rename(columns={'Quantity': 'Total Quantity'})
product_sales_total['Demand Rank'] = product_sales_total['Total Quantity'].rank(ascending=False)

algorithms = [
    KMeans(n_clusters=3, random_state=42),
    GaussianMixture(n_components=3, random_state=42),
    DBSCAN(eps=3, min_samples=2),
    AgglomerativeClustering(n_clusters=3),
    Birch(n_clusters=3)
]

print("Optimization Algorithms and Most Demanded Products:")
for algorithm in algorithms:
    algorithm.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
    algorithm_name = algorithm.__class__.__name__
    if algorithm_name == 'GaussianMixture':
        # For Gaussian Mixture, we can calculate the cluster based on the highest probability
        product_sales_total['Cluster'] = algorithm.predict_proba(product_sales_total[['Total Quantity', 'Demand Rank']].values).argmax(axis=1)
    else:
        product_sales_total['Cluster'] = algorithm.labels_
    most_demanded_products = product_sales_total.groupby('Product Name')['Cluster'].sum().reset_index().sort_values(
        by='Cluster', ascending=False)
    print(f"\nOptimization Algorithm: {algorithm_name}\n")
    print("Most Demanded Products:")
    print(most_demanded_products.to_string(index=False))

# Step 6: Create a Pandas Excel Writer
excel_folder = 'demanded_products'
os.makedirs(excel_folder, exist_ok=True)
excel_filename = f"{excel_folder}/demanded_products_{date.today()}.xlsx"
excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# Step 7: Export Most Demanded Products to Excel
most_demanded_products.to_excel(excel_writer, sheet_name='Demanded Products', index=False)

# Step 8: Create Line Plots for Each Product and Save them as Images
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

# Step 9: Create a Pie Chart for Total Sales Distribution and Save it as an Image
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

# Step 10: Save and Close the Excel File
excel_writer._save()
excel_writer.close()

print("Data analysis and visualization completed.")















# import os
# import pandas as pd
# import matplotlib.pyplot as plt
# from datetime import date
# from sklearn.cluster import KMeans, DBSCAN, AgglomerativeClustering, Birch
# from sklearn.mixture import GaussianMixture

# # Step 1: Data Collection
# sales_data = pd.read_csv('sales_data1.csv')


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

# # Step 5: Apply Machine Learning Optimization Algorithms
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

# # Step 6: Create a Pandas Excel Writer
# excel_folder = 'demanded_products'
# excel_filename = f"{excel_folder}/demanded_products_{date.today()}.xlsx"
# excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# # Step 7: Export Most Demanded Products to Excel
# most_demanded_products.to_excel(excel_writer, sheet_name='Demanded Products', index=False)

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

# # Step 9: Save and Close the Excel File
# # Create a new sheet for algorithm comparison
# # Rest of the code...

# # Step 10: Compare Optimization Algorithms
# algorithm_comparison = pd.DataFrame(columns=['Algorithm', 'Most Demanded Products'])

# # K-Means
# kmeans = KMeans(n_clusters=3)
# kmeans.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
# kmeans_labels = kmeans.labels_
# kmeans_most_demanded = product_sales_total[['Product Name', 'Total Quantity']].copy()
# kmeans_most_demanded['Cluster'] = kmeans_labels
# kmeans_result = pd.DataFrame({'Algorithm': 'KMeans', 'Most Demanded Products': kmeans_most_demanded.to_string(index=False)}, index=[0])
# algorithm_comparison = algorithm_comparison.append(kmeans_result, ignore_index=True)

# # DBSCAN
# dbscan = DBSCAN(eps=0.5, min_samples=5)
# dbscan.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
# dbscan_labels = dbscan.labels_
# dbscan_most_demanded = product_sales_total[['Product Name', 'Total Quantity']].copy()
# dbscan_most_demanded['Cluster'] = dbscan_labels
# dbscan_result = pd.DataFrame({'Algorithm': 'DBSCAN', 'Most Demanded Products': dbscan_most_demanded.to_string(index=False)}, index=[0])
# algorithm_comparison = algorithm_comparison.append(dbscan_result, ignore_index=True)

# # Agglomerative Clustering
# agglomerative = AgglomerativeClustering(n_clusters=3)
# agglomerative.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
# agglomerative_labels = agglomerative.labels_
# agglomerative_most_demanded = product_sales_total[['Product Name', 'Total Quantity']].copy()
# agglomerative_most_demanded['Cluster'] = agglomerative_labels
# agglomerative_result = pd.DataFrame({'Algorithm': 'Agglomerative Clustering', 'Most Demanded Products': agglomerative_most_demanded.to_string(index=False)}, index=[0])
# algorithm_comparison = algorithm_comparison.append(agglomerative_result, ignore_index=True)

# # Birch
# birch = Birch(n_clusters=3)
# birch.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
# birch_labels = birch.labels_
# birch_most_demanded = product_sales_total[['Product Name', 'Total Quantity']].copy()
# birch_most_demanded['Cluster'] = birch_labels
# birch_result = pd.DataFrame({'Algorithm': 'Birch', 'Most Demanded Products': birch_most_demanded.to_string(index=False)}, index=[0])
# algorithm_comparison = algorithm_comparison.append(birch_result, ignore_index=True)

# # Gaussian Mixture Model
# gmm = GaussianMixture(n_components=3)
# gmm.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
# gmm_labels = gmm.predict_proba(product_sales_total[['Total Quantity', 'Demand Rank']].values).argmax(axis=1)
# gmm_most_demanded = product_sales_total[['Product Name', 'Total Quantity']].copy()
# gmm_most_demanded['Cluster'] = gmm_labels
# gmm_result = pd.DataFrame({'Algorithm': 'Gaussian Mixture Model', 'Most Demanded Products': gmm_most_demanded.to_string(index=False)}, index=[0])
# algorithm_comparison = algorithm_comparison.append(gmm_result, ignore_index=True)

# # Step 11: Save Algorithm Comparison to Excel
# excel_writer = pd.ExcelWriter('algorithm_comparison.xlsx', engine='xlsxwriter')
# algorithm_comparison.to_excel(excel_writer, sheet_name='Algorithm Comparison', index=False)

# # Save and close the Excel file for algorithm comparison
# excel_writer.save()
# excel_writer.close()

# # Rest of the code...



# print("Algorithm comparison saved to algorithm_comparison.xlsx")
