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
sales_data = pd.read_csv('deveopedData/analysisDataWithRegion_2023-06-05.csv')

# Step 2: Analyze Weekly Sales Data
sales_data['Date'] = pd.to_datetime(sales_data['Date'])
sales_data = sales_data.groupby(['Date', 'Product Name', 'Region'])['Quantity'].sum().reset_index()

# Step 3: Analyze Weekly Sales Data
sales_data['Week'] = sales_data['Date'].dt.isocalendar().week
weekly_sales = sales_data.groupby(['Product Name', 'Region', 'Week', 'Date'])['Quantity'].sum().reset_index()

# Step 4: Find the Demanded Products for the Last Two Weeks
last_two_weeks = weekly_sales['Week'].unique()[-2:]
product_demand = weekly_sales.groupby(['Product Name', 'Region', 'Week'])['Quantity'].sum().reset_index()
product_demand = product_demand[product_demand['Week'].isin(last_two_weeks)]
most_demanded_products = product_demand.groupby(['Region', 'Product Name'])['Quantity'].sum().reset_index().sort_values(
    by=['Region', 'Quantity'], ascending=[True, False])

# Step 5: Calculate Increase and Decrease in Demand
last_week_demand = product_demand[product_demand['Week'] == last_two_weeks[0]]
current_week_demand = product_demand[product_demand['Week'] == last_two_weeks[1]]
demand_change = pd.merge(current_week_demand, last_week_demand, on=['Product Name', 'Region'], how='left')
demand_change['Change'] = demand_change['Quantity_x'] - demand_change['Quantity_y']
demand_change['Change(%)'] = (demand_change['Change'] / demand_change['Quantity_y']) * 100

# Separate increase and decrease in demand
increase_demand = demand_change[demand_change['Change'] > 0]
decrease_demand = demand_change[demand_change['Change'] < 0]

# Step 8: Create a Pandas Excel Writer
excel_folder = 'demanded_products'
os.makedirs(excel_folder, exist_ok=True)  # Create the demanded_products directory
excel_filename = f"{excel_folder}/demanded_products_with_graph{date.today().strftime('%Y-%m-%d')}.xlsx"
excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# Step 9: Export Increase and Decrease in Demand to Excel
increase_demand.to_excel(excel_writer, sheet_name='Increase in Demand', index=False)
decrease_demand.to_excel(excel_writer, sheet_name='Decrease in Demand', index=False)

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
product_sales_total = weekly_sales.groupby(['Product Name', 'Region'])['Quantity'].sum().reset_index()
product_sales_total = product_sales_total.rename(columns={'Quantity': 'Total Quantity'})
product_sales_total['Demand Rank'] = product_sales_total.groupby('Region')['Total Quantity'].rank(ascending=False)

for algorithm_index, algorithm in enumerate(algorithms):
    algorithm.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
    algorithm_name = algorithm_names[algorithm_index]
    if algorithm_name == 'GaussianMix':
        # For Gaussian Mixture, we can calculate the cluster based on the highest probability
        product_sales_total['Cluster'] = algorithm.predict_proba(product_sales_total[['Total Quantity', 'Demand Rank']].values).argmax(axis=1)
    else:
        product_sales_total['Cluster'] = algorithm.labels_
    # most_demanded_products = product_sales_total.groupby(['Region', 'Product Name'])['Cluster'].sum().reset_index().sort_values(
        # by=['Region', 'Cluster'], ascending=[True, False])
    
    # most_demanded_products = product_sales_total.groupby(['Region', 'Product Name'])['Total Quantity'].sum().reset_index().sort_values(
    # by=['Region', 'Total Quantity'], ascending=[True, False])


    most_demanded_products = product_sales_total.groupby(['Region', 'Product Name'])['Total Quantity'].sum().reset_index().sort_values(
    by=['Region', 'Total Quantity'], ascending=[True, False])
    most_demanded_products = most_demanded_products.rename(columns={'Total Quantity': 'Quantity'})


    sheet_name = f"Algorithm_{algorithm_index}_{algorithm_name}"
    most_demanded_products.to_excel(excel_writer, sheet_name=sheet_name, index=False)

# Step 13: Create Line Plots for Each Product and Region and Save them as Images
os.makedirs('line_plots', exist_ok=True)  # Create the line_plots directory

row_counter = 0  # Counter for Excel rows

for region in weekly_sales['Region'].unique():
    for product in weekly_sales['Product Name'].unique():
        product_sales = weekly_sales[(weekly_sales['Product Name'] == product) & (weekly_sales['Region'] == region)]

        # Prepare data for plotting
        dates = product_sales['Date']
        quantities = product_sales['Quantity']

        # Create a line plot for the sales data
        plt.figure()
        plt.plot(dates, quantities, marker='o')
        plt.title(f'Sales Data for {product} ({region})')
        plt.xlabel('Date')
        plt.ylabel('Quantity Sold')
        plt.xticks(rotation=45)

        # Save the line plot as an image
        image_filename = f"line_plots/{product}_{region}_line_plot.png"
        plt.savefig(image_filename)
        plt.close()

        # Add the line plot image to the Excel file
        worksheet_name = f"{product} Line Plot ({region})"
        worksheet = excel_writer.book.add_worksheet(worksheet_name)
        row_counter += 20  # Increment row counter for each line plot
        worksheet.insert_image(f'A{row_counter}', image_filename)


        # color_index = (color_index + 1) % len(colors)

# Step 14: Create a Bar Chart for Most Demanded Products and Store in Excel
worksheet_graphs = excel_writer.book.add_worksheet('Graphs')

# Prepare the colors for the bar chart
colors = plt.rcParams['axes.prop_cycle'].by_key()['color']
color_index = 0

# Iterate over each region and plot the most demanded products
for region in most_demanded_products['Region'].unique():
    region_products = most_demanded_products[most_demanded_products['Region'] == region]['Product Name']
    product_quantities = most_demanded_products[most_demanded_products['Region'] == region]['Quantity']

    # Create a bar chart for the most demanded products
    plt.bar(region_products, product_quantities, color=colors[color_index])
    plt.title(f'Most Demanded Products in {region}')
    plt.xlabel('Product Name')
    plt.ylabel('Quantity')

    # Save the bar chart as an image
    image_filename = f"bar_charts/{region}_most_demanded_products.png"
    plt.savefig(image_filename)
    plt.close()

    # Add the bar chart image to the Excel file
    worksheet_graphs.insert_image(f'A{color_index*15+1}', image_filename)

    color_index = (color_index + 1) % len(colors)

# Step 15: Save and Close the Excel File
excel_writer._save()
excel_writer.close()

print(f"Exported most demanded products, line plots, and graphs to {excel_filename}")
