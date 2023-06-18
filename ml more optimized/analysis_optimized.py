import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date, datetime
from sklearn.cluster import KMeans
from sklearn.mixture import GaussianMixture
from sklearn.cluster import DBSCAN
from sklearn.cluster import AgglomerativeClustering
from sklearn.cluster import Birch



def read_sales_data(file_path):
    sales_data = pd.read_excel(file_path)
    return sales_data


def analyze_weekly_sales(sales_data):
    sales_data['Date'] = pd.to_datetime(sales_data['Date'])
    sales_data['Week'] = sales_data['Date'].dt.isocalendar().week
    weekly_sales = sales_data.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
    return weekly_sales


def calculate_demand_change(product_demand):
    last_week = product_demand['Week'].max()
    last_two_weeks = [last_week - 1, last_week]
    last_week_demand = product_demand[product_demand['Week'] == last_two_weeks[0]]
    current_week_demand = product_demand[product_demand['Week'] == last_two_weeks[1]]
    demand_change = pd.merge(current_week_demand, last_week_demand, on='Product Name', how='left')
    demand_change['Change'] = demand_change['Quantity_x'] - demand_change['Quantity_y']
    demand_change['Change(%)'] = (demand_change['Change'] / demand_change['Quantity_y']) * 100

    increase_demand = demand_change[demand_change['Change'] > 0]
    decrease_demand = demand_change[demand_change['Change'] < 0]

    return increase_demand, decrease_demand




def find_most_demanded_products(weekly_sales):
    last_two_weeks = weekly_sales['Week'].unique()[-2:]
    product_demand = weekly_sales.groupby(['Product Name', 'Week'])['Quantity'].sum().reset_index()
    product_demand = product_demand[product_demand['Week'].isin(last_two_weeks)]
    most_demanded_products = product_demand.groupby('Product Name')['Quantity'].sum().reset_index().sort_values(
        by='Quantity', ascending=False)

    return most_demanded_products


# def calculate_demand_change(product_demand):
#     last_week_demand = product_demand[product_demand['Week'] == last_two_weeks[0]]
#     current_week_demand = product_demand[product_demand['Week'] == last_two_weeks[1]]
#     demand_change = pd.merge(current_week_demand, last_week_demand, on='Product Name', how='left')
#     demand_change['Change'] = demand_change['Quantity_x'] - demand_change['Quantity_y']
#     demand_change['Change(%)'] = (demand_change['Change'] / demand_change['Quantity_y']) * 100

#     increase_demand = demand_change[demand_change['Change'] > 0]
#     decrease_demand = demand_change[demand_change['Change'] < 0]

#     return increase_demand, decrease_demand


def export_demand_changes_to_excel(increase_demand, decrease_demand, excel_writer):
    increase_demand.to_excel(excel_writer, sheet_name='Increase in Demand', index=False)
    decrease_demand.to_excel(excel_writer, sheet_name='Decrease in Demand', index=False)


def run_clustering_algorithms(product_sales_total, algorithms, algorithm_names):
    for algorithm_index, algorithm in enumerate(algorithms):
        algorithm.fit(product_sales_total[['Total Quantity', 'Demand Rank']].values)
        algorithm_name = algorithm_names[algorithm_index]
        if algorithm_name == 'GaussianMix':
            product_sales_total['Cluster'] = algorithm.predict_proba(
                product_sales_total[['Total Quantity', 'Demand Rank']].values).argmax(axis=1)
        else:
            product_sales_total['Cluster'] = algorithm.labels_
        most_demanded_products = product_sales_total.groupby('Product Name')['Cluster'].sum().reset_index().sort_values(
            by='Cluster', ascending=False)
        sheet_name = f"Algorithm_{algorithm_index}_{algorithm_name}"
        most_demanded_products.to_excel(excel_writer, sheet_name=sheet_name, index=False)


def create_line_plots(weekly_sales, line_plots_folder):
    os.makedirs(line_plots_folder, exist_ok=True)

    for product in weekly_sales['Product Name'].unique():
        product_sales = weekly_sales[weekly_sales['Product Name'] == product]

        dates = product_sales['Date']
        quantities = product_sales['Quantity']

        plt.figure()
        plt.plot(dates, quantities, marker='o')
        plt.title(f'Sales Data for {product}')
        plt.xlabel('Date')
        plt.ylabel('Quantity Sold')
        plt.xticks(rotation=45)

        line_plot_filename = os.path.join(line_plots_folder, f"{product}_line_plot.png")
        plt.savefig(line_plot_filename)
        plt.close()


def create_pie_chart(total_sales, pie_chart_folder):
    os.makedirs(pie_chart_folder, exist_ok=True)

    product_sales_total = total_sales.groupby('Product Name')['Quantity'].sum()
    plt.figure()
    plt.pie(product_sales_total, labels=product_sales_total.index, autopct='%1.1f%%')
    plt.title("Total Sales Distribution")

    pie_chart_filename = os.path.join(pie_chart_folder, "total_sales_pie_chart.png")
    plt.savefig(pie_chart_filename)
    plt.close()


def export_graphs_to_excel(line_plots_folder, pie_chart_folder, excel_writer):
    worksheet = excel_writer.book.add_worksheet("Line Plots")

    for i, image_file in enumerate(os.listdir(line_plots_folder)):
        image_path = os.path.join(line_plots_folder, image_file)
        worksheet.insert_image(i * 10, 0, image_path)

    worksheet = excel_writer.book.add_worksheet("Pie Charts")
    worksheet.insert_image(0, 0, os.path.join(pie_chart_folder, "total_sales_pie_chart.png"))


today_date = datetime.now().strftime('%Y-%m-%d')
file_name = f'./developed_data/sales_data_{today_date}.xlsx'
# Step 1: Data Collection
sales_data = read_sales_data(file_name)

# Step 2: Analyze Weekly Sales Data
weekly_sales = analyze_weekly_sales(sales_data)

weekly_sales['Week'] = weekly_sales['SalesDate'].dt.week

# Step 3: Find the Demanded Products for the Last Two Weeks
most_demanded_products = find_most_demanded_products(weekly_sales)

# Step 4: Calculate Increase and Decrease in Demand
increase_demand, decrease_demand = calculate_demand_change(most_demanded_products)

# Step 5: Create a Pandas Excel Writer
excel_folder = 'demanded_products'
os.makedirs(excel_folder, exist_ok=True)
excel_filename = f"{excel_folder}/demanded_products_{date.today()}.xlsx"
excel_writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# Step 6: Export Increase and Decrease in Demand to Excel
export_demand_changes_to_excel(increase_demand, decrease_demand, excel_writer)

# Step 7: Define the list of algorithms
algorithms = [
    KMeans(n_clusters=3, random_state=42),
    GaussianMixture(n_components=3, random_state=42),
    DBSCAN(eps=3, min_samples=2),
    AgglomerativeClustering(n_clusters=3),
    Birch(n_clusters=3)
]

algorithm_names = {
    0: 'KMeans',
    1: 'GaussianMix',
    2: 'DBSCAN',
    3: 'AggClustering',
    4: 'Birch'
}

# Step 8: Calculate product_sales_total
product_sales_total = weekly_sales.groupby('Product Name')['Quantity'].sum().reset_index()
product_sales_total = product_sales_total.rename(columns={'Quantity': 'Total Quantity'})
product_sales_total['Demand Rank'] = product_sales_total['Total Quantity'].rank(ascending=False)

# Step 9: Run Clustering Algorithms
run_clustering_algorithms(product_sales_total, algorithms, algorithm_names)

# Step 10: Create Line Plots
line_plots_folder = 'line_plots'
create_line_plots(weekly_sales, line_plots_folder)

# Step 11: Create Pie Chart
pie_chart_folder = 'pie_charts'
create_pie_chart(weekly_sales, pie_chart_folder)

# Step 12: Export Line Plots and Pie Chart to Excel
export_graphs_to_excel(line_plots_folder, pie_chart_folder, excel_writer)

# Step 13: Save and Close the Excel File
excel_writer._save()
excel_writer.close()

print(f"Exported demanded products, line plots, and pie charts to {excel_filename}")
