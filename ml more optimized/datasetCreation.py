import pandas as pd
import random
from datetime import datetime, timedelta

class Product:
    def __init__(self, name):
        self.name = name

class Region:
    def __init__(self, name):
        self.name = name

class SalesData:
    def __init__(self, product, region, date, quantity):
        self.product = product
        self.region = region
        self.date = date
        self.quantity = quantity

# Load the dataset containing the available products
products_df = pd.read_excel('../load_dataset/loadDataset.xlsx')
available_products = products_df['Product Name'].tolist()

# Load the dataset containing the available regions
regions_df = pd.read_excel('../load_dataset/regions.xlsx')
available_regions = regions_df['Region'].tolist()

# Get the start date and end date
start_date_str = '2023-01-01'  # Change this to your desired start date
end_date_str = None  # Change this to your desired end date, or leave it as None

# Convert start date and end date to datetime objects
start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
if end_date_str:
    try:
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
    except ValueError:
        end_date = start_date.replace(day=28) + timedelta(days=4)
else:
    next_month = start_date.replace(day=28) + timedelta(days=4)
    end_date = next_month - timedelta(days=next_month.day)

# Generate random sales data for each product and region
sales_data = []
current_date = start_date
while current_date <= end_date:
    for product_name in available_products:
        product = Product(product_name)
        for region_name in available_regions:
            region = Region(region_name)
            quantity_sold = random.randint(0, 100)  # Generate a random quantity sold
            sales_entry = SalesData(product, region, current_date.strftime('%Y-%m-%d'), quantity_sold)
            sales_data.append(sales_entry)
    current_date += timedelta(days=1)

# Create a DataFrame from the sales data
sales_df = pd.DataFrame([{
    'Product Name': entry.product.name,
    'Region': entry.region.name,
    'Date': entry.date,
    'Quantity': entry.quantity
} for entry in sales_data])

# Generate the file name with an underscore and today's date
today_date = datetime.now().strftime('%Y-%m-%d')
file_name = f'./developed_data/sales_data_{today_date}.xlsx'

# Save the sales data to an Excel file
sales_df.to_excel(file_name, index=False)
