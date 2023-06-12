import pandas as pd
import random
from datetime import timedelta, datetime, date

# Load the dataset containing the available products
products_df = pd.read_excel('./load_dataset/loadDataset.xlsx')

# Get the list of available products from the dataset
available_products = products_df['Product Name'].tolist()

# Load the dataset containing the regions
regions_df = pd.read_excel('./load_dataset/regions.xlsx')

# Get the list of regions from the dataset
available_regions = regions_df['Region'].tolist()

# Start date is January 1, 2023
start_date = datetime(2023, 1, 1)

# End date is today's date
end_date = date.today()

# Generate random sales data for each product and region
sales_data = []
current_date = start_date
while current_date.date() <= end_date:
    for product in available_products:
        region = random.choice(available_regions)  # Select a random region
        quantity_sold = random.randint(0, 100)  # Generate a random quantity sold
        sales_data.append({
            'Date': current_date.strftime('%Y-%m-%d'),
            'Product Name': product,
            'Region': region,
            'Quantity': quantity_sold
        })
    current_date += timedelta(days=1)

# Create a DataFrame from the sales data
sales_df = pd.DataFrame(sales_data)

# Save the sales data to a CSV file
sales_df.to_csv('./deveopedData/createdData.csv', index=False)

# Save the regions data to a CSV file
regions_df.to_csv('./deveopedData/regions.csv', index=False)







# import pandas as pd
# import random
# from datetime import timedelta, datetime,date

# # Load the dataset containing the available products
# products_df = pd.read_excel('./load_dataset/loadDataset.xlsx')

# # Get the list of available products from the dataset
# available_products = products_df['Product Name'].tolist()

# # Load the dataset containing the regions
# regions_df = pd.read_excel('./load_dataset/regions.xlsx')

# # Get the list of regions from the dataset
# available_regions = regions_df['Region'].tolist()

# # Start and end dates for the two-month period
# start_date = datetime(2023, 1, 1)
# end_date = start_date + timedelta(days=59)

# # Generate random sales data for each product and region
# sales_data = []
# current_date = start_date
# while current_date <= end_date:
#     for product in available_products:
#         region = random.choice(available_regions)  # Select a random region
#         quantity_sold = random.randint(0, 100)  # Generate a random quantity sold
#         sales_data.append({
#             'Date': current_date.strftime('%Y-%m-%d'),
#             'Product Name': product,
#             'Region': region,
#             'Quantity': quantity_sold
#         })
#     current_date += timedelta(days=1)

# # Create a DataFrame from the sales data
# sales_df = pd.DataFrame(sales_data)

# # Save the sales data to a CSV file
# sales_df.to_csv('./deveopedData/createdData.csv', index=False)

# # Save the regions data to a CSV file
# regions_df.to_csv('./deveopedData/regions.csv', index=False)











# # import pandas as pd
# # import random
# # from datetime import timedelta, datetime

# # # Load the dataset containing the available products
# # products_df = pd.read_excel('./load_dataset/loadDataset.xlsx')

# # # Get the list of available products from the dataset
# # available_products = products_df['Product Name'].tolist()

# # # Start and end dates for the two-month period
# # start_date = datetime(2023, 1, 1)
# # end_date = start_date + timedelta(days=59)

# # # Generate random sales data for each product
# # sales_data = []
# # current_date = start_date
# # while current_date <= end_date:
# #     for product in available_products:
# #         quantity_sold = random.randint(0, 100)  # Generate a random quantity sold
# #         sales_data.append({
# #             'Date': current_date.strftime('%Y-%m-%d'),
# #             'Product Name': product,
# #             'Quantity': quantity_sold
# #         })
# #     current_date += timedelta(days=1)

# # # Create a DataFrame from the sales data
# # sales_df = pd.DataFrame(sales_data)

# # # Save the sales data to an Excel file
# # sales_df.to_csv('./deveopedData/createdData1.csv', index=False)
