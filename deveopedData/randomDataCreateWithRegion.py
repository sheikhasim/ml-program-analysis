import random
from datetime import date, timedelta
import csv

# Define the range of dates
start_date = date(2023, 4, 1)
end_date = date(2023, 6, 6)

# List of product names
product_names = ["Product A", "Product B", "Product C", "Product D", "Product E"]

# List of regions
regions = ["Region 1", "Region 2", "Region 3", "Region 4", "Region 5"]

# Generate random data for 300 entries
data = []
for _ in range(500):
    random_date = start_date + timedelta(days=random.randint(0, (end_date - start_date).days))
    random_product = random.choice(product_names)
    random_quantity = random.randint(1, 50)
    random_region = random.choice(regions)
    entry = [random_date, random_product, random_quantity, random_region]
    data.append(entry)

# Generate the filename
today = date.today().strftime("%Y-%m-%d")
filename = f"analysisDataWithRegion_{today}.csv"

# Write the data to the CSV file
with open(filename, mode="w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["Date", "Product Name", "Quantity", "Region"])  # Write header
    writer.writerows(data)

print(f"Data saved to {filename}.")
