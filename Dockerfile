FROM busybox
CMD echo "After Sales Inventory - ML Analysis"

# Use an official Python runtime as the base image
FROM python:3.9

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file to the container
COPY requirements.txt .

# Install the dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the code to the container
COPY . .

# Set the command to run your application
CMD [ "python", "app.py" ]
