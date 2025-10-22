# 1. Start with an official Python base image
FROM python:3.10-slim

# 2. Set the working directory inside the container
WORKDIR /app

# 3. Copy your requirements file and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 4. Copy all your app code (the .py files) into the container
COPY . .

# 5. Expose the port Render will provide (10000 is common)
EXPOSE 10000

# 6. Tell gunicorn to bind to the $PORT variable provided by Render
#    This is the most important change.
CMD ["gunicorn", "dashboard_app:server", "--bind", "0.0.0.0:$PORT"]