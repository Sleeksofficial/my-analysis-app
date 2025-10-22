# Excel to Interactive Dashboard Generator

This web application instantly transforms raw Excel files into a beautiful, interactive, Power BI-style dashboard.

Upload your `.xlsx` or `.xls` file and get an immediate, detailed analysis and dynamic, colorful charts. This tool is built with Python, Dash (by Plotly), and is fully containerized with Docker, allowing you to run it locally or deploy it to the cloud with ease.

![image](httpst://github.com/user/repo/assets/your_image.png) 
*(Optional: Add a screenshot of your dashboard here!)*

---

## 🚀 Features

* **File Upload:** Simple "drag and drop" or "click to upload" interface for your Excel files.
* **Detailed Written Analysis:** Automatically generates a text-based report summarizing your data, including:
    * Row and column counts
    * Data quality checks (missing values, duplicate rows)
    * Column type identification (numerical vs. categorical)
* **Interactive Dashboard:** A "Power BI-style" dashboard with dynamic, colorful charts that are fully responsive.
* **Dynamic Chart Builder:** Allows you to select different columns for the X and Y axes, choose chart types (bar, line, scatter, etc.), and group data, all in real-time.

## 🛠️ Tech Stack

* **Backend & Dashboard:** Python, Plotly Dash, Dash Bootstrap Components
* **Data Manipulation:** Pandas
* **Containerization:** Docker
* **Web Server:** Gunicorn (for production inside Docker)

## 🏁 Getting Started

To run this application on your local machine, you only need one piece of software installed.

### Prerequisites

* **Docker Desktop:** You must have [Docker Desktop](https://www.docker.com/products/docker-desktop/) installed and **running** on your PC (Windows, Mac, or Linux).

---

## 🏃 How to Run Locally

You can get the entire application running with just two commands in your terminal.

1.  **Clone the Repository:**
    ```bash
    git clone [https://github.com/your-username/your-repository-name.git](https://github.com/your-username/your-repository-name.git)
    cd your-repository-name
    ```

2.  **Build the Docker Image:**
    This command reads the `Dockerfile`, downloads the Python base image, and installs all the dependencies from `requirements.txt`.
    ```bash
    docker build -t analysis-app .
    ```

3.  **Run the Docker Container:**
    This command starts your app.
    ```bash
    docker run -p 8050:8050 -d analysis-app
    ```
    * `-p 8050:8050`: Connects your computer's port 8050 to the app's port inside the container.
    * `-d`: Runs the app in "detached" mode (in the background).

4.  **Access Your App:**
    Open your web browser and go to:
    **[http://127.0.0.1:8050](http://127.0.0.1:8050)**

## 💡 How to Use the App

1.  **Upload Your File:** Drag an Excel file onto the upload box or click to select one.
2.  **View Analysis:** The app will process the file and instantly display the interactive dashboard.
3.  **Explore Data:**
    * Click the **"Detailed Analysis Report"** tab to read the auto-generated summary of your data.
    * Click the **"Interactive Dashboard"** tab to build your own charts.
    * Use the dropdown menus to select different sheets, chart types, and columns for the X and Y axes.

## 📁 Project Structure

Here is a brief overview of the project's file structure:

. ├── dashboard_app.py # The main Dash web application file (handles UI, callbacks, and layout). ├── analysis_module.py # A Python module with all the backend logic (Excel parsing, report generation). ├── Dockerfile # Instructions for Docker to build the app container. ├── requirements.txt # A list of all Python libraries needed for the app. └── README.md # You are here!
