# Multi-Report Generator (PyQt5)
<p align="left">
  <img src="https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/PyQt5-41CD52?style=for-the-badge&logo=qt&logoColor=white" alt="PyQt5">
  <img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas">
  <img src="https://img.shields.io/badge/OpenPyXL-1D68C3?style=for-the-badge&logo=python&logoColor=white" alt="OpenPyXL">
</p>

A Python desktop application for generating multiple business reports from CSV files with **Excel export**. Built with **PyQt5** and **Pandas**, featuring a modern dark-mode GUI using **qdarkstyle**.

## Features

* **Multi-Report Generation:** Generate Inventory, Sales, HR, Marketing, and Feedback reports from CSV files.
* **Excel Export:** Reports are saved automatically as `.xlsx` files.
* **Clean GUI:** Dark-mode interface for easy navigation and usability.
* **Real-Time Logging:** Status messages and errors are displayed in the UI and saved to `automation_log.txt`.
* **Data Aggregation:** Calculates totals, averages, revenue, ROI, and more depending on report type.

## 🛠️ Technologies Used

* **Language:** Python 3.x
* **GUI Framework:** PyQt5
* **Data Analysis:** Pandas
* **Excel Export:** OpenPyXL
* **UI Styling:** qdarkstyle

## Installation and Setup

1. **Clone the Repository:**
    ```bash
    git clone [YOUR_REPO_URL]
    cd [YOUR_REPO_NAME]
    ```

2. **Install Dependencies:**
    ```bash
    pip install pandas pyqt5 qdarkstyle openpyxl
    ```

3. **Run the Application:**
    ```bash
    python main.py
    ```

### Quick Start

1. Click **Browse** to select your CSV file.  
2. Choose the desired report type from the dropdown.  
3. Click **Generate Report** – the Excel file will be saved automatically.  
4. Check the `automation_log.txt` for detailed logs.

## Supported Reports & Required Columns

| Report Type | Required Columns |
|------------|----------------|
| Inventory | Category, Brand, Price, Stock |
| Sales | Date, Product, Quantity, Price |
| HR | Department, Status, Salary, Join Date |
| Marketing | Campaign, Clicks, Conversions, Spend |
| Feedback | Date, Category, Rating |

## Future Enhancements

* Chart & graph visualization
* Batch CSV processing
* Database support (SQLite/MySQL)
* PDF export
* Web-based version

***

### Git Commands for Finalizing Project

```bash
git add .
git commit -m "feat: Final project completion with PyQt5 GUI and multi-report generator"
git push
