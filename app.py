import sys
import os
import pandas as pd
import logging
import qdarkstyle
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox, QComboBox,
    QGroupBox, QTextEdit
)
from PyQt5.QtGui import QIcon


logging.basicConfig(filename="automation_log.txt", level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")



def generate_inventory_report(df):
    df.dropna(subset=["Category", "Price", "Stock"], inplace=True)
    df["TotalValue"] = df["Price"] * df["Stock"]
    return df.groupby(["Category", "Brand"])["TotalValue"].sum().reset_index()


def generate_sales_report(df):
    df["Date"] = pd.to_datetime(df["Date"])
    df["Revenue"] = df["Quantity"] * df["Price"]
    return df.groupby(["Date", "Product"])["Revenue"].sum().reset_index()


def generate_hr_report(df):
    df["Join Date"] = pd.to_datetime(df["Join Date"])
    headcount = df.groupby("Department")[
        "Status"].count().reset_index(name="Headcount")
    salary = df.groupby("Department")[
        "Salary"].mean().reset_index(name="AvgSalary")
    return pd.merge(headcount, salary, on="Department")


def generate_marketing_report(df):
    df["ROI"] = df["Conversions"] / df["Spend"]
    return df.groupby("Campaign")[["Clicks", "Conversions", "Spend", "ROI"]].sum().reset_index()


def generate_feedback_report(df):
    df["Date"] = pd.to_datetime(df["Date"])
    return df.groupby("Category")["Rating"].mean().reset_index(name="AvgRating")



class ReportApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📊 Multi-Report Generator")
        self.setGeometry(200, 200, 600, 300)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        
        file_group = QGroupBox("Select CSV File")
        file_layout = QHBoxLayout()
        self.file_input = QLineEdit()
        browse_btn = QPushButton("Browse")
        browse_btn.setIcon(QIcon.fromTheme("folder"))
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_input)
        file_layout.addWidget(browse_btn)
        file_group.setLayout(file_layout)

        
        type_group = QGroupBox("Choose Report Type")
        type_layout = QHBoxLayout()
        self.type_combo = QComboBox()
        self.type_combo.addItems(
            ["Inventory", "Sales", "HR", "Marketing", "Feedback"])
        type_layout.addWidget(self.type_combo)
        type_group.setLayout(type_layout)

        
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setPlaceholderText(
            "Status messages will appear here...")

        
        generate_btn = QPushButton("Generate Report")
        generate_btn.setStyleSheet("font-weight: bold; padding: 10px;")
        generate_btn.clicked.connect(self.run_report)

        layout.addWidget(file_group)
        layout.addWidget(type_group)
        layout.addWidget(generate_btn)
        layout.addWidget(self.log_output)
        self.setLayout(layout)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select CSV File", "", "CSV Files (*.csv)")
        if file_path:
            self.file_input.setText(file_path)

    def run_report(self):
        file_path = self.file_input.text()
        report_type = self.type_combo.currentText()

        if not os.path.exists(file_path):
            self.log_output.append("❌ CSV file not found.")
            return

        try:
            df = pd.read_csv(file_path)
            if report_type == "Inventory":
                report = generate_inventory_report(df)
            elif report_type == "Sales":
                report = generate_sales_report(df)
            elif report_type == "HR":
                report = generate_hr_report(df)
            elif report_type == "Marketing":
                report = generate_marketing_report(df)
            elif report_type == "Feedback":
                report = generate_feedback_report(df)
            else:
                self.log_output.append("⚠️ Invalid report type selected.")
                return

            output_file = f"{report_type.lower()}_report.xlsx"
            report.to_excel(output_file, index=False)
            self.log_output.append(
                f"✅ {report_type} report saved as {output_file}")
            logging.info(f"{report_type} report generated: {output_file}")

        except Exception as e:
            self.log_output.append(f"❌ Error: {str(e)}")
            logging.error(f"Error generating {report_type} report: {e}")



if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    window = ReportApp()
    window.show()
    sys.exit(app.exec_())
