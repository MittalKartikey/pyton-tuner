import tkinter as tk
from openpyxl import Workbook
from datetime import datetime
from tkinter import ttk

class ExpenseTracker:
    def __init__(self, root):
        self.root = root
        self.root.title("Expense Tracker")

        # Create sections
        self.date_section = tk.Frame(self.root)
        self.date_section.pack(fill="x")

        self.category_section = tk.Frame(self.root)
        self.category_section.pack(fill="x")

        self.amount_section = tk.Frame(self.root)
        self.amount_section.pack(fill="x")

        # Create entry fields
        self.date_label = tk.Label(self.date_section, text="Date (DD):")
        self.date_label.pack(side="left")
        self.date_entry = tk.Entry(self.date_section, width=5)
        self.date_entry.pack(side="left")
        self.date_entry.bind("<Return>", lambda event: self.category_entry.focus())

        self.month_label = tk.Label(self.date_section, text="Month:")
        self.month_label.pack(side="left")
        self.month_var = tk.StringVar()
        self.month_var.set("January")  # default month
        self.month_menu = ttk.Combobox(self.date_section, textvariable=self.month_var, values=["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"])
        self.month_menu.pack(side="left")

        self.year_label = tk.Label(self.date_section, text="Year:")
        self.year_label.pack(side="left")
        self.year_var = tk.StringVar()
        self.year_var.set("2024")  # default year
        self.year_menu = ttk.Combobox(self.date_section, textvariable=self.year_var, values=["2023", "2024", "2025"])
        self.year_menu.pack(side="left")

        self.category_label = tk.Label(self.category_section, text="Category:")
        self.category_label.pack(side="left")
        self.category_entry = tk.Entry(self.category_section)
        self.category_entry.pack(side="left")
        self.category_entry.bind("<Return>", lambda event: self.amount_entry.focus())

        self.amount_label = tk.Label(self.amount_section, text="Amount:")
        self.amount_label.pack(side="left")
        self.amount_entry = tk.Entry(self.amount_section)
        self.amount_entry.pack(side="left")
        self.amount_entry.bind("<Return>", self.submit_form)

        # Create a button to submit the form
        self.submit_button = tk.Button(self.root, text="Submit", command=self.submit_form)
        self.submit_button.pack(fill="x")

        # Create an Excel workbook and worksheet
        self.wb = Workbook()
        self.ws = self.wb.active

        # Create a dictionary to store the data month-wise
        self.monthly_expenses = {}

    def submit_form(self, event=None):
        # Get the values from the entry fields
        day = self.date_entry.get()
        month = self.month_var.get()
        year = self.year_var.get()
        date_str = f"{year}-{self.get_month_number(month)}-{day}"
        category = self.category_entry.get()
        amount = float(self.amount_entry.get())

        # Parse the date string to get the month
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        month_name = date_obj.strftime("%B")

        # Add the data to the dictionary
        if month_name not in self.monthly_expenses:
            self.monthly_expenses[month_name] = []
        self.monthly_expenses[month_name].append({"category": category, "amount": amount})

        # Calculate the total monthly expense
        total_monthly_expense = sum([item["amount"] for item in self.monthly_expenses[month_name]])

        # Add the data to the Excel sheet
        self.ws.append([date_str, category, amount, month_name, total_monthly_expense])

        # Save the Excel file
        self.wb.save("expenses.xlsx")

        # Clear the entry fields
        self.date_entry.delete(0, tk.END)
        self.category_entry.delete(0, tk.END)
        self.amount_entry.delete(0, tk.END)

        # Move focus back to the first field
        self.date_entry.focus()

    def get_month_number(self, month_name):
        month_numbers = {
            "January": "01",
            "February": "02",
            "March": "03",
            "April": "04",
            "May": "05",
            "June": "06",
            "July": "07",
            "August": "08",
            "September": "09",
            "October": "10",
            "November": "11",
            "December": "12"
        }
        return month_numbers[month_name]

root = tk.Tk()
app = ExpenseTracker(root)
root.mainloop()
