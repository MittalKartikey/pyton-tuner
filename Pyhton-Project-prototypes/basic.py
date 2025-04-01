import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, QAbstractItemView, QLineEdit, QLabel, QGridLayout
from PyQt5.QtCore import Qt

class ExpenseTracker(QWidget):
    def __init__(self):
        super().__init__()

        self.expenses = {}

        self.initUI()

    def initUI(self):
        self.setGeometry(300, 300, 800, 600)  # Set the window size to 800x600
        self.setWindowTitle('Expense Tracker')

        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.table = QTableWidget()
        self.table.setRowCount(0)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['Date', 'Category', 'Amount'])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.layout.addWidget(self.table)

        self.add_layout = QGridLayout()
        self.layout.addLayout(self.add_layout)

        self.date_label = QLabel('Date (dd/mm/yyyy)')
        self.add_layout.addWidget(self.date_label, 0, 0)

        self.date_input = QLineEdit()
        self.add_layout.addWidget(self.date_input, 0, 1)

        self.category_label = QLabel('Category')
        self.add_layout.addWidget(self.category_label, 1, 0)

        self.category_input = QLineEdit()
        self.add_layout.addWidget(self.category_input, 1, 1)

        self.amount_label = QLabel('Amount')
        self.add_layout.addWidget(self.amount_label, 2, 0)

        self.amount_input = QLineEdit()
        self.add_layout.addWidget(self.amount_input, 2, 1)

        self.add_button = QPushButton('Add Expense')
        self.add_button.clicked.connect(self.add_expense)
        self.add_layout.addWidget(self.add_button, 3, 0, 1, 2)

        self.total_label = QLabel('Total Expense: ')
        self.layout.addWidget(self.total_label)

        self.update_table()

    def add_expense(self):
        date = self.date_input.text()
        category = self.category_input.text()
        amount = self.amount_input.text()
        if date in self.expenses:
            self.expenses[date].append((category, amount))
        else:
            self.expenses[date] = [(category, amount)]
        self.update_table()
        self.date_input.clear()
        self.category_input.clear()
        self.amount_input.clear()

    def update_table(self):
        self.table.setRowCount(0)
        row = 0
        total = 0
        for date, expenses in self.expenses.items():
            for category, amount in expenses:
                self.table.insertRow(row)
                self.table.setItem(row, 0, QTableWidgetItem(date))
                self.table.setItem(row, 1, QTableWidgetItem(category))
                self.table.setItem(row, 2, QTableWidgetItem(amount))
                total += float(amount)
                row += 1
        self.total_label.setText('Total Expense: ' + str(total))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    tracker = ExpenseTracker()
    tracker.show()
    sys.exit(app.exec_())
