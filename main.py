import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                            QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
                            QDialog, QLineEdit, QLabel, QMessageBox, QFileDialog)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIntValidator
from PyQt6.QtGui import QShortcut, QKeySequence
import pandas as pd
from pymongo import MongoClient
from datetime import datetime
from bson import ObjectId

class OwnerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add/Edit Owner")
        self.setMinimumWidth(300)
        layout = QVBoxLayout()

        self.first_name = QLineEdit()
        self.last_name = QLineEdit()
        self.doc_number = QLineEdit()
        self.ownership_share = QLineEdit()

        layout.addWidget(QLabel("First Name:"))
        layout.addWidget(self.first_name)
        layout.addWidget(QLabel("Last Name:"))
        layout.addWidget(self.last_name)
        layout.addWidget(QLabel("Document Number:"))
        layout.addWidget(self.doc_number)
        layout.addWidget(QLabel("Ownership Share:"))
        layout.addWidget(self.ownership_share)

        save_button = QPushButton("Save")
        save_button.clicked.connect(self.accept)
        layout.addWidget(save_button)

        self.setLayout(layout)

class HouseDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add House")
        self.setMinimumWidth(300)
        layout = QVBoxLayout()

        self.address = QLineEdit()
        self.build_year = QLineEdit()
        self.build_year.setValidator(QIntValidator(1800, 2100))

        layout.addWidget(QLabel("Address:"))
        layout.addWidget(self.address)
        layout.addWidget(QLabel("Build Year:"))
        layout.addWidget(self.build_year)

        buttons = QHBoxLayout()
        save_button = QPushButton("Save")
        cancel_button = QPushButton("Cancel")

        save_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)

        buttons.addWidget(save_button)
        buttons.addWidget(cancel_button)
        layout.addLayout(buttons)

        self.setLayout(layout)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Real Estate Registry")
        self.setGeometry(100, 100, 1200, 800)
        # Добавляем горячую клавишу Ctrl+Q для выхода
        exit_shortcut = QShortcut(QKeySequence('Ctrl+Q'), self)
        exit_shortcut.activated.connect(self.exit_application)
        # MongoDB connection
        try:
            self.client = MongoClient('mongodb://localhost:27017/')
            self.db = self.client['real_estate_registry']
        except Exception as e:
            QMessageBox.critical(self, "Database Error", f"Cannot connect to MongoDB: {str(e)}")
            sys.exit(1)
        exit_shortcut = QShortcut(QKeySequence('Ctrl+Q'), self)
        exit_shortcut.activated.connect(self.exit_application)

        self.setup_ui()

    def setup_ui(self):
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Create button panel
        button_panel = QWidget()
        button_layout = QHBoxLayout(button_panel)

        # Create buttons
        import_button = QPushButton("Import from XLS")
        add_house_button = QPushButton("Add House")
        add_owner_button = QPushButton("Add Owner")
        refresh_button = QPushButton("Refresh Data")
        exit_button = QPushButton("Exit")
        exit_button.setStyleSheet("""
                QPushButton {
                    background-color: #ff4444;
                    color: white;
                    padding: 5px 15px;
                    border: none;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #ff6666;
                }
                QPushButton:pressed {
                    background-color: #cc0000;
                }
            """)
        exit_button.clicked.connect(self.exit_application)

        # Connect button signals
        import_button.clicked.connect(self.import_from_xls)
        add_house_button.clicked.connect(self.add_house)
        add_owner_button.clicked.connect(self.add_owner)
        refresh_button.clicked.connect(self.update_tables)

        # Add buttons to layout
        button_layout.addWidget(import_button)
        button_layout.addWidget(add_house_button)
        button_layout.addWidget(add_owner_button)
        button_layout.addWidget(refresh_button)
        button_layout.addStretch()

        main_layout.addWidget(button_panel)

        # Create tables
        # Houses table
        houses_label = QLabel("Houses:")
        self.houses_table = QTableWidget()
        self.houses_table.setColumnCount(3)
        self.houses_table.setHorizontalHeaderLabels(["ID", "Address", "Build Year"])
        self.houses_table.horizontalHeader().setStretchLastSection(True)

        # Premises table
        premises_label = QLabel("Premises:")
        self.premises_table = QTableWidget()
        self.premises_table.setColumnCount(4)
        self.premises_table.setHorizontalHeaderLabels(["ID", "House ID", "Number", "Area"])
        self.premises_table.horizontalHeader().setStretchLastSection(True)

        # Owners table
        owners_label = QLabel("Owners:")
        self.owners_table = QTableWidget()
        self.owners_table.setColumnCount(6)
        self.owners_table.setHorizontalHeaderLabels(
            ["ID", "First Name", "Last Name", "Document", "Share", "Status"])
        self.owners_table.horizontalHeader().setStretchLastSection(True)

        # Add tables to layout
        main_layout.addWidget(houses_label)
        main_layout.addWidget(self.houses_table)
        main_layout.addWidget(premises_label)
        main_layout.addWidget(self.premises_table)
        main_layout.addWidget(owners_label)
        main_layout.addWidget(self.owners_table)

        # Initial data load
        self.update_tables()
    def exit_application(self):
                reply = QMessageBox.question(self, 'Exit',
                    "Are you sure you want to exit?",
                    QMessageBox.StandardButton.Yes |
                    QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No)

                if reply == QMessageBox.StandardButton.Yes:
                    # Закрываем соединение с базой данных
                    self.client.close()
                    QApplication.quit()

    def closeEvent(self, event):
                reply = QMessageBox.question(self, 'Exit',
                    "Are you sure you want to exit?",
                    QMessageBox.StandardButton.Yes |
                    QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No)

                if reply == QMessageBox.StandardButton.Yes:
                    # Закрываем соединение с базой данных
                    self.client.close()
                    event.accept()
                else:
                    event.ignore()
    def import_from_xls(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")

        if file_name:
            try:
                df = pd.read_excel(file_name)

                for _, row in df.iterrows():
                    house_data = {
                        'address': str(row['address']),
                        'build_year': int(row['build_year'])
                    }

                    house_id = self.db.houses.insert_one(house_data).inserted_id

                    if 'premises' in row:
                        for premise in str(row['premises']).split(';'):
                            premise_data = {
                                'house_id': house_id,
                                'number': premise.strip(),
                                'area': float(row.get('area', 0))
                            }
                            self.db.premises.insert_one(premise_data)

                self.update_tables()
                QMessageBox.information(self, "Success", "Data imported successfully")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Import failed: {str(e)}")

    def add_house(self):
        dialog = HouseDialog(self)
        if dialog.exec():
            try:
                house_data = {
                    'address': dialog.address.text(),
                    'build_year': int(dialog.build_year.text())
                }
                self.db.houses.insert_one(house_data)
                self.update_tables()
                QMessageBox.information(self, "Success", "House added successfully")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to add house: {str(e)}")

    def add_owner(self):
        dialog = OwnerDialog(self)
        if dialog.exec():
            try:
                owner_data = {
                    'first_name': dialog.first_name.text(),
                    'last_name': dialog.last_name.text(),
                    'document': {
                        'number': dialog.doc_number.text(),
                        'issue_date': datetime.now()
                    },
                    'ownership_share': float(dialog.ownership_share.text()),
                    'status': 'active'
                }

                self.db.owners.insert_one(owner_data)
                self.update_tables()
                QMessageBox.information(self, "Success", "Owner added successfully")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to add owner: {str(e)}")

    def update_tables(self):
        try:
            # Update Houses Table
            self.houses_table.setRowCount(0)
            for house in self.db.houses.find():
                row = self.houses_table.rowCount()
                self.houses_table.insertRow(row)
                self.houses_table.setItem(row, 0, QTableWidgetItem(str(house['_id'])))
                self.houses_table.setItem(row, 1, QTableWidgetItem(house['address']))
                self.houses_table.setItem(row, 2, QTableWidgetItem(str(house['build_year'])))

            # Update Premises Table
            self.premises_table.setRowCount(0)
            for premise in self.db.premises.find():
                row = self.premises_table.rowCount()
                self.premises_table.insertRow(row)
                self.premises_table.setItem(row, 0, QTableWidgetItem(str(premise['_id'])))
                self.premises_table.setItem(row, 1, QTableWidgetItem(str(premise['house_id'])))
                self.premises_table.setItem(row, 2, QTableWidgetItem(premise['number']))
                self.premises_table.setItem(row, 3, QTableWidgetItem(str(premise['area'])))

            # Update Owners Table
            self.owners_table.setRowCount(0)
            for owner in self.db.owners.find():
                row = self.owners_table.rowCount()
                self.owners_table.insertRow(row)
                self.owners_table.setItem(row, 0, QTableWidgetItem(str(owner['_id'])))
                self.owners_table.setItem(row, 1, QTableWidgetItem(owner['first_name']))
                self.owners_table.setItem(row, 2, QTableWidgetItem(owner['last_name']))
                self.owners_table.setItem(row, 3, QTableWidgetItem(owner['document']['number']))
                self.owners_table.setItem(row, 4, QTableWidgetItem(str(owner['ownership_share'])))
                self.owners_table.setItem(row, 5, QTableWidgetItem(owner['status']))

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to update tables: {str(e)}")

def main():
    app = QApplication(sys.argv)

    # Set application style
    app.setStyle('Fusion')

    window = MainWindow()
    window.show()

    sys.exit(app.exec())

if __name__ == '__main__':
    main()
