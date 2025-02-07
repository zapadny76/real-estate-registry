import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                            QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
                            QDialog, QLineEdit, QLabel, QMessageBox, QFileDialog,
                            QTabWidget)
from PyQt6.QtCore import Qt, QEvent
from PyQt6.QtGui import QIntValidator, QShortcut, QKeySequence, QCloseEvent
import pandas as pd
from pymongo import MongoClient
from datetime import datetime
from bson import ObjectId

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Real Estate Registry")
        self.setGeometry(100, 100, 1200, 800)

        # MongoDB connection
        try:
            self.client = MongoClient('mongodb://localhost:27017/')
            self.db = self.client['real_estate_registry']
        except Exception as e:
            QMessageBox.critical(self, "Database Error", f"Cannot connect to MongoDB: {str(e)}")
            sys.exit(1)

        # Создаем центральный виджет
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Создаем главный layout
        self.main_layout = QVBoxLayout(self.central_widget)

        # Создаем и добавляем кнопки
        self.create_buttons()

        # Создаем вкладки
        self.create_tabs()

        # Добавляем горячую клавишу
        exit_shortcut = QShortcut(QKeySequence('Ctrl+Q'), self)
        exit_shortcut.activated.connect(self.exit_application)

        # Загружаем начальные данные
        self.update_tables()

        # Настраиваем обработчик поиска
        self.setup_search_handler()

    def create_buttons(self):
        # Создаем контейнер для кнопок
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(10, 10, 10, 10)

        # Создаем кнопки
        import_button = QPushButton("Import from XLS")
        add_house_button = QPushButton("Add House")
        add_owner_button = QPushButton("Add Owner")
        refresh_button = QPushButton("Refresh Data")
        exit_button = QPushButton("Exit")

        # Устанавливаем минимальную ширину для кнопок
        min_button_width = 120
        for button in [import_button, add_house_button, add_owner_button,
                      refresh_button, exit_button]:
            button.setMinimumWidth(min_button_width)

        # Стилизуем кнопки
        button_style = """
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 8px 16px;
                border: none;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """

        exit_button_style = """
            QPushButton {
                background-color: #f44336;
                color: white;
                padding: 8px 16px;
                border: none;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:pressed {
                background-color: #b71c1c;
            }
        """

        # Применяем стили
        for button in [import_button, add_house_button, add_owner_button, refresh_button]:
            button.setStyleSheet(button_style)
        exit_button.setStyleSheet(exit_button_style)

        # Подключаем сигналы
        import_button.clicked.connect(self.import_from_xls)
        add_house_button.clicked.connect(self.add_house)
        add_owner_button.clicked.connect(self.add_owner)
        refresh_button.clicked.connect(self.update_tables)
        exit_button.clicked.connect(self.exit_application)

        # Добавляем кнопки в layout
        button_layout.addWidget(import_button)
        button_layout.addWidget(add_house_button)
        button_layout.addWidget(add_owner_button)
        button_layout.addWidget(refresh_button)
        button_layout.addStretch()
        button_layout.addWidget(exit_button)

        # Добавляем контейнер с кнопками в главный layout
        self.main_layout.addWidget(button_container)
    def create_tabs(self):
        # Создаем виджет вкладок
        self.tab_widget = QTabWidget()

        # Создаем вкладки
        self.houses_tab = QWidget()
        self.premises_tab = QWidget()
        self.owners_tab = QWidget()
        self.vtb_tab = QWidget()
        self.vtb_statements_tab = QWidget()

                # Создаем layouts для вкладок
        houses_layout = QVBoxLayout(self.houses_tab)
        premises_layout = QVBoxLayout(self.premises_tab)
        owners_layout = QVBoxLayout(self.owners_tab)
        vtb_layout = QVBoxLayout(self.vtb_tab)
        vtb_statements_layout = QVBoxLayout(self.vtb_statements_tab)

                # Создаем элементы для вкладки ВТБ
        vtb_buttons_layout = QHBoxLayout()

                # Кнопки для работы с регистрами ВТБ
        import_vtb_button = QPushButton("Импорт из ВТБ")
        export_vtb_button = QPushButton("Экспорт в ВТБ")
        check_vtb_button = QPushButton("Проверка данных")

                # Стиль для кнопок ВТБ
        vtb_button_style = """
        QPushButton {
                background-color: #2196F3;
                color: white;
                padding: 8px 16px;
                border: none;
                border-radius: 4px;
                font-size: 14px;
                min-width: 150px;
                }
          QPushButton:hover {
                background-color: #1976D2;
                }
          QPushButton:pressed {
                background-color: #0D47A1;
                }
                """
                # Создаем элементы для вкладки выписок ВТБ
        statements_tools_layout = QHBoxLayout()

                # Кнопки для работы с выписками
        import_statement_button = QPushButton("Импорт выписки")
        process_statement_button = QPushButton("Обработать выписку")
        generate_report_button = QPushButton("Сформировать отчет")
        search_statement_button = QPushButton("Поиск")

                # Стиль для кнопок выписок
        statement_button_style = """
        QPushButton {
                background-color: #673AB7;
                color: white;
                padding: 8px 16px;
                border: none;
                border-radius: 4px;
                font-size: 14px;
                min-width: 150px;
                }
        QPushButton:hover {
                background-color: #5E35B1;
                }
        QPushButton:pressed {
                background-color: #512DA8;
                }
                """
                # Применяем стили к кнопкам
        import_vtb_button.setStyleSheet(vtb_button_style)
        export_vtb_button.setStyleSheet(vtb_button_style)
        check_vtb_button.setStyleSheet(vtb_button_style)

        for button in [import_statement_button, process_statement_button,
                  generate_report_button, search_statement_button]:
              button.setStyleSheet(statement_button_style)

                # Подключаем сигналы
        import_vtb_button.clicked.connect(self.import_vtb_data)
        export_vtb_button.clicked.connect(self.export_vtb_data)
        check_vtb_button.clicked.connect(self.check_vtb_data)

        import_statement_button.clicked.connect(self.import_statement)
        process_statement_button.clicked.connect(self.process_statement)
        generate_report_button.clicked.connect(self.generate_statement_report)
        search_statement_button.clicked.connect(self.search_in_statements)

                # Добавляем кнопки ВТБ в layout
        vtb_buttons_layout.addWidget(import_vtb_button)
        vtb_buttons_layout.addWidget(export_vtb_button)
        vtb_buttons_layout.addWidget(check_vtb_button)
        vtb_buttons_layout.addStretch()

                # Добавляем кнопки выписок в layout
        statements_tools_layout.addWidget(import_statement_button)
        statements_tools_layout.addWidget(process_statement_button)
        statements_tools_layout.addWidget(generate_report_button)
        statements_tools_layout.addWidget(search_statement_button)
        statements_tools_layout.addStretch()

                # Создаем поисковую панель для выписок
        search_layout = QHBoxLayout()
        self.statement_search_input = QLineEdit()
        self.statement_search_input.setPlaceholderText("Поиск по ФИО или номеру счета...")
        self.statement_search_input.setMinimumWidth(300)
        search_layout.addWidget(self.statement_search_input)
        self.houses_table = QTableWidget()
        self.houses_table.setColumnCount(3)
        self.houses_table.setHorizontalHeaderLabels(["ID", "Address", "Build Year"])
        self.houses_table.horizontalHeader().setStretchLastSection(True)

        self.premises_table = QTableWidget()
        self.premises_table.setColumnCount(4)
        self.premises_table.setHorizontalHeaderLabels(["ID", "House ID", "Number", "Area"])
        self.premises_table.horizontalHeader().setStretchLastSection(True)

        self.owners_table = QTableWidget()
        self.owners_table.setColumnCount(6)
        self.owners_table.setHorizontalHeaderLabels(
                    ["ID", "First Name", "Last Name", "Document", "Share", "Status"])
        self.owners_table.horizontalHeader().setStretchLastSection(True)

                        # Создаем таблицу ВТБ
        self.vtb_table = QTableWidget()
        self.vtb_table.setColumnCount(7)
        self.vtb_table.setHorizontalHeaderLabels([
                    "ID", "Номер счета", "ФИО владельца",
                    "Дата операции", "Тип операции", "Сумма", "Статус"
                    ])
        self.vtb_table.horizontalHeader().setStretchLastSection(True)

                        # Создаем таблицу выписок
        self.statements_table = QTableWidget()
        self.statements_table.setColumnCount(8)
        self.statements_table.setHorizontalHeaderLabels([
                    "ID", "ФИО", "Номер счета", "Дата операции",
                    "Тип операции", "Сумма", "Назначение платежа", "Статус"
                    ])
        self.statements_table.horizontalHeader().setStretchLastSection(True)

                        # Добавляем таблицы на вкладки
        houses_layout.addWidget(self.houses_table)
        premises_layout.addWidget(self.premises_table)
        owners_layout.addWidget(self.owners_table)
        vtb_layout.addLayout(vtb_buttons_layout)
        vtb_layout.addWidget(self.vtb_table)

                        # Добавляем элементы на вкладку выписок
        vtb_statements_layout.addLayout(statements_tools_layout)
        vtb_statements_layout.addLayout(search_layout)
        vtb_statements_layout.addWidget(self.statements_table)

                        # Добавляем вкладки в виджет вкладок
        self.tab_widget.addTab(self.houses_tab, "Houses")
        self.tab_widget.addTab(self.premises_tab, "Premises")
        self.tab_widget.addTab(self.owners_tab, "Owners")
        self.tab_widget.addTab(self.vtb_tab, "Регистры ВТБ")
        self.tab_widget.addTab(self.vtb_statements_tab, "Выписки ВТБ (физ. лица)")

                        # Стилизуем вкладки
        tab_style = """
                            QTabWidget::pane {
                                border: 1px solid #cccccc;
                                background: white;
                            }
                            QTabWidget::tab-bar {
                                left: 5px;
                            }
                            QTabBar::tab {
                                background: #f0f0f0;
                                border: 1px solid #cccccc;
                                padding: 8px 12px;
                                margin-right: 2px;
                            }
                            QTabBar::tab:selected {
                                background: #4CAF50;
                                color: white;
                            }
                            QTabBar::tab:hover {
                                background: #45a049;
                                color: white;
                            }
                        """
        self.tab_widget.setStyleSheet(tab_style)
                        # Добавляем виджет вкладок в главный layout
        self.main_layout.addWidget(self.tab_widget)
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

    def import_vtb_data(self):
        try:
            file_name, _ = QFileDialog.getOpenFileName(
                self,
                "Выберите файл регистра ВТБ",
                "",
                "Excel Files (*.xlsx *.xls);;CSV Files (*.csv)"
            )

            if file_name:
                if file_name.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file_name)
                else:
                    df = pd.read_csv(file_name, encoding='utf-8')

                self.vtb_table.setRowCount(0)

                                        # Заполняем таблицу данными
                for index, row in df.iterrows():
                    self.vtb_table.insertRow(index)
                    for col, value in enumerate(row):
                        self.vtb_table.setItem(
                            index, col, QTableWidgetItem(str(value))
                        )

                QMessageBox.information(
                    self,
                    "Успех",
                    "Данные из регистра ВТБ успешно импортированы"
                )

        except Exception as e:
          QMessageBox.critical(
                self,
                "Ошибка",
                f"Ошибка при импорте данных ВТБ: {str(e)}"
          )

    def import_statement(self):
        try:
          file_name, _ = QFileDialog.getOpenFileName(
          self,
          "Выберите файл выписки",
          "",
          "Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;PDF Files (*.pdf)"
          )

          if not file_name:
             return

          if file_name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_name)
          elif file_name.endswith('.csv'):
            df = pd.read_csv(file_name, encoding='utf-8')
          elif file_name.endswith('.pdf'):
            QMessageBox.information(
                self,
                "PDF Import",
                "Обработка PDF файлов будет добавлена в следующей версии"
            )
            return
                                    # Очищаем таблицу
          self.statements_table.setRowCount(0)

                                    # Заполняем таблицу данными
          for index, row in df.iterrows():
              self.statements_table.insertRow(index)
              for col, value in enumerate(row):
                  self.statements_table.setItem(
                      index, col, QTableWidgetItem(str(value))
                  )
          QMessageBox.information(
              self,
              "Успех",
              "Выписка успешно импортирована"
          )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Ошибка",
                f"Ошибка при импорте выписки: {str(e)}"
            )
    def export_vtb_data(self):
        try:
          file_name, _ = QFileDialog.getSaveFileName(
              self,
              "Сохранить регистр ВТБ",
              "",
              "Excel Files (*.xlsx);;CSV Files (*.csv)"
          )

          if file_name:
              data = []
              for row in range(self.vtb_table.rowCount()):
                  row_data = []
                  for col in range(self.vtb_table.columnCount()):
                      item = self.vtb_table.item(row, col)
                      row_data.append(item.text() if item else "")
                  data.append(row_data)

              df = pd.DataFrame(
                  data,
                  columns=[
                      "ID", "Номер счета", "ФИО владельца",
                      "Дата операции", "Тип операции", "Сумма", "Статус"
                  ]
              )                try:
                        file_name, _ = QFileDialog.getSaveFileName(
                            self,
                            "Сохранить регистр ВТБ",
                            "",
                            "Excel Files (*.xlsx);;CSV Files (*.csv)"
                        )

                        if file_name:
                            data = []
                            for row in range(self.vtb_table.rowCount()):
                                row_data = []
                                for col in range(self.vtb_table.columnCount()):
                                    item = self.vtb_table.item(row, col)
                                    row_data.append(item.text() if item else "")
                                data.append(row_data)

                            df = pd.DataFrame(
                                data,
                                columns=[
                                    "ID", "Номер счета", "ФИО владельца",
                                    "Дата операции", "Тип операции", "Сумма", "Статус"
                                ]
                            )
                                                    if file_name.endswith('.xlsx'):
                                                        df.to_excel(file_name, index=False)
                                                    else:
                                                        df.to_csv(file_name, index=False, encoding='utf-8')

                                                    QMessageBox.information(self, "Успех", "Данные успешно экспортированы")

                                            except Exception as e:
                                                QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте данных: {str(e)}")

                                        def process_statement(self):
                                            try:
                                                if self.statements_table.rowCount() == 0:
                                                    QMessageBox.warning(
                                                        self,
                                                        "Предупреждение",
                                                        "Нет данных для обработки. Сначала импортируйте выписку."
                                                    )
                                                    return

                                                # Здесь будет логика обработки выписки

                                                QMessageBox.information(
                                                    self,
                                                    "Обработка выписки",
                                                    "Выписка успешно обработана"
                                                )

                                            except Exception as e:
                                                QMessageBox.critical(
                                                    self,
                                                    "Ошибка",
                                                    f"Ошибка при обработке выписки: {str(e)}"
                                                )

                                        def generate_statement_report(self):
                                            try:
                                                if self.statements_table.rowCount() == 0:
                                                    QMessageBox.warning(
                                                        self,
                                                        "Предупреждение",
                                                        "Нет данных для формирования отчета"
                                                    )
                                                    return

                                                file_name, _ = QFileDialog.getSaveFileName(
                                                    self,
                                                    "Сохранить отчет",
                                                    "",
                                                    "Excel Files (*.xlsx);;PDF Files (*.pdf)"
                                                )

                                                if not file_name:
                                                    return

                                                data = []
                                                headers = []
                                                for col in range(self.statements_table.columnCount()):
                                                    headers.append(self.statements_table.horizontalHeaderItem(col).text())

                                                for row in range(self.statements_table.rowCount()):
                                                    row_data = []
                                                    for col in range(self.statements_table.columnCount()):
                                                        item = self.statements_table.item(row, col)
                                                        row_data.append(item.text() if item else "")
                                                    data.append(row_data)

                                                df = pd.DataFrame(data, columns=headers)

                                                if file_name.endswith('.xlsx'):
                                                    df.to_excel(file_name, index=False)
                                                elif file_name.endswith('.pdf'):
                                                    # Здесь можно добавить генерацию PDF
                                                    pass

                                                QMessageBox.information(
                                                    self,
                                                    "Успех",
                                                    "Отчет успешно сформирован"
                                                )

                                            except Exception as e:
                                                QMessageBox.critical(
                                                    self,
                                                    "Ошибка",
                                                    f"Ошибка при формировании отчета: {str(e)}"
                                                )

                                        def check_vtb_data(self):
                                            try:
                                                errors = []
                                                for row in range(self.vtb_table.rowCount()):
                                                    account = self.vtb_table.item(row, 1)
                                                    if account and not account.text().isdigit():
                                                        errors.append(f"Строка {row + 1}: Неверный формат номера счета")

                                                    amount = self.vtb_table.item(row, 5)
                                                    if amount:
                                                        try:
                                                            float(amount.text().replace(',', '.'))
                                                        except ValueError:
                                                            errors.append(f"Строка {row + 1}: Неверный формат суммы")

                                                if errors:
                                                    error_text = "\n".join(errors)
                                                    QMessageBox.warning(
                                                        self,
                                                        "Результаты проверки",
                                                        f"Найдены ошибки в данных:\n{error_text}"
                                                    )
                                                else:
                                                    QMessageBox.information(
                                                        self,
                                                        "Результаты проверки",
                                                        "Ошибок в данных не обнаружено"
                                                    )

                                            except Exception as e:
                                                QMessageBox.critical(
                                                    self,
                                                    "Ошибка",
                                                    f"Ошибка при проверке данных: {str(e)}"
                                                )

                                        def search_in_statements(self):
                                            search_text = self.statement_search_input.text().lower()

                                            if not search_text:
                                                for row in range(self.statements_table.rowCount()):
                                                    self.statements_table.setRowHidden(row, False)
                                                return

                                            for row in range(self.statements_table.rowCount()):
                                                row_matches = False
                                                for col in range(self.statements_table.columnCount()):
                                                    item = self.statements_table.item(row, col)
                                                    if item and search_text in item.text().lower():
                                                        row_matches = True
                                                        break
                                                self.statements_table.setRowHidden(row, not row_matches)

                                        def setup_search_handler(self):
                                            self.statement_search_input.textChanged.connect(self.search_in_statements)

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

                                        def closeEvent(self, a0: QCloseEvent | None) -> None:
                                            if a0 is not None:
                                                if self.confirm_exit():
                                                    self.client.close()
                                                    a0.accept()
                                                else:
                                                    a0.ignore()

                                        def confirm_exit(self) -> bool:
                                            reply = QMessageBox.question(self, 'Exit',
                                                "Are you sure you want to exit?",
                                                QMessageBox.StandardButton.Yes |
                                                QMessageBox.StandardButton.No,
                                                QMessageBox.StandardButton.No)
                                            return reply == QMessageBox.StandardButton.Yes

                                        def exit_application(self) -> None:
                                            if self.confirm_exit():
                                                self.client.close()
                                                QApplication.quit()

                                    def main():
                                        app = QApplication(sys.argv)
                                        app.setStyle('Fusion')

                                        window = MainWindow()
                                        window.show()

                                        sys.exit(app.exec())

                                    if __name__ == '__main__':
                                        main()
