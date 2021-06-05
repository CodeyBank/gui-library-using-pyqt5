import datetime
import time
from PyQt5.QtCore import *
# from PyQt5.QtGui import *
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
import sys
from PyQt5.uic import loadUiType
import pymysql
from xlrd import *
from xlsxwriter import *

pymysql.install_as_MySQLdb()
import MySQLdb


ui, _ = loadUiType('library.ui')
login, _ = loadUiType('login.ui')


# set up a class and pass the main window and UI file already loaded as arguments
class Login(QWidget, login):
    def __init__(self):
        # initialise window widgets
        QWidget.__init__(self)
        # super(Login, self).__init__()
        self.setupUi(self)
        self.db = MySQLdb.connect(host='localhost', user='root', password='Thebossm@#995', db="library", port=3310)
        self.cur = self.db.cursor()  # Create a cursor
        self.pushButton.clicked.connect(self.loginHandler)
        style = open('themes/ElegantDark.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def loginHandler(self):
        username = self.usernameLEdit.text()
        password = self.password.text()
        sql = '''SELECT * FROM users '''
        self.cur.execute(sql)
        data = self.cur.fetchall()
        for row in data:
            if username == row[1] and password == row[3]:
                print("User found")
                self.window2 = MainApp()
                self.close()
                self.window2.show()
            else:
                self.label.setText('Invalid credentials. Please try again')


# set up a class and pass the main window and UI file already loaded as arguments
class MainApp(QMainWindow, ui):
    def __init__(self):
        # initialise window widgets
        QMainWindow.__init__(self)
        self.setupUi(self)
        # initial UI fixes
        self.handle_ui_changes()  # when class is instantiated this method is run
        self.handle_buttons()
        self.elegant_dark_theme()
        self.anim = QPropertyAnimation(self.groupBox_3, b"geometry")

        # Create connection to database
        self.db = MySQLdb.connect(host='localhost', user='root', password='Thebossm@#995', db="library", port=3310)
        self.cur = self.db.cursor()  # Create a cursor
        self.statusBar().showMessage('Established connection to database!', 5000)

        #  update all comboboxes in the settings tab
        self.show_category()
        self.show_author()
        self.show_publisher()

        # Update all combo boxes
        self.show_category_combobox()
        self.show_author_combobox()
        self.show_publisher_combobox()

        # count how many times a button is clicked
        self.click_count = 0

        # update all tables on startup
        self.show_all_clients()
        # self.show_all_books()
        self.show_all_operations()

    # Any event that occurs in the application calls all of these methods
    def handle_buttons(self):
        # control themes group box
        self.pushButton_5.clicked.connect(self.show_themes)
        self.pushButton_8.clicked.connect(self.doAnimClose)

        # Connecting each navigation button to respective tabs
        self.day_to_day_btn.clicked.connect(self.open_day_to_day_tab)
        self.open_books_btn.clicked.connect(self.open_books_tab)
        self.open_users_btn.clicked.connect(self.open_users_tab)
        self.settings_btn.clicked.connect(self.open_settings_tab)
        self.clients_btn.clicked.connect(self.open_clients_tab)

        # Book tab
        self.save_NewBook.clicked.connect(self.add_new_book)
        self.search_btn.clicked.connect(self.search_books)
        self.save_btn.clicked.connect(self.edit_books)
        self.delete_btn.clicked.connect(self.delete_books)

        # Settings tab. add button handlers
        self.add_categBtn.clicked.connect(self.add_category)
        self.btn_addAuthor.clicked.connect(self.add_author)
        self.addPub_btn.clicked.connect(self.add_publisher)

        # Users tab
        self.add_user_btn.clicked.connect(self.add_new_user)
        self.login_btn.clicked.connect(self.login)
        self.edit_user_data_btn.clicked.connect(self.edit_user)

        # UI buttons
        self.aqua_theme_btn.clicked.connect(self.aqua_theme)
        self.amoled_theme_btn.clicked.connect(self.amoled_theme)
        self.material_dark_btn.clicked.connect(self.material_dark_theme)
        self.ubuntu_theme_btn.clicked.connect(self.ubuntu_theme)
        self.elegant_dark_theme_btn.clicked.connect(self.elegant_dark_theme)
        self.console_theme_btn.clicked.connect(self.console_style)
        self.manjaro_btn.clicked.connect(self.manjaromix_theme)

        # Clients Tab
        self.addNewClientBtn.clicked.connect(self.add_new_client)
        self.clientSearchBtn.clicked.connect(self.search_clients)
        self.saveClientData.clicked.connect(self.edit_client_details)
        self.deleteClientData.clicked.connect(self.delete_client)

        # day to day operations
        self.pushButton_6.clicked.connect(self.operations)

        # Export buttons
        self.pushButton_3.clicked.connect(self.exportOperations)
        self.pushButton.clicked.connect(self.exportBooks)
        self.pushButton_2.clicked.connect(self.exportClients)

    def handle_ui_changes(self):
        self.hiding_themes()
        self.main_tab_widget.tabBar().setVisible(False)
        # self.books_tab_widget.tabBar().setVisible(False)
        self.resize_tHeaders()
        # header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        # self.operations_table.resizeColumnsToContents()

    def resize_tHeaders(self):
        t1 = range(0, 6)
        t2 = range(0, 7)
        t3 = range(0, 4)
        header_ops = self.operations_table.horizontalHeader()
        header_allbooks = self.allbooks_table.horizontalHeader()
        header_allclients = self.allclients_table.horizontalHeader()

        for item in t1:
            header_ops.setSectionResizeMode(item, QtWidgets.QHeaderView.ResizeToContents)
        for item1 in t2:
            header_allbooks.setSectionResizeMode(item1, QtWidgets.QHeaderView.ResizeToContents)
        for item2 in t3:
            header_allclients.setSectionResizeMode(item2, QtWidgets.QHeaderView.ResizeToContents)

    def show_themes(self):
        # count = 1  # set count to for first run
        self.click_count += 1
        if self.click_count == 1:
            self.groupBox_3.show()
            self.doAnim()
        elif self.click_count == 2:
            self.doAnimClose()
            # self.hiding_themes()

        else:
            pass

    def doAnim(self):
        self.anim.setDuration(200)
        self.anim.setStartValue(QRect(100, 0, 1, 495))
        self.anim.setEndValue(QRect(100, 0, 400, 495))
        self.anim.start()

    def doAnimClose(self):
        self.anim = QPropertyAnimation(self.groupBox_3, b"geometry")
        self.anim.setDuration(200)
        self.anim.setStartValue(QRect(100, 0, 400, 495))
        self.anim.setEndValue(QRect(100, 0, 1, 495))
        self.anim.start()
        self.click_count = 0  # reset count value

    def hiding_themes(self):
        self.groupBox_3.hide()

    ########################################################################################################
    ########################### Opening tabs ###############################################################

    def open_day_to_day_tab(self):
        self.main_tab_widget.setCurrentIndex(0)

    def open_books_tab(self):
        self.main_tab_widget.setCurrentIndex(1)

    def open_users_tab(self):
        self.main_tab_widget.setCurrentIndex(2)

    def open_clients_tab(self):
        self.main_tab_widget.setCurrentIndex(3)

    def open_settings_tab(self):
        self.main_tab_widget.setCurrentIndex(4)

    ########################################################################################################
    ###########################     Operations    ##########################################################
    def operations(self):
        book_title = self.lineEdit.text()
        book_type = self.comboBox.currentText()
        burrowed_days = self.comboBox_2.currentIndex() + 1
        client_name = self.lineEdit_2.text()
        date = datetime.date.today()
        now = time.localtime()
        operation_time = time.strftime("%H:%M:%S", now)
        to = date + datetime.timedelta(days=int(burrowed_days))
        print(date)
        print(to)
        print(operation_time)

        self.cur.execute('''
            INSERT INTO dayoperations (book_name, client_name, type, days, operation_time, date, to_date)
             VALUES (%s, %s, %s, %s, %s, %s, %s)
        ''', (book_title, client_name, book_type, burrowed_days, operation_time, date, to))

        self.db.commit()
        self.show_all_operations()
        self.statusBar().showMessage('Operation added', 3000)

    def show_all_operations(self):
        self.operations_table.setRowCount(0)
        self.operations_table.insertRow(0)
        self.cur.execute('''
        SELECT book_name, client_name, type, operation_time, date, to_date FROM dayoperations
        ''')
        data = self.cur.fetchall()
        if data:
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.operations_table.setItem(row, column, QTableWidgetItem(str(item)))
                current_row = self.operations_table.rowCount()
                self.operations_table.insertRow(current_row)

    ########################################################################################################
    ###########################     Books    ###############################################################
    def show_all_books(self):
        self.allbooks_table.setRowCount(0)
        self.allbooks_table.insertRow(0)

        self.cur.execute('''
        SELECT book_code, book_name, book_category, book_description, book_author, book_publisher, book_price FROM book
        ''')
        data = self.cur.fetchall()
        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.allbooks_table.setItem(row, column, QTableWidgetItem(str(item)))
            current_row = self.allbooks_table.rowCount()
            self.allbooks_table.insertRow(current_row)

    def add_new_book(self):
        # self.db = MySQLdb.connect(host='localhost', user='root', passwd='Thebossm@#995', db="library", port=3310)
        # self.cur = self.db.cursor()  # Create a cursor
        book_title = self.bookTitle_LEdit.text()
        book_code = self.bookCode_LEdit.text()
        book_description = self.textEdit.toPlainText()
        book_category = self.categ_comboBox.currentText()
        book_author = self.author_comboBox.currentText()
        book_publisher = self.pub_comboBox.currentText()
        book_price = self.price_LEdit.text()

        self.cur.execute("""
        INSERT INTO book (book_name, book_description, book_code, book_category, book_author, book_publisher, book_price)
        VALUES (%s, %s, %s, %s, %s, %s, %s) 
        """, (book_title, book_description, book_code, book_category, book_author, book_publisher, book_price,))

        self.db.commit()

        self.show_all_books()

        # clear entry fields
        self.textEdit.setPlainText('')
        self.bookTitle_LEdit.setText('')
        self.bookCode_LEdit.setText('')
        self.categ_comboBox.setCurrentIndex(-1)  # returns the combobox selector to default value, 0
        self.author_comboBox.setCurrentIndex(-1)
        self.pub_comboBox.setCurrentIndex(-1)
        self.price_LEdit.setText('')

        # show message on the status bar
        self.statusBar().showMessage("New book added", 3000)

    def search_books(self):
        # noinspection PyBroadException

        book_title = self.search_query.text()

        sql = '''SELECT * FROM book WHERE book_name = %s'''
        self.cur.execute(sql, [book_title])

        data = self.cur.fetchone()
        if data:
            self.book_titleEdit.setText(data[1])
            self.description_LEdit.setPlainText(data[2])
            self.code_LEdit.setText(data[3])
            self.edit_cat_combo.setCurrentText(str(data[4]))
            self.edit_author_combo.setCurrentText(str(data[5]))
            self.edit_pub_combo.setCurrentText(str(data[6]))
            self.price_LEdit_2.setText(str(data[7]))
            self.statusBar().showMessage('Found result!', 3000)
        else:
            self.book_titleEdit.setText('')
            self.description_LEdit.setPlainText('')
            self.code_LEdit.setText('')
            self.edit_cat_combo.setCurrentIndex(-1)
            self.edit_author_combo.setCurrentIndex(-1)
            self.edit_pub_combo.setCurrentIndex(-1)
            self.price_LEdit_2.setText('')
            self.statusBar().showMessage('Data not found on DB !', 3000)

    def edit_books(self):
        book_title = self.book_titleEdit.text()
        book_code = self.code_LEdit.text()
        book_description = self.description_LEdit.toPlainText()
        book_category = self.edit_cat_combo.currentText()
        book_author = self.edit_author_combo.currentText()
        book_publisher = self.edit_pub_combo.currentText()
        book_price = self.price_LEdit_2.text()

        search_book_title = self.search_query.text()
        self.cur.execute('''
        UPDATE book SET 
            book_name=%s,
            book_description=%s, 
            book_code=%s, 
            book_category=%s, 
            book_author=%s,
            book_publisher=%s, 
            book_price=%s  
        WHERE book_name = %s
        ''', (book_title, book_description, book_code, book_category, book_author,
              book_publisher, book_price, search_book_title,))
        self.db.commit()
        self.show_all_books()
        self.statusBar().showMessage("Database updated successfully", 3000)

    def delete_books(self):
        book_title = self.search_query.text()

        warning = QMessageBox.warning(self, "Delete Book", "Are you sure you want to delete this book?",
                                      QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes:
            sql = """DELETE FROM book WHERE book_name=%s"""
            self.cur.execute(sql, book_title)
            self.db.commit()

            self.statusBar().showMessage("Record deleted successfully", 3000)
            self.show_all_books()

    ########################################################################################################
    ########################### Users      ###############################################################
    def add_new_user(self):
        username = self.add_username.text()
        email = self.add_email.text()
        password = self.add_password.text()
        repeat_password = self.repeat_password.text()

        if password == repeat_password:
            self.cur.execute("""
                INSERT INTO users 
                (user_name, user_email, user_password)
                VALUES (%s, %s, %s)
            """, (username, email, password))
            self.db.commit()
            self.statusBar().showMessage('New user Added!', 3000)

        else:
            self.password_error_label.setText('Passwords do not match !')

    def login(self):
        username = self.edit_username.text()
        password = self.edit_password.text()

        sql = '''SELECT * FROM users '''
        self.cur.execute(sql)
        data = self.cur.fetchall()
        for row in data:
            if username == row[1] and password == row[3]:
                # print("User found")
                self.statusBar().showMessage('User Account found', 3000)
                self.edit_groupBox.setEnabled(True)
                self.new_username.setText(row[1])
                self.new_email.setText(row[2])
                self.new_password.setText(row[3])
                break
            else:
                self.statusBar().showMessage("Invalid username or password. Try again", 5000)

    def edit_user(self):
        username = self.new_username.text()
        password = self.new_password.text()
        email = self.new_email.text()
        repeat_pword = self.repeat_new_password.text()

        old_username = self.edit_username.text()

        if password == repeat_pword:
            # sql = '''UPDATE users SET user_name=%s. user_email=%s. user_password=%s WHERE user_name=%s'''
            self.cur.execute('''UPDATE users SET user_name=%s, user_email=%s, user_password=%s WHERE user_name=%s''',
                             (username, email, password, old_username))
            self.db.commit()
            self.statusBar().showMessage('User data updated successfully', 3000)
            self.edit_groupBox.setEnabled(False)
        else:
            # print('New passwords do not match')
            self.statusBar().showMessage('New passwords do not match', 3000)

    ########################################################################################################
    ########################### Clients      ###############################################################

    def add_new_client(self):
        client_name = self.newClientName.text()
        client_email = self.newClientEmail.text()
        client_national_id = self.newClientNID.text()
        client_address = self.newClientAddress.toPlainText()

        # print(client_address, client_email, '\n' ,client_national_id, client_name)
        self.cur.execute('''
        INSERT INTO clients(client_name, client_email, national_id, perm_address) VALUES(%s, %s, %s, %s)
        ''', (client_name, client_email, client_national_id, client_address))

        self.db.commit()
        self.show_all_clients()
        self.statusBar().showMessage('New client successfully added', 3000)

    def edit_client_details(self):
        searchQuery = self.searchClientData.text()
        newClientName = self.editClientName.text()
        newClientEmail = self.editClientEmail.text()
        newClientNID = self.editClientNID.text()
        newClientPermAddress = self.newClientAddress_2.toPlainText()

        self.cur.execute("""UPDATE clients SET 
            client_name=%s,
            client_email=%s,
            national_id=%s,
            perm_address=%s 
            WHERE national_id=%s""",
                         (newClientName, newClientEmail, newClientNID, newClientPermAddress, searchQuery))

        self.db.commit()
        self.show_all_clients()
        self.statusBar().showMessage('Client details updated successfully', 3000)

    def show_all_clients(self):
        self.cur.execute("""
            SELECT client_name, client_email, national_id, perm_address FROM clients""")
        data = self.cur.fetchall()

        if data:  # check if there is data in the database
            self.allclients_table.setRowCount(0)  # This clears the table
            self.allclients_table.insertRow(0)  # start inserting at the first column
            for row, form in enumerate(data):
                # print(row, form)
                for column, item in enumerate(form):
                    # print(column, item)
                    # this adds items into the table widget at specified column and row
                    self.allclients_table.setItem(row, column, QTableWidgetItem(str(item)))
                    # column += 1

                row_position = self.allclients_table.rowCount()
                self.allclients_table.insertRow(row_position)

    def delete_client(self):
        search_NID = self.searchClientData.text()

        warning = QMessageBox.warning(self, "Delete Client", "Are you sure you want to delete this Client?",
                                      QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes:
            sql = """DELETE FROM clients WHERE national_id=%s"""
            self.cur.execute(sql, search_NID)
            self.db.commit()
            self.show_all_clients()
            self.statusBar().showMessage("Client details deleted successfully", 3000)

    def search_clients(self):
        searchQuery = self.searchClientData.text()

        self.cur.execute("""
            SELECT * FROM clients WHERE national_id=%s
        """, searchQuery)
        data = self.cur.fetchone()
        print(data)
        if data:
            self.editClientName.setText(data[1])
            self.editClientEmail.setText(data[2])
            self.editClientNID.setText(data[3])
            self.newClientAddress_2.setPlainText(data[4])
        else:
            self.statusBar().showMessage('Client not found, please try again', 3000)

    ########################################################################################################
    ###########################  Settings    ###############################################################

    def add_category(self):
        category_name = self.setings_categ_LEdit.text()
        if category_name:
            self.cur.execute("INSERT INTO category (category_name) VALUES (%s)", (category_name,))
            self.db.commit()
            self.statusBar().showMessage("New Category successfully added !", 3000)
            self.show_category_combobox()
            self.show_category()
            self.setings_categ_LEdit.setText('')  # clear the line edit
        else:
            self.statusBar().showMessage('Entry field should not be empty', 3000)

    def show_category(self):
        self.cur.execute('''SELECT category_name from category ''')
        data = self.cur.fetchall()
        # print(data)

        if data:  # check if there is data in the database
            self.categ_tableWidget.setRowCount(0)  # This clears the table
            self.categ_tableWidget.insertRow(0)
            for row, form in enumerate(data):
                # print(row, form)
                for column, item in enumerate(form):
                    # print(column, item)
                    # this adds items into the table widget at specified column and row
                    self.categ_tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1

                row_position = self.categ_tableWidget.rowCount()
                self.categ_tableWidget.insertRow(row_position)

    def add_author(self):
        author_name = self.settings_authorLedit.text()
        if author_name:
            self.cur.execute("INSERT INTO authors (author_name) VALUES (%s)", (author_name,))
            self.db.commit()
            self.settings_authorLedit.setText('')  # clear the line edit
            self.show_author()
            self.show_author_combobox()
            self.statusBar().showMessage("New author successfully added !", 3000)
        else:
            self.statusBar().showMessage('Entry field should not be empty', 3000)

    def show_author(self):
        self.cur.execute('''SELECT author_name from authors ''')
        data = self.cur.fetchall()
        # print(data)

        if data:  # check if there is data in the database
            self.author_tableWidget.setRowCount(0)  # so that each time the script is run, extra rows are not added
            self.author_tableWidget.insertRow(0)
            for row, form in enumerate(data):
                # print(row, form)
                for column, item in enumerate(form):
                    # print(column, item)
                    # this adds items into the table widget at specified column and row
                    self.author_tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1

                row_position = self.author_tableWidget.rowCount()
                self.author_tableWidget.insertRow(row_position)

    def add_publisher(self):
        publisher_name = self.pubName_LEdit.text()
        if publisher_name:
            self.cur.execute("INSERT INTO publisher (publisher_name) VALUES (%s)", (publisher_name,))
            self.db.commit()
            self.pubName_LEdit.setText('')  # clear the line edit
            self.show_publisher()
            self.show_publisher_combobox()
            self.statusBar().showMessage("New publisher successfully added !", 3000)
        else:
            self.statusBar().showMessage('Entry field should not be empty', 3000)

    def show_publisher(self):  # display publisher table details on table widget
        self.cur.execute('''SELECT publisher_name from publisher ''')
        data = self.cur.fetchall()
        # print(data)

        if data:  # check if there is data in the database
            self.pub_tableWidget.setRowCount(0)  # so that each time the script is run, extra rows are not added, clear
            self.pub_tableWidget.insertRow(0)  # start inserting data at the 0th index of the table widget
            for row, form in enumerate(data):  # enumerate returns an iterable object(tuple) with count value
                # print(row, form)
                for column, item in enumerate(form):
                    # print(column, item)
                    # this adds items into the table widget at specified column and row
                    self.pub_tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1

                row_position = self.pub_tableWidget.rowCount()
                self.pub_tableWidget.insertRow(row_position)  # sets the row for the next entry

    ###############################################################################################################
    ######################################## Show settings data in combobox ######################################
    def show_category_combobox(self):
        self.cur.execute('''SELECT category_name FROM category''')
        data = self.cur.fetchall()

        self.categ_comboBox.clear()  # clear the combobox first to avoid appending the data to existing data
        for category in data:  # iterate through the received data from the data base
            self.categ_comboBox.addItem(category[0])  # add item to the combo box
            self.edit_cat_combo.addItem(category[0])

    def show_author_combobox(self):
        self.cur.execute('''SELECT author_name FROM authors''')
        data = self.cur.fetchall()

        self.author_comboBox.clear()
        for author in data:
            self.edit_author_combo.addItem(author[0])
            self.author_comboBox.addItem(author[0])

    def show_publisher_combobox(self):
        self.cur.execute('''SELECT publisher_name FROM publisher''')
        data = self.cur.fetchall()

        self.pub_comboBox.clear()  # clear the combobox before appending data to it
        for publisher in data:
            self.pub_comboBox.addItem(publisher[0])
            self.edit_pub_combo.addItem(publisher[0])

    #######################################################################################################################
    ############################################## handling popups ######################################################
    # def show_pop_up(self, **kwargs):
    #     message = kwargs['message'] if 'message' in kwargs else 'Confirm your action'
    #     title = kwargs['title'] if 'title' in kwargs else 'Error'
    #     icon = kwargs['type'] if 'type' in kwargs else 'information'
    #     response_type = kwargs['response'] if 'response' in kwargs else 'ok'
    #
    #     popup = QMessageBox()
    #     popup.setWindowTitle(title)
    #     popup.setText(message)
    #
    #     # choose the icon type
    #     if icon == 'information':
    #         popup.setIcon(QMessageBox.Information)
    #     elif icon == 'warning':
    #         popup.setIcon(QMessageBox.Warning)
    #     elif icon == 'critical':
    #         popup.setIcon(QMessageBox.Critical)
    #     elif icon == 'question':
    #         popup.setIcon(QMessageBox.Question)
    #     else:
    #         print('Error')
    #
    #     # chose response types
    #     if response_type == 'yesno':
    #         popup.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    #     if response_type == 'ignoreRetry':
    #         popup.setStandardButtons(QMessageBox.Ignore | QMessageBox.Abort)
    #
    #     x = popup.exec_()

    ###############################################################################################################
    ########################################     Export to Excel   ################################################

    def exportOperations(self):
        self.cur.execute('''
        SELECT book_name, client_name, type, operation_time, date, to_date FROM dayoperations
        ''')
        data = self.cur.fetchall()
        # create a new excel workbook
        wb = Workbook('Operations.xlsx')
        # create a new sheet in the workbook to store data
        sheet_1 = wb.add_worksheet("Today's data")

        # Create columns and give them names
        sheet_1.write(0, 0, 'Book Title')
        sheet_1.write(0, 1, 'Client Name')
        sheet_1.write(0, 2, 'Type')
        sheet_1.write(0, 3, 'Time of Operation')
        sheet_1.write(0, 4, 'Date Burrowed')
        sheet_1.write(0, 0, 'Day to return')

        # write data in the created columns using a nested for
        row_number = 1  # remember row 1 has already been used for the column headers
        for row in data:
            column_number = 0  # Each time this loop is iterated start at the first column
            for item in row:
                sheet_1.write(row_number, column_number, str(item))
                column_number += 1  # increase the column number for each line of entry
            row_number += 1   # add data to the next row after each iteration of the loop is completed
            wb.close()  # book must be closed
            self.statusBar().showMessage('Data successfully exported to excel sheet', 5000)

    def exportBooks(self):

        self.cur.execute('''
        SELECT book_code, book_name, book_category, book_description, book_author, book_publisher, book_price FROM book
        ''')
        data = self.cur.fetchall()
        # create a new excel workbook
        wb = Workbook('Books.xlsx')
        # create a new sheet in the workbook to store data
        sheet_1 = wb.add_worksheet("Today's data")

        # Create columns and give them names
        sheet_1.write(0, 0, 'Book Code')
        sheet_1.write(0, 1, 'Book Name')
        sheet_1.write(0, 2, 'Book Category')
        sheet_1.write(0, 3, 'Book Description')
        sheet_1.write(0, 4, 'Book Author')
        sheet_1.write(0, 5, 'Book Publisher')
        sheet_1.write(0, 6, 'Book Price')

        # write data in the created columns using a nested for
        row_number = 1  # remember row 1 has already been used for the column headers
        for row in data:
            column_number = 0  # Each time this loop is iterated start at the first column
            for item in row:
                sheet_1.write(row_number, column_number, str(item))
                column_number += 1  # increase the column number for each line of entry
            row_number += 1  # add data to the next row after each iteration of the loop is completed
            wb.close()  # book must be closed
            self.statusBar().showMessage('Data successfully exported to excel sheet', 5000)

    def exportClients(self):
        self.cur.execute("""
            SELECT client_name, client_email, national_id, perm_address FROM clients""")
        data = self.cur.fetchall()

        # create a new excel workbook
        wb1 = Workbook('Clients.xlsx')
        # create a new sheet in the workbook to store data
        sheet_1 = wb1.add_worksheet("Today's data")

        # Create columns and give them names
        sheet_1.write(0, 0, 'Client Name')
        sheet_1.write(0, 1, 'Client Email')
        sheet_1.write(0, 2, 'National ID')
        sheet_1.write(0, 3, 'Permanent Address')
        sheet_1.write(0, 4, 'Date Burrowed')

        # write data in the created columns using a nested for
        row_number = 1  # remember row 1 has already been used for the column headers
        for row in data:
            column_number = 0  # Each time this loop is iterated start at the first column
            for item in row:
                sheet_1.write(row_number, column_number, str(item))
                column_number += 1  # increase the column number for each line of entry
            row_number += 1  # add data to the next row after each iteration of the loop is completed
            wb1.close()  # book must be closed
            self.statusBar().showMessage('Data successfully exported to excel sheet', 5000)

    ###############################################################################################################
    ########################################     UI Themes   ####################################################
    def aqua_theme(self):
        style = open('themes/Aqua.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def console_style(self):
        style = open('themes/ConsoleStyle.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def elegant_dark_theme(self):
        style = open('themes/ElegantDark.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def material_dark_theme(self):
        style = open('themes/MaterialDark.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def manjaromix_theme(self):
        style = open('themes/ManjaroMix.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def ubuntu_theme(self):
        style = open('themes/Ubuntu.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def amoled_theme(self):
        style = open('themes/AMOLED.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    # # def __del__(self):
    #     self.db.close()
    #     print('Connection closed')


def main():
    """
    # Only one instance of the QApplication is needed per app. sys.argv allows you pass command line arguments for the
    # app. If you don't need this, leave it empty.as qApplication([])
    The app.exec() creates an event loop. any statement before this statement is executed before the app starts running
    Anything after it would not run until the app is exited.

    """

    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()
