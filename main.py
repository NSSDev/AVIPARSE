from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QAbstractItemView
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time
from tkinter import messagebox
from selenium.webdriver.common.by import By
import os
from datetime import datetime
import pytz
from PIL import Image
from pytesseract import pytesseract
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import os.path


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(428, 321)
        MainWindow.setStyleSheet("background-color: rgb(53, 53, 53);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.listWidget = QtWidgets.QListWidget(self.centralwidget)
        self.listWidget.setGeometry(QtCore.QRect(20, 60, 391, 151))
        font = QtGui.QFont()
        font.setFamily("Rockwell Condensed")
        self.listWidget.setFont(font)
        self.listWidget.setStyleSheet("color: rgb(255, 255, 255);")
        self.listWidget.setObjectName("listWidget")
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 40, 81, 16))
        font = QtGui.QFont()
        font.setFamily("Segoe Print")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setStyleSheet("color: rgb(255, 255, 255);")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(160, 0, 111, 41))
        font = QtGui.QFont()
        font.setFamily("Niagara Solid")
        font.setPointSize(36)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(380, 260, 47, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setStyleSheet("color: rgb(255, 255, 255);")
        self.label_3.setObjectName("label_3")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(90, 230, 241, 41))
        font = QtGui.QFont()
        font.setFamily("Segoe Print")
        font.setPointSize(11)
        self.pushButton.setFont(font)
        self.pushButton.setStyleSheet("background-color: rgb(0, 170, 0);\n"
                                      "color: rgb(255, 255, 255);")
        self.pushButton.setObjectName("pushButton")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 428, 21))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.action = QtWidgets.QAction(MainWindow)
        font = QtGui.QFont()
        font.setFamily("Segoe Print")
        self.action.setFont(font)
        self.action.setObjectName("action")
        self.action_2 = QtWidgets.QAction(MainWindow)
        font = QtGui.QFont()
        font.setFamily("Segoe Print")
        self.action_2.setFont(font)
        self.action_2.setObjectName("action_2")
        self.menu.addAction(self.action)
        self.menu.addAction(self.action_2)
        self.menubar.addAction(self.menu.menuAction())
        self.action_2.triggered.connect(self.exit_action)
        self.action.triggered.connect(self.author)
        self.pushButton.clicked.connect(self.main)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "AVIPARSE"))
        __sortingEnabled = self.listWidget.isSortingEnabled()
        self.listWidget.setSortingEnabled(False)
        item = self.listWidget.item(0)
        item.setText(_translate("MainWindow", "Транспорт и перевозки"))
        item = self.listWidget.item(1)
        item.setText(_translate("MainWindow", "Ремонт, строительство"))
        item = self.listWidget.item(2)
        item.setText(_translate("MainWindow", "Красота, здоровье"))
        item = self.listWidget.item(3)
        item.setText(_translate("MainWindow", "Ремонт техники"))
        item = self.listWidget.item(4)
        item.setText(_translate("MainWindow", "Уборка, дезинфекция"))
        item = self.listWidget.item(5)
        item.setText(_translate("MainWindow", "Оборудование, производство"))
        item = self.listWidget.item(6)
        item.setText(_translate("MainWindow", "Обучение, курсы"))
        item = self.listWidget.item(7)
        item.setText(_translate("MainWindow", "Деловые услуги"))
        item = self.listWidget.item(8)
        item.setText(_translate("MainWindow", "Установка техники"))
        item = self.listWidget.item(9)
        item.setText(_translate("MainWindow", "Мастер на час"))
        item = self.listWidget.item(10)
        item.setText(_translate("MainWindow", "Сад, участок"))
        item = self.listWidget.item(11)
        item.setText(_translate("MainWindow", "Бытовые услуги"))
        item = self.listWidget.item(12)
        item.setText(_translate("MainWindow", "Праздники, мероприятия"))
        item = self.listWidget.item(13)
        item.setText(_translate("MainWindow", "IT, интернет, телеком"))
        item = self.listWidget.item(14)
        item.setText(_translate("MainWindow", "Реклама, полиграфия"))
        item = self.listWidget.item(15)
        item.setText(_translate("MainWindow", "Главная страница"))
        self.listWidget.setSortingEnabled(__sortingEnabled)
        self.label.setText(_translate("MainWindow", "УСЛУГИ:"))
        self.label_2.setText(_translate("MainWindow", "AVIPARSE"))
        self.label_3.setText(_translate("MainWindow", "V0.1.1"))
        self.pushButton.setText(_translate("MainWindow", "Получить данные"))
        self.menu.setTitle(_translate("MainWindow", "Меню"))
        self.action.setText(_translate("MainWindow", "Автор"))
        self.action_2.setText(_translate("MainWindow", "Закрыть"))

    def author(self):
        messagebox.showinfo(title="AVIPARSE v0.1.1", message="Разработчик программы:\nTelegram: @HasimNas")

    def exit_action(self):
        MainWindow.close()

    def select_file_directory(self):
        self.file_path = QFileDialog.getExistingDirectory(None, caption='Выберите папку для сохранения картинок',
                                                          directory=os.getcwd())

    def select_json_directory(self):
        self.json_path = QFileDialog.getExistingDirectory(None, caption='Выберите путь с .json файлом', directory=os.getcwd())

    def choose_category(self):
        self.select = self.listWidget.currentRow()
        if self.select == 0:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/transport_perevozki-ASgBAgICAUSYC8SfAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 1:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/remont_stroitelstvo-ASgBAgICAUSYC8CfAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 2:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/krasota_zdorove-ASgBAgICAUSYC6qfAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 3:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/remont_i_obsluzhivanie_tehniki-ASgBAgICAUSYC7T3AQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 4:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/uborka_klining-ASgBAgICAUSYC7L3AQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 5:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/oborudovanie_proizvodstvo-ASgBAgICAUSYC7SfAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 6:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/obuchenie_kursy-ASgBAgICAUSYC7afAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 7:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/delovye_uslugi-ASgBAgICAUSYC7KfAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 8:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/ustanovka_tehniki-ASgBAgICAUSYC9r1AQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 9:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/master_na_chas-ASgBAgICAUSYC7zvAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 10:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/sad_blagoustroystvo-ASgBAgICAUSYC8KfAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 11:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/bytovye_uslugi-ASgBAgICAUSYC7CfAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 12:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/prazdniki_meropriyatiya-ASgBAgICAUSYC7yfAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 13:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/it_internet_telekom-ASgBAgICAUSYC6afAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 14:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug/reklama_poligrafiya-ASgBAgICAUSYC76fAQ?context=H4sIAAAAAAAA_0q0MrKqLraysFJKK8rPDUhMT1WyLrYyNLNSKipNKspMTizJLwrPTElPLQGLG1gplSQWpaeWwFQaWCkpWdcCAgAA__-OZLTGRgAAAA&p=2"
        if self.select == 15:
            self.url = "https://www.avito.ru/rossiya/predlozheniya_uslug?cd=1&isVertical=1"

    def collect_info(self):
        for r, d, f in os.walk("C:\\"):
            for files in f:
                if files == "geckodriver.exe":
                    self.driver_path = os.path.join(r, files)
        options = webdriver.FirefoxOptions()
        options.add_argument("--headless")
        self.driver = webdriver.Firefox(executable_path=self.driver_path, options=options)
        self.driver.maximize_window()
        counter = 0
        page_counter = 0
        while counter != 500:
            try:
                counter += 1
                page_counter += 1
                self.driver.get(url=f"{self.url[:-1]}{page_counter}")
                time.sleep(1)
                if self.url[-9:] == "authsrc=w":
                    continue
                if page_counter == 95:
                    page_counter = 0
                time.sleep(7)
                if self.select == 15:
                    items = self.driver.find_elements(By.CLASS_NAME, "styles-list-M1e9B")
                else:
                    items = self.driver.find_elements(By.CLASS_NAME, "items-items-kAJAg")
                category_name = self.driver.find_element(By.CLASS_NAME, "page-title-text-tSffu").text
                # до этого определяем какую подкатегорию выбрал пользователь
                for item in items:
                    links = item.find_elements(By.TAG_NAME, "a")
                    page_url = [link.get_attribute('href') for link in links]
                    # получаем каждую ссылку на услугу
                    for item_url in page_url:  # если в конце ссылки имеется указанный атрибут,то такую ссылку пропускаем
                        if item_url[-9:] == "authsrc=w":
                            continue
                        self.driver.get(url=item_url)
                        # переходим на ссылку с услугой
                        id = item_url[-10:]
                        if self.driver.find_element(By.XPATH, "//div[@data-marker='seller-info/label']").is_displayed():

                            status = self.driver.find_element(By.XPATH, "//div[@data-marker='seller-info/label']").text
                            region = self.driver.find_element(By.XPATH,
                                                              "//span[@class='style-item-address__string-wt61A']").text
                            # если там отображено кнопка с телефоном,то продолжаем
                            self.driver.find_element(By.XPATH,
                                                     "//button[@class='button-button-eBrUW button-button_phone-_Yo3v button-button_card-AkthM button-button-CmK9a button-size-s-r9SeD button-success-_ytiQ']").click()
                            time.sleep(2)
                            image = self.driver.find_element(By.XPATH, "//img[@alt='phone']").get_attribute('src')
                            self.driver.get(image)
                            if counter == 30:
                                time.sleep(30)
                            if self.driver.current_url[21] == "i":
                                continue
                            self.driver.save_screenshot(f"{self.file_path}/{id}.png")
                            time.sleep(1)
                            numbers_to_text = Image.open(f"{self.file_path}/{id}.png")
                            numbers_to_text.save(f"{self.file_path}/{id}.png")
                            moscow_time = datetime.now(pytz.timezone('Europe/Moscow')).strftime(
                                "%d/%m/%Y, %H:%M:%S")
                            path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
                            pytesseract.tesseract_cmd = path_to_tesseract
                            number = pytesseract.image_to_string(numbers_to_text)


                            scope = ["https://spreadsheets.google.com/feeds",
                                     'https://www.googleapis.com/auth/spreadsheets',
                                     "https://www.googleapis.com/auth/drive.file",
                                     "https://www.googleapis.com/auth/drive"]
                            creds = ServiceAccountCredentials.from_json_keyfile_name(f"{self.json_path}/for_project.json", scope)
                            spreadsheetId = '1TTB7xi0PgMS1Y0BpKHOgy3jLD7LTZdgJAGn5ojBsseQ'
                            sheetName = 'ContactsData'
                            client = gspread.authorize(creds)
                            sh = client.open_by_key(spreadsheetId)
                            df2 = pd.DataFrame(
                                {'col1': [id], 'col2': [region], 'col3': [status], 'col4': [category_name],
                                 'col5': [number], 'col6': [moscow_time], 'col7':[item_url]})
                            worksheet = sh.worksheet(sheetName)
                            values = df2.values.tolist()
                            x = [item for item in worksheet.col_values(1) if item]
                            if id in x:
                                continue
                            else:
                                sh.values_append(sheetName, {'valueInputOption': 'USER_ENTERED'}, {'values': values})
                                continue



            except NoSuchElementException:
                pass

    def main(self):
        self.select_file_directory()
        if not self.file_path:
            messagebox.showerror(title="AVIPARSE",message="Вы не выбрали папку для сохранения картинок !")
            return
        self.select_json_directory()
        if not self.json_path:
            messagebox.showerror(title="AVIPARSE",message="Вы не выбрали путь с .json файлом !")
            return
        self.choose_category()  # проверка на выбор подкатегории
        self.collect_info()  # основной метод для парсинга данных
        messagebox.showinfo(title="AVIPARSE",
                            message="Программа закончила свою работу")  # по завершению закрываем selenium
        self.driver.close()


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
