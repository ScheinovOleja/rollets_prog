import configparser
import csv
import datetime
import os
import re
import smtplib
import sys
import xml.etree.cElementTree
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from docxtpl import DocxTemplate
from pony.orm import db_session
from win32com import client

from design import Ui_MainWindow
from import_to_db import add_to_db
from models import DataNumDoc, AutomationData, run


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        if not os.path.isfile('config.cfg'):
            QMessageBox.critical(self, 'Ошибка', f'Не найден файл config.cfg!\nПоместите его в папку с программой и '
                                                 f'повторите попытку!')
            sys.exit(0)
        self.manager = configparser.ConfigParser()
        self.manager.read(f'{os.getcwd()}\\config.cfg', encoding='utf-8')
        self.xml_file = None
        self.csv_file = None
        self.path_png = None
        self.setWindowIcon(QIcon('order.ico'))
        self.buttonBox = QButtonGroup()
        self.buttonBox_2 = QButtonGroup()
        run()
        self.init_ui()
        self.connect_ui()
        self.location_on_the_screen()
        self.show()

    def location_on_the_screen(self):
        screen = QGuiApplication.screenAt(QCursor().pos())
        fg = self.frameGeometry()
        fg.moveCenter(screen.geometry().center())
        self.move(fg.topLeft())

    def init_ui(self):
        self.buttonBox.addButton(self.ui.radioButton)
        self.ui.radioButton.setChecked(True)
        self.buttonBox.addButton(self.ui.radioButton_2)
        self.buttonBox_2.addButton(self.ui.radioButton_4)
        self.buttonBox_2.addButton(self.ui.radioButton_3)
        self.add_item_to_combobox()
        self.ui.dateEdit.setMaximumDate(datetime.datetime.now().date())
        self.ui.dateEdit.setDate(datetime.datetime.now().date())
        self.ui.radioButton.setText(self.manager['manager_1']['full_name_manager'])
        self.ui.radioButton_2.setText(self.manager['manager_2']['full_name_manager'])
        self.get_num_doc()

    def connect_ui(self):
        self.ui.pushButton_3.clicked.connect(self.add_services)
        self.ui.treeWidget.itemDoubleClicked.connect(self.delete_item_1)
        self.ui.treeWidget_2.itemDoubleClicked.connect(self.delete_item_2)
        self.ui.pushButton_4.clicked.connect(self.add_automatic)
        self.ui.action_automatic.triggered.connect(self.add_automatic_to_db)
        self.ui.dateEdit.dateChanged.connect(self.get_num_doc)
        self.ui.pushButton_2.clicked.connect(self.send_to_email)
        self.ui.pushButton_5.clicked.connect(self.send_commercial_to_mail)

    def all_clear(self):
        self.ui.lineEdit.clear()
        self.ui.lineEdit_2.clear()
        self.ui.lineEdit_3.clear()
        self.ui.lineEdit_4.clear()
        self.ui.doubleSpinBox.setValue(0.0)
        self.ui.doubleSpinBox_2.setValue(0.0)
        self.ui.doubleSpinBox_3.setValue(0.0)
        self.ui.doubleSpinBox_4.setValue(0.0)
        self.ui.spinBox_2.setValue(0)
        self.ui.spinBox_3.setValue(0)
        self.ui.treeWidget.clear()
        self.ui.treeWidget_2.clear()

    @db_session
    def add_item_to_combobox(self):
        self.ui.comboBox.clear()
        all_automatic = AutomationData.select().order_by(AutomationData.id)[:]
        for item in all_automatic:
            self.ui.comboBox.addItem(f'{item.name_automatic}/{item.price_automatic}')

    def add_automatic_to_db(self):
        try:
            file = QFileDialog.getOpenFileName(self, 'Open file', f'{os.getcwd()}', 'CSV Files (*.csv)')[0]
            if file == '':
                return QMessageBox.information(self, 'Неувязочка!', f'Вы не выбрали ни одного файла!')
            else:
                add_to_db(file, False)
                self.add_item_to_combobox()
                QMessageBox.about(self, "Отлично!", "Вы успешно загрузили данные об автоматике!")
        except Exception as exc:
            QMessageBox.critical(self, 'Ошибка', f'Произошла ошибка!\n{exc}')
            return

    def load_xml(self):
        try:
            self.xml_file = QFileDialog.getOpenFileName(self, 'Загрузить роллеты!', f'{os.getcwd()}',
                                                        'XML Files (*.xml)')[0]
            if self.xml_file == '':
                QMessageBox.information(self, 'Неувязочка!', f'Вы не выбрали ни одного файла!')
            else:
                QMessageBox.about(self, "Отлично!", "Вы успешно загрузили данные о роллетах!")
        except Exception as exc:
            QMessageBox.critical(self, 'Ошибка!', f'Произошла ошибка!\n{exc}')

    def load_csv(self):
        try:
            self.csv_file = QFileDialog.getOpenFileName(self, 'Загрузить цены!', f'{os.getcwd()}',
                                                        'CSV Files (*.csv)')[0]
            if self.csv_file == '':
                QMessageBox.information(self, 'Неувязочка!', f'Вы не выбрали ни одного файла!')
            else:
                QMessageBox.about(self, "Отлично!", "Вы успешно загрузили данные о ценах!")
        except Exception as exc:
            QMessageBox.critical(self, 'Ошибка!', f'Произошла ошибка!\n{exc}')

    def add_automatic(self):
        if self.ui.spinBox_3.value() == 0:
            return QMessageBox.information(self, 'Неувязочка!', f'Зачем добавлять нулевое количество товара?)')
        item = self.ui.comboBox.currentText()
        index = self.ui.comboBox.currentIndex()
        self.ui.comboBox.removeItem(index)
        automatic = item.split('/')[0]
        price = float(item.split('/')[1])
        count = self.ui.spinBox_3.value()
        rowcount = self.ui.treeWidget_2.topLevelItemCount()
        self.ui.treeWidget_2.addTopLevelItem(QTreeWidgetItem(rowcount))
        self.ui.treeWidget_2.topLevelItem(rowcount).setText(0, automatic)
        self.ui.treeWidget_2.topLevelItem(rowcount).setText(1, str(count))
        self.ui.treeWidget_2.topLevelItem(rowcount).setText(2, str(round(price * count, 2)))
        self.ui.spinBox_3.setValue(0)

    def add_services(self):
        text = self.ui.lineEdit_4.text()
        if text == '':
            self.ui.doubleSpinBox_4.setValue(0.00)
            return QMessageBox.information(self, 'Неувязочка!', f'Вы не ввели название услуги!')
        price = self.ui.doubleSpinBox_4.text()
        if price == '0,00':
            self.ui.lineEdit_4.clear()
            return QMessageBox.information(self, 'Неувязочка!', f'Вы не ввели стоимость услуги!')
        rowcount = self.ui.treeWidget.topLevelItemCount()
        self.ui.treeWidget.addTopLevelItem(QTreeWidgetItem(rowcount))
        self.ui.treeWidget.topLevelItem(rowcount).setText(0, text)
        self.ui.treeWidget.topLevelItem(rowcount).setText(1, price.replace(',', '.'))
        self.ui.lineEdit_4.clear()
        self.ui.doubleSpinBox_4.setValue(0.00)

    def delete_item_1(self):
        item = self.ui.treeWidget.currentItem()
        self.ui.treeWidget.takeTopLevelItem(self.ui.treeWidget.indexOfTopLevelItem(item))

    @db_session
    def delete_item_2(self):
        item = self.ui.treeWidget_2.currentItem()
        self.ui.comboBox.addItem(f'{item.text(0)}/{float(item.text(2)) / int(item.text(1))}')
        self.ui.treeWidget_2.takeTopLevelItem(self.ui.treeWidget_2.indexOfTopLevelItem(item))

    @db_session
    def get_num_doc(self) -> int:
        date = DataNumDoc.select().where(date=self.ui.dateEdit.date().toString("dd.MM.yyyy")).order_by(DataNumDoc.id)[:]
        if len(date) == 0:
            self.ui.spinBox_4.setValue(1)
            return self.ui.spinBox_4.value()
        else:
            self.ui.spinBox_4.setValue(len(date) + 1)
            return self.ui.spinBox_4.value()

    @staticmethod
    def counting_price_rolls(csv_file):
        with open(csv_file, 'r', encoding='cp1251') as file:
            reader = csv.reader(file, delimiter=';')
            i = 0
            all_prices = []
            for item in reader:
                if len(item) == 0:
                    continue
                else:
                    try:
                        if item[10] == '':
                            continue
                        i += 1
                        if i < 3:
                            continue
                        try:
                            all_prices.append(float(item[10].replace(',', '.').replace(u'\xa0', '')))
                        except Exception as exc:
                            all_prices.pop(-1)
                            break
                    except Exception as exc:
                        continue
            return all_prices

    @db_session
    def unload_all_rolls(self):
        all_rolls = []
        tree = xml.etree.cElementTree.parse(self.xml_file)
        test = tree.findall('Orders/Orders_Rec/Rolletes/Rolletes_Rec')
        i = 0
        total_price = 0
        for item in test:
            i += 1
            width = float(item.attrib["WIDTH"].replace(',', '.'))
            height = float(item.attrib["HEIGHT"].replace(',', '.'))
            price = self.counting_price_rolls(self.csv_file)
            if not item.attrib['WIND_SPEED_'] == 'null':
                wind_speed = f"({item.attrib['WIND_SPEED_']} м/с)"
            else:
                wind_speed = ''
            data = {
                'num': i, 'cols': [
                    int(round(width * 1000, 0)),
                    int(round(height * 1000, 0)),
                    item.attrib["PROFILE_"],
                    item.attrib["SHAFT_"],
                    item.attrib["BASKET_"],
                    item.attrib["BLOCK_"],
                    item.attrib["DRIVE_"],
                    f"{item.attrib['WIND_ZONE_DESCRIPTION_']}"
                    f"{wind_speed}",
                    item.attrib["CNT"],
                    price[i - 1]
                ],
                'gear': item.attrib["GEAR_"]
            }
            all_rolls.append(data)
            total_price += price[i - 1]
        return all_rolls, round(total_price, 2)

    def get_all_services(self):
        root = self.ui.treeWidget.invisibleRootItem()
        child_count = root.childCount()
        all_services = []
        for i in range(child_count):
            item = root.child(i)
            data = {
                'name': item.text(0),
                'price': item.text(1).replace(',', '.')
            }
            all_services.append(data)
        return all_services

    def get_all_automatic(self):
        root = self.ui.treeWidget_2.invisibleRootItem()
        child_count = root.childCount()
        all_automatic = []
        total_price_automatic = 0
        for i in range(child_count):
            item = root.child(i)
            data = {
                'name': item.text(0),
                'count': item.text(1),
                'price': item.text(2).replace(',', '.')
            }
            all_automatic.append(data)
            total_price_automatic += float(item.text(2))
        return all_automatic, total_price_automatic

    def get_price_services(self):
        root = self.ui.treeWidget_2.invisibleRootItem()
        child_count = root.childCount()
        total_price_services = 0
        for i in range(child_count):
            item = root.child(i)
            total_price_services += float(item.text(1))
        return total_price_services

    def get_manager(self):
        buttons = self.buttonBox.buttons()
        if buttons[0].isChecked():
            return {
                'full_name_manager': self.manager['manager_1']['full_name_manager'],
                'email_manager': self.manager['manager_1']['email_manager'],
                'phone_manager': self.manager['manager_1']['phone_manager']
            }
        else:
            return {
                'full_name_manager': self.manager['manager_2']['full_name_manager'],
                'email_manager': self.manager['manager_2']['email_manager'],
                'phone_manager': self.manager['manager_2']['phone_manager']
            }

    def get_person(self):
        buttons = self.buttonBox_2.buttons()
        if buttons[0].isChecked():
            return True
        else:
            return False

    def get_all_info(self):
        date = self.ui.dateEdit.date().toString("dd.MM.yyyy")
        num = self.ui.spinBox_4.value()
        if num < 10:
            num = f'0{num}'
        try:
            all_rolls, total_price_rolls = self.unload_all_rolls()
        except TypeError:
            QMessageBox.information(self, 'Неувязочка!', f'Вы не загрузили XML файл!')
            return
        all_services = self.get_all_services()
        all_automatic, total_price_automatic = self.get_all_automatic()
        total_price_services = self.get_price_services()
        total_price_rolls += total_price_automatic
        manager = self.get_manager()
        return date, num, all_rolls, round(total_price_rolls,
                                           2), all_services, all_automatic, manager, total_price_services

    @db_session
    def new_doc(self, date, num):
        if DataNumDoc.get(date=date, num_doc=num):
            pass
        else:
            DataNumDoc(date=date, num_doc=num)

    @db_session
    def check_all_object(self):
        if self.ui.lineEdit.text() == '':
            QMessageBox.information(self, 'Неувязочка', f'Вы не ввели ФИО заказчика!')
            return False
        if self.ui.lineEdit_2.text() == '':
            QMessageBox.information(self, 'Неувязочка', f'Вы не ввели адрес заказчика!')
            return False
        if self.ui.lineEdit_3.text() == '':
            QMessageBox.information(self, 'Неувязочка', f'Вы не ввели email заказчика!')
            return False
        if not re.match(r'[\w.-]+@[\w.-]+(\.[\w]+)+', self.ui.lineEdit_3.text()):
            QMessageBox.critical(self, 'Ошибка', f'Вы ввели неверный email')
            return
        if self.ui.doubleSpinBox_3.value() == 0.0:
            QMessageBox.information(self, 'Неувязочка', f'Вы не ввели итоговую стоимость заказа!!')
            return False
        if self.ui.spinBox_2.value() == 0:
            QMessageBox.information(self, 'Неувязочка', f'Вы не выставили срок производства!')
            return False
        self.load_xml()
        if self.xml_file == '':
            return False
        self.load_csv()
        if self.csv_file == '':
            return False
        return True

    @staticmethod
    def round_to_ten(num):
        remaining = num % 10
        if remaining in range(0, 3):
            return num - remaining
        return num + (0 - remaining), round(remaining, 2)

    @staticmethod
    def get_percent(num, percent):
        return round((percent * 100) / num, 2)

    def print_to_pdf(self, doc_name, pdf_name):
        try:
            word = client.gencache.EnsureDispatch('Word.Application')
            word.Visible = False
            if os.path.exists(pdf_name):
                os.remove(pdf_name)
            word_doc = word.Documents.Open(doc_name)
            word_doc.SaveAs(pdf_name, FileFormat=17)
            word_doc.Close()
        except Exception as exc:
            QMessageBox.critical(self, 'Ошибка', f'Произошла ошибка!Пришлите следующий текст разработчику - \n{exc}')
            return

    def import_to_pdf(self):
        self.add_item_to_combobox()
        if not self.check_all_object():
            return None
        date, num, all_rolls, total_price_rolls, all_services, all_automatic, manager, total_price_services = self.get_all_info()
        total_price = round(total_price_rolls + self.ui.doubleSpinBox.value() + self.ui.doubleSpinBox_2.value(), 2)
        my_total_price = self.ui.doubleSpinBox_3.value() + total_price_services
        difference_price = round(total_price - my_total_price, 2)
        percent_total_discount = self.get_percent(total_price, difference_price)
        context = {
            'number': num,
            'date': date,
            'full_name': self.ui.lineEdit.text(),
            'address': self.ui.lineEdit_2.text(),
            'email': self.ui.lineEdit_3.text(),
            'tbl_contents': [rol for rol in all_rolls],
            'num_rolls': len(all_rolls),
            'total_price_rolls': total_price_rolls,
            'delivery_price': self.ui.doubleSpinBox.value(),
            'installation_price': self.ui.doubleSpinBox_2.value(),
            'discount': percent_total_discount,
            'discount_recalculation': difference_price,
            'total_price': my_total_price,
            'vat': round(my_total_price * 0.2, 2),
            'additional_services': [service for service in all_services],
            'all_automatic': [automatic for automatic in all_automatic],
            'expiration_date': self.ui.spinBox_2.value(),
            'full_name_manager': manager['full_name_manager'],
            'phone_manager': manager['phone_manager'],
            'email_manager': manager['email_manager']
        }
        return context

    def send_commercial_to_mail(self):
        self.add_item_to_combobox()
        context = self.import_to_pdf()
        with open('texts/texts_commercial.txt', 'r', encoding='utf-8') as file:
            text = ''
            texts = file.readlines()
            for line in texts:
                text += line
        if context is None:
            return
        file, check = self.create_pdf(context, 'samples/template_commercial.docx', 'Unload_commercial')
        if check:
            files = self.get_open_files_and_dirs('Загрузить фотографии!', f'{os.getcwd()}', 'Image Files (*.png *.jpeg)')
            files.insert(0, file)
            manager = self.get_manager()
            msg = MIMEMultipart()
            msg['From'] = f"{self.manager['send_from']['email']}"
            msg['To'] = f'{context["email"]}, {manager["email_manager"]}'
            msg['Subject'] = f'Коммерческое предложение по рольставням и рулонным воротам по заказу № ВАР ' \
                             f'{context["number"]}/{context["date"]}. ABC-SAFETY'
            for i, file in enumerate(files):
                try:
                    if os.path.isdir(file):
                        for file_from_path in os.listdir(file):
                            files.append(f'{file}/{file_from_path}')
                        files.pop(i)
                        continue
                    attachment = MIMEApplication(open(file, "rb").read())
                    new_file = os.path.basename(file)
                    attachment.add_header('Content-Disposition', 'attachment', filename=new_file)
                    msg.attach(attachment)
                except Exception as exc:
                    QMessageBox.critical(
                        self, 'Ошибка', f'Произошла ошибка!\nНевозможно прикрепить данный файл к письму!\n{exc}')
                    return
            body = text
            msg.attach(MIMEText(body, 'plain'))
            try:
                server = smtplib.SMTP(self.manager['send_from']['smtp_server'], self.manager['send_from']['smtp_port'])
                server.starttls()
                server.login(self.manager['send_from']['email'], self.manager['send_from']['password'])
            except Exception as exc:
                QMessageBox.critical(
                    self, 'Ошибка', f'Произошла ошибка!\nНевозможно зайти в данный аккаунт!\n{exc}')
                return
            try:
                server.send_message(msg)
                server.quit()
            except Exception as exc:
                QMessageBox.critical(self, 'Ошибка', f'Произошла ошибка!\nНевозможно отправить сообщение!\n{exc}')
                return
            QMessageBox.information(self, 'Успех!', f'Письмо успешно отправлено на почту {context["email"]}!')
            return
        else:
            return

    def get_open_files_and_dirs(self, caption='', directory='', filter='', initialFilter='', options=None):

        def update_text():
            selected = []
            for index in view.selectionModel().selectedRows():
                selected.append('"{}"'.format(index.data()))
            lineEdit.setText(' '.join(selected))

        dialog = QFileDialog(self, windowTitle=caption)
        dialog.setFileMode(dialog.ExistingFiles)
        if options:
            dialog.setOptions(options)
        dialog.setOption(dialog.DontUseNativeDialog, True)
        if directory:
            dialog.setDirectory(directory)
        if filter:
            dialog.setNameFilter(filter)
            if initialFilter:
                dialog.selectNameFilter(initialFilter)
        dialog.accept = lambda: QDialog.accept(dialog)
        stackedWidget = dialog.findChild(QStackedWidget)
        view = stackedWidget.findChild(QListView)
        view.selectionModel().selectionChanged.connect(update_text)
        lineEdit = dialog.findChild(QLineEdit)
        dialog.directoryEntered.connect(lambda: lineEdit.setText(''))
        dialog.exec_()
        return dialog.selectedFiles()

    def create_pdf(self, context, file_name, save_path='Unload'):
        try:
            doc = DocxTemplate(file_name)
        except Exception as exc:
            QMessageBox.critical(self, 'Ошибка', f'Не найдено файла {os.path.basename(file_name)}\n{exc}')
            return None, False
        person = self.get_person()
        if person:
            prefix = 'individual'
        else:
            prefix = 'entity'
        doc.render(context)
        doc.save('template-final.docx')
        path = str(Path(os.getcwd() + '\\template-final.docx'))
        if not os.path.isdir(str(Path(os.getcwd() + f'\\{save_path}'))):
            os.mkdir(os.getcwd() + f'\\{save_path}')
        if not os.path.isdir(str(Path(os.getcwd() + f'\\{save_path}\\{context["date"].replace(".", "_")}'))):
            os.mkdir(str(Path(os.getcwd() + f'\\{save_path}\\{context["date"].replace(".", "_")}')))
        if not os.path.isdir(str(Path(os.getcwd() + f'\\{save_path}\\{context["date"].replace(".", "_")}\\{prefix}'))):
            os.mkdir(str(Path(os.getcwd() + f'\\{save_path}\\{context["date"].replace(".", "_")}\\{prefix}')))
        new_path = str(
            Path(
                os.getcwd() + f'\\{save_path}\\{context["date"].replace(".", "_", 3)}\\{prefix}\\{context["number"]}.pdf'))
        self.print_to_pdf(path, new_path)
        os.remove(path)
        QMessageBox.about(self, 'Успех!', f'Ваш документ сформирован и находится по пути:\n{new_path}')
        self.new_doc(context["date"], context["number"])
        self.get_num_doc()
        os.startfile(new_path)
        test = QMessageBox.question(self, "Нравится?", "Вас устраивает такой вариант?",
                                    QMessageBox.Yes | QMessageBox.No)
        if QMessageBox.Yes == test:
            self.all_clear()
            return new_path, True
        elif QMessageBox.No == test:
            QMessageBox.about(self, 'Ну и как это понимать?', f'Ну и ладно(')
            return None, False
        else:
            return None, False

    def send_to_email(self):
        self.add_item_to_combobox()
        context = self.import_to_pdf()
        if context is None:
            return
        person = self.get_person()
        if person:
            sample_file = 'samples\\образец_физ_лица.pdf'
            with open('texts/texts_order_individual.txt', 'r', encoding='utf-8') as file:
                text = ''
                texts = file.readlines()
                for line in texts:
                    text += line
            file, check = self.create_pdf(context, 'samples/template_individual.docx')
        else:
            sample_file = 'samples\\образец_юр_лица.pdf'
            with open('texts/texts_order_entity.txt', 'r', encoding='utf-8') as file:
                text = ''
                texts = file.readlines()
                for line in texts:
                    text += line
            file, check = self.create_pdf(context, 'samples/template_entity.docx')
        if check:
            msg = MIMEMultipart()
            msg['From'] = f"{self.manager['send_from']['email']}"
            msg['To'] = f'{context["email"]}'
            msg['Subject'] = 'Бланк заказа на подпись!'
            files = [file, sample_file]
            try:
                for f in files:
                    attachment = MIMEApplication(open(f, "rb").read())
                    new_file = os.path.basename(f)
                    attachment.add_header('Content-Disposition', 'attachment', filename=new_file)
                    msg.attach(attachment)
            except Exception as exc:
                QMessageBox.critical(
                    self, 'Ошибка', f'Произошла ошибка!\nНевозможно прикрепить данный файл к письму!\n{exc}')
                return
            body = text
            msg.attach(MIMEText(body, 'plain'))
            try:
                server = smtplib.SMTP(self.manager['send_from']['smtp_server'], self.manager['send_from']['smtp_port'])
                server.starttls()
                server.login(self.manager['send_from']['email'], self.manager['send_from']['password'])
            except Exception as exc:
                QMessageBox.critical(
                    self, 'Ошибка', f'Произошла ошибка!\nНевозможно зайти в данный аккаунт!\n{exc}')
                return
            try:
                server.send_message(msg)
                server.quit()
            except Exception as exc:
                QMessageBox.critical(self, 'Ошибка', f'Произошла ошибка!\nНевозможно отправить сообщение!\n{exc}')
                return
            QMessageBox.information(self, 'Успех!', f'Письмо успешно отправлено на почту {context["email"]}!')
            return
        else:
            return


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
