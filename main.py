import configparser
import datetime
import os
import sys
import xml.etree.cElementTree as ET

import mammoth as mammoth
import pdfkit
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from docxtpl import DocxTemplate
from pony.orm import db_session

from design import Ui_MainWindow
from import_to_db import add_to_db
from models import DataNumDoc, AutomationData, Rollers


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.manager = configparser.ConfigParser()
        self.manager.read(f'{os.getcwd()}/config.cfg')
        self.xml_file = None
        self.init_ui()
        self.connect_ui()
        self.location_on_the_screen()

    def location_on_the_screen(self):
        screen = QGuiApplication.screenAt(QCursor().pos())
        fg = self.frameGeometry()
        fg.moveCenter(screen.geometry().center())
        self.move(fg.topLeft())

    def init_ui(self):
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
        self.ui.pushButton.clicked.connect(self.import_to_pdf)
        self.ui.pushButton_4.clicked.connect(self.add_automatic)

    @db_session
    def add_item_to_combobox(self):
        all_automatic = AutomationData.select().order_by(AutomationData.id)[:]
        for item in all_automatic:
            self.ui.comboBox.addItem(f'{item.name_automatic}/{item.price_automatic}')

    def add_to_db(self):
        try:
            file = QFileDialog.getOpenFileName(self, 'Open file', f'{os.getcwd()}', 'CSV Files (*.csv)')[0]
            if file == '':
                return QMessageBox.information(self, 'Неувязочка!', f'Вы не выбрали ни одного файла!')
            else:
                add_to_db(file)
                QMessageBox.about(self, "Отлично!", "Вы успешно загрузили данные о ценах!")
        except Exception as exc:
            QMessageBox.critical(self, 'Ошибка', f'Произошла ошибка!\n{exc}')
            return

    def load_xml(self):
        try:
            self.xml_file = QFileDialog.getOpenFileName(self, 'Open file', f'{os.getcwd()}', 'XML Files (*.xml)')[0]
            if self.xml_file == '':
                QMessageBox.information(self, 'Неувязочка!', f'Вы не выбрали ни одного файла!')
            else:
                QMessageBox.about(self, "Отлично!", "Вы успешно загрузили данные о роллетах!")
        except Exception as exc:
            QMessageBox.critical(self, 'Ошибка!', f'Произошла ошибка!\n{exc}')

    def add_automatic(self):
        if self.ui.spinBox_3.value() == 0:
            return QMessageBox.information(self, 'Неувязочка!', f'Зачем добавлять нулевое количество товара?)')
        item = self.ui.comboBox.currentText()
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

    def delete_item_2(self):
        item = self.ui.treeWidget_2.currentItem()
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

    @db_session
    def unload_all_rolls(self):
        all_rolls = []
        tree = ET.parse(self.xml_file)
        test = tree.findall('Orders/Orders_Rec/Rolletes/Rolletes_Rec')
        i = 0
        total_price = 0
        for item in test:
            i += 1
            width = float(item.attrib["WIDTH"].replace(',', '.'))
            height = float(item.attrib["HEIGHT"].replace(',', '.'))
            try:
                roll = Rollers.get(code=int(item.attrib["PROFILE_CODE_"]))
            except Exception as exc:
                QMessageBox.information(self, 'Неувязочка', f'Вы не загрузили цены!')
                return
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
                    roll.price
                ],
                'gear': item.attrib["GEAR_"]
            }
            all_rolls.append(data)
            total_price += roll.price
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

    def get_manager(self):
        if self.ui.radioButton.isChecked():
            manager = {
                'full_name_manager': self.manager['manager_1']['full_name_manager'],
                'email_manager': self.manager['manager_1']['email_manager'],
                'phone_manager': self.manager['manager_1']['phone_manager']
            }
        else:
            manager = {
                'full_name_manager': self.manager['manager_2']['full_name_manager'],
                'email_manager': self.manager['manager_2']['email_manager'],
                'phone_manager': self.manager['manager_2']['phone_manager']
            }
        return manager

    def get_all_info(self):
        date = self.ui.dateEdit.date().toString("dd.MM.yyyy")
        num = self.get_num_doc()
        if num < 10:
            num = f'0{num}'
        try:
            all_rolls, total_price_rolls = self.unload_all_rolls()
        except TypeError as exc:
            QMessageBox.information(self, 'Неувязочка!', f'Вы не загрузили XML файл!')
            return
        all_services = self.get_all_services()
        all_automatic, total_price_automatic = self.get_all_automatic()
        total_price_rolls += total_price_automatic
        manager = self.get_manager()
        return date, num, all_rolls, total_price_rolls, all_services, all_automatic, manager

    @db_session
    def new_doc(self, date, num):
        DataNumDoc(date=date, num_doc=num)

    @db_session
    def check_all_object(self):
        if len(Rollers.select().order_by(Rollers.id)[:]) == 0:
            QMessageBox.information(self, 'Неувязочка', f'Вы не загрузили информацию о ценах!\n'
                                                        f'Нажмите меню "действия" и выберите пункт "Загрузить цены"!')
            return False
        if self.ui.lineEdit.text() == '':
            QMessageBox.information(self, 'Неувязочка', f'Вы не ввели ФИО заказчика!')
            return False
        if self.ui.lineEdit_2.text() == '':
            QMessageBox.information(self, 'Неувязочка', f'Вы не ввели адрес заказчика!')
            return False
        if self.ui.lineEdit_3.text() == '':
            QMessageBox.information(self, 'Неувязочка', f'Вы не ввели email заказчика!')
            return False
        # if self.ui.doubleSpinBox.value() == 0.00:
        #     QMessageBox.information(self, 'Неувязочка', f'Вы не ввели стоимость доставки!')
        #     return False
        # if self.ui.doubleSpinBox_2.value() == 0.00:
        #     QMessageBox.information(self, 'Неувязочка', f'Вы не ввели стоимость установки!')
        #     return False
        if self.ui.spinBox_2.value() == 0:
            QMessageBox.information(self, 'Неувязочка', f'Вы не выставили срок производства!')
            return False
        self.load_xml()
        if self.xml_file == '':
            return False
        return True

    def import_to_pdf(self):
        if not self.check_all_object():
            return
        try:
            doc = DocxTemplate("template.docx")
        except Exception as exc:
            QMessageBox.critical(self, 'Ошибка', f'Не найдено файла template.docx\n{exc}')
            return
        date, num, all_rolls, total_price_rolls, all_services, all_automatic, manager = self.get_all_info()
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
            'discount': self.ui.spinBox.value(),
            'discount_recalculation': round(total_price_rolls * self.ui.spinBox.value() / 100, 2),
            'additional_discount': self.ui.doubleSpinBox_3.value(),
            'total_price': round
            (total_price_rolls + self.ui.doubleSpinBox.value() + self.ui.doubleSpinBox_2.value() - round(
                total_price_rolls * self.ui.spinBox.value() / 100, 2) - self.ui.doubleSpinBox_3.value(), 2),
            'vat': round
            (round(total_price_rolls + self.ui.doubleSpinBox.value() + self.ui.doubleSpinBox_2.value() - round(
                total_price_rolls * self.ui.spinBox.value() / 100, 2) - self.ui.doubleSpinBox_3.value(), 2) * 0.2, 2),
            'additional_services': [service for service in all_services],
            'all_automatic': [automatic for automatic in all_automatic],
            'expiration_date': self.ui.spinBox_2.value(),
            'full_name_manager': manager['full_name_manager'],
            'phone_manager': manager['phone_manager'],
            'email_manager': manager['email_manager'],
        }
        doc.render(context)
        path = os.getcwd() + '/unload/template-final.docx'
        doc.save(f"template-final.docx")
        self.new_doc(date, num)

        try:
            config = pdfkit.configuration(wkhtmltopdf=f'{os.getcwd()}/wkhtmltopdf/bin/wkhtmltopdf.exe')
        except Exception:
            QMessageBox.critical(self, 'Ошибка', f'Не найден файл\n{os.getcwd()}/\nwkhtmltopdf/bin/wkhtmltopdf.exe')
            return
        with open(f"template-final.docx", 'rb') as docx:
            document = mammoth.convert_to_html(docx)
            html_doc = document.value
        with open('test.html', "w", encoding='UTF8') as html:
            html.writelines(html_doc)
        if not os.path.isdir(f'{os.getcwd()}/Technical_Tasks'):
            os.mkdir(f'{os.getcwd()}/Technical_Tasks')
        pdfkit.from_file(f'test.html',
                         f'{os.getcwd()}/Technical_Tasks/test.pdf',
                         configuration=config, options={'encoding': "UTF-8"})
        os.remove('test.html')
        os.remove('template-final.docx')
        QMessageBox.about(self, 'Успех!', f'Ваш документ сформирован и находится по пути: {os.path.abspath(path)}')
        os.startfile(f'{os.getcwd()}/Technical_Tasks/test.pdf')
        self.get_num_doc()
        # return True, f'{os.getcwd()}\\Technical_Tasks\\{context["number"].replace("/", "_")}.pdf'


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())
