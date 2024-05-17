import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, QComboBox, QLabel, QTextEdit, QHeaderView, QScrollArea, QStyledItemDelegate, QLineEdit, QMessageBox, QHBoxLayout, QInputDialog
from PyQt5.QtGui import QFont, QPixmap, QRegExpValidator, QColor, QIcon
from PyQt5.QtCore import Qt, QRegExp, QDate, QDateTime
import win32com.client as win32
import os
import re
import random
import json
filesave = os.path.join(os.getenv('LOCALAPPDATA'), 'EmailHelper', 'EmailHelper_data.pkl')

class DataEntryWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle('Створення щотижневого листа для магазинів та кураторів')
        self.setMinimumSize(1300, 900)  

        self.filials_data = {
            "Центр-1": {
                "адреса": ["вул. Будівельна 19","Квартал 103, виділ 9, 1","Гната Юри вул, 20","Родини Бунґе вул, 8","Руденка Миколи  бул, 14 М","Богуна Івана вул, 2а","Київська вул, 1/102","Київська вул, 36","Неб.Сотні пр-т,24/83","Соборна вул, 140А"],
                "email": [""],
                "emailсс": ["copy1@example.com", "copy2@example.com"]  # Добавлено значение "emailсс"
            },
            "Центр-2": {
                "адреса": ["Архипенко Олександра вул, 6","Бальзака вул, 91/29а","Володимира Івасюка пр-т, 12П","Володимира Івасюка пр-т, 8А","Закревського вул, 61/2","Лаврухіна вул, 4","Милославська вул, 10а","Рональда Рейгана вул, 8","Степана Бандери пр-т, 36","Червоної Калини пр-т, 43/2","Червоної Калини пр-т, 75/2","Погребський Шлях, 19"],
                "email": [""],
                "emailсс": [""]  # Добавлено значение "emailсс"
            },
            "Центр-3": {
                "адреса": ["Берестейський пр, 87","Берестейський пр-т, 47","Берестейський пр-т, 94/1","Берковецька вул, 6","Борщагівська вул, 154а","Підлісна вул, 1","Чоколівський б-р, 6","Чоколівський бул, 28/1","Чорнобильська вул, 3","Київська вул, 10"],
                "email": [""],
                "emailсс": [""]  # Добавлено значение "emailсс"
            },
            "Центр-4": {
                "адреса": ["Братиславська вул, 14Б","Вершигори вул, 1","Воскресенський пр-т, 36","Дарницький бул, 8а","Кибальчича вул, 11А","Лісовий пр, 39","Райдужна вул, 15","Стальського вул, 22/10","Фінський пров, 3","Якова Гніздовського вул. 1А"],
                "email": [""],
                "emailсс": [""]  # Добавлено значение "emailсс"
            },
            "Центр-5": {
                "адреса": ["Арх-ра Вербицького вул, 1","Вербицького вул, 30","Григоренка вул, 23","Мішуги вул, 4","Ревуцького вул, 12/1А","Срібнокільська вул, 3-Г","Харківське ш, 144","Харківське ш, 168","Харківське шосе, 1В"],
                "email": [""],
                "emailсс": [""]  # Добавлено значение "emailсс"
            },
            "Центр-6":{
                "адреса": ["Березнева вул, 12А","Дніпровська наб. вул, 12","Дніпровська набережна,33а","Драгоманова вул, 10","Здолбунівська вул, 4","Інженерна вул, 1","Павла Тичини вул., 1в","Русанівська наб, 10","Шептицького вул, 22","Шептицького вул, 4 А"],
                "email": [""],
                "emailсс": [""]  # Добавлено значение "emailсс"
            },
            "Центр-7": {
                "адреса": ["Бережанська вул, 22","Вишгородська вул, 21","Дорогожицька вул, 2","Западинська вул, 15 А","Івашкевича вул, 6-8 А","Литовський пр-т, 4а","Порика вул, 5а","Правди просп, 66","Правди пр-т, 58","Скляренка вул, 17","Щербаківського вул, 56/7"],
                "email": [""],
                "emailсс": [""]  # Добавлено значение "emailсс"
            },
            "Центр-8": {
                "адреса": ["Андрія Верхогляда вул, 16","Антоновича вул, 165","Кільцева дорога вул, 1","Коновальця вул, 26 А","Липківського вул. 1А","Малевича вул, 107","Самійла Кішки вул, 7","Столичне шс, 103","Філатова вул, 7","Ватутіна вул, 170"],
                "email": [""],
                "emailсс": [""]  # Добавлено значение "emailсс"
            },
            "Центр-9": {
                "адреса": ["Володимира Івасюка пр-т, 46","Володимира Івасюка пр-т,27Б","Героїв полку «Азов» вул, 5","Героїв полку «Азов» вул,34","Оболонський пр-т, 19","Оболонський пр-т. 1Б","Оболонський пр-т. 21Б","Петра Калнишевського вул, 2","Північна вул, 6"],
                "email": [""],
                "emailсс": [""]  # Добавлено значение "emailсс"
            },
            "Центр-10": {
                "адреса": ["Басейна вул, 6","Берестейський пр-т, 24, 26","Білоруська вул, 2","Глибочицька вул, 32Б","Гончара вул, 96","Загорівська вул. 17-20","Оленівська вул, 3","Сагайдачного вул, 41","Січових Стрільців вул, 37/41","Спортивна пл, 1 А","Ярославська вул, 56а"],
                "email": [""],
                "emailсс": [""]  # Добавлено значение "emailсс"
            },
            "Північ-1": {
                "адреса": ["Катеринославський бул, 1","Кондратюка вул, 4","Слобожанський пр,76/78","Слобожанський пр-т, 31Д","Крушельницької пров, 6А","Європейська вул, 18А","Фабра Андрія вул, 7","Гетьманська 47А","Горького вул, 166","Шевченка пр-т, 9"],
                "email": ["fu012-0026-head@fozzy.ua","fu012-0232-head@fozzy.ua","dp-gor166-head@fozzy.ua","DP-EKATE1-head@fozzy.ua","dp-fabra7-head@fozzy.ua","dp-miro20-head@fozzy.ua","dp-gaze76-head@fozzy.ua","dp-komun4-head@fozzy.ua","dp-lmaki6-head@fozzy.ua","dp-gaze31-head@fozzy.ua"],
                "emailсс": ["li.demianchuk@fozzy.ua","i.rashkov@fozzy.ua","a.fedorchenko@fozzy.ua","s.chegryn@fozzy.ua","y.tiutiunnik@fozzy.ua","i.omelchenko@fozzy.ua","a.podenko@fozzy.ua","g.krupnyk@fozzy.ua","y.makarova@fozzy.ua","a.korotkyi@fozzy.ua","t.grynkova@fozzy.ua","t.shystia@fozzy.ua","t.grynkova@fozzy.ua","t.shystia@fozzy.ua","zh.los@fozzy.ua","v.soian@fozzy.ua","n.pokotys@fozzy.ua","o.rimanova@fozzy.ua","i.denysova@fozzy.ua","s.seima@fozzy.ua","ir.pryshchepa@fozzy.ua"]  # Добавлено значение "emailсс"
            },
            "Північ-2": {
                "адреса": ["Вокзальна пл, 13", "Гагаріна пр, 3", "Незалежності  вул, 36", "Новокримська вул, 3а", "Пастера вул, 6А", "Слави бул, 5", "Тополина вул, 1", "Пет. Сагайдачного, 20Б", "Чарівна вул, 74"],
                "email": ["dp-topol1-head@fozzy.ua","dp-petr13-head@fozzy.ua","dp-gagar3-head@fozzy.ua","dp-novok3-head@fozzy.ua","dp-paste6-head@fozzy.ua","dp-slavy5-head@fozzy.ua","zp-boro20-head@fozzy.ua","zp-char74-head@fozzy.ua","dp-tito36-head@fozzy.ua"],
                "emailсс": ["o.shaposhnyk@fozzy.ua","g.yakymenko@fozzy.ua","s.zemziulin@fozzy.ua","m.zalizniak@fozzy.ua","a.roldugin@fozzy.ua","m.soian@fozzy.ua","r.saienko@fozzy.ua","y.omelianenko@fozzy.ua","yu.avdieiev@fozzy.ua","zh.los@fozzy.ua","v.kykot@fozzy.ua","t.shystia@fozzy.ua","v.soian@fozzy.ua","t.grynkova@fozzy.ua","o.rimanova@fozzy.ua","g.komarenko@fozzy.ua","t.chernomorets@fozzy.ua","y.kuzyk@fozzy.ua","ol.maksymenko@fozzy.ua","te.golovko@fozzy.ua","g.nikogosian@fozzy.ua","i.kozyriev@fozzy.ua","d.pysmennyi@fozzy.ua","n.pokotys@fozzy.ua","i.denysova@fozzy.ua","s.seima@fozzy.ua","y.povazhuk@fozzy.ua","ir.pryshchepa@fozzy.ua"]  # Добавлено значение "emailсс"
            },
            "Захід-1": {
                "адреса": ["Б.Хмельницького вул, 214","Гетьмана І. Мазепи вул, 1Б","Мазепи вул, 11","Під Дубом вул, 7 Б","Соборна пл, 14-15","Чорновола пр-т, 77","Шевченка вул, 358а","Шевченка вул, 60","Широка вул, 87"],
                "email": [""],
                "emailсс": []  # Добавлено значение "emailсс"
            },
            "Захід-2": {
                "адреса": ["Академіка Сахарова вул, 45","Виговського вул, 100","Володимира Великого вул,26А","Городоцька вул, 179","Кульпарківська вул, 226 А","Кульпарківська вул, 93А","Наукова вул, 35 А","Пасічна вул, 164","Садова вул, 2А","Стрийська вул, 45","Чер.Калини пр, 62"],
                "email": [""],
                "emailсс": []  # Добавлено значение "emailсс"
            },
            "Південь-1": {
                "адреса": ["Генуезька вул, 24 Б","Генуезька вул, 5","Довженка вул, 4","Люстдорфська дорога вул, 54","Новощіпний ряд, 2","Семафорний пров, 4","Філатова вул, 1","Фонтанська дор, 39","Французький бул, 16","Черняховського вул, 1"],
                "email": [""],
                "emailсс": []  # Добавлено значение "emailсс"
            },
            "Південь-2": {
                "адреса": ["Жемчужна вул. 5","Шевченка вул, 228","Вільямса вул, 75","Героїв Крут вул, 17/1","Корольова вул, 44","Небесної Сотні пр-т, 14","Небесної сотні пр-т, 2","Небесної Сотні пр-т, 5А","Петрова вул, 51","Фонтанська дор, 58/1","Піонерна вул, 1"],
                "email": [""],
                "emailсс": []  # Добавлено значение "emailсс"
            },
            "Південь-3": {
                "адреса": ["Бочарова вул, 13 А","Бочарова вул, 44","Г-в оборони Одеси вул,98Б","Єврейська вул, 50/1","Катерининська вул, 27/1","Кримська вул, 71","Семена Палія вул, 125-Б","Семена Палія вул, 72","Семена Палія вул, 93А","Григорівсь-го Десанту пр-т,3"],
                "email": [""],
                "emailсс": []  # Добавлено значение "emailсс"
            },
        }
        self.layout = QVBoxLayout()

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        widget_inside_scroll = QWidget()
        scroll_area.setWidget(widget_inside_scroll)

        vertical_layout = QVBoxLayout(widget_inside_scroll)
        
        horizontal_layout = QHBoxLayout()

        # Получите путь к изображению
        relative_path = "bushes.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)

        # Создайте QLabel для изображения
        image_label = QLabel()
        pixmap = QPixmap(image_path) 
        new_width = 30  # Новая ширина
        new_height = 30  # Новая высота
        scaled_pixmap = pixmap.scaled(new_width, new_height)
        image_label.setPixmap(scaled_pixmap)

        label = QLabel("Обери свій кущ:")
        label.setFont(QFont("Calibri", 16))

        self.kust_combobox = QComboBox()
        self.kust_combobox.setFont(QFont("Calibri", 14))
        self.kust_combobox.addItems(self.filials_data.keys())
        self.kust_combobox.currentIndexChanged.connect(self.on_combobox_changed)
        self.kust_combobox.setFixedWidth(200)  # Устанавливаем фиксированную ширину
        self.previous_index = self.kust_combobox.currentIndex()

        horizontal_layout.addWidget(image_label)
        horizontal_layout.addWidget(label)
        horizontal_layout.addWidget(self.kust_combobox)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        #################################################################

        # Получите путь к изображению
        relative_path = "calendar.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)

        # Создайте QLabel для изображения
        image_label = QLabel()
        pixmap = QPixmap(image_path) 
        new_width = 30  # Новая ширина
        new_height = 30  # Новая высота
        scaled_pixmap = pixmap.scaled(new_width, new_height)
        image_label.setPixmap(scaled_pixmap)

        horizontal_layout = QHBoxLayout()

        label = QLabel("Обери тиждень за який пишеш звіт...")
        label.setFont(QFont("Calibri", 16))
        label.setAlignment(Qt.AlignLeft)

        # Визначаємо поточний номер тижня
        current_week = QDate.currentDate().weekNumber()

        # Формуємо список номерів тижнів для випадаючого списку
        weeks = [str(current_week[0] - i) for i in range(3, 0, -1)]

        # Створення випадаючого списку з номерами тижнів
        self.week_combobox = QComboBox()
        self.week_combobox.setFont(QFont("Calibri", 12))
        self.week_combobox.setFixedWidth(50)  # Задаємо фіксовану ширину для поля вводу
        self.week_combobox.addItems(weeks)  # Додаємо номери тижнів у випадаючий список
        self.week_combobox.setCurrentIndex(2)  # Задаємо початкове значення
        self.week_combobox.currentIndexChanged.connect(self.on_combobox_week_changed)

        horizontal_layout.addWidget(image_label)
        horizontal_layout.addWidget(label)
        horizontal_layout.addWidget(self.week_combobox)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)
        
        horizontal_layout = QHBoxLayout()

        # Создание кнопки
        self.save_button = QPushButton("Зберегти")
        self.save_button.setFont(QFont("Calibri", 14))
        self.save_button.setFixedWidth(297)                                 

        # Установка иконки

        relative_path = "save.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)

        icon = QIcon(image_path)  # Укажите путь к вашему файлу изображения
        self.save_button.setIcon(icon)
        self.save_button.clicked.connect(lambda: self.save_data(filesave))
        horizontal_layout.addWidget(self.save_button)

        self.import_button = QPushButton("Відкрити збережені дані")
        self.import_button.setFont(QFont("Calibri", 14))
        self.import_button.setFixedWidth(297)
        self.import_button.clicked.connect(lambda: self.load_data(filesave))
        # Установка иконки
        relative_path = "import.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)

        icon = QIcon(image_path)  # Укажите путь к вашему файлу изображения
        self.import_button.setIcon(icon)
        horizontal_layout.addWidget(self.import_button)
        horizontal_layout.setAlignment(Qt.AlignLeft)
        vertical_layout.addLayout(horizontal_layout)

        self.send_email_button = QPushButton("Переглянути лист")
        self.send_email_button.setFont(QFont("Calibri", 14))
        self.send_email_button.setFixedWidth(600)
        self.send_email_button.clicked.connect(self.send_email)
        relative_path = "mail.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)
        icon = QIcon(image_path)  # Укажите путь к вашему файлу изображения
        self.send_email_button.setIcon(icon)
        vertical_layout.addWidget(self.send_email_button)

        self.random_button = QPushButton("TEST! Заповнити дані")
        self.random_button.setFont(QFont("Calibri", 14))
        self.random_button.setFixedWidth(600)
        self.random_button.clicked.connect(self.randomize_empty_cells)
        vertical_layout.addWidget(self.random_button)

        self.remove_button = QPushButton("TEST! Очистити всі дані")
        self.remove_button.setFont(QFont("Calibri", 14))
        self.remove_button.setFixedWidth(600)
        self.remove_button.clicked.connect(self.remove_data)
        vertical_layout.addWidget(self.remove_button)

        ################################################################
        vertical_layout.addWidget(QLabel("-----------------------------------------------------"))

        horizontal_layout = QHBoxLayout()

        # Получите путь к изображению
        relative_path = "hello.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)

        # Создайте QLabel для изображения
        image_label = QLabel()
        pixmap = QPixmap(image_path) 
        new_width = 30  # Новая ширина
        new_height = 30  # Новая высота
        scaled_pixmap = pixmap.scaled(new_width, new_height)
        image_label.setPixmap(scaled_pixmap)

        # Создайте QLabel для текста
        text_label = QLabel("<b>Привітання і рев'ю</b>")
        text_label.setFont(QFont("Calibri", 16))
        text_label.setAlignment(Qt.AlignLeft)

        # Добавьте изображение и текст в горизонтальный макет
        horizontal_layout.addWidget(image_label)
        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        ###############################################
        
        horizontal_layout = QHBoxLayout()

        # Поле ввода с предустановленным текстом
        self.greeting_review_textedit = QTextEdit()
        self.greeting_review_textedit.setFont(QFont("Calibri", 12))
        self.greeting_review_textedit.setPlaceholderText("Введи привітання та короткий опис тижня...")
        self.greeting_review_textedit.setFixedWidth(500)
        horizontal_layout.addWidget(self.greeting_review_textedit)

        text_label = QLabel("Привітайся, опиши коротко загальну ситуацію на минулому тижні,\nпохвали магазини за особливі успіхи в роботі (якщо є приклади)")
        text_label.setFont(QFont("Calibri", 12))
        text_label.setAlignment(Qt.AlignLeft)
        text_label.setFixedWidth(500)
        text_label.setStyleSheet("color: #666666;")

        # Добавьте изображение и текст в горизонтальный макет
        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        ###############################################


        # vertical_layout.addWidget(QLabel("-----------------------------------------------------"))

        horizontal_layout = QHBoxLayout()

        # Получите путь к изображению
        relative_path = "image002.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)

        # Создайте QLabel для изображения
        image_label = QLabel()
        pixmap = QPixmap(image_path) 
        new_width = 30  # Новая ширина
        new_height = 30  # Новая высота
        scaled_pixmap = pixmap.scaled(new_width, new_height)
        image_label.setPixmap(scaled_pixmap)

        # Создайте QLabel для текста
        text_label = QLabel("<b>Загальні показники:</b>")
        text_label.setFont(QFont("Calibri", 16))
        text_label.setAlignment(Qt.AlignLeft)

        # Добавьте изображение и текст в горизонтальный макет
        horizontal_layout.addWidget(image_label)
        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        self.general_table = QTableWidget()
        self.general_table.setColumnCount(9)
        self.general_table.setHorizontalHeaderLabels([
            'Філіал', 'Пенетрація план', 'Пенетрація факт', '% виконання плану',
            'План виторгу', 'Факт виторгу', '% виконання плану','ДП план', 'ДП факт'
        ])
        self.general_table.verticalHeader().setVisible(False)
        self.general_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)


        vertical_layout.addWidget(self.general_table)
        self.importButton_general = QPushButton('Імпорт з буферу')
        self.importButton_general.clicked.connect(self.import_to_general_table)
        vertical_layout.addWidget(self.importButton_general)
        ###############################################
        
        horizontal_layout = QHBoxLayout()

        # Поле ввода с предустановленным текстом
        self.general_comments_textedit = QTextEdit()
        self.general_comments_textedit.setFont(QFont("Calibri", 12))
        self.general_comments_textedit.setPlaceholderText("Введи коментарі про загальні показники...")
        self.general_comments_textedit.setMaximumWidth(600)
        self.general_comments_textedit.setFixedHeight(150)
        horizontal_layout.addWidget(self.general_comments_textedit)

        text_label = QLabel("Опиши результати по пенетрації, виторгу та доступності.\nЯкщо знаєш чому магазин не виконав або перевиконав план - вкажи це.")
        text_label.setFont(QFont("Calibri", 12))
        text_label.setAlignment(Qt.AlignLeft)
        text_label.setFixedWidth(600)
        text_label.setStyleSheet("color: #666666;")

        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        ###############################################

        # vertical_layout.addWidget(QLabel("-----------------------------------------------------"))

        horizontal_layout = QHBoxLayout()

        # Получите путь к изображению
        relative_path = "image003.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)

        # Создайте QLabel для изображения
        image_label = QLabel()
        pixmap = QPixmap(image_path) 
        new_width = 30  # Новая ширина
        new_height = 30  # Новая высота
        scaled_pixmap = pixmap.scaled(new_width, new_height)
        image_label.setPixmap(scaled_pixmap)

        # Создайте QLabel для текста
        text_label = QLabel("<b>Списання:</b>")
        text_label.setFont(QFont("Calibri", 16))
        text_label.setAlignment(Qt.AlignLeft)

        # Добавьте изображение и текст в горизонтальный макет
        horizontal_layout.addWidget(image_label)
        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        self.write_off_table = QTableWidget()
        self.write_off_table.setColumnCount(4)
        self.write_off_table.setHorizontalHeaderLabels([
            'Філіал', 'План', 'Факт', '% виконання'
        ])
        self.write_off_table.verticalHeader().setVisible(False)
        self.write_off_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        vertical_layout.addWidget(self.write_off_table)

        self.importButton_write_off = QPushButton('Імпорт з буферу')
        self.importButton_write_off.clicked.connect(self.import_to_write_off_table)
        vertical_layout.addWidget(self.importButton_write_off)

        ##############################################
        
        horizontal_layout = QHBoxLayout()

        self.write_off_comments_textedit = QTextEdit()
        self.write_off_comments_textedit.setFont(QFont("Calibri", 12))
        self.write_off_comments_textedit.setPlaceholderText("Введи коментарі про списання...")
        self.write_off_comments_textedit.setFixedHeight(150)
        self.write_off_comments_textedit.setMaximumWidth(600)
        horizontal_layout.addWidget(self.write_off_comments_textedit)

        text_label = QLabel("Опиши ситуацію зі списанням. Підкажи магазинам на що треба звернути увагу на\nпоточному тижні.")
        text_label.setFont(QFont("Calibri", 12))
        text_label.setAlignment(Qt.AlignLeft)
        text_label.setFixedWidth(600)
        text_label.setStyleSheet("color: #666666;")

        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        ###############################################

        # vertical_layout.addWidget(QLabel("-----------------------------------------------------"))

        horizontal_layout = QHBoxLayout()

        # Получите путь к изображению
        relative_path = "image005.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)

        # Создайте QLabel для изображения
        image_label = QLabel()
        pixmap = QPixmap(image_path) 
        new_width = 30  # Новая ширина
        new_height = 30  # Новая высота
        scaled_pixmap = pixmap.scaled(new_width, new_height)
        image_label.setPixmap(scaled_pixmap)

        # Создайте QLabel для текста
        text_label = QLabel("<b>Списання банану:</b>")
        text_label.setFont(QFont("Calibri", 16))
        text_label.setAlignment(Qt.AlignLeft)

        # Добавьте изображение и текст в горизонтальный макет
        horizontal_layout.addWidget(image_label)
        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        self.write_off_banana_table = QTableWidget()
        self.write_off_banana_table.setColumnCount(5)
        self.write_off_banana_table.setHorizontalHeaderLabels([
            'Філіал', 'Списання план, %', 'Списання факт, %', 'Списання, кг', 'Продажі, кг'
        ])
        self.write_off_banana_table.verticalHeader().setVisible(False)
        self.write_off_banana_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        vertical_layout.addWidget(self.write_off_banana_table)

        self.importButton_write_off_banana = QPushButton('Імпорт з буферу')
        self.importButton_write_off_banana.clicked.connect(self.import_to_write_off_banana_table)
        vertical_layout.addWidget(self.importButton_write_off_banana)

        ##############################################
        
        horizontal_layout = QHBoxLayout()

        self.write_off_banana_textedit = QTextEdit()
        self.write_off_banana_textedit.setFont(QFont("Calibri", 12))
        self.write_off_banana_textedit.setPlaceholderText("Введи коментарі про списання...")
        self.write_off_banana_textedit.setFixedHeight(150)
        self.write_off_banana_textedit.setMaximumWidth(600)
        horizontal_layout.addWidget(self.write_off_banana_textedit)

        text_label = QLabel("Опиши ситуацію з бананом.")
        text_label.setFont(QFont("Calibri", 12))
        text_label.setAlignment(Qt.AlignLeft)
        text_label.setFixedWidth(600)
        text_label.setStyleSheet("color: #666666;")

        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        ###############################################

        horizontal_layout = QHBoxLayout()

        # Получите путь к изображению
        relative_path = "image007.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)

        # Создайте QLabel для изображения
        image_label = QLabel()
        pixmap = QPixmap(image_path) 
        new_width = 30  # Новая ширина
        new_height = 30  # Новая высота
        scaled_pixmap = pixmap.scaled(new_width, new_height)
        image_label.setPixmap(scaled_pixmap)

        # Создайте QLabel для текста
        text_label = QLabel("<b>Зелений базар:</b>")
        text_label.setFont(QFont("Calibri", 16))
        text_label.setAlignment(Qt.AlignLeft)

        # Добавьте изображение и текст в горизонтальный макет
        horizontal_layout.addWidget(image_label)
        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        selected_week = int(self.week_combobox.currentText()) - 1
        self.heavy_greens_table = QTableWidget()
        self.heavy_greens_table.setColumnCount(5)
        self.heavy_greens_table.setHorizontalHeaderLabels([
            'Філіал', f'Продажі {self.week_combobox.currentText()} тиждень, кг',
            f'Продажі {selected_week} тиждень, кг', '% приросту', 'Списання, кг'
        ])
        self.heavy_greens_table.verticalHeader().setVisible(False)
        self.heavy_greens_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)


        vertical_layout.addWidget(self.heavy_greens_table)

        self.importButton_heavy_greens = QPushButton('Імпорт з буферу')
        self.importButton_heavy_greens.clicked.connect(self.import_to_heavy_greens_table)
        vertical_layout.addWidget(self.importButton_heavy_greens)

        ##############################################
        
        horizontal_layout = QHBoxLayout()

        self.heavy_greens_textedit = QTextEdit()
        self.heavy_greens_textedit.setFont(QFont("Calibri", 12))
        self.heavy_greens_textedit.setPlaceholderText("Введи коментарі про зелений базар...")
        self.heavy_greens_textedit.setFixedHeight(150)
        self.heavy_greens_textedit.setMaximumWidth(600)
        horizontal_layout.addWidget(self.heavy_greens_textedit)

        text_label = QLabel("Опиши продажі та списання по зеленому базару.")
        text_label.setFont(QFont("Calibri", 12))
        text_label.setAlignment(Qt.AlignLeft)
        text_label.setFixedWidth(600)
        text_label.setStyleSheet("color: #666666;")

        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        ###############################################
        horizontal_layout = QHBoxLayout()

        # Получите путь к изображению
        relative_path = "image009.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)

        # Создайте QLabel для изображения
        image_label = QLabel()
        pixmap = QPixmap(image_path) 
        new_width = 30  # Новая ширина
        new_height = 30  # Новая высота
        scaled_pixmap = pixmap.scaled(new_width, new_height)
        image_label.setPixmap(scaled_pixmap)

        # Создайте QLabel для текста
        text_label = QLabel("<b>Планові показники:</b>")
        text_label.setFont(QFont("Calibri", 16))
        text_label.setAlignment(Qt.AlignLeft)

        # Добавьте изображение и текст в горизонтальный макет
        horizontal_layout.addWidget(image_label)
        horizontal_layout.addWidget(text_label)
        horizontal_layout.setAlignment(Qt.AlignLeft)

        vertical_layout.addLayout(horizontal_layout)

        self.plan_table = QTableWidget()
        self.plan_table.setColumnCount(6)
        self.plan_table.setHorizontalHeaderLabels([
            'Філіал', 'Пенетрація, %', 'Виторг, грн', 'Середньоденний виторг, грн', 'Списання, %', 'Доступність продажів, %'
        ])
        self.plan_table.verticalHeader().setVisible(False)
        self.plan_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        vertical_layout.addWidget(self.plan_table)

        self.importButton_plan = QPushButton('Імпорт з буферу')
        self.importButton_plan.clicked.connect(self.import_to_plan_table)
        vertical_layout.addWidget(self.importButton_plan)

        # vertical_layout.addWidget(QLabel("-----------------------------------------------------"))

        self.send_email_button = QPushButton("Переглянути лист")
        self.send_email_button.setFont(QFont("Calibri", 14))
        self.send_email_button.setFixedHeight(70)
        self.send_email_button.clicked.connect(self.send_email)
        relative_path = "mail.png"
        current_dir = os.path.dirname(os.path.realpath(__file__))
        image_path = os.path.join(current_dir, relative_path)
        icon = QIcon(image_path)  # Укажите путь к вашему файлу изображения
        self.send_email_button.setIcon(icon)
        vertical_layout.addWidget(self.send_email_button)

        self.layout.addWidget(scroll_area)

        self.setLayout(self.layout)

        self.update_tables(0) 

        self.general_table.itemChanged.connect(self.handle_item_changed)
        self.general_table.itemChanged.connect(self.recalculate_general_data)
        self.write_off_table.itemChanged.connect(self.recalculate_write_off_data)
        self.write_off_banana_table.itemChanged.connect(self.recalculate_write_off_banana_data)
        self.heavy_greens_table.itemChanged.connect(self.recalculate_heavy_greens_data)
        self.plan_table.itemChanged.connect(self.recalculate_plan_data)

        self.general_table.setItemDelegate(NumericDelegate())
        self.write_off_table.setItemDelegate(NumericDelegate())
        self.write_off_banana_table.setItemDelegate(NumericDelegate())
        self.heavy_greens_table.setItemDelegate(NumericDelegate())
        self.plan_table.setItemDelegate(NumericDelegate())

        for row in range(self.general_table.rowCount()):
            item = QTableWidgetItem("93%")
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            if self.general_table.item(row, 7) is None or self.general_table.item(row, 7).text() == "":
                self.general_table.setItem(row, 7, item)
            else:
                # Если ячейка не пуста, ничего не меняем
                pass

        for row in range(self.plan_table.rowCount()):
            item = QTableWidgetItem("93%")
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            if self.plan_table.item(row, 5) is None or self.plan_table.item(row, 5).text() == "":
                self.plan_table.setItem(row, 5, item)
            else:
                # Если ячейка не пуста, ничего не меняем
                pass

    def save_data(self, filename):
        data_to_save = {}
        # Сохранение данных из основной таблицы
        table_data = []
        for row in range(self.general_table.rowCount()):
            row_data = []
            for column in range(self.general_table.columnCount()):
                item = self.general_table.item(row, column)
                if column in [3, 6, 7]:  # Если текущий столбец в списке столбцов для замены на None
                    row_data.append(None)
                else:
                    if item is not None:
                        text = item.text().replace('%', '')  # Удаление символов "%"
                        row_data.append(text)
                    else:
                        row_data.append(None)
            table_data.append(row_data)

        # Сохранение данных из дополнительных полей ввода
        additional_data = {
            'greeting_review': self.greeting_review_textedit.toPlainText(),
            'general_comments': self.general_comments_textedit.toPlainText(),
            'write_off_comments': self.write_off_comments_textedit.toPlainText(),
            'write_off_banana': self.write_off_banana_textedit.toPlainText(),
            'heavy_greens': self.heavy_greens_textedit.toPlainText()
        }

        # Сохранение данных из таблицы self.write_off_table
        write_off_table_data = []
        directory = os.path.dirname(filename)
        if not os.path.exists(directory):
            os.makedirs(directory)
        for row in range(self.write_off_table.rowCount()):
            row_data = []
            for column in range(self.write_off_table.columnCount()):
                item = self.write_off_table.item(row, column)
                if column == 3:  # Если текущий столбец - индекс 3, заменяем на None
                    row_data.append(None)
                else:
                    if item is not None:
                        row_data.append(item.text())
                    else:
                        row_data.append(None)
            write_off_table_data.append(row_data)

        write_off_banana_data = []
        for row in range(self.write_off_banana_table.rowCount()):
            row_data = []
            for column in range(self.write_off_banana_table.columnCount()):
                item = self.write_off_banana_table.item(row, column)
                if item is not None:
                    row_data.append(item.text())
                else:
                    row_data.append(None)
            write_off_banana_data.append(row_data)

        # Сохранение данных из таблицы self.heavy_greens_table
        heavy_greens_table_data = []
        for row in range(self.heavy_greens_table.rowCount()):
            row_data = []
            for column in range(self.heavy_greens_table.columnCount()):
                item = self.heavy_greens_table.item(row, column)
                if column == 3:  # Если текущий столбец - индекс 3, заменяем на None
                    row_data.append(None)
                else:
                    if item is not None:
                        row_data.append(item.text())
                    else:
                        row_data.append(None)
            heavy_greens_table_data.append(row_data)

        # Сохранение данных из таблицы self.plan_table
        plan_table_data = []
        for row in range(self.plan_table.rowCount()):
            row_data = []
            for column in range(self.plan_table.columnCount()):
                item = self.plan_table.item(row, column)
                if column in [3, 5]:  # Если текущий столбец в списке столбцов для замены на None
                    row_data.append(None)
                else:
                    if item is not None:
                        row_data.append(item.text())
                    else:
                        row_data.append(None)
            plan_table_data.append(row_data)

        data_to_save = {
            'table_data': table_data,
            'additional_data': additional_data,
            'write_off_table_data': write_off_table_data,
            'write_off_banana_data': write_off_banana_data,
            'heavy_greens_table_data': heavy_greens_table_data,
            'plan_table_data': plan_table_data  # Добавление данных таблицы self.plan_table
        }
        with open(filename, 'w') as f:
            json.dump(data_to_save, f)
        current_datetime = QDateTime.currentDateTime().toString("dd.MM.yyyy hh:mm:ss")
        self.setWindowTitle(f"Створення щотижневого листа для магазинів та кураторів / Збережено: {current_datetime}")

    def load_data(self, filename):
        try:
            with open(filename, 'r') as f:
                data = json.load(f)
        except FileNotFoundError:
            QMessageBox.warning(self, "Помилка", "Ти ще нічого не зберіг!")
            # Обработайте ошибку здесь, например, уведомив пользователя или выполнив другие действия
            return

        # Загрузка данных из дополнительных полей ввода
        additional_data = data.get('additional_data', {})
        self.greeting_review_textedit.setPlainText(additional_data.get('greeting_review', ''))
        self.general_comments_textedit.setPlainText(additional_data.get('general_comments', ''))
        self.write_off_comments_textedit.setPlainText(additional_data.get('write_off_comments', ''))
        self.write_off_banana_textedit.setPlainText(additional_data.get('write_off_banana', ''))
        self.heavy_greens_textedit.setPlainText(additional_data.get('heavy_greens', ''))

        # Загрузка данных основной таблицы
        table_data = data.get('table_data', [])
        self.general_table.setRowCount(len(table_data))
        for row, row_data in enumerate(table_data):
            for column, value in enumerate(row_data):
                if value is not None:
                    item = QTableWidgetItem(str(value))
                    self.general_table.setItem(row, column, item)
                    self.update_tables(self.general_table)

        # Загрузка данных из таблицы self.write_off_table
        write_off_table_data = data.get('write_off_table_data', [])
        self.write_off_table.setRowCount(len(write_off_table_data))
        for row, row_data in enumerate(write_off_table_data):
            for column, value in enumerate(row_data):
                if value is not None:
                    item = QTableWidgetItem(str(value))
                    self.write_off_table.setItem(row, column, item)
                    self.update_tables(self.write_off_table)

        # Загрузка данных из таблицы self.write_off_banana_table
        write_off_banana_table_data = data.get('write_off_banana_data', [])
        self.write_off_banana_table.setRowCount(len(write_off_banana_table_data))
        for row, row_data in enumerate(write_off_banana_table_data):
            for column, value in enumerate(row_data):
                if value is not None:
                    item = QTableWidgetItem(str(value))
                    self.write_off_banana_table.setItem(row, column, item)
                    self.update_tables(self.write_off_banana_table)

        # Загрузка данных из таблицы self.heavy_greens_table
        heavy_greens_table_data = data.get('heavy_greens_table_data', [])
        self.heavy_greens_table.setRowCount(len(heavy_greens_table_data))
        for row, row_data in enumerate(heavy_greens_table_data):
            for column, value in enumerate(row_data):
                if value is not None and column != 3:  # Пропускаем столбец с индексом 3
                    item = QTableWidgetItem(str(value))
                    self.heavy_greens_table.setItem(row, column, item)
                    self.update_tables(self.heavy_greens_table)

        # Загрузка данных из таблицы self.plan_table
        plan_table_data = data.get('plan_table_data', [])
        self.plan_table.setRowCount(len(plan_table_data))
        for row, row_data in enumerate(plan_table_data):
            for column, value in enumerate(row_data):
                if value is not None and column not in [3, 5]:  # Пропускаем столбцы с индексами 3 и 5
                    item = QTableWidgetItem(str(value))
                    self.plan_table.setItem(row, column, item)
                    self.update_tables(self.plan_table)


    def import_to_general_table(self):
        clipboard = QApplication.clipboard()
        clipboard_text = clipboard.text().strip()  # Убираем лишние пробелы в начале и конце строки
        
        # Проверка на наличие букв в буфере обмена
        if any(char.isalpha() for char in clipboard_text):
            QMessageBox.warning(self, "Помилка", "В буфері присутні букви. Імпортувати можна тільки цифри та відсотки.")
            return

        # Проверка на наличие только цифр и допустимых спец. символов
        valid_characters = r'[\d\.,%\s]'
        if not all(re.match(valid_characters, char) for char in clipboard_text):
            QMessageBox.warning(self, "Помилка", "В скопійованих рядках є недопустимі символи. Імпортувати можна тільки цифри та відсотки.")
            return

        # Замена точки на запятую и удаление знака процента
        clipboard_text = clipboard_text.replace('.', ',').replace('%', '')

        rows = clipboard_text.split('\n')

        # Проверяем, что количество строк в буфере обмена соответствует количеству строк в таблице
        if len(rows) != self.general_table.rowCount():
            QMessageBox.warning(self, "Помилка", "Кількість скопійованих рядків не співпадає с кількістю магазинів.")
            return

        # Получаем названия столбцов
        header_labels = [self.general_table.horizontalHeaderItem(i).text() for i in range(self.general_table.columnCount())]

        # Создаем список столбцов для выбора
        column_items = [label for label in header_labels if header_labels.index(label) not in [0, 3, 6, 7]]

        # Показываем диалоговое окно с выбором столбца
        column_name, ok = QInputDialog.getItem(self, "Вибір стовпчика", "Оберіть стовпчик для імпорту данних:", column_items, 0, False)
        if not ok:
            return

        # Определяем индекс выбранного столбца
        column_index = header_labels.index(column_name)

        # Заполняем только выбранный столбец
        for i, row in enumerate(rows):
            cols = row.split('\t')
            col = cols[0] if len(cols) > 0 else ''  # Используем индекс 0 для выбора первого столбца
            item = QTableWidgetItem(col)
            self.general_table.setItem(i, column_index, item)  # Установка значения в выбранный столбец
    
    def import_to_write_off_table(self):
        clipboard = QApplication.clipboard()
        clipboard_text = clipboard.text().strip()  # Убираем лишние пробелы в начале и конце строки

        # Проверка на наличие букв в буфере обмена
        if any(char.isalpha() for char in clipboard_text):
            QMessageBox.warning(self, "Помилка", "В буфере обмена присутствуют буквы.")
            return

        # Проверка на наличие только цифр и допустимых спец. символов
        valid_characters = r'[\d\.,%\s]'
        if not all(re.match(valid_characters, char) for char in clipboard_text):
            QMessageBox.warning(self, "Помилка", "В скопійованих рядках є недопустимі символи. Імпортувати можна тільки цифри та відсотки.")
            return

        # Удаляем точки и проценты
        clipboard_text = clipboard_text.replace('.', '').replace('%', '')

        # Проверка на пустые строки
        if not clipboard_text.strip():
            QMessageBox.warning(self, "Помилка", "В буфере обмена відсутні дані для імпорту.")
            return

        rows = clipboard_text.split('\n')

        # Проверяем, что количество строк в буфере обмена соответствует количеству строк в таблице
        if len(rows) != self.write_off_table.rowCount():
            QMessageBox.warning(self, "Помилка", "Кількість скопійованих рядків не співпадає с кількістю магазинів.")
            return

        # Получаем названия столбцов
        header_labels = [self.write_off_table.horizontalHeaderItem(i).text() for i in range(self.write_off_table.columnCount())]

        # Создаем список столбцов для выбора
        column_items = [label for label in header_labels if header_labels.index(label) in [1, 2]]

        # Показываем диалоговое окно с выбором столбца
        column_name, ok = QInputDialog.getItem(self, "Вибір стовпчика", "Оберіть стовпчик для імпорту данних:", column_items, 0, False)
        if not ok:
            return

        # Определяем индекс выбранного столбца
        column_index = header_labels.index(column_name)

        # Заполняем только выбранный столбец
        for i, row in enumerate(rows):
            cols = row.split('\t')
            col = cols[0] if len(cols) > 0 else ''  # Используем индекс 0 для выбора первого столбца
            item = QTableWidgetItem(col)
            self.write_off_table.setItem(i, column_index, item)  # Установка значения в выбранный столбец

    def import_to_write_off_banana_table(self):
        clipboard = QApplication.clipboard()
        clipboard_text = clipboard.text().strip()  # Убираем лишние пробелы в начале и конце строки

        # Проверка на наличие букв в буфере обмена
        if any(char.isalpha() for char in clipboard_text):
            QMessageBox.warning(self, "Помилка", "В буфере обмена присутствуют буквы.")
            return

        # Проверка на наличие только цифр и допустимых спец. символов
        valid_characters = r'[\d\.,%\s]'
        if not all(re.match(valid_characters, char) for char in clipboard_text):
            QMessageBox.warning(self, "Помилка", "В скопійованих рядках є недопустимі символи. Імпортувати можна тільки цифри та відсотки.")
            return

        # Удаляем точки и проценты
        clipboard_text = clipboard_text.replace('.', '').replace('%', '')

        # Проверка на пустые строки
        if not clipboard_text.strip():
            QMessageBox.warning(self, "Помилка", "В буфере обмена відсутні дані для імпорту.")
            return

        rows = clipboard_text.split('\n')

        # Проверяем, что количество строк в буфере обмена соответствует количеству строк в таблице
        if len(rows) != self.write_off_banana_table.rowCount():
            QMessageBox.warning(self, "Помилка", "Кількість скопійованих рядків не співпадає с кількістю магазинів.")
            return

        # Получаем названия столбцов
        header_labels = [self.write_off_banana_table.horizontalHeaderItem(i).text() for i in range(self.write_off_banana_table.columnCount())]

        # Создаем список столбцов для выбора
        column_items = [label for label in header_labels if header_labels.index(label) in [1, 2, 3, 4]]

        # Показываем диалоговое окно с выбором столбца
        column_name, ok = QInputDialog.getItem(self, "Вибір стовпчика", "Оберіть стовпчик для імпорту данних:", column_items, 0, False)
        if not ok:
            return

        # Определяем индекс выбранного столбца
        column_index = header_labels.index(column_name)

        # Заполняем выбранные столбцы
        for i, row in enumerate(rows):
            cols = row.split('\t')
            for j, col in enumerate(cols[:4]):  # Используем только первые четыре значения для столбцов
                item = QTableWidgetItem(col)
                self.write_off_banana_table.setItem(i, column_index + j, item)  # Установка значения в выбранные столбцы

    def import_to_heavy_greens_table(self):
        clipboard = QApplication.clipboard()
        clipboard_text = clipboard.text().strip()  # Убираем лишние пробелы в начале и конце строки

        # Проверка на наличие букв в буфере обмена
        if any(char.isalpha() for char in clipboard_text):
            QMessageBox.warning(self, "Помилка", "В буфере обмена присутствуют буквы.")
            return

        # Проверка на наличие только цифр и допустимых спец. символов
        valid_characters = r'[\d\.,%\s]'
        if not all(re.match(valid_characters, char) for char in clipboard_text):
            QMessageBox.warning(self, "Помилка", "В скопійованих рядках є недопустимі символи. Імпортувати можна тільки цифри та відсотки.")
            return

        # Удаляем точки и проценты
        clipboard_text = clipboard_text.replace('.', '').replace('%', '')

        # Проверка на пустые строки
        if not clipboard_text.strip():
            QMessageBox.warning(self, "Помилка", "В буфере обмена відсутні дані для імпорту.")
            return

        rows = clipboard_text.split('\n')

        # Проверяем, что количество строк в буфере обмена соответствует количеству строк в таблице
        if len(rows) != self.heavy_greens_table.rowCount():
            QMessageBox.warning(self, "Помилка", "Кількість скопійованих рядків не співпадає с кількістю магазинів.")
            return

        # Получаем названия столбцов
        header_labels = [self.heavy_greens_table.horizontalHeaderItem(i).text() for i in range(self.heavy_greens_table.columnCount())]

        # Создаем список столбцов для выбора
        column_items = [label for label in header_labels if header_labels.index(label) in [1, 2, 4]]

        # Показываем диалоговое окно с выбором столбца
        column_name, ok = QInputDialog.getItem(self, "Вибір стовпчика", "Оберіть стовпчик для імпорту данних:", column_items, 0, False)
        if not ok:
            return

        # Определяем индекс выбранного столбца
        column_index = header_labels.index(column_name)

        # Заполняем выбранные столбцы
        for i, row in enumerate(rows):
            cols = row.split('\t')
            for j, col in enumerate(cols[:3]):  # Используем только первые три значения для столбцов
                item = QTableWidgetItem(col)
                self.heavy_greens_table.setItem(i, column_index + j, item)  # Установка значения в выбранные столбцы

    def import_to_plan_table(self):
        clipboard = QApplication.clipboard()
        clipboard_text = clipboard.text().strip()  # Убираем лишние пробелы в начале и конце строки

        # Проверка на наличие букв в буфере обмена
        if any(char.isalpha() for char in clipboard_text):
            QMessageBox.warning(self, "Помилка", "В буфере обмена присутствуют буквы.")
            return

        # Проверка на наличие только цифр и допустимых спец. символов
        valid_characters = r'[\d\.,%\s]'
        if not all(re.match(valid_characters, char) for char in clipboard_text):
            QMessageBox.warning(self, "Помилка", "В скопійованих рядках є недопустимі символи. Імпортувати можна тільки цифри та відсотки.")
            return

        # Удаляем пробелы
        clipboard_text = clipboard_text.replace(' ', '')

        # Удаляем точки и проценты
        clipboard_text = clipboard_text.replace('.', '').replace('%', '')

        # Проверка на пустые строки
        if not clipboard_text.strip():
            QMessageBox.warning(self, "Помилка", "В буфере обмена відсутні дані для імпорту.")
            return

        rows = clipboard_text.split('\n')

        # Проверяем, что количество строк в буфере обмена соответствует количеству строк в таблице
        if len(rows) != self.plan_table.rowCount():
            QMessageBox.warning(self, "Помилка", "Кількість скопійованих рядків не співпадає с кількістю магазинів.")
            return

        # Получаем названия столбцов
        header_labels = [self.plan_table.horizontalHeaderItem(i).text() for i in range(self.plan_table.columnCount())]

        # Создаем список столбцов для выбора
        column_items = [label for label in header_labels if header_labels.index(label) in [1, 2, 4]]

        # Показываем диалоговое окно с выбором столбца
        column_name, ok = QInputDialog.getItem(self, "Вибір стовпчика", "Оберіть стовпчик для імпорту данних:", column_items, 0, False)
        if not ok:
            return

        # Определяем индекс выбранного столбца
        column_index = header_labels.index(column_name)

        # Заполняем выбранные столбцы
        for i, row in enumerate(rows):
            cols = row.split('\t')
            for j, col in enumerate(cols[:3]):  # Используем только первые три значения для столбцов
                item = QTableWidgetItem(col)
                self.plan_table.setItem(i, column_index + j, item)  # Установка значения в выбранные столбцы

    def on_combobox_changed(self, index):
        # Если это первое событие изменения выбора, сохраняем начальное значение индекса
        if self.previous_index is None:
            self.previous_index = index
            return  # Выходим из метода, так как нет предыдущего индекса для сравнения

        # Получаем текущий выбор в выпадающем списке
        current_index = self.kust_combobox.currentIndex()

        # Если выбор изменился
        if current_index != self.previous_index:
            # Отображаем диалоговое окно подтверждения
            reply = QMessageBox.question(self, 'Зміна куща', 'Ти впевнений, що хочеш змінити кущ?\nВсі внесені дані будуть видалені!',
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

            # Если пользователь подтвердил выбор
            if reply == QMessageBox.Yes:
                self.general_table.setRowCount(0)
                self.write_off_table.setRowCount(0)
                self.write_off_banana_table.setRowCount(0)
                self.heavy_greens_table.setRowCount(0)
                self.plan_table.setRowCount(0)
                self.update_tables(current_index)
                self.previous_index = current_index  # Обновляем предыдущий индекс выбора
            else:
                # Возвращаем предыдущий выбор в выпадающем списке
                self.kust_combobox.setCurrentIndex(self.previous_index)
        else:
            # Если выбор не изменился, обновляем предыдущий выбор
            self.previous_index = current_index
    
    def on_combobox_week_changed(self, index):
        selected_week = int(self.week_combobox.currentText()) - 1
        self.heavy_greens_table.setHorizontalHeaderLabels([
            'Філіал', f'Продажі {self.week_combobox.currentText()} тиждень, кг',
            f'Продажі {selected_week} тиждень, кг', '% приросту', 'Списання, кг'
        ])

    def recalculate_general_data(self, item):
        if item.column() in [1, 2, 4, 5, 8]:
            if item.text() is not None:
                value = item.text()
                if value and '%' not in value:
                    if value.count(',') <= 1:
                        if value == ',':
                            item.setText("")
                        else:
                            value_with_comma = value.replace('.', ',')
                            item.setText(value_with_comma)
                    else:
                        item.setText("")

            if item.column() in [1, 2, 8]:
                if item.text() is not None:
                    value = item.text()
                    if value and '%' not in value:
                        value_with_comma = value.replace('.', ',')
                        item.setText(f"{value_with_comma}%")

            if item.column() in [1, 2, 4, 5, 8]:
                row = item.row()
                penetration_item = self.general_table.item(row, 1)
                actual_item = self.general_table.item(row, 2)
                retention_item = self.general_table.item(row, 4)
                conversion_item = self.general_table.item(row, 5)
                availability_item = self.general_table.item(row, 7)

                if penetration_item is not None and actual_item is not None and item.column() in [1, 2]:
                    penetration_text = penetration_item.text().replace('%', '') if penetration_item.text() is not None else ''
                    actual_text = actual_item.text().replace('%', '') if actual_item.text() is not None else ''

                    if all(text for text in [penetration_text, actual_text]):
                        penetration = float(penetration_text.replace(',', '.')) if penetration_text and penetration_text != 'None' else 0
                        actual = float(actual_text.replace(',', '.')) if actual_text and actual_text != 'None' else 0
                        if penetration != 0:
                            percentage = round((actual / penetration) * 100, 2)
                        else:
                            percentage = 0
                        percentage_text = f"{percentage:.2f}%".replace('.', ',')

                        penetration_plan_item = QTableWidgetItem(percentage_text)
                        penetration_plan_item.setFlags(penetration_plan_item.flags() & ~Qt.ItemIsEditable)

                        if percentage > 100:
                            penetration_plan_item.setBackground(QColor(204, 255, 204))
                        else:
                            penetration_plan_item.setBackground(QColor(255, 204, 204))

                        self.general_table.setItem(row, 3, penetration_plan_item)
                    else:
                        self.general_table.setItem(row, 3, QTableWidgetItem(""))

                elif retention_item is not None and conversion_item is not None and item.column() in [4, 5]:
                    retention_text = retention_item.text().replace('%', '').replace(' ', '') if retention_item.text() is not None else ''
                    conversion_text = conversion_item.text().replace('%', '').replace(' ', '') if conversion_item.text() is not None else ''

                    if all(text for text in [retention_text, conversion_text]):
                        retention = float(retention_text.replace(',', '.')) if retention_text else 0
                        conversion = float(conversion_text.replace(',', '.')) if conversion_text else 0

                        if retention != 0:
                            result = round(conversion / retention * 100, 2)
                        else:
                            result = 0.0
                        result_text = f"{result:.2f}%".replace('.', ',')

                        result_item = QTableWidgetItem(result_text)
                        result_item.setFlags(result_item.flags() & ~Qt.ItemIsEditable)

                        if result > 100:
                            result_item.setBackground(QColor(204, 255, 204))
                        else:
                            result_item.setBackground(QColor(255, 204, 204))

                        self.general_table.setItem(row, 6, result_item)
                    else:
                        self.general_table.setItem(row, 6, QTableWidgetItem(""))

                elif item.column() == 8:
                    availability = item.text().replace('%', '') if item.text() is not None else ''

                    if availability is not None and availability != 'None':
                        availability_value = float(availability.replace(',', '.')) if availability else 0
                        if availability_value > 93:
                            item.setBackground(QColor(204, 255, 204))
                        else:
                            item.setBackground(QColor(255, 204, 204))
                    else:
                        item.setBackground(QColor(Qt.white))

        for row in range(self.general_table.rowCount()):
            if self.general_table.item(row, 7) is None or self.general_table.item(row, 7).text() == "":
                item = QTableWidgetItem("93%")
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.general_table.setItem(row, 7, item)
                
    def handle_item_changed(self, item):
        if item.column() in [4, 5]:
            # Получаем текст из ячейки
            value = item.text()
            # Проверяем, что ячейка не пустая и не содержит символ процента
            if value and '%' not in value:
                # Проверка на количество запятых
                if value.count(',') <= 1:
                    # Удаляем пробелы из значения
                    value = value.replace(' ', '')
                    # Проверяем, содержит ли значение только запятую, без числа
                    if value != ',':
                        # Преобразуем значение в число
                        if value is not None and value != 'None':
                            number = float(value.replace(',', '.'))
                        else:
                            number = 0.0
                        # Форматируем число в строку без дробной части, если она равна нулю
                        formatted_value = '{:,.0f}'.format(number) if number.is_integer() else '{:,.2f}'.format(number)
                        # Разделяем число на разряды с помощью пробела
                        formatted_value = formatted_value.replace(',', ' ').replace('.', ',')
                        # Устанавливаем отформатированное значение в ячейку
                        item.setText(formatted_value)
                    else:
                        item.setText("")  # Очистить ячейку, если значение только запятая
                else:
                    item.setText("")  # Очистить ячейку, если больше одной запятой
            else:
                item.setText("")  # Очистить ячейку, если в ней есть символ процента

    def recalculate_write_off_data(self, item):
        if item.column() in [1, 2]:  
            row = item.row()
            column1_item = self.write_off_table.item(row, 1)
            column2_item = self.write_off_table.item(row, 2)
            
            if all(item is not None for item in [column1_item, column2_item]):
                column1_text = column1_item.text()
                column2_text = column2_item.text()
                
                if all(text for text in [column1_text, column2_text]):
                    # Удаляем знак процента, если он присутствует в значении ячейки
                    column1_value = float(column1_text.replace(',', '.').replace('%', ''))  
                    column2_value = float(column2_text.replace(',', '.').replace('%', ''))  
                    percentage = round((column1_value / column2_value) * 100, 2)
                    
                    # Добавляем данные с пользовательской ролью (UserRole) в ячейку
                    result_item = QTableWidgetItem()
                    result_item.setData(Qt.DisplayRole, f"{percentage:.2f}%")
                    result_item.setData(Qt.UserRole, percentage)  # Данные без знака процента
                    result_item.setFlags(result_item.flags() & ~Qt.ItemIsEditable)
                    
                    if percentage > 100:
                        result_item.setBackground(QColor(204, 255, 204))  
                    else:
                        result_item.setBackground(QColor(255, 204, 204))  
                        
                    self.write_off_table.setItem(row, 3, result_item)
                else:
                    self.write_off_table.setItem(row, 3, QTableWidgetItem(""))
            else:
                self.write_off_table.setItem(row, 3, QTableWidgetItem(""))
        
        if item.column() in [1, 2] and item.text():  # Если пользователь изменил значение в столбце 1 или 2 и ячейка не пустая
            value = item.text()  # Получаем текст из ячейки
            if '%' not in value:  # Проверяем, что знак процента отсутствует
                item.setText(f"{value}%")  # Добавляем знак процента к числу

    def recalculate_write_off_banana_data(self, item):
        value = None  # Инициализируем переменную value

        if item.text():
            value = item.text()
            # Проверяем, является ли введенное значение числом с плавающей точкой (дробным)
            try:
                # Используем запятую для разделения дробной части
                float_value = float(value.replace(',', '.'))
                # Проверяем, является ли число целым
                is_integer = float_value.is_integer()
                if not is_integer:  # Если число не целое, округляем до двух знаков после запятой
                    value = f"{float_value:.2f}".replace('.', ',')  # Используем запятую вместо точки
                item.setText(value)  # Обновляем текст в ячейке
            except ValueError:
                pass  # Пропускаем, если введенное значение не является числом

        if item.column() == 1:  # Проверяем, что это первый столбец
            row = item.row()
            column1_item = item
            column2_item = self.write_off_banana_table.item(row, 2)

            if column1_item and column1_item.text() and column2_item and column2_item.text():
                column1_value = float(column1_item.text().replace(',', '.').replace('%', ''))
                column2_value = float(column2_item.text().replace(',', '.').replace('%', ''))

                if column2_value < column1_value:
                    column2_item.setBackground(QColor(204, 255, 204))  # Окрашиваем второй столбец в зеленый цвет
                else:
                    column2_item.setBackground(QColor(255, 204, 204))  # Окрашиваем второй столбец в красный цвет

            elif not column1_item.text() and column2_item:  # Если значение в первом столбце удалено, сбросить окрашивание
                column2_item.setBackground(QColor(255, 255, 255))  # Сбрасываем цвет фона

            if value and '%' not in value: 
                item.setText(f"{value}%")  # Добавляем знак процента, если его нет

        elif item.column() == 2:  # Проверяем, что это второй столбец
            row = item.row()
            column1_item = self.write_off_banana_table.item(row, 1)
            column2_item = item

            if column1_item and column1_item.text() and column2_item and column2_item.text():
                column1_value = float(column1_item.text().replace(',', '.').replace('%', ''))
                column2_value = float(column2_item.text().replace(',', '.').replace('%', ''))

                if column2_value < column1_value:
                    column2_item.setBackground(QColor(204, 255, 204))  # Окрашиваем второй столбец в зеленый цвет
                else:
                    column2_item.setBackground(QColor(255, 204, 204))  # Окрашиваем второй столбец в красный цвет

            elif not column2_item.text() and column1_item:  # Если значение во втором столбце удалено, сбросить окрашивание
                column2_item.setBackground(QColor(255, 255, 255))  # Сбрасываем цвет фона

            if value and '%' not in value: 
                item.setText(f"{value}%")  # Добавляем знак процента, если его нет

    def recalculate_heavy_greens_data(self, item):
        value = None  # Инициализируем переменную value

        if item.text():
            value = item.text()
            # Проверяем, является ли введенное значение числом с плавающей точкой (дробным)
            try:
                # Используем запятую для разделения дробной части
                float_value = float(value.replace(',', '.'))
                # Проверяем, является ли число целым
                is_integer = float_value.is_integer()
                if not is_integer:  # Если число не целое, округляем до двух знаков после запятой
                    value = f"{float_value:.2f}".replace('.', ',')  # Используем запятую вместо точки
                item.setText(value)  # Обновляем текст в ячейке
            except ValueError:
                pass  # Пропускаем, если введенное значение не является числом

        if item.column() in [1, 2]:  
            row = item.row()
            column1_item = self.heavy_greens_table.item(row, 1)
            column2_item = self.heavy_greens_table.item(row, 2)
            
            if all(item is not None for item in [column1_item, column2_item]):
                column1_text = column1_item.text()
                column2_text = column2_item.text()
                
                if all(text for text in [column1_text, column2_text]):
                    # Удаляем знак процента, если он присутствует в значении ячейки
                    column1_value = float(column1_text.replace(',', '.').replace('%', ''))  
                    column2_value = float(column2_text.replace(',', '.').replace('%', ''))  
                    percentage = round(((column1_value / column2_value) * 100) - 100, 2)
                    
                    # Добавляем данные с пользовательской ролью (UserRole) в ячейку
                    result_item = QTableWidgetItem()
                    result_item.setData(Qt.DisplayRole, f"{percentage:.2f}%")
                    result_item.setData(Qt.UserRole, percentage)  # Данные без знака процента
                    result_item.setFlags(result_item.flags() & ~Qt.ItemIsEditable)
                    
                    if percentage > 0:
                        result_item.setBackground(QColor(204, 255, 204))  
                    else:
                        result_item.setBackground(QColor(255, 204, 204))  
                        
                    self.heavy_greens_table.setItem(row, 3, result_item)
                else:
                    self.heavy_greens_table.setItem(row, 3, QTableWidgetItem(""))
            else:
                self.heavy_greens_table.setItem(row, 3, QTableWidgetItem(""))

    def recalculate_plan_data(self, item):
        if item.column() == 2:  # Проверяем, что столбец является 2
            value = item.text()
            if value:
                # Удаляем пробелы из введенного значения
                value = value.replace(' ', '')
                try:
                    # Проверяем, является ли число дробным
                    is_float = ',' in value
                    # Добавляем разделение чисел на разряды с использованием пробела и запятой для разделителя дробной части
                    value_with_spaces = '{:,.2f}'.format(float(value.replace(',', '.'))).replace(',', ' ').replace('.', ',') if is_float else '{:,.0f}'.format(float(value.replace(',', '.'))).replace(',', ' ').replace('.', ',')
                    column2_value = float(value.replace(',', '.'))  # Преобразуем текст в число
                    result = round(column2_value / 7)  # Выполняем вычисление по формуле и округляем
                    # Форматируем результат вычислений: добавляем дробную часть только для дробных чисел
                    result_formatted = '{:,.2f}'.format(result).replace(',', ' ').replace('.', ',') if is_float else '{:,.0f}'.format(result).replace(',', ' ').replace('.', ',')
                    # Создаем элемент для ячейки со значением результата и разделением на разряды
                    result_item = QTableWidgetItem(result_formatted)
                    result_item.setFlags(result_item.flags() & ~Qt.ItemIsEditable)  # Делаем ячейку нередактируемой
                    self.plan_table.setItem(item.row(), 3, result_item)  # Устанавливаем элемент в столбец 5
                    item.setText(value_with_spaces)  # Возвращаем значение с разделенными разрядами
                except ValueError:
                    pass  # Пропускаем, если введенное значение не является числом

        if item.column() in [1, 4]: 
            value = item.text()
            if value and '%' not in value:  # Проверяем, что переменная value определена
                item.setText(f"{value}%")  # Добавляем знак процента, если его нет

        for row in range(self.plan_table.rowCount()):
            item = QTableWidgetItem("93%")
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            if self.plan_table.item(row, 5) is None or self.plan_table.item(row, 5).text() == "":
                self.plan_table.setItem(row, 5, item)
            else:
                # Если ячейка не пуста, ничего не меняем
                pass

  
    def update_tables(self, index):
        kust = self.kust_combobox.currentText()
        filials = self.filials_data.get(kust, {}).get('адреса', [])
        
        # Для таблицы общих показателей
        self.general_table.setRowCount(len(filials))
        for i, filial in enumerate(filials):
            item = QTableWidgetItem(filial)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Устанавливаем флаг "нередактируемости"
            self.general_table.setItem(i, 0, item)
        
        # Для таблицы списаний
        self.write_off_table.setRowCount(len(filials))
        for i, filial in enumerate(filials):
            item = QTableWidgetItem(filial)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Устанавливаем флаг "нередактируемости"
            self.write_off_table.setItem(i, 0, item)

        # Для таблицы списаний банана
        self.write_off_banana_table.setRowCount(len(filials))
        for i, filial in enumerate(filials):
            item = QTableWidgetItem(filial)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Устанавливаем флаг "нередактируемости"
            self.write_off_banana_table.setItem(i, 0, item)

        # Для таблицы зеленого базара
        self.heavy_greens_table.setRowCount(len(filials))
        for i, filial in enumerate(filials):
            item = QTableWidgetItem(filial)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Устанавливаем флаг "нередактируемости"
            self.heavy_greens_table.setItem(i, 0, item)

        # Для таблицы с планом
        self.plan_table.setRowCount(len(filials))
        for i, filial in enumerate(filials):
            item = QTableWidgetItem(filial)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Устанавливаем флаг "нередактируемости"
            self.plan_table.setItem(i, 0, item)

        # Устанавливаем высоту таблиц в соответствии с количеством строк
        self.set_table_height(self.general_table)
        self.set_table_height(self.write_off_table)
        self.set_table_height(self.write_off_banana_table)
        self.set_table_height(self.heavy_greens_table)
        self.set_table_height(self.plan_table)

    def set_table_height(self, table):
        table_height = table.horizontalHeader().height()  # Получаем высоту заголовка таблицы
        table_height += table.rowHeight(0) * table.rowCount()  # Высота каждой строки * количество строк
        table.setFixedHeight(table_height + 10)

    def ask_to_save_data(self):
        reply = QMessageBox.question(None, "Збереження даних", "Бажаєте зберегти дані перед відкриттям форми листа?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        return reply == QMessageBox.Yes

    def send_email(self):
        if self.ask_to_save_data():
            self.save_data(filesave)
        general_comments = self.general_comments_textedit.toPlainText()
        write_off_comments = self.write_off_comments_textedit.toPlainText()
        greeting_review = self.greeting_review_textedit.toPlainText()
        write_off_banana = self.write_off_banana_textedit.toPlainText()
        heavy_greens_сomments = self.heavy_greens_textedit.toPlainText()

        general_html = """
        <table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0
        width="100%" style='width:100.0%;border-collapse:collapse;border:none;
        mso-border-alt:solid windowtext .5pt;mso-yfti-tbllook:1184;mso-padding-alt:
        0cm 5.4pt 0cm 5.4pt'>
        <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:29.95pt'>
        <td width="28%" rowspan=2 style='width:28.3%;border:solid windowtext 1.0pt;
        mso-border-alt:solid windowtext .5pt;background:#215E99;mso-background-themecolor:
        text2;mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;
        height:29.95pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center;'><b><span lang=UK
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'>Філіал<o:p></o:p></span></b></p>
        </td>
        <td width="27%" colspan=3 style='width:27.08%;border:solid windowtext 1.0pt;
        border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
        solid windowtext .5pt;background:#215E99;mso-background-themecolor:text2;
        mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;height:29.95pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><span class=spelle><b><span
        lang=UK style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'>Пенетрація</span></b></span><b><span
        lang=EN-US style='font-family:"Helvetica",sans-serif;color:white;
        mso-themecolor:background1;mso-ansi-language:EN-US'><o:p></o:p></span></b></p>
        </td>
        <td width="31%" colspan=3 style='width:31.84%;border:solid windowtext 1.0pt;
        border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
        solid windowtext .5pt;background:#215E99;mso-background-themecolor:text2;
        mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;height:29.95pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><span class=spelle><b><span
        lang=RU style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:RU'>Виторг</span></b></span><b><span
        lang=RU style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:RU'><o:p></o:p></span></b></p>
        </td>
        <td width="12%" colspan=2 style='width:12.78%;border:solid windowtext 1.0pt;
        border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
        solid windowtext .5pt;background:#215E99;mso-background-themecolor:text2;
        mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;height:29.95pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><span class=spelle><b><span
        lang=RU style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:RU'>Доступн</span></b></span><span
        class=spelle><b><span lang=UK style='font-family:"Helvetica",sans-serif;
        color:white;mso-themecolor:background1;mso-ansi-language:UK'>ість</span></b></span><b><span
        lang=UK style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'> продажів</span></b><b><span
        lang=EN-US style='font-family:"Helvetica",sans-serif;color:white;
        mso-themecolor:background1;mso-ansi-language:EN-US'><o:p></o:p></span></b></p>
        </td>
        </tr>
        <tr style='mso-yfti-irow:1;height:7.55pt'>
        <td width="7%" style='width:7.86%;border-top:none;border-left:none;
        border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
        mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
        mso-border-alt:solid windowtext .5pt;background:#215E99;mso-background-themecolor:
        text2;mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;
        height:7.55pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><b><span lang=UK
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'>План<o:p></o:p></span></b></p>
        </td>
        <td width="8%" style='width:8.66%;border-top:none;border-left:none;
        border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
        mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
        mso-border-alt:solid windowtext .5pt;background:#215E99;mso-background-themecolor:
        text2;mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;
        height:7.55pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><b><span lang=UK
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'>Факт<o:p></o:p></span></b></p>
        </td>
        <td width="10%" style='width:10.56%;border-top:none;border-left:none;
        border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
        mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
        mso-border-alt:solid windowtext .5pt;background:#215E99;mso-background-themecolor:
        text2;mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;
        height:7.55pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><b><span lang=UK
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'>% виконання<o:p></o:p></span></b></p>
        </td>
        <td width="10%" style='width:10.1%;border-top:none;border-left:none;
        border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
        mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
        mso-border-alt:solid windowtext .5pt;background:#215E99;mso-background-themecolor:
        text2;mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;
        height:7.55pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><b><span lang=UK
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'>План</span></b><b><span lang=EN-US
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:EN-US'><o:p></o:p></span></b></p>
        </td>
        <td width="10%" style='width:10.4%;border-top:none;border-left:none;
        border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
        mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
        mso-border-alt:solid windowtext .5pt;background:#215E99;mso-background-themecolor:
        text2;mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;
        height:7.55pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><b><span lang=UK
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'>Факт</span></b><b><span lang=EN-US
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:EN-US'><o:p></o:p></span></b></p>
        </td>
        <td width="11%" style='width:11.34%;border-top:none;border-left:none;
        border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
        mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
        mso-border-alt:solid windowtext .5pt;background:#215E99;mso-background-themecolor:
        text2;mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;
        height:7.55pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><b><span lang=UK
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'>% виконання</span></b><b><span
        lang=EN-US style='font-family:"Helvetica",sans-serif;color:white;
        mso-themecolor:background1;mso-ansi-language:EN-US'><o:p></o:p></span></b></p>
        </td>
        <td width="6%" style='width:6.34%;border-top:none;border-left:none;
        border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
        mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
        mso-border-alt:solid windowtext .5pt;background:#215E99;mso-background-themecolor:
        text2;mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;
        height:7.55pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><b><span lang=UK
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'>План<o:p></o:p></span></b></p>
        </td>
        <td width="6%" style='width:6.44%;border-top:none;border-left:none;
        border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
        mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
        mso-border-alt:solid windowtext .5pt;background:#215E99;mso-background-themecolor:
        text2;mso-background-themetint:191;padding:0cm 5.4pt 0cm 5.4pt;
        height:7.55pt'>
        <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
        mso-margin-bottom-alt:auto;text-align:center'><b><span lang=UK
        style='font-family:"Helvetica",sans-serif;color:white;mso-themecolor:
        background1;mso-ansi-language:UK'>Факт<o:p></o:p></span></b></p>
        </td>
        </tr>
        """
        for row in range(0, self.general_table.rowCount()):  # Начинаем с 1, чтобы пропустить первую строку с заголовками
            general_html += "<tr>"
            for column in range(self.general_table.columnCount()):
                item = self.general_table.item(row, column)
                if item is not None:
                    general_html += f"<td style='border: 1px solid black; color:black; font-family: Helvetica, sans-serif; '>{item.text()}</td>"
                else:
                    general_html += "<td style='border: 1px solid black; font-family: Helvetica, sans-serif; '></td>"
            general_html += "</tr>"
        general_html += "</table>"

        # Сначала создадим список кортежей, содержащих данные о строках и их значениях в третьем столбце
        rows_data = []
        for row in range(self.write_off_table.rowCount()):
            item = self.write_off_table.item(row, 3)  # третий столбец
            if item is not None and item.text().strip():  # Проверяем, что элемент не None и текст не пустой после удаления лишних пробелов
                text = item.text()
                if '%' in text:  # Если встречается символ процента
                    value = float(text.replace('%', '')) / 100  # Убираем символ процента и делим на 100
                else:
                    value = float(text)
                rows_data.append((row, value))

        # Отсортируем список кортежей по значению в третьем столбце в убывающем порядке
        sorted_rows_data = sorted(rows_data, key=lambda x: x[1], reverse=True)

        # Теперь формируем HTML таблицу, используя отсортированные данные
        write_off_html = "<table border='1' cellpadding=2 style='border-collapse: collapse; font-family: Helvetica, sans-serif; width:100%'>"
        write_off_html += "<tr>"
        for column in range(self.write_off_table.columnCount()):
            # Добавляем стили для заголовков
            write_off_html += f"<th style='border: 1px solid black; background-color: #C65911; color: white;'>{self.write_off_table.horizontalHeaderItem(column).text()}</th>"
        write_off_html += "</tr>"
        for row, _ in sorted_rows_data:
            write_off_html += "<tr>"
            for column in range(self.write_off_table.columnCount()):
                item = self.write_off_table.item(row, column)
                if item is not None:
                    cell_value = item.text()
                    # Добавляем заливку ячейки в четвертом столбце в соответствии с условием
                    if column == 3:  # четвертый столбец
                        if '%' in cell_value:  # Если встречается символ процента
                            value = float(cell_value.replace('%', ''))  # Убираем символ процента
                        else:
                            value = float(cell_value)
                        if value >= 100:
                            write_off_html += f"<td style='border: 1px solid black; background-color: rgb(217, 242, 208);'>{cell_value}</td>"
                        else:
                            write_off_html += f"<td style='border: 1px solid black; background-color: rgb(250, 226, 213);'>{cell_value}</td>"
                    else:
                        write_off_html += f"<td style='border: 1px solid black;'>{cell_value}</td>"
                else:
                    write_off_html += "<td style='border: 1px solid black;'></td>"
            write_off_html += "</tr>"
        write_off_html += "</table>"

        write_off_banana_html = ""
        # Создайте список кортежей, содержащих данные из таблицы
        table_data = []
        for row in range(self.write_off_banana_table.rowCount()):
            row_data = []
            for column in range(self.write_off_banana_table.columnCount()):
                item = self.write_off_banana_table.item(row, column)
                if item is not None:
                    row_data.append(item.text())
                else:
                    row_data.append("")
            table_data.append(row_data)

        # Отсортируйте данные по значению в столбце с индексом 2
        sorted_table_data = sorted(table_data, key=lambda x: float(x[2].replace(',', '').rstrip('%')) if x[2] else 0)

        # Сгенерируйте HTML из отсортированных данных
        for i, row_data in enumerate(sorted_table_data):
            write_off_banana_html += "<tr>"
            for j, item in enumerate(row_data):
                try:
                    value = float(item.replace(',', '').rstrip('%'))
                    if j == 2 and j > 0:  # Если это сортируемый столбец и не первый
                        left_value = float(sorted_table_data[i][j - 1].replace(',', '').rstrip('%'))
                        if value < left_value:
                            cell_style = "background-color: rgb(217, 242, 208);"
                        else:
                            cell_style = ""
                    else:
                        cell_style = ""
                except ValueError:
                    cell_style = ""
                write_off_banana_html += f"<td style='border: 1px solid black; font-family: Helvetica, sans-serif; {cell_style}'>{item}</td>"
            write_off_banana_html += "</tr>"
        write_off_banana_html += "</table>"

        heavy_greens_html = "<table cellpadding=3 style='border-collapse: collapse;'>"

        # Формирование заголовков таблицы
        heavy_greens_html += "<tr>"
        for column in range(self.heavy_greens_table.columnCount()):
            header_item = self.heavy_greens_table.horizontalHeaderItem(column)
            if header_item is not None:
                heavy_greens_html += f"<th style='border: 1px solid black; background-color: #538135; color: white; font-family: Helvetica, sans-serif; font-weight: bold; text-align: center;'>{header_item.text()}</th>"
        heavy_greens_html += "</tr>"

        # Формирование строк таблицы
        for row in range(self.heavy_greens_table.rowCount()):
            is_empty_row = False  # Флаг для обнаружения пустой строки
            row_html = "<tr>"
            for column in range(self.heavy_greens_table.columnCount()):
                item = self.heavy_greens_table.item(row, column)
                if item is not None and item.text().strip():  # Проверяем, не пустая ли ячейка и содержит ли она текст
                    row_html += f"<td style='border: 1px solid black; font-family: Helvetica, sans-serif;'>{item.text()}</td>"
                else:
                    is_empty_row = True  # Если хотя бы одна ячейка пустая, помечаем строку как пустую
                    break  # Прерываем цикл, чтобы не добавлять пустую строку в таблицу
            if not is_empty_row:
                heavy_greens_html += row_html + "</tr>"
        heavy_greens_html += "</table>"

        plan_table_html = "<table cellpadding=3 style='border-collapse: collapse;'>"

        # Формирование заголовков
        plan_table_html += "<tr>"
        for column in range(self.plan_table.columnCount()):
            header_item = self.plan_table.horizontalHeaderItem(column)
            if header_item is not None:
                header_text = header_item.text()
                if column == 0:
                    plan_table_html += f"<th style='border: 1px solid black; background-color: #7030A0; color: white; font-family: Helvetica, sans-serif; text-align: center; width: 221.15pt;'>{header_text}</th>"
                else:
                    plan_table_html += f"<th style='border: 1px solid black; background-color: #7030A0; color: white; font-family: Helvetica, sans-serif; text-align: center;'>{header_text}</th>"
            else:
                plan_table_html += "<th style='border: 1px solid black; background-color: #7030A0; color: white; font-family: Helvetica, sans-serif; text-align: center;'></th>"
        plan_table_html += "</tr>"

        # Формирование содержимого таблицы
        for row in range(self.plan_table.rowCount()):
            plan_table_html += "<tr>"
            for column in range(self.plan_table.columnCount()):
                item = self.plan_table.item(row, column)
                if item is not None:
                    plan_table_html += f"<td style='border: 1px solid black; font-family: Helvetica, sans-serif;'>{item.text()}</td>"
                else:
                    plan_table_html += "<td style='border: 1px solid black; font-family: Helvetica, sans-serif;'></td>"
            plan_table_html += "</tr>"
        plan_table_html += "</table>"
        
        current_dir = os.path.dirname(os.path.realpath(__file__))
        
        relative_path = "image001.png"
        image_path = os.path.join(current_dir, relative_path)
        relative_path = "image002.png"
        image_before_general_comments_path = os.path.join(current_dir, relative_path)
        relative_path = "image003.png"
        image_write_off_path = os.path.join(current_dir, relative_path)
        relative_path = "image004.png"
        image_write_off_pic_path = os.path.join(current_dir, relative_path)
        relative_path = "image005.png"
        image_write_off_banana_path = os.path.join(current_dir, relative_path)
        relative_path = "image006.png"
        image_write_off_banana_pic_path = os.path.join(current_dir, relative_path)
        relative_path = "image007.png"
        image_heavy_greens_pic_path = os.path.join(current_dir, relative_path)
        relative_path = "image008.png"
        image_razdelitel_path = os.path.join(current_dir, relative_path)
        relative_path = "image009.png"
        image_plan_path = os.path.join(current_dir, relative_path)
        relative_path = "image010.png"
        image_last_path = os.path.join(current_dir, relative_path)

        # Проверяем, существуют ли файлы по указанным путям
        if os.path.exists(image_path):
            self.image_html = f"<img src='file://{image_path}' width='632' align=center/>"
        else:
            self.image_html = "<b>ERROR!! Что-то пошло не так. Проверь наличие картинки с названием image001.png</b>"

        if os.path.exists(image_before_general_comments_path):
            self.image_before_general_comments_html = f"<img src='file://{image_before_general_comments_path}' width='25' align=left hspace='5'/>"
        else:
            self.image_before_general_comments_html = "<b>ERROR!! Что-то пошло не так. Проверь наличие картинки с названием image002.png</b>"

        if os.path.exists(image_write_off_path):
            self.image_write_off_html = f"<img src='file://{image_write_off_path}' width='25' align=left hspace='3' vspace='5'/>"
        else:
            self.image_write_off_html = "<b>ERROR!! Что-то пошло не так. Проверь наличие картинки с названием image003.png</b>"

        if os.path.exists(image_write_off_pic_path):
            self.image_write_off_pic_html = f"<img src='file://{image_write_off_pic_path}' width='25' align=left hspace='3' vspace='5'/>"
        else:
            self.image_write_off_pic_html = "<b>ERROR!! Что-то пошло не так. Проверь наличие картинки с названием image003.png</b>"

        email_body = f"""
        
        <html xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:w="urn:schemas-microsoft-com:office:word"
        xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
        xmlns="http://www.w3.org/TR/REC-html40">

        <body bgcolor="#538135" lang=ru-UA link="#0563C1" vlink="#954F72"
        style='tab-interval:36.0pt;word-wrap:break-word;'>
        <p align=center>{self.image_html}</p>
        <div align=center>
        <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=5 width=825
        style='font-family:Helvetica;color:black;width:667.45pt;border-collapse:collapse;mso-yfti-tbllook:1184;
        mso-padding-alt:0cm 0cm 0cm 0cm'>
        <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:19.5pt'>
            <td width=825 valign=top style='width:667.45pt;background:white;padding:
            0cm 5.4pt 0cm 5.4pt;height:19.5pt'>
            <p class=MsoNormal align=center style='text-align:center'><span
            style='font-family:"Helvetica",sans-serif;color:black'><b>{greeting_review}</span></p></b>
            <p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
            style='font-family:Helvetica;color:black;mso-no-proof:yes'><!--[if gte vml 1]><v:shape
            id="Рисунок_x0020_8" o:spid="_x0000_i1032" type="#_x0000_t75" style='width:23.25pt;
            height:23.25pt;visibility:visible;mso-wrap-style:square'>
            <v:imagedata src="{image_before_general_comments_path}" o:title=""/>
            </v:shape><![endif]--><![if !vml]><img width=31 height=31
            src="{image_before_general_comments_path}" v:shapes="Рисунок_x0020_8"><![endif]></span></b><span
            style='font-family:"Helvetica",sans-serif;color:black'>{general_comments}</span>
            {general_html}</div>
            <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto font-family:"Helvetica",sans-serif;color:black'><!--[if gte vml 1]><v:shape
            id="Рисунок_x0020_10" o:spid="_x0000_s1026" type="#_x0000_t75" style='position:absolute;
            margin-left:0;margin-top:0;width:18.75pt;height:18.75pt;z-index:251659264;
            visibility:visible;mso-wrap-style:square;mso-width-percent:0;
            mso-height-percent:0;mso-wrap-distance-left:2.25pt;mso-wrap-distance-top:3.75pt;
            mso-wrap-distance-right:2.25pt;mso-wrap-distance-bottom:3.75pt;
            mso-position-horizontal:left;mso-position-horizontal-relative:text;
            mso-position-vertical:absolute;mso-position-vertical-relative:line;
            mso-width-percent:0;mso-height-percent:0;mso-width-relative:page;
            mso-height-relative:page' o:allowoverlap="f">
            <v:imagedata src="{image_write_off_path}" o:title=""/>
            <w:wrap type="square" anchory="line"/>
            </v:shape><![endif]--><![if !vml]><img width=25 height=25
            src="{image_write_off_path}" align=left hspace=3 vspace=5 v:shapes="Рисунок_x0020_10"><![endif]>{write_off_comments}<span style='color:black'><o:p>&nbsp;</o:p></span></b></p>
            <table class=MsoTableGrid border=none cellspacing=0 cellpadding=0
            style='border-collapse:collapse;border:none;mso-border-alt:none;
            mso-yfti-tbllook:1184;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
            <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>
                <td width=437 valign=top style='width:368.05pt;border:none;
                mso-border-alt:none;padding:0cm 5.4pt 0cm 5.4pt'>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:
                auto'><span lang=RU style='font-family:"Helvetica",sans-serif;color:black;
                mso-ansi-language:RU'>{write_off_html}<o:p></o:p></span></p>
                </td>
                <td width=437 style='width:288.1pt;border:none;
                border-left:none;mso-border-left-alt:none;mso-border-alt:
                none .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
                <p class=MsoNormal align=center style='mso-margin-top-alt:auto;mso-margin-bottom-alt:
                auto;text-align:center'><span lang=RU style='font-family:"Helvetica",sans-serif;
                color:black;mso-ansi-language:RU;mso-no-proof:yes'><!--[if gte vml 1]><v:shapetype
                id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t"
                path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
                <v:stroke joinstyle="miter"/>
                <v:formulas>
                <v:f eqn="if lineDrawn pixelLineWidth 0"/>
                <v:f eqn="sum @0 1 0"/>
                <v:f eqn="sum 0 0 @1"/>
                <v:f eqn="prod @2 1 2"/>
                <v:f eqn="prod @3 21600 pixelWidth"/>
                <v:f eqn="prod @3 21600 pixelHeight"/>
                <v:f eqn="sum @0 0 1"/>
                <v:f eqn="prod @6 1 2"/>
                <v:f eqn="prod @7 21600 pixelWidth"/>
                <v:f eqn="sum @8 21600 0"/>
                <v:f eqn="prod @7 21600 pixelHeight"/>
                <v:f eqn="sum @10 21600 0"/>
                </v:formulas>
                <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
                <o:lock v:ext="edit" aspectratio="t"/>
                </v:shapetype><v:shape id="Рисунок_x0020_1" o:spid="_x0000_i1025" type="#_x0000_t75"
                alt="Изображение"
                style='width:185.25pt;height:185.25pt;visibility:visible;mso-wrap-style:square'>
                <v:imagedata src="{image_write_off_pic_path}" o:title="Изображение"/>
                </v:shape><![endif]--><![if !vml]><img width=247 height=247
                src="{image_write_off_pic_path}"
                alt="Изображение"
                v:shapes="Рисунок_x0020_1"><![endif]></span><span lang=RU style='font-family:
                "Helvetica",sans-serif;color:black;mso-ansi-language:RU'><o:p></o:p></span></p>
                </td>
            </tr>
            </table>
            <p class=MsoNormal><span style='font-family:"Helvetica",sans-serif;
            color:black'><!--[if gte vml 1]><v:shape id="Рисунок_x0020_312" o:spid="_x0000_i1029"
            type="#_x0000_t75" alt="" style='width:24pt;height:24pt'>
            <v:imagedata src="{image_write_off_banana_path}"
            o:href="cid:image005.png@01DAA168.34D7A770"/>
            </v:shape><![endif]--><![if !vml]><img width=32 height=32
            src="{image_write_off_banana_path}"
            style='height:.333in;width:.333in' v:shapes="Рисунок_x0020_312"><![endif]>{write_off_banana}</span></p>
              <table class=MsoTableGrid border=0 cellspacing=0 cellpadding=0
            style='border-collapse:collapse;border:none;mso-yfti-tbllook:1184;
            mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:none;mso-border-insidev:
            none'>
            <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>
                <td width=267 style='width:199.9pt;padding:0cm 5.4pt 0cm 5.4pt'>
                <p class=MsoNormal align=center style='mso-margin-top-alt:auto;mso-margin-bottom-alt:
                auto;text-align:center'><span lang=RU style='font-family:"Helvetica",sans-serif;
                color:black;mso-ansi-language:RU;mso-no-proof:yes'><!--[if gte vml 1]><v:shapetype
                id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t"
                path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
                <v:stroke joinstyle="miter"/>
                <v:formulas>
                <v:f eqn="if lineDrawn pixelLineWidth 0"/>
                <v:f eqn="sum @0 1 0"/>
                <v:f eqn="sum 0 0 @1"/>
                <v:f eqn="prod @2 1 2"/>
                <v:f eqn="prod @3 21600 pixelWidth"/>
                <v:f eqn="prod @3 21600 pixelHeight"/>
                <v:f eqn="sum @0 0 1"/>
                <v:f eqn="prod @6 1 2"/>
                <v:f eqn="prod @7 21600 pixelWidth"/>
                <v:f eqn="sum @8 21600 0"/>
                <v:f eqn="prod @7 21600 pixelHeight"/>
                <v:f eqn="sum @10 21600 0"/>
                </v:formulas>
                <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
                <o:lock v:ext="edit" aspectratio="t"/>
                </v:shapetype><v:shape id="Рисунок_x0020_1" o:spid="_x0000_i1025" type="#_x0000_t75"
                alt="Изображение"
                style='width:185.25pt;height:185.25pt;visibility:visible;mso-wrap-style:square'>
                <v:imagedata src="{image_write_off_banana_pic_path}" o:title="Изображение"/>
                </v:shape><![endif]--><![if !vml]><img width=247 height=247
                src="{image_write_off_banana_pic_path}"
                alt="Изображение"
                v:shapes="Рисунок_x0020_1"><![endif]></span><span lang=RU style='font-family:
                "Helvetica",sans-serif;color:black;mso-ansi-language:RU'><o:p></o:p></span></p>
                </td>
                <td width=608 valign=top style='width:455.95pt;padding:0cm 5.4pt 0cm 5.4pt'>
                <table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
                style='border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;
                mso-yfti-tbllook:1184;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
                <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
                <td width=295 rowspan=2 style='width:221.2pt;border:solid windowtext 1.0pt;
                mso-border-alt:solid windowtext .5pt;background:#FFC000;padding:0cm 5.4pt 0cm 5.4pt'>
                <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
                mso-margin-bottom-alt:auto;text-align:center'><span class=SpellE><b><span
                lang=RU style='font-family:"Helvetica",sans-serif;color:black;mso-ansi-language:
                RU'>Філіал</span></b></span><b><span lang=RU style='font-family:"Helvetica",sans-serif;
                color:black;mso-ansi-language:RU'><o:p></o:p></span></b></p>
                </td>
                <td width=217 colspan=3 style='width:162.55pt;border:solid windowtext 1.0pt;
                border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
                solid windowtext .5pt;background:#FFC000;padding:0cm 5.4pt 0cm 5.4pt'>
                <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
                mso-margin-bottom-alt:auto;text-align:center'><span class=SpellE><b><span
                lang=RU style='font-family:"Helvetica",sans-serif;color:black;mso-ansi-language:
                RU'>Списання</span></b></span><b><span lang=RU style='font-family:"Helvetica",sans-serif;
                color:black;mso-ansi-language:RU'><o:p></o:p></span></b></p>
                </td>
                <td width=81 rowspan=2 style='width:60.6pt;border:solid windowtext 1.0pt;
                border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
                solid windowtext .5pt;background:#FFC000;padding:0cm 5.4pt 0cm 5.4pt'>
                <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
                mso-margin-bottom-alt:auto;text-align:center'><span class=SpellE><b><span
                lang=RU style='font-family:"Helvetica",sans-serif;color:black;mso-ansi-language:
                RU'>Продажі</span></b></span><b><span lang=RU style='font-family:"Helvetica",sans-serif;
                color:black;mso-ansi-language:RU'>, кг<o:p></o:p></span></b></p>
                </td>
                </tr>
                <tr style='mso-yfti-irow:1'>
                <td width=85 style='width:63.8pt;border-top:none;border-left:none;
                border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
                mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
                mso-border-alt:solid windowtext .5pt;background:#FFC000;padding:0cm 5.4pt 0cm 5.4pt'>
                <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
                mso-margin-bottom-alt:auto;text-align:center'><b><span lang=RU
                style='font-family:"Helvetica",sans-serif;color:black;mso-ansi-language:
                RU'>План<o:p></o:p></span></b></p>
                </td>
                <td width=66 style='width:49.6pt;border-top:none;border-left:none;
                border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
                mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
                mso-border-alt:solid windowtext .5pt;background:#FFC000;padding:0cm 5.4pt 0cm 5.4pt'>
                <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
                mso-margin-bottom-alt:auto;text-align:center'><b><span lang=RU
                style='font-family:"Helvetica",sans-serif;color:black;mso-ansi-language:
                RU'>Факт<o:p></o:p></span></b></p>
                </td>
                <td width=66 style='width:49.15pt;border-top:none;border-left:none;
                border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
                mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
                mso-border-alt:solid windowtext .5pt;background:#FFC000;padding:0cm 5.4pt 0cm 5.4pt'>
                <p class=MsoNormal align=center style='mso-margin-top-alt:auto;
                mso-margin-bottom-alt:auto;text-align:center'><b><span lang=RU
                style='font-family:"Helvetica",sans-serif;color:black;mso-ansi-language:
                RU'>в кг<o:p></o:p></span></b></p>
                </td>
                </tr>
            {write_off_banana_html}
            </table>
            <p class=MsoNormal><b><span style='font-family:"Helvetica",sans-serif;
            color:black'><!--[if gte vml 1]><v:shape id="Рисунок_x0020_1" o:spid="_x0000_i1031"
            type="#_x0000_t75" alt="" style='width:24.75pt;height:24.75pt'>
            <v:imagedata src="{image_heavy_greens_pic_path}"
            o:href="cid:image007.png@01DAA218.EA767EC0"/>
            </v:shape><![endif]--><![if !vml]><img width=33 height=33
            src="{image_heavy_greens_pic_path}"
            style='height:.343in;width:.500in' v:shapes="Рисунок_x0020_1"><![endif]></span></b>{heavy_greens_сomments}{heavy_greens_html}</span></p>

            <p class=MsoNormal><span style='font-family:"Helvetica",sans-serif;
            color:black'><!--[if gte vml 1]><v:shape id="Рисунок_x0020_315" o:spid="_x0000_i1032"
            type="#_x0000_t75" alt="" style='width:650.25pt;height:36pt'>
            <v:imagedata src="{image_razdelitel_path}"
            o:href="cid:image007.png@01DAA168.34D7A770"/>
            </v:shape><![endif]--><![if !vml]><img width=751 height=48
            src="{image_razdelitel_path}"
            style='height:.5in;width:7.822in' v:shapes="Рисунок_x0020_315"><![endif]></span><span
            style='font-family:"Helvetica",sans-serif'><o:p></o:p></span></p>
            <p class=MsoNormal><span style='font-family:"Helvetica",sans-serif;
            color:black'><!--[if gte vml 1]><v:shape id="Рисунок_x0020_316" o:spid="_x0000_i1033"
            type="#_x0000_t75" alt="" style='width:23.25pt;height:23.25pt'>
            <v:imagedata src="{image_plan_path}"
            o:href="cid:image009.png@01DAA168.34D7A770"/>
            </v:shape><![endif]--><![if !vml]><img width=31 height=31
            src="{image_plan_path}"
            style='height:.322in;width:.322in' v:shapes="Рисунок_x0020_316"><![endif]>&nbsp;<b>Планові
            показники</b> на поточний тиждень:</span><span style='font-family:"Helvetica",sans-serif'><o:p></o:p></span></p>
            {plan_table_html}
            <p class=MsoNormal><span style='font-family:"Helvetica",sans-serif'><o:p>&nbsp;</o:p></span></p>
            <p class=MsoNormal><b><span style='font-family:"Helvetica",sans-serif;
            color:black'>Всім дякую за роботу та приділену увагу дільниці. <o:p></o:p></span></b></p>
            <p class=MsoNormal><b><span style='font-family:"Helvetica",sans-serif;
            color:black'>Гарних продажів і ще кращих показників!<o:p></o:p></span></b></p>
            <p class=MsoNormal><b><span style='font-family:"Helvetica",sans-serif;
            color:black'>Пишіть/дзвоніть!</span></b><b><span style='font-family:"Helvetica",sans-serif'><o:p></o:p></span></b></p>
            <p class=MsoNormal align=center style='text-align:center'><span
            style='font-family:"Helvetica",sans-serif;color:black'><!--[if gte vml 1]><v:shape
            id="Рисунок_x0020_317" o:spid="_x0000_i1034" type="#_x0000_t75" alt=""
            style='width:426pt;height:226.5pt'>
            <v:imagedata src="{image_last_path}"
            o:href="cid:image009.png@01DAA168.34D7A770"/>
            </v:shape><![endif]--><![if !vml]><img width=568 height=302
            src="{image_last_path}"
            style='height:3.145in;width:5.916in' v:shapes="Рисунок_x0020_317"><![endif]></span><span
            style='font-family:"Helvetica",sans-serif'><o:p></o:p></span></p>
            </td>
        </tr>
        <tr style='mso-yfti-irow:1;mso-yfti-lastrow:yes;height:19.5pt'>
            <td width=825 valign=top style='width:618.45pt;background:white;padding:
            0cm 5.4pt 0cm 5.4pt;height:19.5pt'>
            <p class=MsoNormal align=center style='text-align:center'><b><span
            style='font-family:"Helvetica",sans-serif;color:black'><o:p>&nbsp;</o:p></span></b></p>
            </td>
        </tr>
        </table>
        </td>
        </tr>
        </table>

        </div>
        </body>

        </html>

        """
        selected_week = self.week_combobox.currentText()

        kust = self.kust_combobox.currentText()
        emails = self.filials_data.get(kust, {}).get('email', [])
        emails_cc = self.filials_data.get(kust, {}).get('emailсс', [])

        email_str = ';'.join(emails)
        email_cc_str = ';'.join(emails_cc)

        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = f'Результати роботи овочевої дільниці за {selected_week} тиждень'
        mail.To = email_str
        mail.CC = email_cc_str  # Укажите адреса электронной email, разделенные точкой с запятой
        mail.HTMLBody = email_body
        mail.Display()
    
    def randomize_empty_cells(self):
        # Для self.general_table
        for row in range(self.general_table.rowCount()):
            for col in range(self.general_table.columnCount()):
                if col in [1, 2, 4, 5, 8] and not self.general_table.item(row, col):
                    if col in [1, 2, 4, 5]:
                        value = round(random.uniform(21, 46), 2) if col in [1, 2] else round(random.uniform(100000, 2000000), 2)
                    else:
                        value = random.randint(90, 99)
                    item = QTableWidgetItem(str(value))
                    self.general_table.setItem(row, col, item)

         # Для self.write_off_table
        for row in range(self.write_off_table.rowCount()):
            for col in range(self.write_off_table.columnCount()):
                if col in [1, 2] and not self.write_off_table.item(row, col):
                    value = random.randint(1, 15)
                    item = QTableWidgetItem(str(value))
                    self.write_off_table.setItem(row, col, item)

        # Для self.write_off_banana_table
        for row in range(self.write_off_banana_table.rowCount()):
            for col in range(self.write_off_banana_table.columnCount()):
                if col in [1, 2, 3, 4] and not self.write_off_banana_table.item(row, col):
                    if col == 1:
                        value = round(random.uniform(0.5, 1.5), 2)
                    elif col == 2:
                        value = round(random.uniform(0.5, 2), 2)
                    elif col == 3:
                        value = random.randint(1, 100)
                    else:
                        value = random.randint(500, 3000)
                    item = QTableWidgetItem(str(value))
                    self.write_off_banana_table.setItem(row, col, item)

        # Для self.heavy_greens_table
        for row in range(self.heavy_greens_table.rowCount()):
            for col in range(self.heavy_greens_table.columnCount()):
                if col in [1, 2, 4] and not self.heavy_greens_table.item(row, col):
                    if col in [1, 2]:
                        value = random.randint(5, 100)
                    else:
                        value = random.randint(2, 20)
                    item = QTableWidgetItem(str(value))
                    self.heavy_greens_table.setItem(row, col, item)

        # Для self.plan_table
        for row in range(self.plan_table.rowCount()):
            for col in range(self.plan_table.columnCount()):
                if col in [1, 2, 4] and not self.plan_table.item(row, col):
                    if col == 1:
                        value = round(random.uniform(21, 46), 2)
                    elif col == 2:
                        value = round(random.uniform(100000, 2000000), 2)
                    else:
                        value = random.randint(3, 10)
                    item = QTableWidgetItem(str(value))
                    self.plan_table.setItem(row, col, item)

        self.greeting_review_textedit.setText("Привіт! Подивимось на результати роботи і показники овочевої дільниці за 18 тиждень.")
        self.general_comments_textedit.setText("Загальні показники такі: пенетрація виконана тільки у Чарівної. Виторг у Новокримської, Слави, Тополиної та Чарівної. ДП виконана у всіх крім Вокзальної, Гагаріна та Незалежності.")
        self.write_off_comments_textedit.setText("Списання виконано тільки у Тополиної, а Чарівна знаходиться близько до цілі. На минулому тижні майже ніхто не оформлював рекламації на якість, але списували неймовірні кількості товару. Всім треба звернути увагу на приймання товару по якості та оформлення рекламацій, а також оформлення рекламацій якщо товар не витримує в ТЗ.")
        self.write_off_banana_textedit.setText("Списання банану не виконано у Тополиної, Незалежності і Слави. Тополину і Незалежності вкотре прошу звернути увагу на роботу з бананом та рекламаціями. Слави мали непідтверджену рекламацію.")
        self.heavy_greens_textedit.setText("Додаю таблицю з продажами та списаннями по ваговій зелені. Слави мають найбільшу долю списання, а Новокримська найменшу.")

    def remove_data(self):
        self.general_table.setRowCount(0)
        self.write_off_table.setRowCount(0)
        self.write_off_banana_table.setRowCount(0)
        self.heavy_greens_table.setRowCount(0)
        self.plan_table.setRowCount(0)

        self.greeting_review_textedit.clear()
        self.general_comments_textedit.clear()
        self.write_off_comments_textedit.clear()
        self.write_off_banana_textedit.clear()
        self.heavy_greens_textedit.clear()
        self.update_tables(self.kust_combobox.currentIndex())
        
class NumericDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = super().createEditor(parent, option, index)
        if isinstance(editor, QLineEdit):
            # Регулярное выражение для ввода чисел, включая целые и дробные с использованием запятой
            validator = QRegExpValidator(QRegExp("[0-9,]*"), editor)
            editor.setValidator(validator)
        return editor
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DataEntryWindow()
    window.show()
    sys.exit(app.exec_())
