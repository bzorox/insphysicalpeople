import os
import sys
import pandas as pd
import re
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, 
                            QPushButton, QFileDialog, QProgressBar, QMessageBox, 
                            QHBoxLayout, QFrame, QSizePolicy, QCheckBox)
from PyQt5.QtCore import Qt, QFile, QTextStream, QPropertyAnimation, QEasingCurve, QThread, pyqtSignal
from PyQt5.QtGui import QIcon, QPixmap, QFont
import folium
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable
import time

class GeocoderThread(QThread):
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(list)

    def __init__(self, addresses):
        super().__init__()
        self.addresses = addresses

    def run(self):
        geolocator = Nominatim(user_agent="insurance_app")
        geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, max_retries=2)
        
        results = []
        total = len(self.addresses)
        
        for i, address in enumerate(self.addresses, 1):
            try:
                location = geocode(address, timeout=10)
                if location:
                    results.append((location.latitude, location.longitude))
                else:
                    results.append((None, None))
            except (GeocoderTimedOut, GeocoderUnavailable) as e:
                results.append((None, None))
                time.sleep(2)
            
            progress = int((i / total) * 100)
            self.progress_signal.emit(progress)
        
        self.finished_signal.emit(results)

class StyledButton(QPushButton):
    def __init__(self, text):
        super().__init__(text)
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("styledButton")
        self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        self._original_width = self.sizeHint().width()

    def enterEvent(self, event):
        animation = QPropertyAnimation(self, b"minimumWidth")
        animation.setDuration(150)
        animation.setStartValue(self.width())
        animation.setEndValue(self._original_width + 10)
        animation.setEasingCurve(QEasingCurve.OutQuad)
        animation.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        animation = QPropertyAnimation(self, b"minimumWidth")
        animation.setDuration(150)
        animation.setStartValue(self.width())
        animation.setEndValue(self._original_width)
        animation.setEasingCurve(QEasingCurve.OutQuad)
        animation.start()
        super().leaveEvent(event)

class InsuranceDataProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Обработка страховых данных")
        self.setMinimumSize(800, 700)
        
        if hasattr(sys, '_MEIPASS'):
            self.setWindowIcon(QIcon(os.path.join(sys._MEIPASS, 'assets', 'logo.png')))
        else:
            self.setWindowIcon(QIcon('assets/logo.png'))
        
        self.initUI()
        self.file_path = None
        self.save_path = None
        self.geocoder_thread = None
        
    def initUI(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        
        self.load_styles()
        
        # Header with logo and title
        header = QHBoxLayout()
        logo = QLabel()
        if hasattr(sys, '_MEIPASS'):
            pixmap = QPixmap(os.path.join(sys._MEIPASS, 'assets', 'logo.png')))
        else:
            pixmap = QPixmap('assets/logo.png')
        logo.setPixmap(pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        header.addWidget(logo)
        
        title = QLabel("Обработка страховых данных")
        title.setObjectName("titleLabel")
        header.addWidget(title)
        header.addStretch()
        main_layout.addLayout(header)
        
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setObjectName("separator")
        main_layout.addWidget(separator)
        
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(20, 15, 20, 15)
        content_layout.setSpacing(15)
        
        # File selection
        file_frame = self.create_section_frame("Файл Excel:", "Не выбран", "Выбрать файл", self.browse_file)
        self.file_path_label = file_frame.findChild(QLabel, "valueLabel")
        content_layout.addWidget(file_frame)
        
        # Save folder selection
        save_frame = self.create_section_frame("Папка для сохранения:", "Не выбрана", "Выбрать папку", self.choose_save_path)
        self.save_path_label = save_frame.findChild(QLabel, "valueLabel")
        content_layout.addWidget(save_frame)
        
        # Options checkboxes
        options_frame = QFrame()
        options_frame.setObjectName("optionsFrame")
        options_layout = QVBoxLayout(options_frame)
        options_layout.setContentsMargins(15, 15, 15, 15)
        
        self.create_map_check = QCheckBox("Создать интерактивную карту")
        self.create_map_check.setChecked(True)
        self.create_map_check.stateChanged.connect(self.toggle_geocode)
        options_layout.addWidget(self.create_map_check)
        
        self.geocode_check = QCheckBox("Геокодировать адреса (Nominatim)")
        self.geocode_check.setChecked(True)
        options_layout.addWidget(self.geocode_check)
        
        content_layout.addWidget(options_frame)
        
        # Progress indicators
        self.progress_label = QLabel("Готов к работе")
        self.progress_label.setObjectName("progressLabel")
        content_layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        content_layout.addWidget(self.progress_bar)
        
        # Process button
        self.process_btn = StyledButton("Обработать данные")
        self.process_btn.setObjectName("processButton")
        self.process_btn.clicked.connect(self.process_data)
        self.process_btn.setEnabled(False)
        content_layout.addWidget(self.process_btn, stretch=1)
        
        main_layout.addLayout(content_layout)
        
        # Footer
        footer = QLabel("Разработано: стажёр Мальцев Максим")
        footer.setObjectName("footerLabel")
        footer.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(footer)
    
    def toggle_geocode(self):
        self.geocode_check.setEnabled(self.create_map_check.isChecked())
    
    def create_section_frame(self, title, default_value, button_text, callback):
        frame = QFrame()
        frame.setObjectName("sectionFrame")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)
        
        title_label = QLabel(title)
        title_label.setObjectName("sectionTitle")
        layout.addWidget(title_label)
        
        value_label = QLabel(default_value)
        value_label.setObjectName("valueLabel")
        value_label.setWordWrap(True)
        layout.addWidget(value_label)
        
        btn = StyledButton(button_text)
        btn.clicked.connect(callback)
        layout.addWidget(btn)
        
        return frame
    
    def load_styles(self):
        style_file = QFile("style.qss")
        if style_file.open(QFile.ReadOnly | QFile.Text):
            stream = QTextStream(style_file)
            self.setStyleSheet(stream.readAll())
    
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel", "", "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.file_path = file_path
            self.file_path_label.setText(os.path.basename(file_path))
            self.file_path_label.setToolTip(file_path)
            self.check_ready()
    
    def choose_save_path(self):
        save_path = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения")
        
        if save_path:
            self.save_path = save_path
            self.save_path_label.setText(os.path.basename(save_path))
            self.save_path_label.setToolTip(save_path)
            self.check_ready()
    
    def check_ready(self):
        self.process_btn.setEnabled(bool(self.file_path and self.save_path))
    
    def normalize_address(self, address):
        """Улучшенная нормализация адресов"""
        if not isinstance(address, str):
            return address
            
        patterns = [
            # Квартиры и помещения
            r',\s*(?:кв\.?\s*\d+[а-яa-z]?(?:\/\d+[а-яa-z]?)?|квартир?а\s*\d+[а-яa-z]?|'
            r'пом\.?\s*\d+[а-яa-z]?|помещ(ение)?\.?\s*\d+[а-яa-z]?)(?:\s*,?\s*(?:ком\.?|комнат[аы]?)\s*\d+[а-яa-z]?)?\b',
            # Офисы
            r',\s*(?:оф\.?\s*\d+[а-яa-z]?|офис\s*\d+[а-яa-z]?)\b',
            # Литеры
            r',\s*литер[аы]?\s*[А-Яа-яA-Za-z]\b'
        ]
        
        for pattern in patterns:
            address = re.sub(pattern, '', address, flags=re.IGNORECASE)
            
        # Удаление двойных запятых и пробелов
        address = re.sub(r',\s*,', ',', address)
        address = re.sub(r'\s{2,}', ' ', address)
        return address.strip(' ,')
    
    def create_folium_map(self, df):
        """Создание интерактивной карты с кружками"""
        if df.empty or not all(col in df.columns for col in ['lat', 'lon']):
            return None
        
        # Фильтруем только адреса с координатами
        df = df.dropna(subset=['lat', 'lon'])
        if df.empty:
            return None
        
        # Создаем карту с центром на первом адресе
        first_loc = df.iloc[0]
        m = folium.Map(location=[first_loc['lat'], first_loc['lon']], zoom_start=12)
        
        # Определяем цвет круга в зависимости от суммы
        def get_color(total):
            if total > 7_000_000_000:
                return '#FF0000'  # Красный
            elif total > 6_000_000_000:
                return '#FF4500'  # Оранжево-красный
            elif total > 5_000_000_000:
                return '#FF8C00'  # Темно-оранжевый
            elif total > 2_000_000_000:
                return '#FFA500'  # Оранжевый
            elif total > 500_000_000:
                return '#FFD700'  # Золотой
            elif total > 300_000_000:
                return '#FFFF00'  # Желтый
            elif total > 100_000_000:
                return '#ADFF2F'  # Зелено-желтый
            else:
                return '#32CD32'  # Лаймовый
        
        # Добавляем круги на карту
        for _, row in df.iterrows():
            folium.CircleMarker(
                location=[row['lat'], row['lon']],
                radius=8 + (row['Сумма кумуляции'] / 100_000_000) ** 0.3,  # Динамический радиус
                popup=f"Адрес: {row['Адрес']}<br>Сумма: {row['Сумма кумуляции']:,.2f} руб.",
                color=get_color(row['Сумма кумуляции']),
                fill=True,
                fill_color=get_color(row['Сумма кумуляции'])
            ).add_to(m)
        
        return m
    
    def process_data(self):
        try:
            self.progress_label.setText("Прогресс: Чтение файла...")
            self.progress_bar.setValue(5)
            QApplication.processEvents()
            
            # Чтение файла Excel
            df = pd.read_excel(self.file_path)
            
            # Проверка необходимых столбцов
            required_columns = ['object', 'money', 'date_end', 'adress']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                raise ValueError(f"Отсутствуют необходимые столбцы: {', '.join(missing_columns)}")
            
            self.progress_label.setText("Прогресс: Обработка данных...")
            self.progress_bar.setValue(20)
            QApplication.processEvents()
            
            # Фильтрация строк
            mask1 = ~df['object'].str.contains(
                'Гражданская ответсвенность|Гражданская ответственность', 
                case=False, 
                na=False
            )
            mask2 = pd.to_datetime(df['date_end'], dayfirst=True) >= datetime(2025, 5, 31)
            filtered_df = df[mask1 & mask2].copy()
            
            # Нормализация адресов
            filtered_df['normalized_adress'] = filtered_df['adress'].apply(self.normalize_address)
            
            # Сохранение отфильтрованных данных
            filtered_data_path = os.path.join(self.save_path, 'filtered_data.xlsx')
            filtered_df.to_excel(filtered_data_path, index=False)
            
            self.progress_label.setText("Прогресс: Расчет кумуляции...")
            self.progress_bar.setValue(40)
            QApplication.processEvents()
            
            # Расчет кумуляции по нормализованным адресам
            result_df = filtered_df.groupby('normalized_adress')['money'].sum().reset_index()
            result_df.columns = ['Адрес', 'Сумма кумуляции']
            
            # Разделение по суммам
            thresholds = [
                ('>100M', 100_000_000),
                ('>300M', 300_000_000),
                ('>500M', 500_000_000),
                ('>2B', 2_000_000_000),
                ('>5B', 5_000_000_000),
                ('>6B', 6_000_000_000),
                ('>7B', 7_000_000_000)
            ]
            
            # Геокодирование, если нужно
            if self.create_map_check.isChecked() and self.geocode_check.isChecked():
                self.progress_label.setText("Прогресс: Геокодирование адресов...")
                self.progress_bar.setValue(50)
                QApplication.processEvents()
                
                self.geocoder_thread = GeocoderThread(result_df['Адрес'].tolist())
                self.geocoder_thread.progress_signal.connect(self.progress_bar.setValue)
                
                def on_geocoding_finished(coords):
                    result_df['lat'] = [c[0] for c in coords]
                    result_df['lon'] = [c[1] for c in coords]
                    self.finalize_processing(result_df, thresholds)
                
                self.geocoder_thread.finished_signal.connect(on_geocoding_finished)
                self.geocoder_thread.start()
                return
            else:
                self.finalize_processing(result_df, thresholds)
                
        except Exception as e:
            self.progress_label.setText("Прогресс: Ошибка")
            self.progress_bar.setValue(0)
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка:\n{str(e)}")
    
    def finalize_processing(self, result_df, thresholds):
        try:
            self.progress_label.setText("Прогресс: Сохранение результатов...")
            self.progress_bar.setValue(80)
            QApplication.processEvents()
            
            # Сохранение результатов
            output_path = os.path.join(self.save_path, 'результаты_обработки.xlsx')
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                result_df.to_excel(writer, sheet_name='Все адреса', index=False)
                
                for name, threshold in thresholds:
                    sheet_df = result_df[result_df['Сумма кумуляции'] > threshold]
                    sheet_df.to_excel(writer, sheet_name=name, index=False)
            
            # Создание карты, если нужно
            if self.create_map_check.isChecked():
                self.progress_label.setText("Прогресс: Генерация карты...")
                self.progress_bar.setValue(90)
                QApplication.processEvents()
                
                map_html_path = os.path.join(self.save_path, 'карта_кумуляции.html')
                folium_map = self.create_folium_map(result_df)
                
                if folium_map:
                    folium_map.save(map_html_path)
                    self.progress_label.setText(f"Готово! Карта сохранена в {map_html_path}")
                else:
                    self.progress_label.setText("Готово! Не удалось создать карту (нет координат)")
            
            self.progress_bar.setValue(100)
            
            QMessageBox.information(
                self, 
                "Успех", 
                f"Обработка завершена успешно!\n\n"
                f"Результаты сохранены в:\n{output_path}"
            )
            
        except Exception as e:
            self.progress_label.setText("Прогресс: Ошибка")
            self.progress_bar.setValue(0)
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при сохранении:\n{str(e)}")
        
        finally:
            self.progress_bar.setValue(0)
            self.progress_label.setText("Готов к новой обработке")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    font = QFont("Segoe UI", 12)
    app.setFont(font)
    window = InsuranceDataProcessor()
    window.show()
    sys.exit(app.exec_())
