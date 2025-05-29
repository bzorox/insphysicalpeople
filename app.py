import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import logging
from datetime import datetime
from PIL import Image, ImageTk
import webbrowser
import tempfile
import folium
from folium.plugins import MarkerCluster
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable
import time

# Configure logging
logging.basicConfig(filename='insurance_app.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def resource_path(relative_path):
    """Get absolute path to resource for PyInstaller and development"""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)

# Compile regex patterns once at module level
PATTERN = r"""
    \b(?:[гГ][рР]?[аА]?[жЖ]?[дД]?[аА]?[нН]?[сС]?[кК]?[аАяЯ]?[яЯ]?\s*)?
    (?:ответ[сстССТ]?[тТ]?[вВ]?[еЕ]?[нН]?[нН]?[оО]?[сС]?[тТ]?[ьЬ]?|ГО|г\.о\.)\b
    |\b(?:страхование\s*)?(?:гражданской|гр[ао]жданской|гржданской)\s*(?:ответственности|ответсвенности|ответсвенноти|ответственноти)\b
    |\b(?:ответственность|ответсвенность|ответсвенноть|ответственноть)\b
    |\bГО\b|\bг\.о\.\b
"""
REGEX = re.compile(PATTERN, flags=re.IGNORECASE | re.VERBOSE)

ADDRESS_CLEANING_PATTERN = r"""
    (?:,\s*|\s+)(?:кв\.?\s*\d+[а-яА-Я]?\b|квартира\s*\d+[а-яА-Я]?\b|кв\.?\s*№\s*\d+[а-яА-Я]?\b|
    квартира\s*№\s*\d+[а-яА-Я]?\b|оф\s*\d+[а-яА-Я]?\b|оф\.\s*\d+[а-яА-Я]?\b|офис\s*\d+[а-яА-Я]?\b|
    офис\s*№\s*\d+[а-яА-Я]?\b|оф\.\s*№\s*\d+[а-яА-Я]?\b|оф\s*№\s*\d+[а-яА-Я]?\b|
    пом\.\s*\d+[а-яА-Я-]?\b|помещ\.\s*\d+[а-яА-Я-]?\b|помещение\s*\d+[а-яА-Я]?\b)
"""
ADDRESS_REGEX = re.compile(ADDRESS_CLEANING_PATTERN, flags=re.IGNORECASE | re.VERBOSE)

def clean_text(text):
    """Очистка полей object от ФИО"""
    text = str(text)
    text = re.sub(r"\b[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\b|\b[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]\b|\b[А-ЯЁ][а-яё]+-\b|\b[А-ЯЁ][а-яё]+s\b|\b[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s\b", '', text)
    text = re.sub(r",\d{2}-\d{2}-\d{2},", '', text)
    return text.strip()

def clean_address(address):
    """Очистка адресов от номеров квартир, офисов и помещений"""
    if pd.isna(address):
        return address
    address = str(address)
    # Удаляем указания на квартиры, офисы и помещения
    address = ADDRESS_REGEX.sub('', address)
    # Удаляем возможные двойные запятые и пробелы
    address = re.sub(r',\s*,', ',', address)
    address = re.sub(r'\s{2,}', ' ', address)
    return address.strip(' ,')

def geocode_addresses(address_series, progress_callback=None):
    """Геокодирование адресов с помощью Nominatim"""
    geolocator = Nominatim(user_agent="insurance_app")
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, max_retries=2)
    
    results = []
    total = len(address_series)
    
    for i, address in enumerate(address_series, 1):
        try:
            location = geocode(address, timeout=10)
            if location:
                results.append((location.latitude, location.longitude))
            else:
                results.append((None, None))
        except (GeocoderTimedOut, GeocoderUnavailable) as e:
            logging.warning(f"Geocoding error for {address}: {str(e)}")
            results.append((None, None))
            time.sleep(2)  # Подождать перед следующей попыткой
        
        if progress_callback:
            progress = int((i / total) * 100)
            progress_callback(progress)
    
    return results

def create_map(dataframe, output_folder):
    """Создание интерактивной карты с точками адресов"""
    try:
        # Фильтруем адреса с координатами
        df_with_coords = dataframe.dropna(subset=['lat', 'lon'])
        
        if df_with_coords.empty:
            logging.warning("No addresses with coordinates to plot")
            return None
        
        # Создаем базовую карту с центром на первом адресе
        first_location = df_with_coords.iloc[0]
        m = folium.Map(location=[first_location['lat'], first_location['lon']], zoom_start=12)
        
        # Создаем кластер маркеров для лучшей производительности
        marker_cluster = MarkerCluster().add_to(m)
        
        # Определяем цвет маркера в зависимости от суммы
        def get_color(total):
            if total < 10000:
                return 'green'
            elif 10000 <= total < 50000:
                return 'orange'
            else:
                return 'red'
        
        # Добавляем маркеры для каждого адреса
        for _, row in df_with_coords.iterrows():
            folium.Marker(
                location=[row['lat'], row['lon']],
                popup=f"Адрес: {row['adress']}<br>Сумма: {row['total_money']:,.2f} руб.",
                icon=folium.Icon(color=get_color(row['total_money']))
            ).add_to(marker_cluster)
        
        # Сохраняем карту в HTML файл
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        map_file = os.path.join(output_folder, f"insurance_map_{timestamp}.html")
        m.save(map_file)
        
        return map_file
    except Exception as e:
        logging.error(f"Error creating map: {e}")
        return None

class InsuranceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка страховых данных")
        self.root.geometry("800x600")
        
        self.input_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.progress = tk.DoubleVar()
        self.status = tk.StringVar(value="Готов к работе")
        self.create_map_var = tk.BooleanVar(value=True)
        self.geocode_var = tk.BooleanVar(value=True)
        
        self._setup_ui()
        logging.info("Приложение инициализируется")

    def _setup_ui(self):
        """Setup UI components"""
        self._set_icon()
        style = ttk.Style()
        style.configure("TButton", padding=6, font=('Arial', 10))
        style.configure("TLabel", padding=6, font=('Arial', 10))

        # Logo
        logo_frame = ttk.Frame(self.root)
        logo_frame.pack(pady=5)
        try:
            logo_path = resource_path("app_logo.png")
            img = Image.open(logo_path).resize((150, 150), Image.Resampling.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img)
            ttk.Label(logo_frame, image=self.logo_img).pack()
        except Exception as e:
            logging.error(f"Failed to load logo: {e}")
            ttk.Label(logo_frame, text="Логотип отсутствует").pack()

        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # File input
        ttk.Label(main_frame, text="Файл Excel:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_file, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Обзор", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)

        # Output folder
        ttk.Label(main_frame, text="Папка для сохранения:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_folder, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Обзор", command=self.browse_folder).grid(row=1, column=2, padx=5, pady=5)

        # Map checkbox
        ttk.Checkbutton(main_frame, text="Создать интерактивную карту", variable=self.create_map_var,
                       command=self.toggle_geocode).grid(row=2, column=0, columnspan=3, pady=5, sticky=tk.W)
        
        # Geocode checkbox (only enabled when map is enabled)
        self.geocode_check = ttk.Checkbutton(main_frame, text="Геокодировать адреса (Nominatim)", 
                                           variable=self.geocode_var, state=tk.NORMAL)
        self.geocode_check.grid(row=3, column=0, columnspan=3, pady=5, sticky=tk.W)

        # Progress bar
        ttk.Label(main_frame, text="Прогресс:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Progressbar(main_frame, variable=self.progress, maximum=100, length=400).grid(row=4, column=1, columnspan=2, padx=5, pady=5)

        # Status
        ttk.Label(main_frame, textvariable=self.status).grid(row=5, column=0, columnspan=3, pady=5)

        # Process button
        ttk.Button(main_frame, text="Обработать", command=self.process_data).grid(row=6, column=0, columnspan=3, pady=10)

        # Developer label
        ttk.Label(self.root, text="Разработано: стажёр Мальцев Максим", font=('Arial', 12)).pack(side=tk.BOTTOM, pady=5)

    def toggle_geocode(self):
        """Enable/disable geocode checkbox based on map checkbox"""
        if self.create_map_var.get():
            self.geocode_check.config(state=tk.NORMAL)
        else:
            self.geocode_check.config(state=tk.DISABLED)

    def _set_icon(self):
        """Set application icon"""
        try:
            icon_path = resource_path("app_icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            logging.error(f"Failed to set icon: {e}")

    def browse_file(self):
        """Select input Excel file"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")])
        if file_path:
            self.input_file.set(file_path)
            if not self.output_folder.get():
                self.output_folder.set(os.path.dirname(file_path))
            logging.info(f"Selected input file: {file_path}")

    def browse_folder(self):
        """Select output folder"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder.set(folder_path)
            logging.info(f"Selected output folder: {folder_path}")

    def update_progress(self, value):
        """Update progress bar from geocoding thread"""
        self.progress.set(value)
        self.root.update_idletasks()

    def process_data(self):
        """Process Excel data and save results"""
        input_file = self.input_file.get()
        output_folder = self.output_folder.get()

        if not input_file or not os.path.exists(input_file):
            messagebox.showerror("Ошибка", "Выберите действительный файл Excel")
            logging.error("Invalid or missing input file")
            return

        if not output_folder or not os.path.isdir(output_folder):
            messagebox.showerror("Ошибка", "Выберите действительную папку для сохранения")
            logging.error("Invalid or missing output folder")
            return

        try:
            self.status.set("Чтение файла...")
            self.progress.set(10)
            df = pd.read_excel(input_file, engine='openpyxl')
            logging.info("Excel file read successfully")

            self.status.set("Очистка текста...")
            self.progress.set(20)
            df['object'] = df['object'].apply(clean_text)

            self.status.set("Очистка адресов...")
            self.progress.set(30)
            df['clean_adress'] = df['adress'].apply(clean_address)

            self.status.set("Фильтрация данных...")
            self.progress.set(40)
            df_filtered = df[~df['object'].str.lower().str.strip().apply(lambda x: bool(REGEX.search(x)) if pd.notna(x) else False)]

            df_filtered['date_end'] = pd.to_datetime(df_filtered['date_end'], dayfirst=True, errors='coerce')
            filter_date = pd.to_datetime(datetime.now().date())
            df_filtered = df_filtered[(df_filtered['date_end'] >= filter_date) & df_filtered['date_end'].notna()]

            self.status.set("Подсчет сумм...")
            self.progress.set(60)
            adress_sums = df_filtered.groupby('clean_adress')['money'].sum().reset_index(name='total_money')
            adress_sums = adress_sums.rename(columns={'clean_adress': 'adress'})

            # Геокодирование адресов, если нужно
            if self.create_map_var.get() and self.geocode_var.get():
                self.status.set("Геокодирование адресов... (это может занять время)")
                self.root.update_idletasks()
                
                unique_addresses = adress_sums['adress'].unique()
                coordinates = geocode_addresses(unique_addresses, self.update_progress)
                
                # Создаем словарь адрес -> координаты
                address_coords = {addr: coords for addr, coords in zip(unique_addresses, coordinates)}
                
                # Добавляем координаты в DataFrame
                adress_sums['lat'] = adress_sums['adress'].map(lambda x: address_coords.get(x, (None, None))[0])
                adress_sums['lon'] = adress_sums['adress'].map(lambda x: address_coords.get(x, (None, None))[1])

            self.status.set("Сохранение результатов...")
            self.progress.set(90)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_folder, f"insurance_results_{timestamp}.xlsx")
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df_filtered.to_excel(writer, sheet_name='filtered_data', index=False)
                adress_sums.to_excel(writer, sheet_name='adress_totals', index=False)
            
            # Создание карты, если выбрано
            if self.create_map_var.get():
                self.status.set("Создание карты...")
                map_file = create_map(adress_sums, output_folder)
                if map_file:
                    webbrowser.open(f"file://{map_file}")
            
            self.progress.set(100)
            self.status.set("Готово!")
            messagebox.showinfo("Успешно", f"Результаты сохранены в:\n{output_folder}")
            logging.info(f"Results saved to {output_folder}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки:\n{str(e)}")
            self.status.set("Ошибка")
            self.progress.set(0)
            logging.error(f"Processing error: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = InsuranceApp(root)
    root.mainloop()
