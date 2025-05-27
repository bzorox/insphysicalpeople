import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import logging
from datetime import datetime
from PIL import Image, ImageTk

# Configure logging
logging.basicConfig(filename='insurance_app.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def resource_path(relative_path):
    """Get absolute path to resource for PyInstaller and development"""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)

# Compile regex once at module level
PATTERN = r"""
    \b(?:[гГ][рР]?[аА]?[жЖ]?[дД]?[аА]?[нН]?[сС]?[кК]?[аАяЯ]?[яЯ]?\s*)?
    (?:ответ[сстССТ]?[тТ]?[вВ]?[еЕ]?[нН]?[нН]?[оО]?[сС]?[тТ]?[ьЬ]?|ГО|г\.о\.)\b
    |\b(?:страхование\s*)?(?:гражданской|гр[ао]жданской|гржданской)\s*(?:ответственности|ответсвенности|ответсвенноти|ответственноти)\b
    |\b(?:ответственность|ответсвенность|ответсвенноть|ответственноть)\b
    |\bГО\b|\bг\.о\.\b
"""
REGEX = re.compile(PATTERN, flags=re.IGNORECASE | re.VERBOSE)

def clean_text(text):
    """Очистка полей object от ФИО"""
    text = str(text)
    text = re.sub(r"\b[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\b|\b[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]\b|\b[А-ЯЁ][а-яё]+-\b|\b[А-ЯЁ][а-яё]+s\b|\b[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s\b", '', text)
    text = re.sub(r",\d{2}-\d{2}-\d{2},", '', text)
    return text.strip()

class InsuranceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка страховых данных")
        self.root.geometry("700x550")
        
        self.input_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.progress = tk.DoubleVar()
        self.status = tk.StringVar(value="Готов к работе")
        
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
            logo_path = resource_path("assets/app_logo.png")
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

        # Progress bar
        ttk.Label(main_frame, text="Прогресс:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Progressbar(main_frame, variable=self.progress, maximum=100, length=400).grid(row=2, column=1, columnspan=2, padx=5, pady=5)

        # Status
        ttk.Label(main_frame, textvariable=self.status).grid(row=3, column=0, columnspan=3, pady=5)

        # Process button
        ttk.Button(main_frame, text="Обработать", command=self.process_data).grid(row=4, column=0, columnspan=3, pady=10)

        # Developer label
        ttk.Label(self.root, text="Разработано: стажёр Мальцев Максим", font=('Arial', 12)).pack(side=tk.BOTTOM, pady=5)

    def _set_icon(self):
        """Set application icon"""
        try:
            icon_path = resource_path("assets/app_icon.ico")
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
            self.progress.set(30)
            df['object'] = df['object'].apply(clean_text)

            self.status.set("Фильтрация данных...")
            self.progress.set(50)
            df_filtered = df[~df['object'].str.lower().str.strip().apply(lambda x: bool(REGEX.search(x)) if pd.notna(x) else False)]

            df_filtered['date_end'] = pd.to_datetime(df_filtered['date_end'], dayfirst=True, errors='coerce')
            filter_date = pd.to_datetime(datetime.now().date())
            df_filtered = df_filtered[(df_filtered['date_end'] >= filter_date) & df_filtered['date_end'].notna()]

            self.status.set("Подсчет сумм...")
            self.progress.set(70)
            adress_sums = df_filtered.groupby('adress')['money'].sum().reset_index(name='total_money')

            self.status.set("Сохранение результатов...")
            self.progress.set(90)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_folder, f"insurance_results_{timestamp}.xlsx")
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df_filtered.to_excel(writer, sheet_name='filtered_data', index=False)
                adress_sums.to_excel(writer, sheet_name='adress_totals', index=False)
            
            self.progress.set(100)
            self.status.set("Готово!")
            messagebox.showinfo("Успешно", f"Результаты сохранены в:\n{output_file}")
            logging.info(f"Results saved to {output_file}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки:\n{str(e)}")
            self.status.set("Ошибка")
            self.progress.set(0)
            logging.error(f"Processing error: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = InsuranceApp(root)
    root.mainloop()
