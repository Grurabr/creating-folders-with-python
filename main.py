import os
import tkinter as tk
from tkinter import messagebox, scrolledtext
import datetime


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Kansion luominen. EA")

        # Начальные размеры окна (без поля логов)
        self.root.geometry("300x100")

        # Поле для ввода имени
        self.label = tk.Label(self.root, text="Nimikenumero:")
        self.label.pack(pady=5)

        self.entry = tk.Entry(self.root, width=30)
        self.entry.pack(pady=5)

        # Кнопки "ОК" и "Отмена"
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        ok_button = tk.Button(button_frame, text="OK", command=self.start_creation)
        ok_button.pack(side="left", padx=5)

        cancel_button = tk.Button(button_frame, text="Peruuta", command=self.root.quit)
        cancel_button.pack(side="right", padx=5)

        # Создание окна для логов (изначально скрыто)
        self.log_text = scrolledtext.ScrolledText(self.root, width=70, height=20, state="disabled")
        self.log_text.pack(pady=10)
        self.log_text.pack_forget()  # Скрываем поле лога

    def log(self, message):
        """Функция для вывода логов в текстовое поле."""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"{timestamp} - {message}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        print(f"{timestamp} - {message}")  # Для отладки в консоли

    def show_error_message(self, error_message):
        messagebox.showerror("Virhe", f"Virhe:\n{error_message}")

    def start_creation(self):
        """Функция для запуска основной функции после нажатия кнопки 'ОК'."""
        nimike = self.entry.get().strip()
        if not nimike:
            self.show_error_message("Nimikenumero puuttuu.")
            return

        # Очищаем и отображаем поле логов, увеличиваем окно
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")
        self.log_text.pack(pady=10)  # Показываем поле логов

        # Увеличиваем размеры окна
        self.root.geometry("600x400")

        # Запускаем процесс создания папок и ссылок
        self.main(nimike.upper())

    def main(self, nimike):
        try:
            # Переменные для путей
            tiedoston_haku_polku = r"C:\Kuvat"
            kansion_polku = fr"C:\Yhteiset\LAATU\Mittaukset\{nimike}"
            cmmOhjelman_polku = os.path.join(kansion_polku, "CMM Ohjelma", f"{nimike}.PRG")
            romerOhjelmanPolku = os.path.join(kansion_polku, "ROMER Ohjelma", f"{nimike}.mcam")

            # Создание необходимых папок
            self.create_folder_if_not_exists(os.path.join(kansion_polku, "CMM Ohjelma"))
            self.create_folder_if_not_exists(os.path.join(kansion_polku, "CMM Raportti"))
            self.create_folder_if_not_exists(os.path.join(kansion_polku, "ROMER Raportti"))
            self.create_folder_if_not_exists(os.path.join(kansion_polku, "ROMER Ohjelma"))
            self.create_folder_if_not_exists(os.path.join(kansion_polku, "JIGI"))

            # Создание ссылок на файлы, если они существуют
            self.create_shortcut_if_file_exists(tiedoston_haku_polku, kansion_polku, nimike, ".pdf", "KUVA")
            self.create_shortcut_if_file_exists(tiedoston_haku_polku, kansion_polku, nimike, [".step", ".stp"], ".STEP")
            self.create_shortcut_if_file_exists(tiedoston_haku_polku, kansion_polku, f"{nimike}M", ".xls",
                                                "MITTAPÖYTÄКIRJA")

            # Создание ссылки на программу CMM
            self.create_shortcut(kansion_polku, nimike, cmmOhjelman_polku, "OHJELMA")

            # Открыть папку с результатами
            os.startfile(kansion_polku)
            self.log("Operaatio suoritettiin onnistuneesti.")

        except Exception as e:
            self.show_error_message(str(e))
            self.log(f"Virhe: {e}")

    def create_folder_if_not_exists(self, path):
        if not os.path.exists(path):
            os.makedirs(path)
            self.log(f"Kansio luotu: {path}")
        else:
            self.log(f"Kansio on jo olemassa: {path}")

    def create_shortcut_if_file_exists(self, search_path, dest_path, filename, extensions, shortcut_name):
        if isinstance(extensions, str):
            extensions = [extensions]

        for ext in extensions:
            file_path = self.find_file(search_path, filename + ext)
            if file_path:
                self.create_shortcut(dest_path, shortcut_name, file_path, shortcut_name)
                self.log(f"Linkki luotu: {shortcut_name} -> {file_path}")
                break
        else:
            self.log(f"Tiedostoa {filename}, jonka laajennus on {extensions}, ei löydy.")

    def find_file(self, search_path, filename):
        for root, _, files in os.walk(search_path):
            if filename in files:
                return os.path.join(root, filename)
        return None

    def create_shortcut(self, dest_folder, shortcut_name, target_path, link_name):
        try:
            from win32com.client import Dispatch
            shell = Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(os.path.join(dest_folder, f"{link_name}.lnk"))
            shortcut.TargetPath = target_path
            shortcut.Save()
            self.log(f"Pikakuvake luotu {link_name}.lnk")
        except ImportError:
            raise ImportError("Need install pywin32: pip install pywin32")


# Инициализация и запуск окна
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
