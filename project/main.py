# main.py
import tkinter as tk
from app_gui import IikoReportApp # Импортируем класс GUI

def main():
    root = tk.Tk()
    app = IikoReportApp(root) # Создаем экземпляр GUI
    root.mainloop()          # Запускаем главный цикл событий

if __name__ == "__main__":
    main()