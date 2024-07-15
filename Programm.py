import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import matplotlib.pyplot as plt


class StudentAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализ списка студентов")
        self.df = None

        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.configure('TFrame', background='#e0e0e0')
        style.configure('TButton', font=('Arial', 12))
        style.configure('TLabel', background='#e0e0e0', font=('Arial', 12))
        style.configure('Header.TLabel', font=('Arial', 18, 'bold'))

        frame = ttk.Frame(self.root, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        header_label = ttk.Label(frame, text="Анализ списка студентов", style='Header.TLabel')
        header_label.grid(row=0, column=0, pady=10, columnspan=3)

        load_button = ttk.Button(frame, text="Загрузить файл", command=self.load_file)
        load_button.grid(row=1, column=0, pady=10)

        save_button = ttk.Button(frame, text="Сохранить результаты", command=self.save_results)
        save_button.grid(row=1, column=1, pady=10)

        plot_button = ttk.Button(frame, text="Построить график", command=self.plot_data)
        plot_button.grid(row=1, column=2, pady=10)

        filter_button = ttk.Button(frame, text="Фильтровать данные", command=self.filter_data)
        filter_button.grid(row=2, column=0, pady=10, columnspan=3)

        self.result_textbox = tk.Text(frame, wrap='word', height=20, width=80, font=('Arial', 12))
        self.result_textbox.grid(row=3, column=0, pady=10, columnspan=3)
        self.result_textbox.config(state=tk.DISABLED)

        scrollbar = ttk.Scrollbar(frame, command=self.result_textbox.yview)
        scrollbar.grid(row=3, column=2, sticky='nse')
        self.result_textbox['yscrollcommand'] = scrollbar.set

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                self.process_data()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить файл: {e}")

    def save_results(self):
        if self.df is None:
            messagebox.showerror("Ошибка", "Нет данных для сохранения")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.result_textbox.get("1.0", tk.END))
                messagebox.showinfo("Успех", "Результаты успешно сохранены")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")

    def plot_data(self):
        if self.df is None:
            messagebox.showerror("Ошибка", "Нет данных для построения графика")
            return
        try:
            plt.figure(figsize=(10, 6))
            self.df['Балл'].plot(kind='hist', bins=20, edgecolor='black')
            plt.title('Распределение баллов')
            plt.xlabel('Балл')
            plt.ylabel('Количество студентов')
            plt.grid(False)
            plt.show()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось построить график: {e}")

    def filter_data(self):
        if self.df is None:
            messagebox.showerror("Ошибка", "Нет данных для фильтрации")
            return

        filter_window = tk.Toplevel(self.root)
        filter_window.title("Фильтр данных")

        ttk.Label(filter_window, text="Минимальный балл:").grid(row=0, column=0, pady=5)
        min_score_entry = ttk.Entry(filter_window)
        min_score_entry.grid(row=0, column=1, pady=5)

        ttk.Label(filter_window, text="Максимальный балл:").grid(row=1, column=0, pady=5)
        max_score_entry = ttk.Entry(filter_window)
        max_score_entry.grid(row=1, column=1, pady=5)

        def apply_filter():
            try:
                min_score = float(min_score_entry.get()) if min_score_entry.get() else float('-inf')
                max_score = float(max_score_entry.get()) if max_score_entry.get() else float('inf')
                filtered_df = self.df[(self.df['Балл'] >= min_score) & (self.df['Балл'] <= max_score)]
                self.process_data(filtered_df)
                filter_window.destroy()
            except ValueError as e:
                messagebox.showerror("Ошибка", f"Некорректный ввод баллов: {e}")

        ttk.Button(filter_window, text="Применить фильтр", command=apply_filter).grid(row=2, column=0, columnspan=2, pady=10)

    def process_data(self, df=None):
        if df is None:
            df = self.df

        required_columns = ['Имя', 'Балл', 'Статус']
        if not all(col in df.columns for col in required_columns):
            messagebox.showerror("Ошибка", "В файле отсутствуют необходимые столбцы: Имя, Балл, Статус")
            return

        df['Балл'].fillna(df['Балл'].mean(), inplace=True)
        df['Статус'].fillna('Не сдал', inplace=True)

        average_score = df['Балл'].mean()
        median_score = df['Балл'].median()
        pass_count = df[df['Статус'] == 'Сдал'].shape[0]
        fail_count = df[df['Статус'] == 'Не сдал'].shape[0]
        max_score_student = df.loc[df['Балл'].idxmax()]
        min_score_student = df.loc[df['Балл'].idxmin()]
        score_stats = df['Балл'].describe()
        mean_score_by_status = df.groupby('Статус')['Балл'].mean()
        above_average_count = df[df['Балл'] > average_score].shape[0]
        below_average_count = df[df['Балл'] < average_score].shape[0]
        score_distribution = df['Балл'].value_counts().sort_index()

        result_text = (
            f"Средний балл студентов: {average_score:.2f}\n"
            f"Медианный балл студентов: {median_score:.2f}\n\n"
            f"Количество сдавших: {pass_count}\n"
            f"Количество не сдавших: {fail_count}\n\n"
            f"Студент с максимальным баллом: {max_score_student['Имя']} с баллом {max_score_student['Балл']}\n"
            f"Студент с минимальным баллом: {min_score_student['Имя']} с баллом {min_score_student['Балл']}\n\n"
            f"Статистика по баллам:\n{score_stats}\n\n"
            f"Средний балл по статусу:\n{mean_score_by_status}\n\n"
            f"Количество студентов с баллом выше среднего: {above_average_count}\n"
            f"Количество студентов с баллом ниже среднего: {below_average_count}\n\n"
            f"Распределение баллов:\n{score_distribution}"
        )
        self.result_textbox.config(state=tk.NORMAL)
        self.result_textbox.delete(1.0, tk.END)
        self.result_textbox.insert(tk.END, result_text)
        self.result_textbox.config(state=tk.DISABLED)


# Запуск приложения
if __name__ == "__main__":
    root = tk.Tk()
    app = StudentAnalyzerApp(root)
    root.mainloop()
