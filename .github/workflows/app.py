import pandas as pd
from rapidfuzz import process
import tkinter as tk
from tkinter import filedialog, messagebox

# ---------- Load Library ----------
def load_library(path):
    df = pd.read_excel(path)
    df['Название_lower'] = df['Название'].str.lower()
    return df

# ---------- Find Match ONLY in Library ----------
def find_match(name, library_df):
    name = str(name).lower()
    choices = library_df['Название_lower'].tolist()
    match = process.extractOne(name, choices, score_cutoff=80)
    if match:
        return library_df.iloc[match[2]]
    return None

# ---------- Normalize ----------
def normalize(df, library_df, source_name):
    results = []

    for _, row in df.iterrows():
        name = row[0]
        qty = row[3]

        lib_match = find_match(name, library_df)

        if lib_match is None:
            results.append({
                'ID': None,
                'Name': name,
                'Qty': 0,
                'Comment': f'Нет в библиотеке ({source_name})'
            })
        else:
            converted = qty * lib_match['Коэф']
            results.append({
                'ID': lib_match['ID'],
                'Name': lib_match['Название'],
                'Qty': converted,
                'Comment': ''
            })

    return pd.DataFrame(results)

# ---------- Compare ----------
def compare(master_df, estimate_df):
    master_group = master_df.groupby('ID')['Qty'].sum().reset_index()
    estimate_group = estimate_df.groupby('ID')['Qty'].sum().reset_index()

    result = pd.merge(master_group, estimate_group, on='ID', how='outer', suffixes=('_master', '_estimate')).fillna(0)

    comments = []
    diffs = []

    for _, row in result.iterrows():
        diff = row['Qty_master'] - row['Qty_estimate']
        diffs.append(diff)

        if diff > 0:
            comments.append('Перерасход')
        elif diff < 0:
            comments.append('Недобор')
        else:
            comments.append('ОК')

    result['Difference'] = diffs
    result['Comment'] = comments

    return result

# ---------- GUI ----------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Сравнение материалов")

        self.master_file = None
        self.estimate_file = None
        self.library_file = None

        tk.Button(root, text="Загрузить файл Мастера", command=self.load_master).pack(pady=5)
        tk.Button(root, text="Загрузить файл Сметчика", command=self.load_estimate).pack(pady=5)
        tk.Button(root, text="Загрузить Библиотеку", command=self.load_library).pack(pady=5)

        tk.Button(root, text="Обработать", command=self.process).pack(pady=20)

    def load_master(self):
        self.master_file = filedialog.askopenfilename()

    def load_estimate(self):
        self.estimate_file = filedialog.askopenfilename()

    def load_library(self):
        self.library_file = filedialog.askopenfilename()

    def process(self):
        if not all([self.master_file, self.estimate_file, self.library_file]):
            messagebox.showerror("Ошибка", "Загрузите все файлы")
            return

        try:
            lib = load_library(self.library_file)

            master = pd.read_excel(self.master_file)
            estimate = pd.read_excel(self.estimate_file)

            norm_master = normalize(master, lib, "Мастер")
            norm_estimate = normalize(estimate, lib, "Сметчик")

            comparison = compare(norm_master, norm_estimate)

            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
            with pd.ExcelWriter(save_path) as writer:
                norm_master.to_excel(writer, sheet_name='Мастер', index=False)
                norm_estimate.to_excel(writer, sheet_name='Сметчик', index=False)
                comparison.to_excel(writer, sheet_name='Сравнение', index=False)

            messagebox.showinfo("Готово", "Файл успешно создан")

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

# ---------- Run ----------
if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
