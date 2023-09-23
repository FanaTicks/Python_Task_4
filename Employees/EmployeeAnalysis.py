import pandas as pd
import os
import matplotlib.pyplot as plt
from datetime import datetime

class EmployeeAnalysis:
    def __init__(self, csv_path):
        self.csv_path = csv_path  # Шлях до CSV-файлу
        self.dataset = None  # Датасет

    def load_dataset(self):
        if not os.path.exists(self.csv_path):
            print(f"Помилка: Неможливо знайти або відкрити CSV-файл: {self.csv_path}")
            return False

        try:
            # Зазначте параметр dayfirst=True, щоб вказати, що день йде перед місяцем у форматі дати
            self.dataset = pd.read_csv(self.csv_path, encoding='utf-8-sig', sep=';', parse_dates=['Дата народження'], dayfirst=True)
            print("OK. Дані успішно завантажено.")
            return True
        except pd.errors.ParserError:
            print(f"Помилка: Неможливо проаналізувати CSV-файл: {self.csv_path}")
            return False

    def calculate_age(self, dob):
        today = datetime.today()
        age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
        return age

    def gender_distribution(self):
        print("\nЗавдання: Аналіз розподілу за статтю\n" + "-"*40)
        gender_counts = self.dataset['Стать'].value_counts()
        print(gender_counts.to_string())
        gender_counts.plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=['#66b3ff','#99ff99'])
        plt.title('Розподіл за статтю')
        plt.ylabel('')  # Сховати підпис осі Y
        plt.show()

    def age_group_distribution(self):
        print("\nЗавдання: Аналіз розподілу за віковими групами\n" + "-"*40)
        self.dataset['Вік'] = self.dataset['Дата народження'].apply(self.calculate_age)
        age_bins = [0, 18, 45, 70, 200]
        age_labels = ['>18', '18-45', '45-70', '<70']
        self.dataset['Вікова група'] = pd.cut(self.dataset['Вік'], bins=age_bins, labels=age_labels, right=False)
        age_group_counts = self.dataset['Вікова група'].value_counts()
        print(age_group_counts.to_string())
        age_group_counts.plot(kind='bar', color=['#ff9999','#66b3ff','#99ff99','#ffcc99'])
        plt.title('Розподіл за віковими групами')
        plt.show()

    def gender_age_group_distribution(self):
        print("\nЗавдання: Аналіз розподілу за статтю та віковими групами\n" + "-"*40)
        gender_age_groups_counts = self.dataset.groupby(['Стать', 'Вікова група'], observed=True).size().unstack()
        print(gender_age_groups_counts.to_string())
        gender_age_groups_counts.plot(kind='bar', stacked=True, color=['#ff9999','#66b3ff','#99ff99','#ffcc99'])
        plt.title('Розподіл за статтю та віковими групами')
        plt.show()

# Створення екземпляру класу та виконання аналізу
analysis = EmployeeAnalysis("employees.csv")
if analysis.load_dataset():
    analysis.gender_distribution()
    analysis.age_group_distribution()
    analysis.gender_age_group_distribution()
