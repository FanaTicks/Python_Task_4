import csv
from faker import Faker
import random

class EmployeeInfoGenerator:
    # Конструктор класу для ініціалізації основних параметрів
    def __init__(self, output_path, num_records=2000):
        self.output_path = output_path  # шлях до файлу для збереження даних
        self.num_records = num_records  # кількість записів для генерації
        self.fake_generator = Faker('uk_UA')  # ініціалізація генератора фейкових даних з українською локалізацією

    # Метод для створення випадкової інформації про співробітника
    def create_employee_info(self):
        is_male = random.random() < 0.6  # випадковий вибір статі співробітника
        # Вибір імені та прізвища в залежності від статі
        first_name = self.fake_generator.first_name_male() if is_male else self.fake_generator.first_name_female()
        last_name = self.fake_generator.last_name_male() if is_male else self.fake_generator.last_name_female()
        gender = 'чоловік' if is_male else 'жінка'  # встановлення статі

        # Генерація іншої інформації
        middle_name = self.fake_generator.first_name()
        dob = self.fake_generator.date_of_birth(tzinfo=None, minimum_age=15, maximum_age=85).strftime('%d-%m-%Y')
        position = self.fake_generator.job()
        city = self.fake_generator.city()
        address = self.fake_generator.address().replace('\n', ', ')
        phone = self.fake_generator.phone_number()
        email = self.fake_generator.email()

        # Повертаємо список із згенерованою інформацією
        return [last_name, first_name, middle_name, gender, dob, position, city, address, phone, email]

    # Метод для запису згенерованих даних у CSV-файл
    def write_to_csv(self):
        with open(self.output_path, 'w', newline='', encoding='utf-8-sig') as file:
            csv_writer = csv.writer(file, delimiter=';')
            # Запис заголовків у файл
            csv_writer.writerow(["Прізвище", "Ім’я", "По батькові", "Стать", "Дата народження", "Посада",
                                 "Місто проживання", "Адреса проживання", "Телефон", "Email"])
            # Запис згенерованих даних у файл
            for _ in range(self.num_records):
                csv_writer.writerow(self.create_employee_info())

    # Метод для виконання основних операцій
    def execute(self):
        self.write_to_csv()  # виклик методу для запису даних у файл
        print(f"Generated {self.num_records} records in {self.output_path}")  # вивід повідомлення про успішне виконання


# Створення екземпляру класу та виконання основного методу
employee_info_generator = EmployeeInfoGenerator("employees.csv")
employee_info_generator.execute()
