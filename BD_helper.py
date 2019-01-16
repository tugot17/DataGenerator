from random import randint
import xlsxwriter
import datetime
import radar
import roman


def from_male_surnames_create_female_surnames(male_surnames_file="male_surnames.txt"):
    with open(male_surnames_file, 'r') as from_file:
        with open("female_surnames.txt", 'w') as to_file:
            for line in from_file:
                line = line.strip()
                if line[-1] == "i":
                    line = line[:-1] + "a"
                to_file.write('{}\n'.format(line))


def create_students_excel_file(how_many_examples=30):
    with open("Data/male_surnames.txt", 'r', encoding="utf8") as from_file:
        male_surnames = list(from_file)

    with open("Data/male_names.txt", 'r', encoding="utf8") as from_file:
        male_names = list(from_file)

    with open("Data/female_surnames.txt", 'r', encoding="utf8") as from_file:
        female_surnames = list(from_file)

    with open("Data/female_names.txt", 'r', encoding="utf8") as from_file:
        female_names = list(from_file)

    row = 0
    workbook = xlsxwriter.Workbook('Student_surnames.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write(row, 0, row)
    worksheet.write(row, 1, "Surname")
    worksheet.write(row, 2, "Name")
    worksheet.write(row, 3, "Born date")
    worksheet.write(row, 4, "City name")
    worksheet.write(row, 5, "Gender")
    worksheet.write(row, 6, "Class number")

    row += 1

    for x in range(how_many_examples // 2):
        random_class, random_female_surname, random_female_name, random_city, born_year = _get_random_student_data(
            female_surnames, female_names)

        worksheet.write(row, 0, row)
        worksheet.write(row, 1, random_female_surname)
        worksheet.write(row, 2, random_female_name)
        worksheet.write(row, 3, _get_random_day(born_year))
        worksheet.write(row, 4, random_city)
        worksheet.write(row, 5, "K")
        worksheet.write(row, 6, random_class)
        row += 1

        random_class, random_male_surname, random_male_name, random_city, born_year = _get_random_student_data(
            male_surnames, male_names)

        worksheet.write(row, 0, row)
        worksheet.write(row, 1, random_male_surname)
        worksheet.write(row, 2, random_male_name)
        worksheet.write(row, 3, _get_random_day(born_year))
        worksheet.write(row, 4, random_city)
        worksheet.write(row, 5, "M")
        worksheet.write(row, 6, random_class)
        row += 1

        with open('list_of_names.txt', 'a', encoding="utf8") as fileobj:
            fileobj.write('{} {}\n'.format(random_female_name, random_female_surname))
            fileobj.write('{} {}\n'.format(random_male_name, random_male_surname))

    workbook.close()


def create_grades_excel_file():
    surnames = []

    with open("list_of_names.txt", 'r', encoding="utf8") as from_file:
        for line in from_file:
            line = line.split()
            try:
                surname = line[1]
                surnames.append(surname)
            except:
                pass

    row = 0
    workbook = xlsxwriter.Workbook('Grades.xlsx')
    worksheet = workbook.add_worksheet()

    lesson_types = ["Matematyka", "Fizyka", "Chemia", "Informatyka", "Polski", "Historia", "Geografia", "Angielski",
                    "Niemiecki", "Matematyka"]

    grades = [2, 3, 3.5, 4, 4.5, 5]

    for surname in surnames:
        random_marks_amount = randint(0, 10)

        for i in range(random_marks_amount):
            if random_marks_amount % 4 == 0 or random_marks_amount % 7 == 0:
                continue
            random_date = _get_random_day(2017)

            if _is_date_during_holidays(random_date):
                continue

            random_lesson_type = lesson_types[randint(0, len(lesson_types) - 1)]
            random_mark = grades[randint(0, len(grades) - 1)]

            worksheet.write(row, 0, surname)
            worksheet.write(row, 1, random_lesson_type)
            worksheet.write(row, 2, random_mark)
            worksheet.write(row, 3, random_date)
            row += 1

    workbook.close()


def create_teachers_excel_file(teachers_amount=10):
    with open("Data/male_surnames.txt", 'r', encoding="utf8") as from_file:
        male_surnames = list(from_file)

    with open("Data/male_names.txt", 'r', encoding="utf8") as from_file:
        male_names = list(from_file)

    with open("Data/female_surnames.txt", 'r', encoding="utf8") as from_file:
        female_surnames = list(from_file)

    with open("Data/female_names.txt", 'r', encoding="utf8") as from_file:
        female_names = list(from_file)

    row = 0
    workbook = xlsxwriter.Workbook('Nauczyciele.xlsx')
    worksheet = workbook.add_worksheet()

    for x in range(teachers_amount // 2):
        random_female_surname, random_female_name, employment_date, born_year, salary, pensum, telephone, bonus = \
            _get_random_teacher_data(female_surnames, female_names)

        worksheet.write(row, 0, row)
        worksheet.write(row, 1, random_female_surname)
        worksheet.write(row, 2, random_female_name)
        worksheet.write(row, 3, employment_date)
        worksheet.write(row, 4, _get_random_day(born_year))
        worksheet.write(row, 5, "K")
        worksheet.write(row, 6, salary)
        worksheet.write(row, 7, pensum)
        worksheet.write(row, 8, telephone)
        worksheet.write(row, 9, bonus)
        row += 1

        random_male_surname, random_male_name, employment_date, born_year, salary, pensum, telephone, bonus = \
            _get_random_teacher_data(female_surnames, female_names)

        worksheet.write(row, 0, row)
        worksheet.write(row, 1, random_male_surname)
        worksheet.write(row, 2, random_male_name)
        worksheet.write(row, 3, employment_date)
        worksheet.write(row, 4, _get_random_day(born_year))
        worksheet.write(row, 5, "M")
        worksheet.write(row, 6, salary)
        worksheet.write(row, 7, pensum)
        worksheet.write(row, 8, telephone)
        worksheet.write(row, 9, bonus)
        row += 1

    workbook.close()


# <editor-fold desc="Helper methods">
def _get_random_teacher_data(surnames, names):
    random_surname = surnames[randint(0, len(surnames) - 1)].strip()
    random_name = names[randint(0, len(names) - 1)].strip()
    born_year = randint(1970, 1985)
    employment_date = _get_random_day(randint(2010, 2017))
    salary = 3000
    pensum = 300
    telephone = "".join([str(randint(1, 9)) for x in range(9)])
    bonus = randint(0, 1000)

    return random_surname, random_name, employment_date, born_year, salary, pensum, telephone, bonus


def _get_random_student_data(surnames, names):
    school_classes = (
        'Ia', "Ic", "Id", 'IIa', "IIb", "IIc", "IId", 'IIIa', "IIIb", "IIIc", "IIId", 'IVa', "IVb", "IVc", "IVd")
    cities = ["Trzebnica", "Wroclaw", "Olawa", "Trestno", "Radwanice", "Siechnice", "Kielczow", "Smolec", "Wilkszyn",
              "Wilczyce", "Czernica", "Kobierzyce", "Wroclaw", "Wroclaw", "Wroclaw"]
    born_years = [2002, 2001, 2000, 1999]

    random_class = school_classes[randint(0, len(school_classes) - 1)]
    random_surname = surnames[randint(0, len(surnames) - 1)].strip()
    random_name = names[randint(0, len(names) - 1)].strip()
    random_city = cities[randint(0, len(cities) - 1)]
    born_year = born_years[roman.fromRoman(random_class[:len(random_class) - 1]) - 1]

    return random_class, random_surname, random_name, random_city, born_year


def _get_random_day(year=2000):
    born_date = radar.random_datetime(
        start=datetime.datetime(year=year, month=1, day=10),
        stop=datetime.datetime(year=year, month=12, day=31)
    )
    return born_date.strftime('%m/%d/%Y')


def _is_date_during_holidays(date: str):
    month = date[:2]

    if month[0] == 0:
        month = int(month[1])
    else:
        month = int(month)

    return month == 7 or month == 8
# </editor-fold>
