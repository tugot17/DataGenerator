from random import randint
import xlsxwriter
import datetime
import radar


def from_male_surnames_create_femele_surnames(male_surnames_file = "male_surnames.txt"):
    with open(male_surnames_file, 'r') as from_file:
        with open("female_surnames.txt", 'w') as to_file:
            for line in from_file:
                line = line.strip()
                if line[-1] == "i":
                    line = line[:-1] + "a"
                to_file.write('{}\n'.format(line))



def create_excel_file(how_many_examples=30):
    with open("male_surnames.txt", 'r', encoding="utf8") as from_file:
        male_surnames = list(from_file)

    with open("male_names.txt", 'r', encoding="utf8") as from_file:
        male_names = list(from_file)

    with open("female_surnames.txt", 'r', encoding="utf8") as from_file:
        female_surnames = list(from_file)

    with open("female_names.txt", 'r', encoding="utf8") as from_file:
        female_names = list(from_file)

    school_classes = ('1A', '1B', "1C", "1D", '2A', '2B', "2C", "2D", '3A', '3B', "3C", "3D", '4A', '4B', "4C", "4D")


    random_class = school_classes[randint(0,len(school_classes)-1)]
    random_female_surname = female_surnames[randint(0,len(female_surnames)-1)].strip()
    random_female_name = female_names[randint(0,len(female_names)-1)].strip()

    row = 0
    workbook = xlsxwriter.Workbook('Nazwiska1.xlsx')
    worksheet = workbook.add_worksheet()

    for x in range(how_many_examples//2):
        random_class = school_classes[randint(0, len(school_classes) - 1)]
        random_female_surname = female_surnames[randint(0, len(female_surnames) - 1)].strip()
        random_female_name = female_names[randint(0, len(female_names) - 1)].strip()



        worksheet.write(row, 0, row)
        worksheet.write(row, 1, random_female_surname)
        worksheet.write(row, 2, random_female_name)
        worksheet.write(row, 3, get_random_born_day(2000))
        worksheet.write(row, 4, "Wroclaw")
        worksheet.write(row, 5, "K")
        worksheet.write(row, 6, random_class)
        row +=1



        random_class = school_classes[randint(0, len(school_classes) - 1)]
        random_male_surname = male_surnames[randint(0, len(male_surnames) - 1)].strip()
        random_male_name = male_names[randint(0, len(male_names) - 1)].strip().strip()


        worksheet.write(row, 0, row)
        worksheet.write(row, 1, random_male_surname)
        worksheet.write(row, 2, random_male_name)
        worksheet.write(row, 3, get_random_born_day(2000))
        worksheet.write(row, 4, "Wroclaw")
        worksheet.write(row, 5, "M")
        worksheet.write(row, 6, random_class)
        row +=1

    workbook.close()



def get_random_born_day(year=2000):
    born_date = radar.random_datetime(
        start = datetime.datetime(year=year, month=10, day=10),
        stop = datetime.datetime(year=year, month=12, day=31)
    )
    return born_date.strftime('%d/%m/%Y')












