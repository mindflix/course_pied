import datetime
from openpyxl import Workbook

wb = Workbook()
ws1 = wb.worksheets[0]


def get_week_number():
    week_number = datetime.date.today().isocalendar()[1]
    return week_number


def get_days(week_number):
    r = datetime.datetime.strptime("2021-W{}".format(week_number) + "-2", "%Y-W%W-%w")
    days = [
        "Lundi",
        "Mardi",
        "Mercredi",
        "Jeudi",
        "Vendredi",
        "Samedi",
        "Dimanche",
    ]
    print(r)
    return days


def write_in_cell(cell, value):
    ws1.cell(cell.row, cell.column, value)


def write_in_line(data):
    print(f"Workbook->Writing line for {data}")
    table = ws1.iter_cols(min_row=1, max_col=len(data))
    for value in data:
        write_in_cell(next(table)[0], value)
    print(f"Workbook->Line written for {data}")


def main():
    write_in_line(get_days(get_week_number()))
    wb.save("./data/planning.xlsx")


if __name__ == "__main__":
    main()