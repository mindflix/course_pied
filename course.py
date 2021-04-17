import time, json
from openpyxl import Workbook

wb = Workbook()
ws1 = wb.worksheets[0]


class User:
    def __init__(self, user):
        self.user = user
        self.output = []

    def add_time_slot(self, start, end):
        time_slot = str(start) + "-" + str(end)
        data = {"user": self.user, "duration": time_slot}
        self.output.append(data)

    def write_json(self):
        with open("./data/timetable.json", "w") as json_file:
            json.dump(self.output, json_file)


def write_in_cell(cell, value):
    ws1.cell(cell.row, cell.column, value)


def write_in_line(data):
    print(f"Workbook->Writing line for {data}")
    table = ws1.iter_cols(min_row=1, max_col=len(data))
    for value in data:
        write_in_cell(next(table)[0], value)
    print(f"Workbook->Line written for {data}")


def main():
    h1 = User("Nicolas Demol")
    h1.add_time_slot(10, 11)
    h1.add_time_slot(11, 12)
    h1.write_json()
    wb.save("./data/planning.xlsx")


if __name__ == "__main__":
    main()