import time, json
from openpyxl import Workbook

wb = Workbook()
ws1 = wb.worksheets[0]
output = []


class TimeSlot:
    def __init__(self, user):
        self.user = user

    def add_time_slot(self, start, end):
        time_slot = str(start) + "-" + str(end)
        self.write_json(time_slot)

    def write_json(self, time_slot):
        data = {"user": self.user, "duration": time_slot}
        output.append(data)
        with open("./data/timetable.json", "a") as json_file:
            json.dump(output, json_file)


def write_in_cell(cell, value):
    ws1.cell(cell.row, cell.column, value)


def write_in_line(data):
    print(f"Workbook->Writing line for {data}")
    table = ws1.iter_cols(min_row=1, max_col=len(data))
    for value in data:
        write_in_cell(next(table)[0], value)
    print(f"Workbook->Line written for {data}")


def main():
    h1 = TimeSlot("Nicolas Demol")
    h1.add_time_slot(10, 11)
    h1.add_time_slot(11, 12)
    wb.save("./data/planning.xlsx")


if __name__ == "__main__":
    main()