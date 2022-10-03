from openpyxl import load_workbook
import pandas as pd


class TimeSaver:

    def __init__(self, data_path, splited_names, shihva):
        self.data_path = data_path
        self.splited_names = splited_names
        self.shihva_ascii = ord(shihva)

    def save_my_time(self):
        self.search_in_excel()
        self.save_excel_file()

    def search_in_excel(self):
        workbook = load_workbook(filename=self.data_path)
        sheet = workbook.active
        index = 1
        reshumim = []
        while sheet[f"A{index}"].value:
            try:
                parrant, phone_number = str(sheet[f"R{index}"].value).split("-")
            except:
                parrant = "לא נמצא"
                phone_number = sheet[f"R{index}"].value
            current_name = [sheet[f"E{index}"].value, sheet[f"D{index}"].value, "אין", "לא", parrant,
                            phone_number, "אין", sheet[f"F{index}"].value, sheet[f"S{index}"].value]

            for m in self.splited_names.keys():
                joined_name = f"{current_name[0]} {current_name[1]}"
                if joined_name in self.splited_names[m] and \
                        ord(sheet[f"K{index}"].value[0]) == self.shihva_ascii:
                    current_name.insert(2, m)
                    reshumim.append(current_name)

            index += 1

        names = []
        for i in reshumim:
            names.append(f"{i[0]} {i[1]}")

        missed = []
        for i in self.splited_names.values():
            for j in i:
                if j not in names:
                    missed.append(j)

        self.reshumim = reshumim
        self.missed = missed

    def save_excel_file(self):
        l = ["מייל הורה", "תעודת זהות חניך", "אישור הורים יש/איו", "מספר טלפון הורה", "שם הורה", "צמחוני/טבעוני", "בעיות בריאות", "מדריך", "שם משפחה", "שם חניך"]
        l.reverse()
        self.reshumim.insert(0, l)
        a = []
        for j in range(len(self.reshumim[0])):
            l = []

            for i in self.reshumim:
                l.append(i[len(self.reshumim[0]) - j - 1])
            a.append(l)

        a.append(["", "", "נפלו בין הכיסאות:"] + self.missed)

        df = pd.DataFrame(a).T
        df.to_excel(excel_writer="result.xlsx")


def main(data_path, splited_names, age):
    t = TimeSaver(data_path, splited_names, age)
    t.save_my_time()
