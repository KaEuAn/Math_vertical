#coding=utf-8
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

def column_letter(column_number):
    return chr(ord('A')  - 1 + column_number)

fi = open("sheets.json", "r", encoding='utf-8')
sheets = json.load(fi)
fi.close()

class google_tables:
    def __init__(self, docs_name="Рассылки", wks_name="Координаторы", gt_prev=None):
        if gt_prev == None:
            scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
            credentials = ServiceAccountCredentials.from_json_keyfile_name('..//..//keys//math vertical-c8b635a4089b.json', scope)
            self.gt = gspread.authorize(credentials)
        else:
            self.gt = gt_prev
        self.wks = self.gt.open("Рассылки")
        self.sheet = self.wks.worksheet("Координаторы")
        self.row_count = self.sheet.row_count

    def get_cell_list(self, rows = [1, -1], columns = [1, 2]):
        #numeration starts with 1 as like in google tables  
        if rows[1] == -1:
            rows[1] = self.row_count
        cell_string = column_letter(columns[0]) + str(rows[0]) + ':' + column_letter(columns[1]) + str(rows[1])
        cell_table = self.sheet.range(cell_string)
        return cell_table

    def get_cell_table(self, rows = [1, -1], columns = [1, 2]):
        cell_list = self.get_cell_list(rows, columns)
        cell_table = []
        i = 0
        cur_col = 0
        col_count = columns[1] - columns[0] + 1
        while i < len(cell_list):
            cell_table.append([])
            for j in range(col_count):
                cell_table[cur_col].append(cell_list[i])
                i += 1
            cur_col += 1
        return cell_table

 
    def get_mails_by_school(self):
        cell_table = self.get_cell_table(columns = [1,3])
        schools = {}
        print("insert numbers/names of schools:")
        for name in input().split(', '):
            schools[name.upper()] = 0
            
        answer = []
        i = 0
        for i in range(len(cell_table)):
            if cell_table[i][0].value in schools:
                schools[cell_table[i][0].value] += 1
                answer.append(cell_table[i][2].value)
                if schools[cell_table[i][0].value] > 1:
                    print("more than 1 row for 1 school", cell_table[i][0].value)
        for ans in answer:
            print(ans, end= ', ')
        for sch, item in schools.items():
            if item == 0:
                print("no mail for school", sch)
    
    def detect_copy(self):
        cell_table = self.get_cell_table(columns = [1,3])
        mails = {}
        for i in range(len(cell_table)):
            if cell_table[i][0] != '' and cell_table[i][1] != '':
                mails_list = cell_table[i][2].split(', ')
                for mail in mails_list:
                    if mail in mails.keys():
                        print("in row {} detected repeat with row {}, mail {}".format(i + 1, mails[mail], mail))

def get_mails_by_school():
    gt = google_tables(wks_name="участники семинаров")
    gt.get_mails_by_school()

if __name__ == "__main__":
    get_mails_by_school()
