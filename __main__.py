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
    def __init__(self, docs_name="Рассылки", wks_name="Координаторы", gc_prev=None):
        if gc_prev == None:
            scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
            credentials = ServiceAccountCredentials.from_json_keyfile_name('..//..//keys//math vertical-c8b635a4089b.json', scope)
            self.gc = gspread.authorize(credentials)
        else:
            self.gc = gc_prev
        self.wks = self.gc.open(docs_name)
        self.sheet = self.wks.worksheet(wks_name)
        self.row_count = self.sheet.row_count
        print(self.row_count)

    def get_cell_list(self, rows = (1, -1), columns = (1, 2)):
        #numeration starts with 1 as like in google tables  
        if rows[1] == -1:
            rows = (rows[0], self.row_count)
        cell_string = column_letter(columns[0]) + str(rows[0]) + ':' + column_letter(columns[1]) + str(rows[1])
        cell_table = self.sheet.range(cell_string)
        return cell_table

    def get_cell_table(self, rows = (1, -1), columns = (1, 2)):
        cell_list = self.get_cell_list(rows=rows, columns=columns)
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
        cell_table = self.get_cell_table(columns = (1,3))
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
            if cell_table[i][0].value != '' and cell_table[i][1].value != '' and cell_table[i][2].value != '':
                mails_list = cell_table[i][2].value.split(', ')
                for mail in mails_list:
                    if mail in mails.keys():
                        print("in row {} detected repeat with row {}, mail {}, sheet {}".format(
                            i + 1, mails[mail], mail, self.sheet.title))
                    else:
                        mails[mail] = i + 1

def get_mails_by_school():
    gt = google_tables(wks_name="участники семинаров")
    gt.get_mails_by_school()

def check_copies():
    gc = None
    for sheet in sheets:
        gt = google_tables(gc_prev = gc, wks_name=sheet)
        gt.detect_copy()
        gc = gt.gc

if __name__ == "__main__":
    #get_mails_by_school()
    check_copies()