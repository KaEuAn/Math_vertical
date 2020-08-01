#coding=utf-8
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from collections import Iterable

def column_letter(column_number):
    return chr(ord('A')  - 1 + column_number)


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
    
    def get_input(self):
        dic = {}
        print("insert numbers/names of schools:")
        for name in input().split(', '):
            dic[name.upper()] = 0        
        return dic

 
    def get_mails_by_school(self):
        cell_table = self.get_cell_table(columns = (1,3))
        schools = self.get_input()
            
        answer = []
        i = 0
        for i in range(len(cell_table)):
            current_school = cell_table[i][0].value.upper()
            if current_school in schools:
                schools[current_school] += 1
                answer.append(cell_table[i][2].value)
                if schools[current_school] > 1:
                    print("more than 1 row for 1 school", cell_table[i][0].value)
        for ans in answer:
            print(ans, end= ', ')
        print()
        for sch, item in schools.items():
            if item == 0:
                print("no mail for school", sch)
    
    def detect_copy(self, working_columns = [1,3]):
        cell_table = self.get_cell_table(columns = working_columns)
        mails = {}
        for i in range(len(cell_table)):
            if cell_table[i][0].value != '' and cell_table[i][1].value != '' and cell_table[i][2].value != '':
                mails_list = cell_table[i][2].value.split(', ')
                for mail in mails_list:
                    if mail.upper() in mails.keys():
                        print("in row {} detected repeat with row {}, mail {}, sheet {}".format(
                            i + 1, mails[mail.upper()], mail, self.sheet.title))
                    else:
                        mails[mail.upper()] = i + 1
                        
    def detect_if_in_table(self):
        new_mails = self.get_input()
        cell_table = self.get_cell_table(columns = [1,3])
        mails = set()
        for i in range(len(cell_table)):
            if cell_table[i][0].value != '' and cell_table[i][1].value != '' and cell_table[i][2].value != '':
                mails_list = cell_table[i][2].value.split(', ')
                for mail in mails_list:
                    mails.add(mail)
        for new_mail in new_mails:
            print(new_mail in mails, new_mail)
            

class MyTables:

    def __init__(self, sheets=None, doc_name=None):
        if sheets is None:
            fi = open("sheets.json", "r", encoding='utf-8')
            self.sheets = json.load(fi)
            fi.close()
        else:
            self.sheets = sheets
            
        if doc_name is None:
            self.doc_name = "Рассылки"
        else:
            self.doc_name = doc_name
            
    
    def check_copies_summer(self):
        pass
            
    
    def get_mails_by_school(self, name="Координаторы"):
        gt = google_tables(docs_name = self.doc_name, wks_name=name)
        gt.get_mails_by_school()
        
    def do_for_all_sheets(self, function, *args, **kwargs):
        gc = None
        for sheet in self.sheets:
            gt = google_tables(docs_name = self.doc_name, gc_prev = gc, wks_name=sheet)
            function(gt, *args, **kwargs)
            gc = gt.gc    
    
    def check_copies(self):
        return self.do_for_all_sheets(google_tables.detect_copy)
            
    def check_detection(self):
        gt = google_tables(wks_name="участники семинаров")
        gt.detect_if_in_table()
        
    def check_kruzki(self):
        gt = google_tables(wks_name='Кружки')
        gt.detect_copy([3,5])
    
if __name__ == "__main__":
    table = MyTables(["Лист1"], "Ведомость")
    #check_kruzki()
    table.check_copies()
    #check_copies()