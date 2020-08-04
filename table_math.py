#coding=utf-8
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from collections import Iterable
import time

def column_letter(column_number):
    if column_number == -1:
        return 'ZZZZZZ'
    
    res = ''
    if column_number > 26:
        nnum = column_number + 1
        nnum //= 26
        res += chr(ord('A')  - 1 + nnum)
        column_number -= nnum * 26     
    res += chr(ord('A')  - 1 + column_number)
    return res


class google_tables:
    def __init__(self, docs_name="Рассылки", wks_name="Координаторы", gc_prev=None):
        if gc_prev == None:
            scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
            credentials = ServiceAccountCredentials.from_json_keyfile_name('..//..//keys//math vertical-c8b635a4089b.json', scope)
            self.gc = gspread.authorize(credentials)
        else:
            self.gc = gc_prev
        print("open doc {}".format(docs_name))
        self.wks = self.gc.open(docs_name)
        self.sheet = self.wks.worksheet(wks_name)
        self.row_count = self.sheet.row_count
        print(self.row_count)
        
    def update_sheet_column(self, column, values, start_row):
        end_row = start_row + len(values) - 1
        if isinstance(column, int):
            column = column_letter(column)
        st = column + str(start_row) + ':' + column + str(end_row)
        values = list(map(list, values))
        self.sheet.update(st, values)

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
                cell_table[cur_col].append(cell_list[i].value)
                i += 1
            cur_col += 1
        self.table = cell_table
        return cell_table
    
    def get_input(self):
        dic = {}
        print("insert numbers/names of schools:")
        for name in input().split(', '):
            dic[name.upper()] = 0        
        return dic

 
    def get_mails_by_school(self):
        self.get_cell_table(columns = (1,3))
        schools = self.get_input()
            
        answer = []
        i = 0
        for i in range(len(self.table)):
            current_school = self.table[i][0].upper()
            if current_school in schools:
                schools[current_school] += 1
                answer.append(self.table[i][2])
                if schools[current_school] > 1:
                    print("more than 1 row for 1 school", self.table[i][0])
        for ans in answer:
            print(ans, end= ', ')
        print()
        for sch, item in schools.items():
            if item == 0:
                print("no mail for school", sch)
    
    def detect_copy(self, working_columns = [1,3]):
        self.get_cell_table(columns = working_columns)
        mails = {}
        for i in range(len(self.table)):
            if self.table[i][0] != '' and self.table[i][1] != '' and self.table[i][2] != '':
                mails_list = self.table[i][2].split(', ')
                for mail in mails_list:
                    if mail.upper() in mails.keys():
                        print("in row {} detected repeat with row {}, mail {}, sheet {}".format(
                            i + 1, mails[mail.upper()], mail, self.sheet.title))
                    else:
                        mails[mail.upper()] = i + 1
                        
    def detect_if_in_table(self):
        new_mails = self.get_input()
        self.get_cell_table(columns = [1,3])
        mails = set()
        for i in range(len(self.table)):
            if self.table[i][0] != '' and self.table[i][1] != '' and self.table[i][2] != '':
                mails_list = self.table[i][2].split(', ')
                for mail in mails_list:
                    mails.add(mail)
        for new_mail in new_mails:
            print(new_mail in mails, new_mail)

            
    def get_marks(self, working_columns = [2,4], mails_index = 0, marks_index = 1):
        self.get_cell_table(columns = working_columns)
        mails = {}
        for i in range(1, len(self.table)):
            if self.table[i][mails_index] != '' and self.table[i][marks_index] != '':
                mails_list = self.table[i][mails_index].split(', ')
                for mail in mails_list:
                    Umail = mail.upper()
                    if Umail in mails.keys():
                        print("mail repeat {} with name {}".format(Umail, self.table[i][2]))
                        mails[Umail].add(self.table[i][marks_index])
                    else:
                        mails[Umail] = set()
                        mails[Umail].add(self.table[i][marks_index])
                        
        return mails

class MyTables:

    def __init__(self, sheets=None, doc_name="Рассылки"):
        if sheets is None:
            fi = open("sheets.json", "r", encoding='utf-8')
            self.sheets = json.load(fi)
            fi.close()
        else:
            self.sheets = sheets
            
        self.doc_name = doc_name
            
    def do_for_all_sheets(self, function, *args, **kwargs):
        self.gc = None
        result = []
        for sheet in self.sheets:
            self.gt = google_tables(docs_name = self.doc_name, gc_prev = self.gc, wks_name=sheet)
            result.append(function(self.gt, *args, **kwargs))
            self.gc = self.gt.gc
        return result
        
    def proceed_function(self, function, *args, **kwargs):
        return self.do_for_all_sheets(function, *args, **kwargs)
    
    #can delete all this functions, have one proceed
    def get_mails_by_school(self):
        self.do_for_all_sheets(google_tables.get_mails_by_school)
        
    def check_copies(self):
        return self.do_for_all_sheets(google_tables.detect_copy)
            
    def check_detection(self):
        return self.do_for_all_sheets(google_tables.detect_if_in_table)
        
    def check_kruzki(self):
        self.proceed_function(google_tables.detect_copy, [3,5])
    #old functions end    
        
        
class FormTable(MyTables):
    def __init__(self, doc_name="Рассылки", sheets=["Ответы на форму (1)"]):
        super(FormTable, self).__init__(sheets, doc_name)
        
    def get_marks(self, **params):
        marks_by_mail = self.proceed_function(google_tables.get_marks, **params)[0]
        self.marks_by_mails = {}
        for key, value in marks_by_mail.items():
            self.marks_by_mails[key] = 0
            for val in value:
                self.marks_by_mails[key] = max(int(val.split(' / ')[0]), self.marks_by_mails[key])
                self.max_mark = int(val.split(' / ')[1])               
        
class RecordTable(MyTables):
    def __init__(self, doc_name="Ведомость", sheets=["Лист1"]):
        super(RecordTable, self).__init__(sheets, doc_name)
        
    def get_table(self):
        self.proceed_function(google_tables.get_cell_table, columns=(1, 55))[0]
        self.row_by_mail = {}
        for i in range(4, len(self.gt.table)):
            for j in (2, 3):
                for mail in self.gt.table[i][j].split(', '):
                    mail = mail.upper()
                    if mail not in self.row_by_mail.keys():
                        self.row_by_mail[mail] = i
        
def load_table(column_by_sheet):
    rt = RecordTable()
    rt.get_table()
    for sheet, item in column_by_sheet.items():
        column, params, is_done = item
        if is_done:
            continue
        ft = FormTable(sheet)
        ft.get_marks(**params)
        
        #get net column
        new_column = []
        for i in range(247):
            new_column.append([0])
        new_column[0][0] = ft.max_mark
        for mail, mark in ft.marks_by_mails.items():
            if mail not in rt.row_by_mail.keys():
                print("can't find mail {} with mark {}".format(mail, mark))
            else:
                row = rt.row_by_mail[mail]
                new_column[row - 3][0] = mark  
        rt.gt.update_sheet_column(column, new_column, 4)
        time.sleep(30)
        
def load_simple_columns():
    with open("column_by_sheet.json", "r", encoding='utf-8') as fi:
        column_by_sheet = json.load(fi)
    load_table(column_by_sheet)
    
if __name__ == "__main__":
    load_simple_columns()
    #check_kruzki()
    #check_copies()