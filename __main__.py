#coding=utf-8
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def column_letter(column_number):
    return chr(ord('A')  - 1 + column_number)

class google_tables:
    def __init__(self, docs_name="Рассылки", wks_name="Координаторы", gc_prev=None):
        if gc_prev == None:
            scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
            credentials = ServiceAccountCredentials.from_json_keyfile_name('..//keys//math vertical-c8b635a4089b.json', scope)
            self.gc = gspread.authorize(credentials)
        else:
            self.gc = gc_prev
        self.wks = self.gc.open("Рассылки")
        self.sheet = self.wks.worksheet("Координаторы")
        self.row_count = self.sheet.row_count

    def get_cell_list(self, rows = [1, -1], columns = [1, 2]):
        #numeration starts with 1 as like in google tables  
        if rows[1] == -1:
            rows[1] = self.row_count
        cell_string = column_letter(columns[0]) + rows[0] + column_letter(columns[1]) + rows[1]
        cell_list = self.sheet.range(cell_string)
        return cell_list

    def get_cell_table():
        pass
 
    def get_mails_by_school(self):
        cell_list = self.get_cell_list(columns = [1,3])
        schools = {}
        print("insert numbers/names of schools:")
        for name in input().split(', '):
            schools[name.upper()] = 0
            
        answer = []
        i = 0
        while i < len(cell_list):
            if cell_list[i].value in schools:
                schools[cell_list[i].value] += 1
                answer.append(cell_list[i + 2].value)
                if schools[cell_list[i].value] > 1:
                    print("more than 1 row for 1 school", cell_list[i].value)
            i += 3
        for ans in answer:
            print(ans, end= ', ')
        for sch, item in schools.items():
            if item == 0:
                print("no mail for school", sch)
    
    def remove_copy(self):
        cell_list = self.sheet.range('A1:C' + str(self.row_count))


if __name__ == "__main__":
    gt = google_tables(wks_name="участники семинаров")
    gt.remove_copy()