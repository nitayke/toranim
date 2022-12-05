import openpyxl
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as mb
from docx import Document
from docx.shared import Pt
from datetime import datetime
import os
from glob import glob

VERSIONS_FOLDER = 'xl_versions'
RESULTS_FOLDER = 'תוצאות'
TEMPLATE_FILE = 'template.docx'
XL_NAME = 'תורנים.xlsx'
TABLE_SIZE = 9 # x 9
SHISHI_INDEX = (1, 4)
# regular, shishi
TABLE_CELLS = '''15 16 17 24 25
11 12 20 21
42 43 44 51 52
38 39 47 48
36 37 45 46
67 68 76 77
65 66 74 75
63 64 72 73

13 14 22 23
40 41 49 50'''.split('\n')

REGULAR = MITUTA = INFO = 0
SHISHI = HAVRUTA = ERROR = 1


class Excel:
    def __init__(self):
        self.wb = openpyxl.load_workbook(XL_NAME)
        self.ws = self.wb.active
        self.toranim_data = {} # {name: [all, shishi]}
        self.havruta_data = []
    
    def extract(self):
        for i in self.ws['A2:D' + str(self.ws.max_row)]:
            self.toranim_data[i[0].value] = [i[2].value, i[3].value] # regular and shishi
        
        for i in self.ws['F2:G' + str(self.ws.max_row // 2)]:
            if not i[0].value:
                break
            self.havruta_data.append([i[0].value, i[1].value])
        return self.havruta_data, self.toranim_data

    def get_havruta(self, name):
        for i in self.havruta_data:
            if name == i[0]:
                return i[1]
            elif name == i[1]:
                return i[0]
        return ''
    
    def update(self, toranim_data):
        if not os.path.exists(VERSIONS_FOLDER):
            os.mkdir(VERSIONS_FOLDER)
        self.wb.save(f'{VERSIONS_FOLDER}/{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.xlsx')
        for i, name in enumerate(toranim_data):
            self.ws.cell(i+2, 3).value = toranim_data[name][REGULAR]
            self.ws.cell(i+2, 4).value = toranim_data[name][SHISHI]
        self.wb.save(XL_NAME)


class Tkinter:
    strvar_nums = []

    @classmethod
    def start(cls):
        cls.root = tk.Tk()
        cls.root.title('שיבוץ תורנים')
        frame = tk.Frame(cls.root, padx=60, pady=60)
        frame.pack()
        ttk.Button(frame, text='סבב רגיל', command=lambda: Calculate().calculate()).pack()
        ttk.Button(frame, text='סבב מיוחד', command=cls.special_sevev).pack()
        ttk.Button(frame, text='שחזר פעם אחרונה', command=cls.restore).pack()
        cls.root.mainloop()

    @classmethod
    def remove_frame(cls):
        cls.root.winfo_children()[0].destroy()

    @classmethod
    def get_int_counts(cls):
        return [int(i.get()) for i in cls.strvar_nums if i]
    
    @classmethod
    def special_sevev(cls):
        cls.remove_frame()
        frame = tk.Frame(cls.root, padx=60, pady=60)
        frame.pack()

        DAYS = ['ראשון-שני', 'שלישי-רביעי', 'חמישי-שישי', 'שבת']
        cls.strvar_nums = [0] * 12
        for i in range(4):
            tk.Label(frame,text=DAYS[3-i]).grid(row=0, column=i)
            for j in range(3):
                if i == 0 and j == 2 or i == 3 and j == 0:
                    continue
                frame1 = tk.Frame(frame)
                frame1.grid(row=j+1, column=i)
                cls.strvar_nums[j*4+i] = tk.StringVar()
                cls.strvar_nums[j*4+i].set(str(4 + int(i==0)))
                spinbox=tk.Spinbox(frame1, from_=0, to=50, textvariable=cls.strvar_nums[j*4+i])
                spinbox.pack()
    
        ttk.Button(cls.root, text='סבבה', command=lambda:
         Calculate().calculate()).pack()
    
    @classmethod
    def get_sums(cls):
        if cls.strvar_nums: # special sevev
            int_nums = cls.get_int_counts()
            shishis_sum = sum([int_nums[i] for i in SHISHI_INDEX])
            rest_sum = sum(int_nums) - shishis_sum
            return [rest_sum, shishis_sum]
        else:
            return [34, 8]

    @classmethod
    def close(cls):
        cls.root.destroy()

    @classmethod
    def show(cls, type, msg):
        if type == ERROR:
            mb.showerror('קרתה תקלה', msg)
        else:
            mb.showinfo('כל הכבוד', msg)

    @classmethod
    def restore(cls):
        xl_folder, word_foler = glob(VERSIONS_FOLDER + '/*'), glob(RESULTS_FOLDER + '/*')
        if not xl_folder or not word_foler:
            cls.show(ERROR, "אין לך מה לשחזר")
        else:
            os.remove(XL_NAME)
            os.remove(max(word_foler, key=os.path.getctime)) # latest file in folder
            os.rename(max(xl_folder, key=os.path.getctime), XL_NAME)
            cls.show(INFO, 'שוחזר בהצלחה')

class Calculate:
    def __init__(self):
        self.excel = Excel()
        self.havrutot, self.toranim = self.excel.extract()
        self.results = [[[], []], [[], []]] # [regular: [lonely, havruta], shishi: [lonely, havruta]]
        self.count = [0, 0] # regular and shishi
        self.min_lists = [[], []] # regular and shishi

    def get_min_list(self, type):
        res = []
        min_shishi = min([i[SHISHI] for i in self.toranim.values()])
        min_regular = min([i[REGULAR] for i in self.toranim.values()])
        for i in self.toranim:
            if type == SHISHI:
                if self.toranim[i][REGULAR] == min_regular and self.toranim[i][SHISHI] == min_shishi:
                    res.append(i)
            else:
                if self.toranim[i][REGULAR] == min_regular:
                    res.append(i)
        if not res: # if there's no one in united min list
            for i in self.toranim:
                if self.toranim[i][SHISHI] == min_shishi:
                    res.append(i)
        return res

    def add(self, name, has_havruta, type):
        self.results[type][has_havruta].append(name)
        self.count[type] -= 1
        self.toranim[name][REGULAR] += 1
        if type == SHISHI:
            self.toranim[name][SHISHI] += 1
        self.min_lists[type].remove(name)

    def add_last_toran(self, type):
        for i in self.get_min_list(type):
            if self.excel.get_havruta(i) == '':
                self.add(i, MITUTA, type)
                return True
        return False

    def util1(self, type):
        name = self.min_lists[type][0]
        if self.count[type] == 1:
            self.add_last_toran(type)
            return
        havruta = self.excel.get_havruta(name)
        if havruta in self.min_lists[type]:
            self.add(havruta, HAVRUTA, type)
            self.add(name, HAVRUTA, type)
        else:
            self.add(name, MITUTA, type)

    def get_odd_count(self):
        if Tkinter.strvar_nums:
            nums = Tkinter.get_int_counts()
            self.odds_count = sum([i % 2 for i in nums])
        else:
            self.odds_count = 2

    def util(self, type):
        self.min_lists[type] = self.get_min_list(type)
        if len(self.min_lists[type]) < self.count[type]:
            for i in self.min_lists[type]:
                self.util1(type)
            self.min_lists[type] = self.get_min_list(type)
        while self.count[type] > 0:
            self.util1(type)
            # odds_count = days that need a single toran
            if self.count[type] and len(self.results[type][MITUTA]) < self.odds_count:
                self.add_last_toran(type)
            self.min_lists[type] = self.get_min_list(type)

    def calculate(self):
        self.count = Tkinter.get_sums()
        self.get_odd_count()
        self.util(SHISHI)
        self.util(REGULAR)
        self.save_results()
        Tkinter.show(INFO,
        'התוצאה שמורה ב"תוצאות" והקובץ "תורנים" התעדכן. אם אתה רוצה להתחרט אז תיכנס שוב לתוכנה ותלחץ על "שחזר".')
        Tkinter.close()

    def save_results(self):
        word = Word()
        word.update_table_cells()
        word.fill_table(self.results)
        word.save()
        self.excel.update(self.toranim)


class Word:
    def __init__(self):
        self.doc = Document('template.docx')
        self.table = self.doc.tables[0]

    def update_table_cells(self):
        self.table_cells = [[], []] # regular, shishi
        index = REGULAR
        for i in TABLE_CELLS:
            if i == '':
                index = SHISHI
                continue
            self.table_cells[index].append([int(i) for i in i.split(' ')])
        if Tkinter.strvar_nums: # special sevev
            int_counts = Tkinter.get_int_counts()
            shishi_counts = [int_counts[i] for i in SHISHI_INDEX]
            rest_counts = []
            for key, value in enumerate(int_counts):
                if key not in SHISHI_INDEX:
                    rest_counts.append(value)
            counts = [rest_counts, shishi_counts]
            for i in range(len(self.table_cells)): # regular and shishi
                for j in range(len(self.table_cells[i])): # groups
                    if counts[i][j] < len(self.table_cells[i][j]):
                        self.table_cells[i][j] = self.table_cells[i][j][:counts[i][j]]
                    else:
                        self.table_cells[i][j] = (self.table_cells[i][j] * (counts[i][j] //
                         len(self.table_cells[i][j]) + 1))[:counts[i][j]]

    def fill_table(self, results):
        counts = [[0, 0], [0, 0]] # regular [lonely, havruta], shishi [lonely, havruta]
        for i in range(len(self.table_cells)): # 2 iterations - regular and shishi
            for j in self.table_cells[i]: # over groups.
                flag = False
                for iteration, k in enumerate(j):
                    cell = self.table.rows[k // TABLE_SIZE].cells[k % TABLE_SIZE]
                    if cell.text:
                        cell.text += ',\n'
                    # flag is True when one of havruta was inserted
                    if ( counts[i][HAVRUTA] < len(results[i][HAVRUTA]) and iteration < len(j) - 1 ) or flag:
                        cell.text += results[i][HAVRUTA][counts[i][HAVRUTA]]
                        flag = not flag
                        counts[i][HAVRUTA] += 1
                    else:
                        cell.text += results[i][MITUTA][counts[i][MITUTA]]
                        counts[i][MITUTA] += 1
                    cell.paragraphs[0].runs[0].font.size = Pt(12)

    def save(self):
        if not os.path.exists(RESULTS_FOLDER):
            os.mkdir(RESULTS_FOLDER)
        self.doc.save(f'{RESULTS_FOLDER}/{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.docx')


if __name__ == '__main__':
    if not os.path.exists(TEMPLATE_FILE) or not os.path.exists(XL_NAME):
        Tkinter.show(ERROR, 'חסרים לך קבצים (צריך שיהיה לך: template.docx, תורנים.xlsx)')
    else:
        Tkinter.start()
