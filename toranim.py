import openpyxl
import tkinter as tk
from tkinter import ttk
from docx import Document
from datetime import datetime

FOLDER = 'C:/Users/user/Desktop/toranim/'
VERSIONS_FOLDER = 'versions'
XL_NAME = 'תורנים.xlsx'
TABLE_SIZE = 9 # x9
# regular, shishi
TABLE_CELLS = '''9, 10, 18, 19
11, 12, 20, 21
15, 16, 17, 24, 25
36, 37, 45, 46
38, 39, 47, 48
42, 43, 44, 51, 52
63, 64, 72, 73
65, 66, 74, 75
67, 68, 76, 77

13, 14, 22, 23
40, 41, 49, 50'''.split('\n')

REGULAR = MITUTA = 0
SHISHI = HAVRUTA = 1


class Excel:
    def __init__(self):
        self.wb = openpyxl.load_workbook(FOLDER + XL_NAME)
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
        self.wb.save(f'{FOLDER}{VERSIONS_FOLDER}/{datetime.now().strftime("%m-%d-%Y_%H-%M-%S")}.xlsx')
        for i, name in enumerate(toranim_data):
            self.ws.cell(i+2, 3).value = toranim_data[name][REGULAR]
            self.ws.cell(i+2, 4).value = toranim_data[name][SHISHI]
        self.wb.save(FOLDER + XL_NAME)


class Tkinter:
    root = tk.Tk()
    @classmethod
    def start(cls):
        cls.root.title('שיבוץ תורנים')
        frame = tk.Frame(cls.root, padx=60, pady=60)
        frame.pack()
        c = Calculate()
        ttk.Button(frame, text='סבב רגיל', command=lambda: c.regular_sevev()).pack()
        # ttk.Button(frame, text='סבב מיוחד', command=lambda: c.special_sevev()).pack()
        cls.root.mainloop()

    @classmethod
    def remove_frame(cls):
        cls.root.winfo_children()[0].destroy()

    @classmethod
    def add_special_sevev(cls):
        frame = tk.Frame(cls.root, padx=60, pady=60)
        frame.pack()
        label = tk.Label(frame, text='תורנות רגילה:')
        label.pack()
        textvar = tk.StringVar(frame)
        textvar.set('38')
        spinbox = ttk.Spinbox(frame, from_=0, to=50,
        textvariable=textvar)
        spinbox.pack()
        label = tk.Label(frame, text='תורנות שישי:')
        label.pack()
        textvar1 = tk.StringVar(frame)
        textvar1.set('8')
        spinbox = ttk.Spinbox(frame, from_=0, to=50,
        textvariable=textvar1)
        spinbox.pack()
        c = Calculate()
        ttk.Button(frame, text='סבבה',
            command=lambda: c.calculate(
            int(textvar.get()), int(textvar1.get()))).pack()

    @classmethod
    def close(cls):
        cls.root.destroy()


class Calculate:
    def __init__(self):
        self.excel = Excel()
        self.havrutot, self.toranim = self.excel.extract()
        self.results = [[[], []], [[], []]] # [regular: [lonely, havruta], shishi: [lonely, havruta]]
        self.count = [0, 0] # regular and shishi

    def get_united_min_list(self):
        res = []
        min_shishi = min([i[SHISHI] for i in self.toranim.values()])
        min_toranut = min([i[REGULAR] for i in self.toranim.values()])
        for i in self.toranim:
            if self.toranim[i][REGULAR] == min_toranut and self.toranim[i][SHISHI] == min_shishi:
                res.append(i)
        return res

    def get_regular_min_list(self):
        res = []
        min_toranut = min([i[REGULAR] for i in self.toranim.values()]) # TODO: improve double
        for i in self.toranim:
            if self.toranim[i][REGULAR] == min_toranut:
                res.append(i)
        return res

    def add(self, name, has_havruta, shishi):
        self.results[shishi][has_havruta].append(name)
        self.count[shishi] -= 1
        self.toranim[name][REGULAR] += 1
        if shishi:
            self.toranim[name][SHISHI] += 1
    
    def add_last_toran(self, shishi):
        if shishi:
            list = self.get_united_min_list()
        else:
            list = self.get_regular_min_list()
        for i in list.copy(): # TODO: remove copy
            if self.excel.get_havruta(i) == '':
                self.add(i, False, shishi)
                return True
        return False

    def util(self, shishi):
        if len(self.min_lists[shishi]) <= self.count[shishi]:
            self.results[shishi][MITUTA] = self.min_lists[shishi] # TODO: MAYBE NOT SUPPOSED TO BE LONELY
            self.count[shishi] -= len(self.min_lists[shishi])  # need to add more to fill the required
            for i in self.min_lists[shishi]:
                self.toranim[i][shishi] += 1

        for i in self.min_lists[shishi]:
            if self.count[shishi] == 1 and not self.add_last_toran(shishi):
                self.add(i, False, shishi)
            if self.count[shishi] == 0:
                break
            current_havruta = self.excel.get_havruta(i)
            if current_havruta in self.min_lists[shishi]:
                self.add(current_havruta, True, shishi)
                self.add(i, True, shishi)
            else:
                self.add(i, False, shishi)

    def calculate(self, regular_count, shishi_count):
        self.count = [shishi_count, regular_count]
        self.min_lists = [[], self.get_united_min_list()] # regular and united
        # united is where the regular and shishi have minimum values
        self.util(SHISHI)
        self.min_lists[0] = self.get_regular_min_list()
        self.util(REGULAR)
        self.put_results()
        self.excel.update(self.toranim)
        Tkinter.close()

    def put_results(self):
        word = Word(self.count)
        word.fill_table(self.results)
        word.update()

    def regular_sevev(self):
        self.calculate(38, 8)

    def special_sevev(self):
        Tkinter.remove_frame()
        Tkinter.add_special_sevev()


class Word:
    def __init__(self, count):
        self.template_doc = Document(FOLDER + 'template.docx')
        self.table = self.template_doc.tables[0]
        self.result = Document(FOLDER + 'תוצאה.docx')
        self.regular_count, self.shishi_count = count
        self.table_cells = [[], []] # regular, shishi
        index = 0
        for i in TABLE_CELLS:
            if i == '':
                index = 1
                continue
            self.table_cells[index].append([int(i) for i in i.split(', ')])

    def fill_table(self, results):
        counts = [[0, 0], [0, 0]] # regular [lonely, havruta], shishi [lonely, havruta]

        lengths = ((len(results[REGULAR][MITUTA]), len(results[REGULAR][HAVRUTA])),
            ((len(results[SHISHI][MITUTA]), len(results[SHISHI][HAVRUTA]))))
        
        for i in range(len(self.table_cells)): # 2 iterations - regular and shishi
            for j in self.table_cells[1-i]: # over groups. I want to put shishi first
                flag = False
                counter = 0
                for k in j:
                    if counts[i][HAVRUTA] < lengths[i][HAVRUTA] and counter < len(j) - 1 or flag:
                        self.table.rows[k // TABLE_SIZE].cells[k % TABLE_SIZE].text = results[i][HAVRUTA][counts[i][HAVRUTA]]
                        flag = not flag
                        counts[i][HAVRUTA] += 1
                    else:
                        self.table.rows[k // TABLE_SIZE].cells[k % TABLE_SIZE].text = results[i][MITUTA][counts[i][MITUTA]]
                        counts[i][MITUTA] += 1
                    counter += 1

    def update(self):
        self.result.add_paragraph(datetime.now().strftime('%m/%d/%Y, %H:%M:%S'))
        self.result.element.body.append(self.template_doc.element.body[0])
        self.result.save(FOLDER+'תוצאה.docx')



if __name__ == '__main__':
    Tkinter.start()