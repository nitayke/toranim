import openpyxl
from tkinter import *
from tkinter import ttk
from docx import Document, enum

FOLDER = 'C:/Users/user/Desktop/toranim/'
TABLE_SIZE = 9 # x9
# shishi, regular
TABLE_CELLS = '''13, 14, 22, 23
40, 41, 49, 50

9, 10, 18, 19
11, 12, 20, 21
15, 16, 17, 24, 25
36, 37, 45, 46
38, 39, 47, 48
42, 43, 44, 51, 52
63, 64, 72, 73
65, 66, 74, 75
67, 68, 76, 77'''.split('\n')


class Excel:
    def __init__(self, filename):
        wb = openpyxl.load_workbook(FOLDER + filename)
        self.ws = wb.active
        self.toranim_data = {}
        self.havruta_data = []
    
    def extract(self):
        for i in self.ws['A2:D' + str(self.ws.max_row)]:
            self.toranim_data[i[0].value] = [i[2].value, i[3].value]
        
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


class Tkinter:
    root = Tk()
    @classmethod
    def start(cls):
        cls.root.title('שיבוץ תורנים')
        frame = Frame(cls.root, padx=60, pady=60)
        frame.pack()
        c = Calculate()
        ttk.Button(frame, text='סבב רגיל', command=lambda: c.regular_sevev()).pack()
        ttk.Button(frame, text='סבב מיוחד', command=lambda: c.special_sevev()).pack()
        cls.root.mainloop()

    @classmethod
    def remove_frame(cls):
        cls.root.winfo_children()[0].destroy()

    @classmethod
    def add_special_sevev(cls):
        frame = Frame(cls.root, padx=60, pady=60)
        frame.pack()
        label = Label(frame, text='תורנות רגילה:')
        label.pack()
        textvar = StringVar(frame)
        textvar.set('38')
        spinbox = Spinbox(frame, from_=0, to=50,
        textvariable=textvar)
        spinbox.pack()
        label = Label(frame, text='תורנות שישי:')
        label.pack()
        textvar1 = StringVar(frame)
        textvar1.set('8')
        spinbox = Spinbox(frame, from_=0, to=50,
        textvariable=textvar1)
        spinbox.pack()
        c = Calculate()
        Button(frame, text='סבבה',
            command=lambda: c.calculate(
            int(textvar.get()), int(textvar1.get()))).pack()

    @classmethod
    def close(cls):
        cls.root.destroy()


class Calculate:
    def __init__(self):
        self.excel = Excel('תורנים.xlsx')
        self.havrutot, self.toranim = self.excel.extract()
        self.regular_result, self.shishi_result = [[], []], [[], []]
        self.regular_count, self.shishi_count = 0, 0

    def get_united_min_list(self):
        res = []
        min_shishi = min([i[1] for i in self.toranim.values()])
        min_toranut = min([i[0] for i in self.toranim.values()])
        for i in self.toranim:
            if self.toranim[i][0] == min_toranut and self.toranim[i][1] == min_shishi:
                res.append(i)
        return res

    def get_regular_min_list(self):
        res = []
        min_toranut = min([i[0] for i in self.toranim.values()]) # TODO: improve double
        for i in self.toranim:
            if self.toranim[i][0] == min_toranut:
                res.append(i)
        return res

    def add_regular(self, name, has_havruta):
        self.regular_result[has_havruta].append(name)
        self.regular_count -= 1
        self.toranim[name][0] += 1

    def add_shishi(self, name, has_havruta):
        self.shishi_result[has_havruta].append(name)
        self.shishi_count -= 1
        self.toranim[name][0] += 1
        self.toranim[name][1] += 1
    
    def add_last_toran(self, shishi):
        if shishi:
            list = self.get_united_min_list()
            func = self.add_shishi
        else:
            list = self.get_regular_min_list()
            func = self.add_regular
        for i in list.copy():
            if self.excel.get_havruta(i) == '':
                func(i, False)
                return True
        return False

    def calculate(self, regular_count, shishi_count):
        self.regular_count, self.shishi_count = regular_count, shishi_count
        self.united_min_list = self.get_united_min_list()  # toranut and shishi has minimum values
        
        if len(self.united_min_list) <= self.shishi_count:
            self.shishi_result[0] = self.united_min_list
            self.shishi_count -= len(self.united_min_list)  # need to add more to fill the required

        for i in self.united_min_list:
            if self.shishi_count == 1 and not self.add_last_toran(True):
                self.add_shishi(i, False)
            if self.shishi_count == 0:
                break
            current_havruta = self.excel.get_havruta(i)
            if current_havruta in self.united_min_list:
                self.add_shishi(current_havruta, True)
                self.add_shishi(i, True)
            else:
                self.add_shishi(i, False)

        self.regular_min_list = self.get_regular_min_list()

        if len(self.regular_min_list) <= self.regular_count:
            self.regular_result[0] = self.regular_min_list
            self.regular_count -= len(self.regular_min_list)


        for i in self.regular_min_list:
            if regular_count == 1 and not self.add_last_toran(False):
                self.add_regular(i, False)
            if not regular_count:
                break
            current_havruta = self.excel.get_havruta(i)
            if current_havruta in self.regular_min_list:
                self.add_regular(current_havruta, True)
                self.add_regular(i, True)
            else:
                self.add_regular(i, False)


        self.put_results()

        Tkinter.close()

    def put_results(self):
        word = Word(self.shishi_count, self.regular_count)
        word.fill_table(self.shishi_result, self.regular_result)

    def regular_sevev(self):
        self.calculate(38, 8)

    def special_sevev(self):
        Tkinter.remove_frame()
        Tkinter.add_special_sevev()

class Word:
    def __init__(self, shishi_count, regular_count):
        self.template_doc = Document(FOLDER + 'template.docx')
        self.table = self.template_doc.tables[0]
        self.result = Document(FOLDER+'תוצאה.docx')
        self.shishi_count, self.regular_count = shishi_count, regular_count
        self.table_cells = [[], []] # shishi, regular
        index = 0
        for i in TABLE_CELLS:
            if i == '':
                index = 1
                continue
            self.table_cells[index].append([int(i) for i in i.split(', ')])
        print(self.table_cells)

    def fill_table(self, shishi_result, regular_result):
        counts = [0, 0, 0, 0] # shishi [lonely, havruta] regular [lonely, havruta]
        # for i in range(self.shishi_count):
        for i in self.table_cells: # 2 iterations - shishi and regular
            for j in i: # over groups
                for k in j:
                    if counts[1] + 1 < len(shishi_result[1]):
                        self.table.rows[k // TABLE_SIZE].cells[k % TABLE_SIZE].text = shishi_result[1][counts[1]]
                    counts[1] += 1
        # for i in range(2): # shishi
        #     for k in range(2):
        #         if counts[1]+1 < len(shishi_result[1]): # can be shorter
        #             self.table.rows[i*3 + k + 1].cells[4].text = shishi_result[1][counts[1]]
        #             self.table.rows[i*3 + k + 1].cells[5].text = shishi_result[1][counts[1]+1]
        #             counts[1] += 2
        #         else:
        #             self.table.rows[i*3 + k + 1].cells[4].text = shishi_result[0][counts[0]]
        #             self.table.rows[i*3 + k + 1].cells[5].text = shishi_result[0][counts[0]+1]
        #             counts[0] += 2
        
        # for i in range(3): # every week
        #     for j in (0, 2): # rishon and shlishi
        #         for k in range(2): # every line in week
        #             if counts[3]+1 < len(regular_result[1]):
        #                 self.table.rows[i*3 + k + 1].cells[j].text = regular_result[1][counts[3]]
        #                 self.table.rows[i*3 + k + 1].cells[j+1].text = regular_result[1][counts[3]+1]
        #                 counts[3] += 2
        #             else:
        #                 self.table.rows[i*3 + k + 1].cells[j].text = regular_result[0][counts[2]]
        #                 self.table.rows[i*3 + k + 1].cells[j+1].text = regular_result[0][counts[2]+1]
        #                 counts[2] += 2

        # for i in range(2): # shabbes
        #     for j in range(3):
        #         if counts[3]+1 < len(regular_result[1]):
        #             self.table.rows[i*3+1].cells[6+j].text = regular_result[1][counts[3]+j]
        #             counts[3] += 1
        #         else:
        #             self.table.rows[i*3+1].cells[6+j].text = regular_result[1][counts[2]+j]
        #             counts[2] += 1
        #     for j in range(2):
        #         if counts[3]+1 < len(regular_result[1]):
        #             self.table.rows[i*3+2].cells[6+j].text = regular_result[1][counts[3]+j]
        #             counts[3] += 1
        #         else:
        #             self.table.rows[i*3+2].cells[6+j].text = regular_result[1][counts[2]+j]
        #             counts[2] += 1

                # for j in range(len(regular_result[1]) - counts[3]):
                #     self.table.rows[i*3 + k + 1].cells[6].text = regular_result[1][counts[3]+j]
                #     if counts[3]+1 < len(regular_result[1]): # can be shorter
                #         self.table.rows[i*3 + k + 1].cells[6].text = regular_result[1][counts[3]]
                #         self.table.rows[i*3 + k + 1].cells[7].text = regular_result[1][counts[3]+1]
                #         counts[3] += 2
                #     else:
                #         self.table.rows[i*3 + k + 1].cells[6].text = regular_result[0][counts[2]]
                #         self.table.rows[i*3 + k + 1].cells[7].text = regular_result[0][counts[2]+1]
                #         counts[2] += 2
        
        self.result.element.body.append(self.template_doc.element.body[0])
        self.result.save(FOLDER+'תוצאה.docx')




if __name__ == '__main__':
    Tkinter.start()