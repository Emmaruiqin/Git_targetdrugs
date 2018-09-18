from tkinter import *
import tkinter.filedialog as tkfd
import os
import chemotarget_analysis_report_20180914

class reportanalysis():
    def __init__(self, root):
        self.root = root
        self.root.title = '个体化项目（化疗、靶向）报告分析界面'
        self.root.geometry('550x350')
        self.framework()

    def framework(self):
        self.file_opt = options ={}
        options['defaultextension'] = '.xlsm'
        options['filetypes'] = [('all files', '*'), ('xls file', '.xls')]
        options['initialdir'] = 'E:\\化疗套餐报告自动化'
        options['multiple'] = True
        options['parent'] = root
        options['title'] = ' 选择文件'

        pcr_label = Label(self.root, text='输入结果文件', bg='#87CEEB', font='Arial', fg='white', width=15, height=2)
        pcr_label.grid(row=1, column=0)

        pcr_text = Text(self.root, height=4, width=20)
        pcr_text.grid(row=1, column=1)

        def import_file():
            filelist = tkfd.askopenfilenames(**self.file_opt)
            global files
            files = [i.split('/')[-1] for i in filelist]
            filestuple = tuple(files)
            pcr_text.insert(INSERT, filestuple)

        pcr_button = Button(self.root, text='选择文件', command=import_file, bg='#87CEEB', font='Arial', fg='white', width=8,
                            height=2)
        pcr_button.grid(row=1, column=2)

        def cmd_analysis():
            resulfiles = chemotarget_analysis_report_20180914.main(Expresultfiles=files)

        result_button = Button(root, text='提交', command=cmd_analysis, bg='red', font='Arial', width=15, height=2)
        result_button.grid(row=4, column=1)


if __name__ == '__main__':
    root = Tk()
    app = reportanalysis(root)
    root.mainloop()