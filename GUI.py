import os
import tkinter as tk
from tkinter.filedialog import askopenfilename
from DataModelDoc import *
import json


class NewprojectApp:
    def __init__(self, master=None):

        try:
            with open('temp.json') as json_file:
                data = json.load(json_file)
        except:
            data = {'path_source': '',
                    'path_types': '',
                    'new_name': '',
                    'username': ''}
        # build ui
        self.frame1 = tk.Frame(master)
        self.entry1 = tk.Entry(self.frame1)
        self.entry1.configure(width='100')
        self.entry1.grid(column='2', row='1')
        self.entry1.delete('0', 'end')
        self.entry1.insert('0', data['path_source'])
        self.button1 = tk.Button(self.frame1)
        self.button1.configure(text='Browse')
        self.button1.grid(column='1', row='1', sticky='e')
        self.frame1.columnconfigure('1', minsize='150')
        self.button1.configure(command=self.load_csv)
        self.label1 = tk.Label(self.frame1)
        self.label1.configure(anchor='n', compound='left', cursor='arrow', font='TkDefaultFont')
        self.label1.configure(text='CSV file')
        self.label1.grid(column='1', row='1', sticky='w')
        self.entry2 = tk.Entry(self.frame1)
        self.entry2.configure(width='100')
        self.entry2.grid(column='2', row='2')
        self.entry2.delete('0', 'end')
        self.entry2.insert('0', data['path_types'])
        self.frame1.rowconfigure('2', minsize='00')
        self.button2 = tk.Button(self.frame1)
        self.button2.configure(justify='left', takefocus=True, text='Browse')
        self.button2.grid(column='1', row='2', sticky='e')
        self.button2.configure(command=self.load_xml)
        self.label2 = tk.Label(self.frame1)
        self.label2.configure(text='XML data types')
        self.label2.grid(column='1', row='2', sticky='w')
        self.entry3 = tk.Entry(self.frame1)
        self.entry3.configure(width='40')
        self.entry3.grid(column='2', row='3', sticky='w')
        self.entry3.delete('0', 'end')
        self.entry3.insert('0', data['new_name'])
        self.label3 = tk.Label(self.frame1)
        self.label3.configure(text='New File Name')
        self.label3.grid(column='1', row='3', sticky='w')
        self.entry7 = tk.Entry(self.frame1)
        self.entry7.configure(width='40')
        self.entry7.grid(column='2', row='3', sticky='e')
        self.entry7.delete('0', 'end')
        self.entry7.insert('0', data['username'])
        self.label7 = tk.Label(self.frame1)
        self.label7.configure(compound='bottom', justify='left', text='User Name')
        self.label7.grid(column='2', row='3')
        self.button8 = tk.Button(self.frame1)
        self.button8.configure(anchor='n', compound='center', height='3', text='\n R U N')
        self.button8.configure(width='15')
        self.button8.grid(column='2', row='10')
        self.frame1.rowconfigure('10', minsize='100')
        self.button8.configure(command=self.run_script)
        self.frame1.configure(height='200', width='500')
        self.frame1.grid(column='0', row='0')

        # Main widget
        self.mainwindow = self.frame1

    def load_csv(self):
        self.path_source = askopenfilename()
        # _text_ = '''test'''
        self.entry1.delete('0', 'end')
        self.entry1.insert('0', self.path_source)
        # print(self.path_source)

    # def load_filename(self):
    #     self.path_base = askopenfilename()
        # root.filename = tk.StringVar()
        # root.filename.set(askopenfilename())
    def load_xml(self):
        self.types_path = askopenfilename()
        self.entry2.delete('0', 'end')
        self.entry2.insert('0', self.types_path)


    def run(self):
        self.mainwindow.mainloop()

    def run_script(self):
        filename = self.entry1.get().rsplit('/', 1)[0] + '/' + self.entry3.get() + '.xlsx'
        Data_Process(self.entry1.get(), filename, self.entry2.get(), self.entry7.get())
        temp = {'path_source': self.entry1.get(),
                'path_types': self.entry2.get(),
                'new_name': self.entry3.get(),
                'username': self.entry7.get()}
        with open('temp.json', 'w') as fp:
            json.dump(temp, fp)
        os.startfile(filename)
        root.quit()

if __name__ == '__main__':
    root = tk.Tk()
    root.title('Data Model Documentation')
    # root.wm_attributes('-toolwindow', 'True')
    app = NewprojectApp(root)
    app.run()



setuptools
pandas
pytz
openpyxl