import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox



class Application(tk.Tk):
    filename=''
    def __init__(self):
        super().__init__()  # 有点相当于tk.Tk()
        self.createWidgets()

    def createWidgets(self):
        self.title('上传的xls路径不能有中文')
        self.columnconfigure(0, minsize=500)
        # 定义一些变量
        self.entryvar = tk.StringVar()

        # 先定义顶部Frame，用来放置下面的部件
        topframe = tk.Frame(self, height=80)
        topframe.pack(side=tk.TOP)

        # 顶部区域（四个部件）
        glabel = tk.Label(topframe, text='请导入xls文件:')
        gentry = tk.Entry(topframe, textvariable=self.entryvar)
        gbutton = tk.Button(topframe, command=self.__openfile, text='导入文件')
        gbutton2 = tk.Button(topframe, command=self.my_command, text='处理文件')

        # -- 放置位置
        glabel.grid(row=0, column=0, sticky=tk.W)
        gentry.grid(row=0, column=1)
        gbutton.grid(row=0, column=2)
        gbutton2.grid(row=0, column=3)


    def __openfile(self):
        '''打开文件的逻辑'''

        self.filename = filedialog.askopenfilename(title='打开xls文件', filetype=[('xls', '*.xls')])  # 打开文件对话框
        self.entryvar.set(self.filename)  # 设置变量entryvar，等同于设置部件Entry
        print('成功导入文件: {}'.format(self.filename))
        if not self.filename:
            messagebox.showwarning('警告', message='未选择文件！')  # 弹出消息提示框




    def my_command(self):
        '''处理逻辑'''
        # self.data = pd.read_csv(self.filename)

        import warnings

        warnings.filterwarnings('ignore')
        import MYprogram
        MYprogram.run(self.filename)
        print('处理完成')




if __name__ == '__main__':
    # 实例化Application
    app = Application()


    # 主消息循环:
    app.mainloop()

