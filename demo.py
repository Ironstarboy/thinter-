import tkinter as tk
from tkinter import filedialog
window = tk.Tk()
window.geometry("500x300")

def open_file():
    filename = filedialog.askopenfilename(title='打开csv文件', filetype=[('csv', '*.csv')]).pack()  # 打开文件对话框
    # entry_filename.insert('insert', filename)

# 设置button按钮接受功能
button_import = tk.Button(window, text="导入文件", command=open_file)

# 设置entry
entry_filename = tk.Entry(window, width=30, font=("宋体", 10, 'bold'))
entry_filename.pack()
# 尝试输出
def print_file():
    a = entry_filename.get()  #用get提取entry中的内容
    print(a)
tk.Button(window, text="输出", command=print_file).pack()
window.mainloop()
