code01批量创建Excel
code02打开excel表格并输入内容
app = xw.App(visible=True, add_book=False)

App()是xlwings模块中的函数
参数：
visible：用于设置excel程序窗口的可见性。如果为True，表示显示excel窗口；如果为False，表示隐藏excel窗口。
add_book：用于设置启动excel窗口后是否新建工作簿，如果Ture，表示新建一个工作簿，如果为False表示不新建工作簿

workbook = app.books.add()
add（）为books对象的函数