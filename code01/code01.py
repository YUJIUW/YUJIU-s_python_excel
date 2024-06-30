import xlwings as xw

app = xw.App(visible=True, add_book=False)
for i in range(1, 21):
    workbook = app.books.add()
    workbook.save(f'E:\\python_excel\\01\\分公司{i}.xlsx')
    workbook.close()# 关闭工作簿
app.quit()#关闭excel程序
