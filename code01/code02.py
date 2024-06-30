import xlwings as xw

path = 'E:\\python_excel\\01\\分公司1.xlsx'
app = xw.App(visible=True, add_book=False)  # 打开excel窗口
workbook = app.books.open(path)  # 打开对应文件
worksheet = workbook.sheets.add('产品统计表')  # 填写一个新的sheet表“产品统计表”
worksheet.range('A1').value = "编号"  # “产品统计表”单元格A1处输入内容
for i in range(2, 22):
    worksheet.range(f'A{i}').value = f'{i-1}'
workbook.save()
workbook.close()
app.quit()
