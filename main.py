import openpyxl
import sqlite3
# from openpyxl import load_workbook

def dateChanger(source_string):
	''' Функция преобразоваения даты '''
	new_string = source_string
	if "datetime.datetime(" in new_string:
		while "datetime.datetime(" in new_string:
			# Находим дату и разбиваем на составляющие
			lol = new_string[(new_string.find("datetime.datetime(")+len("datetime.datetime(")):new_string.find(")")].split(', ') # "datetime.datetime(".__sizeof__()
			# Формируем новую дату
			new_date_format = lol[2] + "-" + lol[1] + "-" + lol[0]
			# Модифицируем строку
			new_string = new_string[:new_string.find("datetime.datetime(")] + new_date_format + new_string[new_string.find("),")+1:]
	print(new_string)
		


# открываем файл excel
wb = openpyxl.load_workbook(filename = 'contr.xlsx')

#sheet_ranges = wb['range names']

# получаем имена страниц в виде списка
pages = wb.sheetnames

# берем первую страницу
one_page = wb[pages[0]]
temp_string_index = ''

temp_list = []
# Храним наши значения в словаре
full_dict = dict()
# Перебираем нужное нам количество строк, опуская строки заголовков
for row in range(2,6):
	# Перебираем те столбцы, данные из которых нас интересуют
	for i in 'CDEFGJP':
		temp_string_index = i + str(row)
		temp_list.append(one_page[temp_string_index].value)
	# Добавляем ключ-значение
	full_dict[row] = temp_list
	# Затираем временное значение
	temp_list = []

connect = sqlite3.connect("YourDataBase.dbl")
cursor = connect.cursor()

data = ""
list_data = list()
for i in full_dict:
	data =  (str(full_dict[i]))
	list_data = data.strip('[]')

	# Вырезаем часть с датой, разбиваем ее по запятой
	lol = list_data[(list_data.find("datetime.datetime(")+len("datetime.datetime(")):list_data.find(")")].split(', ') # "datetime.datetime(".__sizeof__()
	# 
	# Пока просто выводим данные
	dateChanger(list_data)



for i in range(2,5):

	# query_str = """INSERT INTO organisations VALUES (NULL, """ + data + """)"""
	# print(self.query_str)
	#cursor.execute(query_str)
	pass

# Этот комментарий был добавлен и отправлен на сервер

#connect.commit()



