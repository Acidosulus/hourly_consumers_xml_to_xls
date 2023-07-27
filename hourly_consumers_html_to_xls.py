import xmltodict
from click import echo, style
from os import listdir
from os.path import isfile, join
import pandas
import glob
import time
import sys
import os





def sx(source_string='', left_split='', right_split='', index=1):
	if source_string.count(
			left_split) < index:
		return ""
	lc_str = source_string
	for i in range(0, index):
		lc_str = lc_str[lc_str.find(left_split) + len(left_split):len(lc_str)]
	return lc_str[0:lc_str.find(right_split)]

def str_to_file(file_path:str, st:str): # записывает строку в текстовый файл
	file = open(file_path, mode='w', encoding='utf-8')
	file.write(st)
	file.close()

#  procedure append element into list if it don't exist into list
def append_if_not_exists(element, ll:list):
	if element not in ll:
		ll.append(element)


def file_to_str(file_path:str):		 # считывает текстовый файл в строку
	with open(file_path, "r", encoding="1251") as myfile:
		data = ' '.join(myfile.readlines())
	myfile.close()
	return data


def is_odd(number:int):
	return number%2==0

def Get_Hour_by_half_hour_number(halfhour:int) -> int:
	halfhour += ( 0 if is_odd(halfhour) else 1 )
	return halfhour//2

class loader_hourly_displays_from_html:
	def __init__(self, pc_path_to_text_file:str):
		self.source_directory_path = pc_path_to_text_file
		source_strings_list = []
		source_strings = ''
		self.halfhour_list = []
		self.data = []
		self.dates = []
		print(self.source_directory_path)
		file_contents = file_to_str(self.source_directory_path)
		if '<TH class=style20>R-, квар</TH>' in file_contents: # формат загрузки файлов 11077_176385515_08_2022_1200.html
			source = sx(file_contents,'<TBODY>','</TBODY>')
			echo(style(text='Вариант 2', fg='cyan'))
			keys_reference = {'00:30':1, '01:00':1, '01:30':2, '02:00':2, '02:30':3, '03:00':3, '03:30':4, '04:00':4, '04:30':5, '05:00':5, '05:30':6, '06:00':6, '06:30':7, '07:00':7, '07:30':8, '08:00':8, '08:30':9, '09:00':9, '09:30':10, '10:00':10, '10:30':11, '11:00':11, '11:30':12, '12:00':12, '12:30':13, '13:00':13, '13:30':14, '14:00':14, '14:30':15, '15:00':15, '15:30':16, '16:00':16, '16:30':17, '17:00':17, '17:30':18, '18:00':18, '18:30':19, '19:00':19, '19:30':20, '20:00':20, '20:30':21, '21:00':21, '21:30':22, '22:00':22, '22:30':23, '23:00':23, '23:30':24, '00:00':24}
			self.point = ''
			for i in range(1,source.count('<TR>')+1):
				lc = sx(source, '<TR>', '</TR>', i)
				if i==1: #первая строчка пропускается, но с её берётся дата, поскольку данные столбца даты должны быть сдвинуты на одну позицию вниз, посколько 00:00 считается последней получасовкой суток
					last_date = sx(lc, '<TD class=style21>', '</TD>', 7)
					continue
				ld = {'date':last_date ,'halfhour':sx(lc, '<TD class=style21>', '</TD>', 6), 'value':sx(lc, '<TD class=style21>', '</TD>', 2)}
				echo(style(text=ld, fg='cyan'))
				self.halfhour_list.append(ld)
				last_date = sx(lc, '<TD class=style21>', '</TD>', 7)
		if '<SCRIPT language=javascript src="../../../../res/js/flot/jquery.flot.selection.js" type=text/javascript></SCRIPT>' in file_contents:
			echo(style(text='Вариант 3', fg='cyan'))
			keys_reference = {'00:30':1, '01:00':1, '01:30':2, '02:00':2, '02:30':3, '03:00':3, '03:30':4, '04:00':4, '04:30':5, '05:00':5, '05:30':6, '06:00':6, '06:30':7, '07:00':7, '07:30':8, '08:00':8, '08:30':9, '09:00':9, '09:30':10, '10:00':10, '10:30':11, '11:00':11, '11:30':12, '12:00':12, '12:30':13, '13:00':13, '13:30':14, '14:00':14, '14:30':15, '15:00':15, '15:30':16, '16:00':16, '16:30':17, '17:00':17, '17:30':18, '18:00':18, '18:30':19, '19:00':19, '19:30':20, '20:00':20, '20:30':21, '21:00':21, '21:30':22, '22:00':22, '22:30':23, '23:00':23, '23:30':24, '00:00':24}
			source = sx(file_contents,'<body','</body>')
			#print(source)
			self.point = ''
			for i in range(1,source.count('<tr>')+1):
				lc = sx(source, '<tr>', '</tr>', i)
				print()
				lc_date = sx(lc, "<td class='style21'>", '</td>', 7)
				lc_halfhour = sx(lc, "<td class='style21'>", '</td>', 6)
				lc_value = sx(lc, "<td class='style21'>", '</td>', 2)
				print(f'{lc_date}    {lc_halfhour}    {lc_value}')
				ld = {'date':lc_date ,'halfhour':lc_halfhour, 'value':lc_value}
				echo(style(text=ld, fg='green'))
				self.halfhour_list.append(ld)
		else:
			if 'UTC(мс)' in file_contents: # формат загрузки файла с именем М15247_29857958_08_2022_40.html
				source = sx(file_contents,'<TBODY>','</TBODY>')
				echo(style(text='Вариант 1', fg='cyan'))
				keys_reference = {'00:30':1, '01:00':1, '01:30':2, '02:00':2, '02:30':3, '03:00':3, '03:30':4, '04:00':4, '04:30':5, '05:00':5, '05:30':6, '06:00':6, '06:30':7, '07:00':7, '07:30':8, '08:00':8, '08:30':9, '09:00':9, '09:30':10, '10:00':10, '10:30':11, '11:00':11, '11:30':12, '12:00':12, '12:30':13, '13:00':13, '13:30':14, '14:00':14, '14:30':15, '15:00':15, '15:30':16, '16:00':16, '16:30':17, '17:00':17, '17:30':18, '18:00':18, '18:30':19, '19:00':19, '19:30':20, '20:00':20, '20:30':21, '21:00':21, '21:30':22, '22:00':22, '22:30':23, '23:00':23, '23:30':24, '00:00':24}
				self.point = ''
				for i in range(1,source.count('<TR>')+1):
					lc = sx(source, '<TR>', '</TR>', i)
					ld = {'date':sx(lc, '<TD class=style21>', '</TD>', 7) ,'halfhour':sx(lc, '<TD class=style21>', '</TD>', 6), 'value':sx(lc, '<TD class=style21>', '</TD>', 2)}
					echo(style(text=ld, fg='green'))
					self.halfhour_list.append(ld)


		for row in self.halfhour_list:
			append_if_not_exists(row['date'], self.dates) # уникальные даты
		for date in self.dates:
			for row in self.halfhour_list:
				if len(row['date'])==0:
					continue
				hour = keys_reference[row['halfhour']]
				ld = {'date':date, 'hour':hour, 'value':0.00000}
				append_if_not_exists(ld, self.data)
		for row in self.halfhour_list:
			for rezrow in self.data:
				if len(rezrow['date'])==0:
					continue
				if rezrow['date']==row['date'] and rezrow['hour']==keys_reference[row['halfhour']]:
					rezrow['value'] += float(row['value'])/2
		echo(style(text=self.data, fg='bright_white'))

		return
				
	def data_to_excel_merged(self, pc_file_name:str):
		echo(style(text=pc_file_name, fg='blue', bg='white'))
		ll_array = []
		for row in self.data:
			if len(row['date'])==0:
				continue
			ll_array.append([self.point,  time.strftime('%d.%m.%Y',time.strptime(row['date'],('%d.%m.%y' if len(row['date'])==8 else '%d.%m.%Y'))), f"{row['hour']}", row['value']])
		ll_columns = ['ПУ','Дата','Часы','Значение']
		df = pandas.DataFrame(ll_array, columns=ll_columns )
		writer = pandas.ExcelWriter(self.source_directory_path+f' {pc_file_name}', engine='xlsxwriter')
		#writer = pandas.ExcelWriter(os.path.dirname(self.source_directory_path)+f'\\{self.point}_{pc_file_name}', engine='xlsxwriter')
		

		df.to_excel(writer, sheet_name='Sheet1', startrow=2)
		sheet = writer.sheets['Sheet1']
		sheet.autofilter(2,0,2,25)
		sheet.set_column('A:Z', 15)
		writer.save()

	def data_to_excel_halfhour(self, pc_file_name:str):
		echo(style(text=pc_file_name, fg='blue', bg='white'))
		ll_array = []
		for row in self.halfhour_list:
			if len(row['date'])==0:
				continue
			ll_array.append([self.point, row['date'], row['halfhour'], float(row['value'])])
		ll_columns = ['ПУ','Дата','Получасовка','Значение']
		df = pandas.DataFrame(ll_array, columns=ll_columns )
		writer = pandas.ExcelWriter(self.source_directory_path+f' {pc_file_name}', engine='xlsxwriter')
		#writer = pandas.ExcelWriter(os.path.dirname(self.source_directory_path)+f'\\{self.point}_{pc_file_name}', engine='xlsxwriter')
		df.to_excel(writer, sheet_name='Sheet1', startrow=2)
		sheet = writer.sheets['Sheet1']
		sheet.autofilter(2,0,2,25)
		sheet.set_column('A:Z', 15)
		writer.save()


loader = loader_hourly_displays_from_html(sys.argv[1])
loader.data_to_excel_merged('Объединенные получасовки.xlsx')
loader.data_to_excel_halfhour('Получасовки.xlsx')


#for i in range(1,49):
#	print(f'{i} => {Get_Hour_by_half_hour_number(i)}')

#for i in range(1,25):
#	for j in range(1,3):
#		print(f'i={i}   j={j}        {i*2+j-2} -> {Get_Hour_by_half_hour_number(i*2+j-2)}')


#xlms.data_to_excel_halfhour('Получасовки.xlsx')
#xlms.data_to_excel_merged('Объединенные получасовки.xlsx')
#xlms.data_to_excel_merged_by_columns('Объединённые ПУ.xlsx')


