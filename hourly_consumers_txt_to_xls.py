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


def is_odd(number:int):
	return number%2==0

def Get_Hour_by_half_hour_number(halfhour:int) -> int:
	halfhour += ( 0 if is_odd(halfhour) else 1 )
	return halfhour//2

class loader_hourly_displays_from_txt:
	def __init__(self, pc_path_to_text_file:str):
		self.source_directory_path = pc_path_to_text_file
		source_strings_list = []
		source_strings = ''
		self.halfhour_list = []
		file = open(pc_path_to_text_file, "r")
		source_strings_list= file.readlines()
		file.close
		for line in source_strings_list:
			source_strings += line.strip() + chr(13)
		print(source_strings)

		self.point = sx(source_strings, '< REC=POINT; COUNT=', ';')
		echo(style(text='Прибор учёта:  ', fg = 'green') +  style(text=self.point, fg = 'bright_green'))

		for line in source_strings_list:
			if not '< REC=DAY; DATE=' in line:
				continue
			row_date = sx(line, '< REC=DAY; DATE=', ';') # дата показаний строчки
			echo(style(text='Дата:  ', fg = 'yellow') +  style(text=row_date, fg = 'bright_yellow'))
			row_displays = ',' + sx(line, 'P48=', '; ST48=') + ','
			for halfhour_number in range(1,49):
				lc_display = sx(row_displays,',',',',halfhour_number)
				ld = {'date':row_date, 'halfhour':halfhour_number, 'value':lc_display}
				self.halfhour_list.append(ld)
				echo(style(text = ld, fg = 'green'))
		unique_dates = []
		for row in self.halfhour_list:
			append_if_not_exists(row['date'] , unique_dates)
		print(unique_dates)
		self.data = []
		data_dict = {}
		for date in unique_dates:
			for hhour in range(1,49):
				ld = {}
				for row in self.halfhour_list:
					if row['date'] == date and row['halfhour'] == hhour:
						ld = row
						break
				lc = f"/{date}_{str(Get_Hour_by_half_hour_number(hhour))}/"
				if lc in data_dict:
					data_dict[lc] += float(ld['value'])
				else:
					data_dict[lc] = float(ld['value'])
		print(data_dict)
		for key in data_dict:
			ld = {'date':'20'+sx(key,'/','_')[0:2]+sx(key,'/','_')[2:6],  'hour':sx(key,'_','/'), 'value':data_dict[key]}
			self.data.append(ld)
			print(ld)
				
	def data_to_excel_merged(self, pc_file_name:str):
		echo(style(text=pc_file_name, fg='blue', bg='white'))
		ll_array = []
		for row in self.data:
			ll_array.append([self.point,  time.strftime('%d.%m.%Y',time.strptime(row['date'],'%Y%m%d')), f"{row['hour']}:00", row['value']])
		ll_columns = ['ПУ','Дата','Часы','Значение']
		df = pandas.DataFrame(ll_array, columns=ll_columns )
		#writer = pandas.ExcelWriter(os.path.dirname(self.source_directory_path)+f'\\{self.point}_{pc_file_name}', engine='xlsxwriter')
		writer = pandas.ExcelWriter(self.source_directory_path+f' {pc_file_name}', engine='xlsxwriter')
		df.to_excel(writer, sheet_name='Sheet1', startrow=2)
		sheet = writer.sheets['Sheet1']
		sheet.autofilter(2,0,2,25)
		sheet.set_column('A:Z', 15)
		writer.save()

	def data_to_excel_halfhour(self, pc_file_name:str):
		echo(style(text=pc_file_name, fg='blue', bg='white'))
		ll_array = []
		for row in self.halfhour_list:
			ll_array.append([self.point, row['date'], row['halfhour'], float(row['value'])])
		ll_columns = ['ПУ','Дата','Получасовка','Значение']
		df = pandas.DataFrame(ll_array, columns=ll_columns )
		#writer = pandas.ExcelWriter(os.path.dirname(self.source_directory_path)+f'\\{self.point}_{pc_file_name}', engine='xlsxwriter')
		writer = pandas.ExcelWriter(self.source_directory_path+f' {pc_file_name}', engine='xlsxwriter')
		df.to_excel(writer, sheet_name='Sheet1', startrow=2)
		sheet = writer.sheets['Sheet1']
		sheet.autofilter(2,0,2,25)
		sheet.set_column('A:Z', 15)
		writer.save()


loader = loader_hourly_displays_from_txt(sys.argv[1])
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


