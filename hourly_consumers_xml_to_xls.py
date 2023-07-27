from lib2to3.pgen2.token import AMPEREQUAL
import xmltodict
from click import echo, style
from os import listdir
from os.path import isfile, join
import pandas
import glob
import time
import sys
import pickle


def str_to_file(file_path:str, st:str): # записывает строку в текстовый файл
	file = open(file_path, mode='w', encoding='utf-8')
	file.write(st)
	file.close()

#  procedure append element into list if it don't exist into list
def append_if_not_exists(element, ll:list):
	if element not in ll:
		ll.append(element)

def sx(source_string='', left_split='', right_split='', index=1):
	if source_string.count(
			left_split) < index:
		return ""
	lc_str = source_string
	for i in range(0, index):
		lc_str = lc_str[lc_str.find(left_split) + len(left_split):len(lc_str)]
	return lc_str[0:lc_str.find(right_split)]



class load_xml_set:
	def __init__(self, path_to_xmls:str):
		path_to_xmls = path_to_xmls + ('\\' if path_to_xmls[len(path_to_xmls)-1]!='\\' else '')
		self.source_directory_path = path_to_xmls
		#file_list = [f for f in listdir(path_to_xmls) if isfile(join(path_to_xmls, f))]
		file_list = glob.glob(f'{path_to_xmls}\\*.xml')
		self.xlms = []
		self.data = []
		self.halfhour_data = []
		for file in file_list:
			#echo(style(text = 'Файл:', fg = 'white') + style(text=file, fg='black', bg='yellow'))
			hd = hourlyxml(file)
			self.xlms.append(hd)
		self.points = [] #список уникальных номеров приборов учёта из загруженных данных
		for day in self.xlms:
			for point in day.points:
				ldic = {'code':point.code, 'name':point.name, 'channel':point.channel, 'desc':point.desc}
				append_if_not_exists(ldic, self.points)

		self.dates = [] # список уникальных дат и часов
		for day in self.xlms:
			for hour in range(1,25):
				ld = {'day':day.date, 'hour':hour}
				append_if_not_exists(ld, self.dates)
		#print(self.dates)


		for point in self.points: # уникальные точки учёта
			for xm in self.xlms: # загруженные файлы
				for code in xm.points: # приборы с днями в загруженном файле
					if code.code == point['code']:
						for row in code.data: # данные расхода прибора
							for hour in range(1,25):
								if int(row['hour']) == int(hour):
									row_data = {'code':point['code'], 'name':point['name'], 'channel':point['channel'],'desc': point['desc'], 'date':row['date'], 'hour':row['hour'], 'value':row['value']}
									#echo(style(text=row_data, fg='yellow'))
									self.data.append(row_data)

		lc = ''
		for xm in self.xlms: # загруженные файлы
			for code in xm.points: # приборы с днями в загруженном файле
					for row in code.halfhour_data: # данные расхода прибора
						lc += code.code +'   ' + code.name + '   '+  code.channel + '   ' + code.desc + '  ' + row['date'] + '  ' + row['shour'] + '  ' + row['sminute'] + '   ' + str(row['value']) + chr(13)
						ld = {'code':code.code,'name':code.name,'channel':code.channel,'desc':code.desc,'date':row['date'],'hour':row['shour'],'minute':row['sminute'],'value':row['value']}
						#echo(style(text=ld, fg='bright_green'))
						self.halfhour_data.append(ld)
		#str_to_file('full.txt', lc)

		self.data_by_hours = []
		#ld = {'date':'', 'hour':''}
		#for point in self.points:
		#	ld[point['code']]=point['code']
		#self.data_by_hours.append(ld)

		for row in self.data:
			row['key']=f'{row["code"]}|{row["date"]}|{row["hour"]}'


		for dt in self.dates: #  уникальные сочетания дат и часов
			ld = {'date':dt['day'],'hour':dt['hour']}
			for point in self.points: # уникальные номера счётчиков
				lc_key = f"{point['code']}|{dt['day']}|{dt['hour']}"
				row = next(x for x in self.data if x["key"] == lc_key)
				ld[point['code']]=row['value']
				#echo(style(text=row, fg='bright_red'))
				

			self.data_by_hours.append(ld)

		#lc = ''
		#for row in self.data_by_hours:
		#	lc += str(row) + chr(13)
		#str_to_file('for_report.txt', lc)



	def data_to_excel_merged_by_columns(self, pc_file_name:str):
		echo(style(text=pc_file_name, fg='blue', bg='white'))
		ll_array = []
		for row in self.data_by_hours:
			#ld = [row['date'], row['hour']]
			ld = [time.strftime('%d.%m.%Y',time.strptime(row['date'],'%Y%m%d')), f"{row['hour']}:00"]
			for point in self.points:
				ld.append(row[point['code']])
			#echo(style(text=ld, fg='green'))
			ll_array.append(ld)
		ll_columns = ['Дата','Час']
		for point in self.points:
			#echo(style(text=point['code'], fg='blue'))
			ll_columns.append(point['code'])
		df = pandas.DataFrame(ll_array, columns=ll_columns )
		writer = pandas.ExcelWriter(self.source_directory_path + pc_file_name, engine='xlsxwriter')
		df.to_excel(writer, sheet_name='Sheet1', startrow=2)
		sheet = writer.sheets['Sheet1']
		sheet.autofilter(2,0,2,25)
		sheet.set_column('A:Z', 15)
		writer.save()


	def data_to_excel_merged(self, pc_file_name:str):
		echo(style(text=pc_file_name, fg='blue', bg='white'))
		ll_array = []
		for row in self.data:
			ld = [row['code'], row['name'], row['channel'], row['desc'], row['date'], row['hour'], row['value']]
			#echo(style(text=ld, fg='yellow'))
			ll_array.append(ld)
		ll_columns = ['ПУ','Название','Канал','Описание','Дата','Часы','Значение']
		df = pandas.DataFrame(ll_array, columns=ll_columns )
		writer = pandas.ExcelWriter(self.source_directory_path + pc_file_name, engine='xlsxwriter')
		df.to_excel(writer, sheet_name='Sheet1', startrow=2)
		sheet = writer.sheets['Sheet1']
		sheet.autofilter(2,0,2,25)
		sheet.set_column('A:Z', 15)
		writer.save()

	def data_to_excel_halfhour(self, pc_file_name:str):
		echo(style(text=pc_file_name, fg='blue', bg='white'))
		ll_array = []
		for row in self.halfhour_data:
			ld = [row['code'], row['name'], row['channel'], row['desc'], row['date'], row['hour'], row['minute'], float(row['value'].replace(',','.'))]
			#echo(style(text=ld, fg='cyan'))
			ll_array.append(ld)
		ll_columns = ['ПУ','Название','Канал','Описание','Дата','Часы','Минуты','Значение']
		df = pandas.DataFrame(ll_array, columns=ll_columns )
		writer = pandas.ExcelWriter(self.source_directory_path + pc_file_name, engine='xlsxwriter')
		df.to_excel(writer, sheet_name='Sheet1', startrow=2)
		sheet = writer.sheets['Sheet1']
		sheet.autofilter(2,0,2,25)
		sheet.set_column('A:Z', 15)
		writer.save()


class hourlyxml:
	def __init__(self, file:str):
		self.file_path = file
		self.load_data()

	def load_data(self):
		file = open(self.file_path, 'r', errors = 'ignore')
		self.points = [] # список объектов приборов учёта с их показаниям за день
		source = file.read()
		self.date = sx(source,'<day>','</day>')
		self.inn  = sx(sx(source,'<sender>','/<sender>'),'<inn>','</inn>')
		self.name  = sx(sx(source,'<sender>','/<sender>'),'<name>','</name>')

		echo(style(text = 'Дата:', fg = 'green')+style(text = self.date, fg = 'bright_green'))
		echo(style(text = 'ИНН:', fg = 'green')+style(text = self.inn, fg = 'bright_green'))
		echo(style(text = 'Название:', fg = 'green')+style(text = self.name, fg = 'bright_green'))		
		echo(style(text = 'Приборы учёта:', fg = 'bright_red'))
		
		for i in range(0,source.count('<area>')):
			area = sx(source,'<area>','</area>',i+1)
			for j in range(0,area.count('<measuringpoint')):
				mpoint = sx(area, '<measuringpoint', '</measuringpoint>',j+1)
				hpl = hourlypointxml(mpoint, self.date)
				self.points.append(hpl)
				hpl.print_information()


class hourlypointxml:
	def __init__(self, source:dict, cdate:str):
		self.name = sx(source, 'name="', '"')
		self.code = sx(source, 'code="', '"')
		print('==============',self.code)
		self.channel = sx(source, '<measuringchannel code="','"')
		self.desc = sx(source, 'desc="', '"')
		#echo(style(text=f'====> {self.desc}', bg='blue', fg='bright_white'))
		self.halfhour_data = []
		self.data = []
		print(f"Количество каналов: {source.count('<measuringchannel')}")
		for measure_number in range(0,source.count('<measuringchannel')):
			measure = '<measuringchannel' + sx(source,'<measuringchannel','</measuringchannel>',measure_number+1) + '</measuringchannel>'
			print()
			if 'Реактивная энергия' not in measure:
				echo(style(text = f'{measure_number} не реактивная энергия - чтение данных', bg='blue', fg='bright_white'))
				#echo(style(text = measure, fg='bright_green'))
				print(f"Количество периодов: {measure.count('<period')}")
				for k in range(0,measure.count('<period')):
					period = sx(measure, '<period', '</period>',k+1)
					lcvalue = sx(period,'<value>','</value>')
					if len(lcvalue)==0:
						lcvalue = sx(period,'<value status="0">','</value>')
					append_if_not_exists({'date':cdate, 'shour':sx(period,'start="','"')[0:2], 'sminute':sx(period,'end="','"')[2:4], 'value':lcvalue} , self.halfhour_data)
			else:
				echo(style(text = f'{measure_number} реактивная энергия - пропуск', bg='blue', fg='bright_red'))
			
			append_if_not_exists({'date':cdate, 'shour':sx(period,'start="','"')[0:2], 'sminute':sx(period,'end="','"')[2:4], 'value':lcvalue} , self.halfhour_data)


		for hour in range(1,25): # объединяем получасовки в часы
			ln_value = 0
			for data in self.halfhour_data:
				if int(data['shour'])==hour-1:
					#print(data)
					#if type(data['value']) != dict:
					if len(data['value'])==0:
						data['value']='0'
					ln_value += float(data['value'].replace(',','.'))
					#else:
					#	ln_value += float(data['value']['#text'])
			ld = {'date':cdate, 'hour':str(hour), 'value':round(ln_value,4)}
			#print(hour+1, '  ===>>  ', ld)
			#print()
			self.data.append(ld)
			#exit()
		#print()
		#print(self.data)

	def print_information(self):
		echo(style(text = self.name, fg = 'bright_green') + '	' + style(text = self.code, fg = 'bright_yellow') + '	' + style(text = self.channel, fg = 'bright_cyan') + '	  ' + 
			style(text = self.desc, fg = 'bright_blue'))

	def print_halfhours(self):
		for halhour in self.halfhour_data:
			print(halhour)

	def print_data(self):
		for halhour in self.data:
			print(halhour)


xlms = load_xml_set(sys.argv[1])

with open('xlms.pkl', 'wb') as pickle_out: pickle.dump(xlms, pickle_out)



xlms.data_to_excel_halfhour('Получасовки.xlsx')
xlms.data_to_excel_merged('Объединенные получасовки.xlsx')
xlms.data_to_excel_merged_by_columns('Объединённые ПУ.xlsx')


