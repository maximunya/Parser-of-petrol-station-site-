import requests
from bs4 import BeautifulSoup as BS
import csv
import os, sys, subprocess
import xlsxwriter

URL = 'https://kirishiavtoservis.ru/stations/'
HEADERS = {
	'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.4 Safari/605.1.15',
	'accept': '*/*'
}
FILE_CSV = 'kirishiavtoservis.csv'
FILE_XLSX = 'kirishiavtoservis.xlsx'


def get_html(url, params=None):
	r = requests.get(url, headers=HEADERS, params=params)
	return r


def get_content(html):
	soup = BS(html, 'html.parser')
	items = soup.find_all('div', class_='contacts-table')
	azs = []
	for item in items:
		titles = item.find_all('tr')
		for title in titles:
			prices_box = title.find_all('div', class_='contacts__table-prices-item')
			prices = []
			for el in prices_box:
				prices.append(el.find('strong').get_text())
			azs.append({
				'title': title.find('h5').get_text(),
				'district': item.find('h4').get_text(),
				'address':
				title.find('address', class_='contacts-table__address').get_text(),
				'АИ-98': prices[0].replace('₽', '').replace(',', '.'),
				'АИ-95': prices[1].replace('₽', '').replace(',', '.'),
				'АИ-92': prices[2].replace('₽', '').replace(',', '.'),
				'ДТ': prices[3].replace('₽', '').replace(',', '.'),
				'number': title.find('div', class_='contacts-table__tel').get_text()
			})
	return azs


def save_csv_file(items, path):
	with open(path, 'w', newline='') as file:
		writer = csv.writer(file, delimiter=';')
		writer.writerow([
			'АЗС', 'Район', 'Адрес', 'АИ-98',
			'АИ-95', 'АИ-92', 'ДТ', 'Телефон',
		])
		for item in items:
			writer.writerow([
				item['title'],
				item['district'],
				item['address'],
				item['АИ-98'],
				item['АИ-95'],
				item['АИ-92'],
				item['ДТ'],
				item['number']
			])


def save_xlsx_file(items):
	workbook = xlsxwriter.Workbook('kirishiavtoservis.xlsx')
	worksheet = workbook.add_worksheet()
	bold = workbook.add_format({'bold': True})
	money = workbook.add_format({'num_format': '##.##₽'})
	average = workbook.add_format({'num_format': '##.##₽', 'bold': True})
	center = workbook.add_format({'align': 'center'})
	head = ['АЗС', 'Район', 'Адрес', 'АИ-98', 'АИ-95', 'АИ-92', 'ДТ', 'Телефон']
	col = 0
	row = 1
	for i in head:
		worksheet.write(0, col, i, bold)
		col += 1
	col = 0
	for item in items:
		worksheet.write(row, col, item['title'])
		worksheet.write(row, col + 1, item['district'])
		worksheet.write(row, col + 2, item['address'])
		if item['АИ-98'] != "—":
			worksheet.write(row, col + 3, float(item['АИ-98']), money)
		else:
			worksheet.write(row, col + 3, item['АИ-98'], center)
		if item['АИ-95'] != "—":
			worksheet.write(row, col + 4, float(item['АИ-95']), money)
		else:
			worksheet.write(row, col + 4, item['АИ-95'], center)
		if item['АИ-92'] != "—":
			worksheet.write(row, col + 5, float(item['АИ-92']), money)
		else:
			worksheet.write(row, col + 5, item['АИ-92'], center)
		if item['ДТ'] != "—":
			worksheet.write(row, col + 6, float(item['ДТ']), money)
		else:
			worksheet.write(row, col + 6, item['ДТ'], center)
		worksheet.write(row, col + 7, item['number'])
		row += 1
	row_str = str(row)
	worksheet.write(row, 2, 'СРЕДНЕЕ', bold)
	worksheet.write_formula(row, 3, '=AVERAGE(D2:D' + row_str + ')', average)
	worksheet.write_formula(row, 4, '=AVERAGE(E2:E' + row_str + ')', average)
	worksheet.write_formula(row, 5, '=AVERAGE(F2:F' + row_str + ')', average)
	worksheet.write_formula(row, 6, '=AVERAGE(G2:G' + row_str + ')', average)
	worksheet.write(row + 1, 2, 'МИН.', bold)
	worksheet.write_formula(row + 1, 3, '=MIN(D2:D' + row_str + ')', average)
	worksheet.write_formula(row + 1, 4, '=MIN(E2:E' + row_str + ')', average)
	worksheet.write_formula(row + 1, 5, '=MIN(F2:F' + row_str + ')', average)
	worksheet.write_formula(row + 1, 6, '=MIN(G2:G' + row_str + ')', average)
	worksheet.write(row + 2, 2, 'МАКС.', bold)
	worksheet.write_formula(row + 2, 3, '=MAX(D2:D' + row_str + ')', average)
	worksheet.write_formula(row + 2, 4, '=MAX(E2:E' + row_str + ')', average)
	worksheet.write_formula(row + 2, 5, '=MAX(F2:F' + row_str + ')', average)
	worksheet.write_formula(row + 2, 6, '=MAX(G2:G' + row_str + ')', average)
	workbook.close()


def parse():
	html = get_html(URL)
	if html.status_code == 200:
		azs_list = get_content(html.text)
		save_csv_file(azs_list, FILE_CSV)
		save_xlsx_file(azs_list)
		if sys.platform == 'win32':
			os.startfile(FILE_CSV)
			os.startfile(FILE_XLSX)
		else:
			opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
			subprocess.call([opener, FILE_CSV])
			subprocess.call([opener, FILE_XLSX])
	else:
		print('Error')


parse()
