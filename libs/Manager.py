import xlrd
import json

class Manager(object):

	def __init__(self, xls_filename):
		# The filename for the json file output will be the same of the
		# xls filename
		self.filename = xls_filename
		self.sheet = self.__open_workbook(self.filename)

	def __open_workbook(self, filename, sheet_page=0):
		workbook = xlrd.open_workbook(filename)
		sheet = worksheet = workbook.sheet_by_index(sheet_page)
		return sheet

	def __parse_file_headers(self):
		headers = []
		for cell in range(1, self.sheet.ncols):
			headers.append(
				self.sheet.cell(0, cell).value
			)
		return headers

	# Main function
	# True if you want to generate the local json file
	def parse_file(self, gen_json=False):
		data = []
		headers = self.__parse_file_headers()
		for row in range(1, self.sheet.nrows):
			obj = {}
			for column in range(1, self.sheet.ncols):
				try:
					curr_value = str(self.sheet.cell(row, column)).split('text:')[1].replace("'", '')	
				except Exception as e:
					curr_value = str(self.sheet.cell(row, column)).split('number:')[1].replace("'", '')
				obj[headers[column-1]] = curr_value
			data.append(obj)

		if gen_json:
			status = self.__write_json_file(data)
			if status == 0:
				pass
				# print('parse_file() gen_json => OK')
			elif status == -1:
				pass
				# print('parse_file() gen_json => ERR')

		return data

	def __write_json_file(self, data):
		json_object = json.dumps(data, indent=4)
		try:
			filename = f'{self.filename.split(".")[0]}.json'
			with open(filename, "w") as json_file:
				json_file.write(json_object)
			return 0
		except Exception as e:
			return -1
		return -1