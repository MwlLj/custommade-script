#encoding=utf8
import sys
import os
import time
import datetime
import random
import hashlib
sys.path.append("../../base")
from parse_excel import CParseExcel
import numpy as np
import pandas as pd
import xlrd
import xlsxwriter


class CReadInfo(object):
	def __init__(self, file_path):
		self.m_parse_excel = CParseExcel(file_path)
		self.m_row_datas = []

	def read(self, sheet_callback=None):
		self.m_parse_excel.parse()
		info_dict = self.m_parse_excel.get_info_dict()
		for key, sheet_infos in info_dict.items():
			sheet_len = len(sheet_infos)
			if sheet_len < 1:
				raise SystemExit("[Error] sheet1 is not exist")
			sheet_index = 0
			for sheet_info in sheet_infos:
				sheet_name = sheet_info.get(CParseExcel.SHEET_NAME)
				row_values = sheet_info.get(CParseExcel.ROW_VALUES)
				row_index = 0
				for row_value in row_values:
					row_data = {}
					row_index = row_value.get(CParseExcel.ROW_INDEX)
					col_values = row_value.get(CParseExcel.COL_VALUES)
					col_index = 0
					for col_value in col_values:
						value = col_value.get(CParseExcel.VALUE)
						self.col_value(sheet_index, row_index, col_index, value)
						col_index += 1
					row_index += 1
					self.col_values(sheet_index, sheet_name, row_index, col_values)
					if sheet_callback is not None:
						b = sheet_callback(sheet_index, sheet_name, row_index, col_values)
						if b is False:
							return
					self.m_row_datas.append(row_data)
				sheet_index += 1

	def col_value(self, sheet_index, row_index, col_index, value):
		pass

	def col_values(self, sheet_index, sheet_name, row_index, col_values):
		pass

	def delete_multi_row(self, del_infos, obj_path):
		self.m_parse_excel.delete_multi_row(del_infos, obj_path)

	def get_row_datas(self):
		return self.m_row_datas


class CFindRegionFieldByTitle(CReadInfo):
	def __init__(self, file_path):
		CReadInfo.__init__(self, file_path)
		self.m_region_infos = []
		self.m_file_path = file_path

	def get_region_infos(self):
		def callback(sheet_index, sheet_name, row_index, col_indexs):
			index = 0
			for col_index in col_indexs:
				value = col_index.get(CParseExcel.VALUE)
				if value == "办事处" or value == "所属区域":
					region_info = (sheet_index, sheet_name, row_index, index, col_indexs)
					self.m_region_infos.append(region_info)
				index += 1
			return True
		self.read(callback)
		return self.m_region_infos

	def get_headers(self, sheetname, flag_row_index):
		info = {}
		def callback(sheet_index, sheet_name, row_index, col_indexs):
			if sheet_name != sheetname:
				return True
			if row_index > flag_row_index:
				return False
			info[row_index - 1] = col_indexs
			return True
		self.read(callback)
		return info

	def get_data_after_filter(self, sheetname, path):
		reader = CReadInfo(path)
		info = {}
		def callback(sheet_index, sheet_name, row_index, col_indexs):
			if sheet_name != sheetname:
				return True
			info[row_index] = col_indexs
			return True
		reader.read(callback)
		return info

	def __write(self, write_sheet, datas, row_offset=0, row_height=25, form=None):
		for row, col_values in datas.items():
			for col_value in col_values:
				col = col_value.get(CParseExcel.COL_INDEX)
				value = col_value.get(CParseExcel.VALUE)
				write_sheet.write(row + row_offset, col, value, form)
				write_sheet.set_row(row + row_offset, row_height, form)
				write_sheet.set_tab_color("green")

	def gen(self, obj_path):
		writer = pd.ExcelWriter(obj_path)
		region_infos = self.get_region_infos()
		for sheet_index, sheet_name, row_index, col_index, col_indexs in region_infos:
			excel_ori = pd.read_excel(io=self.m_file_path, sheet_name=sheet_name)
			a = excel_ori.values
			select = a[:, col_index]
			a = a[np.where((select == "青岛办事处") | (select == "上海客户一部") | (select == "上海客户二部") | (select == "上海客户三部") | (select == "南京客户一部") | (select == "南京客户二部") | (select == "南京客户三部") | (select == "山东客户一部") | (select == "山东客户二部") | (select == "商务定制部") | (select == "苏州客户部") | (select == "苏北客户部") | (select == "青岛办事处"))]
			data_df = pd.DataFrame(data=a, copy=True)
			# columns = [index.get(CParseExcel.VALUE) for index in col_indexs]
			# data_df.columns = columns
			data_df.to_excel(writer, sheet_name, index=False, columns=None, header=None)
		writer.save()
		# add headers
		wbt = xlsxwriter.Workbook(obj_path)
		for sheet_index, sheet_name, row_index, col_index, col_indexs in region_infos:
			headers = self.get_headers(sheet_name, row_index)
			context = self.get_data_after_filter(sheet_name, obj_path)
			write_sheet = wbt.add_worksheet(sheet_name)
			header_bold = wbt.add_format({
				'bold': 1,
				# 'fg_color': "#FF0",
				# 'align': 'center',
				# 'valign': 'vcenter',
				# 'text_wrap': 1,
				# "left": 2,
				# "right": 2,
				# "top": 2,
				# "bottom": 2,
				})
			context_bold = wbt.add_format({
				# 'align': 'center',
				# 'valign': 'vcenter',
				# 'text_wrap': 1,
				})
			self.__write(write_sheet, headers, form=header_bold)
			self.__write(write_sheet, context, len(headers) - 1, form=context_bold)
		wbt.close()


if __name__ == "__main__":
	reader = CFindRegionFieldByTitle("./file/test1.xlsx")
	reader.read()
	reader.gen("./obj/test1.xlsx")
