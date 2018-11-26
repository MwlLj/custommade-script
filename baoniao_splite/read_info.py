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


class CReadInfo(object):
	def __init__(self, file_path):
		self.m_parse_excel = CParseExcel(file_path)
		self.m_row_datas = []

	def read(self):
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

	# def col_value(self, sheet_index, row_index, col_index, value):
		# if value == "办事处" or value == "所属区域":
			# region_info = (sheet_index, row_index, col_index)
			# self.m_region_infos.append(region_info)

	def col_values(self, sheet_index, sheet_name, row_index, col_indexs):
		index = 0
		for col_index in col_indexs:
			value = col_index.get(CParseExcel.VALUE)
			if value == "办事处" or value == "所属区域":
				region_info = (sheet_index, sheet_name, row_index, index, col_indexs)
				self.m_region_infos.append(region_info)
			index += 1

	def get_region_infos(self):
		return self.m_region_infos

	def gen(self, obj_path):
		writer = pd.ExcelWriter(obj_path)
		"""
		getter = CGetHeaderInfos(self.m_file_path, reader.get_region_infos())
		getter.read()
		headers = getter.get_header_infos()
		for sheet_index, v in headers.items():
			sheet_name = v.get("sheet_name")
			data = v.get("data")
			for row in data:
				excel_ori = pd.read_excel(io="./file/test1.xlsx", sheet_name=sheet_name)
				data_df = pd.DataFrame(data=excel_ori.values)
				print([index.get(CParseExcel.VALUE) for index in row])
				data_df.columns = [index.get(CParseExcel.VALUE) for index in row]
				data_df.to_excel(writer, sheet_name, index = False)
		"""
		for sheet_index, sheet_name, row_index, col_index, col_indexs in self.get_region_infos():
			excel_ori = pd.read_excel(io=self.m_file_path, sheet_name=sheet_name)
			a = excel_ori.values
			select = a[:, col_index]
			a = a[np.where((select == "青岛办事处") | (select == "上海客户一部") | (select == "上海客户二部") | (select == "上海客户三部") | (select == "南京客户一部") | (select == "南京客户二部") | (select == "南京客户三部") | (select == "山东客户一部") | (select == "山东客户二部") | (select == "商务定制部") | (select == "苏州客户部") | (select == "苏北客户部") | (select == "青岛办事处"))]
			data_df = pd.DataFrame(data=a, copy=True)
			columns = [index.get(CParseExcel.VALUE) for index in col_indexs]
			data_df.columns = columns
			data_df.to_excel(writer, sheet_name, index = False)
		writer.save()


class CGetHeaderInfos(CReadInfo):
	def __init__(self, file_path, region_infos):
		CReadInfo.__init__(self, file_path)
		self.m_region_infos = region_infos
		self.m_header_infos = {}

	def col_values(self, sheet_index, sheet_name, row_index, col_indexs):
		if len(self.m_region_infos) - 1 < sheet_index:
			return
		region_info = self.m_region_infos[sheet_index]
		if row_index < region_info[2]:
			info = self.m_header_infos.get(sheet_index)
			if info is not None:
				# exist
				info["data"].append(col_indexs)
			else:
				# not exist
				info = {}
				info["sheet_name"] = sheet_name
				li = []
				li.append(col_indexs)
				info["data"] = li
			self.m_header_infos[sheet_index] = info

	def get_header_infos(self):
		return self.m_header_infos


class CDeleteRegions(CReadInfo):
	def __init__(self, file_path, region_infos):
		CReadInfo.__init__(self, file_path)
		self.m_region_infos = region_infos
		self.m_delete_indexs = {}

	def __find_region_info(self, sheet_index):
		for info in self.m_region_infos:
			if info[0] == sheet_index:
				return info

	def col_values(self, sheet_index, row_index, col_values):
		region_info = self.__find_region_info(sheet_index)
		reg_row_next = region_info[1] + 1
		reg_col_index = region_info[2]
		li = self.m_delete_indexs.get(sheet_index)
		if li is None:
			li = []
		if row_index > reg_row_next:
			if col_values[reg_col_index].get(CParseExcel.VALUE) not in ["上海客户一部", "上海客户二部", "上海客户三部", "南京客户一部", "南京客户二部", "南京客户三部", "山东客户一部", "山东客户二部", "商务定制部", "苏州客户部", "苏北客户部", "青岛办事处"]:
				li.append(row_index)
		self.m_delete_indexs[sheet_index] = li

	def delete(self, obj_path):
		self.delete_multi_row(self.m_delete_indexs, obj_path)


if __name__ == "__main__":
	reader = CFindRegionFieldByTitle("./file/test1.xlsx")
	reader.read()
	reader.gen("./obj/test1.xlsx")
	# reader = CDeleteRegions("./file/test1.xlsx", reader.get_region_infos())
	# reader.read()
	# reader.delete("./obj/test1.xlsx")
	# print(reader.get_row_datas())
