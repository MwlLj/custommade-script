from read_info import CFindRegionFieldByTitle
import os
import sqlite3
import time


class CCheck(object):
	def __init__(self, check_path):
		self.m_check_path = check_path
		self.m_dbname = "record.db"
		conn = sqlite3.connect(self.m_dbname)
		c = conn.cursor()
		c.execute("create table if not exists path_info(path varchar(256));")
		conn.commit()
		conn.close()
		while True:
			self.check()
			time.sleep(3)

	def check(self):
		dirs = os.listdir(self.m_check_path)
		for d in dirs:
			path = os.path.join(self.m_check_path, d)
			is_file = os.path.isfile(path)
			if is_file is True:
				is_exist = self.path_is_exist(path)
				if is_exist is False:
					self.write_db(path)
					dir_path = path + ".dir"
					if os.path.exists(dir_path) is False:
						os.makedirs(dir_path)
					reader = CFindRegionFieldByTitle(path)
					reader.read()
					reader.gen(os.path.join(dir_path, d))

	def path_is_exist(self, path):
		conn = sqlite3.connect(self.m_dbname)
		c = conn.cursor()
		cursor = c.execute("""
			select count(*) from path_info where path = "{0}";
			""".format(path))
		for row in cursor:
			count = row[0]
			if count > 0:
				conn.close()
				return True
		conn.close()
		return False

	def write_db(self, path):
		conn = sqlite3.connect(self.m_dbname)
		c = conn.cursor()
		c.execute("""insert into path_info values("{0}");""".format(path))
		conn.commit()
		conn.close()


if __name__ == "__main__":
	obj_dir = "./workspace"
	if os.path.exists(obj_dir) is False:
		os.makedirs(obj_dir)
	checker = CCheck(obj_dir)
	checker.check()

