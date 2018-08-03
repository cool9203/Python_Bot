# Python_Bot

```python
class MySql():
	def __init__(self):
		pass


	def __init__(self, ip, uid, pwd, dbName, charset):
		self.SetConnectionString(ip, uid, pwd, dbName, charset)
		self.Connect()


	def __del__(self):
		self.db.close()


	def __enter__(self):
		pass


	def __exit__(self, type, msg, traceback):
		self.__del__()
		return False


	def SetConnectionString(self, ip, uid, pwd, dbName, charset):
		self.ip = ip
		self.uid = uid
		self.pwd = pwd
		self.dbName = dbName
		self.charset = charset
		self.ConnectString = r"Server=" + ip + " ; port=3306 ; database=" + dbName + " ; user id=" + uid + " ; password=" + pwd + " ; charset=" + charset + " ;"


	def GetConnectString(self):
		return self.ConnectString


	def Connect(self):
		try:
			self.db = MySQLdb.connect(self.ip, self.uid, self.pwd, self.dbName, charset = self.charset)
			self.cursor=self.db.cursor()
			self.DllLink()
		except Exception as e:
			raise e


	def SetCommandString(self, commandString):
		self.commandString = commandString


	def Select(self, sqlStr):
		self.cursor.execute(sqlStr)
		return self.cursor.fetchall()


	def Insert(self, tableName, columnName, tempValues, *primaryNumber):
		self.DllCommand(r"SELECT * FROM " + tableName)
		values = MyList()
		if (len(primaryNumber) == 0):
			values = self.CheckRepeat(columnName, tempValues)
		else:
			values = self.CheckRepeat(columnName, tempValues, primaryNumber[0])
		if (len(values) != 0):
			sqlStr = self.GetInsertString(tableName, columnName, values)
			try:
				self.cursor.execute(sqlStr)
				self.db.commit()
			except Exception as e:
				raise e
		

	def GetInsertString(self, tableName, columnName, values):
		string = r"INSERT INTO " + tableName + " ("
		for index in range(len(columnName)):
			string += columnName.Get(index)
			if (index != len(columnName) - 1):
				string += ','

		string += ") VALUES "

		for rowIndex in range(int((len(values) / len(columnName)))):
			string += "("
			for columnIndex in range(len(columnName)):
				if (len(values.Get((rowIndex * len(columnName)) + columnIndex)) == 0):
					string += "null"
				else:
					string += r"'" + str(values.Get((rowIndex * len(columnName)) + columnIndex)) + r"'"
				if (columnIndex != len(columnName) - 1):
					string+=','
			string += ")"
			if (rowIndex != (len(values) / len(columnName)) - 1):
				string += ','
		return string


	def CheckRepeat(self, columnName, tempValues, *primaryNumber):
		values = MyList()
		columnCountLimit = 0
		if (len(primaryNumber) == 0):
			columnCountLimit = len(columnName)
		else:
			columnCountLimit = primaryNumber[0]

		for rowCount in range(int(len(tempValues) / len(columnName))):
			string = ""
			for index in range(int(columnCountLimit)):
				string += str(columnName.Get(index)) + r"='" + str(tempValues.Get(rowCount * len(columnName) + index)) + r"'"
				if (index != columnCountLimit - 1):
					string += r" AND "
			if (self.DllSelect(string) == 0):
				for index in range(len(columnName)):
					values.Push(tempValues.Get(rowCount * len(columnName) + index))
		#self.Write("values.txt", values.Get())
		return values


	def DllLink(self):
		self.dll = Dispatch("ToPython.Application")
		self.dll.SetConnectionString(self.ConnectString)


	def DllCommand(self, dllCommandString):
		result = self.dll.Command(dllCommandString)
		if result == False:
			print("command failed..")
			return False


	def DllSelect(self, dllSelectString):
		result = self.dll.Select(dllSelectString)
		if result == -1:
			print("get Exception..")
			self.Write("DllSelectError.txt",dllSelectString)
		elif result == -2:
			print("get some i don't error....")
			self.Write("DllSelectError.txt",dllSelectString)
		else:
			return result


	#寫入基本檔案，不是DB
	def Write(self, fileName, data):
		try:
			f = open(fileName,'w')
			f.write(str(data))
			f.close()
		except:
			traceback.print_exc()
			wait(20)
			sys.exit()

```
