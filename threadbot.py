# encoding: utf-8
import requests #可以讀和送HttpRequests
import time
from bs4 import BeautifulSoup #HTML的解析器
import re #正規表示式 使用方法 re.compile("要正規的字")
import lxml #xml格式
import sys #systemㄉ
import traceback #debug在用的，可以看堆疊出的程式錯在哪
import os #作業系統
import threading
import MySQLdb
from win32com.client import Dispatch


#實作list型態的多參數傳入
class MyList():
	def __init__(self):
		pass


	def __init__(self, *data):
		self.__list = []
		for tempData in data:
			self.__Push(tempData)


	def __del__(self):
		del self.__list


	def __enter__(self):
		pass


	def __exit__(self, type, msg, traceback):
		self.__del__()
		return False


	def __len__(self):
		return len(self.__list)


	#允許呼叫多參數傳入
	def Push(self, *data):
		for tempData in data:
			self.__Push(tempData)


	#內部的單參數輸入
	def __Push(self, data):
		self.__list.append(data)


	#因為沒有多載(overloading)，所以自行實作類似多載的概念
	#有參數傳入：回傳該參數的指向的位置的內容，若大於list的長度，回傳'none'
	#無參數傳入：回傳整個list
	def Get(self, *assign):
		if (len(assign) == 0):
			return self.__list
		elif (assign[0] >= len(self.__list)):
			return 'none'
		return self.__list[assign[0]]


	#因為沒有多載(overloading)，所以自行實作類似多載的概念
	#有參數傳入：回傳該參數的指向的位置的內容並刪除掉，若大於list的長度，回傳'none'
	#無參數傳入：回傳list的最後一個並刪除掉
	def Pop(self, *assign):
		if (len(assign) == 0):
			if (len(self.__list) == 0):
				return 'none'
			return self.__list.pop()
		elif (assign[0]>=len(self.__list)):
			return 'none'
		return self.__list.pop(assign[0])



#實作簡單的對SQL的處理
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
	


class MyBot():
	#自訂Exception
	class ModelError(Exception):
		pass
	class UrlError(Exception):
		pass


	def __init__(self):
		self.sql = MySql("127.0.0.1", "root", "root", "mydb", "utf8")
		


	def __enter__(self):
		self.InformationPage()


	def __exit__(self, type, msg, traceback):
		return False


	def SetUrl(self, url):
		self.url = url
	

	def IndexPage(self, url):
		try:
			#set proxy
			proxies = {
				"http":"http://10.56.69.199:8080",
				"https":"https://10.56.69.199:8080",
			}
			#get http requests
			req = requests.get(url, proxies = proxies)
			if req.status_code != 200: #如果沒有連線上 6成都是url錯，剩下的就是網路或WebServer的問題
				raise UrlError("input URL error or internet and WebServer no open and not link")

			#給HTML解析器解析
			soup = BeautifulSoup(req.text, 'lxml')
			rowsData = soup.find_all('a',href = re.compile("/s/list"))
			tempList = MyList()
			values = MyList()
			count = 0
			for data in rowsData:
				tempList.Push(r"https://tw.stock.yahoo.com" + data["href"])
			return tempList

		except Exception as e:
			raise e


	#撈資訊的bot
	def InformationPage(self):
		try:
			#set proxy
			proxies = {
				"http":"http://10.56.69.199:8080",
				"https":"https://10.56.69.199:8080",
			}
			
			#get http requests
			req = requests.get(self.url, proxies = proxies)
			if req.status_code != 200: #如果沒有連線上 6成都是url錯，剩下的就是網路或WebServer的問題
				raise UrlError("input URL error or internet and WebServer no open and not link")
				wait(20)
				sys.exit()

			#給HTML解析器解析
			soup = BeautifulSoup(req.text, 'lxml')
			#拿出要爬的東西 - 資料
			rowsData = soup.find_all('td', align=re.compile("center"), nowrap=re.compile(""))
			#拿出要爬的東西 - 欄位名稱
			columnName = []
			temp = soup.find_all('th', nowrap=re.compile(""))
			#先insert公司的資訊進db裡
			self.InsertCompany(rowsData, temp)
			#多塞一個欄位進去columnName
			for index in range(len(temp)):
				data = temp[index]
				if (index == 1):
					columnName.append("日期")
				columnName.append(data.text)

			#把資訊放到self.datalist裡
			self.Information(rowsData, columnName)
			values = MyList()
			for row in self.dataList:
				values.Push(row)
			try:
				#要自己打欄位名稱，
				self.sql.Insert('information', MyList('ID','Date','Time','Transaction','Buy','Sale','Rise_Fall','Number','Close','Open','High','Low'), values, 3)
			except Exception as e:
				
				#整理要寫入txt的字串(這邊算是測試拿出來的資料和是否是要的東西，可當作log檔來存)
				string = ""
				for index in range(len(self.dataList)):
					if index % len(columnName) == 0 and index != 0:
						string+='\n'
					data = self.dataList[index]
					string += columnName[index % len(columnName)] + ':' + data + '\n'
				#寫入檔案
				self.Write("error_data.txt",string)
				
				raise e

		except:
			traceback.print_exc()
			wait(20)
			sys.exit()


	#寫入公司資料到company的table裡
	def InsertCompany(self, rowsData, columnName):
		values = MyList()
		for index in range(len(rowsData)):
			data = rowsData[index]
			if (index % len(columnName) == 0):
				values.Push(self.Split(data.text, 1))
				values.Push(self.Split(data.text, 2))
		try:
			self.sql.Insert('company',MyList('ID', 'Name'), values)
		except Exception as e:
			raise e
		


	#把拿到的資料給整理好
	def Information(self, rowsData, columnName):
		self.dataList = []
		for index in range(len(rowsData)):
			data = rowsData[index]

			if index % (len(columnName) - 1) == 0: #是代號時
				self.dataList.append(self.Split(data.text, 1))
			elif index % (len(columnName) - 1) == 1: #這邊需要多加1個日期
				self.dataList.append(time.strftime("%Y%m%d",time.localtime()))
				self.dataList.append(self.ArrangeString(data.text))
			else:
				self.dataList.append(self.ArrangeString(data.text))


	#整理字串，讓她不要有','&':'
	def ArrangeString(self,data):
		tempString=""
		if (data == "－"):
			return ""
		for temp in data:
			if temp == '▽':
				tempString += '-'
			elif temp != ',' and temp != ':' and temp != '△' and temp != '▲' and temp != '▼':
				tempString += temp
		return tempString


	#model:1 get space before.
	#model:2 get space after.
	#else raise:error
	def Split(self, data, model):
		string = ""
		if model == 1:
			for c in data:
				if c == ' ':
					break
				else:
					string += c

		elif model == 2:
			for c in data:
				if c == ' ':
					string = ""
				else:
					string += c

		else:
			raise self.ModelError('Spilt model input error')
		return string


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

	

class Bot(threading.Thread):
	def __init__(self, url):
		threading.Thread.__init__(self)
		self.url = url


	def run(self):
		bot = MyBot()
		bot.SetUrl(self.url)
		bot.InformationPage()



#實作簡單的ThreadQueue
class Queue(threading.Thread):
	def __init__(self):
		threading.Thread.__init__(self)
		self.queue = []
		self.count = 1


	def __len__(self):
		return len(self.queue)


	def run(self):
		while True:
			time.sleep(1)
			if (len(self.queue) == 0):
				continue
			for index in range(len(self.queue)):
				if (index >= len(self.queue)):
					break
				if (self.queue[index].isAlive() == False):
					self.queue.pop(index)
					print(self.count)
					self.count = self.count + 1


	def push(self, data):
		self.queue.append(data)
		self.queue[len(self.queue)-1].start()


	def pop(self):
		self.queue.pop()


	def wait(self):
		while True:
			if (len(self.queue) == 0):
				break



#dleay time
def wait(s):
	print("等待",s,"秒後結束....")
	time.sleep(s)

def main():
	queue = Queue()
	queue.start()
	try:
		indexBot = MyBot()
		urlList = indexBot.IndexPage(r"https://tw.stock.yahoo.com/h/getclass.php#table1")
		for index in range(len(urlList)):
			if (index == 28):
				break

			while True:
				if (len(queue) < 5):
					break
			temp = Bot(urlList.Get(index))
			queue.push(temp)
		
		queue.wait()
		wait(5)
		sys.exit()
	except Exception as e:
		traceback.print_exc()
		wait(20)



#call main function
if __name__=="__main__":
	main()