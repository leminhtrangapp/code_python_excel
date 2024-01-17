from gui_two_custom import Balance_Widget_Gui_Two
from Function_Gui import Balance,Function_File_And_Data,Function_Combobox,Database_Processing
import os
import customtkinter as ctk
import sqlite3
import xlwings as wx
from openpyxl.utils import get_column_letter
import numpy as np
import shutil
from tkinter import messagebox
import ast
# from threading import Thread

import pywintypes

class Command_Gui_One(Balance_Widget_Gui_Two):

	def __init__(self,root):
		super().__init__(root)

		self.case = 0 #--> trương hợp Coppy

		self.name_table_database_undo_redo = ["undo_table","redo_table"]

		 

		self.button_upfile_data.configure(command=self.up_file_data_gui_one) #----> lệnh cho nút upfile data

		self.button_upfile_official.configure(command=self.up_file_of_gui_one) #----> lệnh cho nút upfile of

		self.button_save.configure(command=self.save_file_of)

		self.button_enter.configure(command=self.enter) #--> lệnh xác nhận

		self.button_undo.configure(command=self.command_undo)

		self.button_redo.configure(command=self.command_redo)

		self.combobox_sheet_data.bind("<<ComboboxSelected>>",self.chane_data_sheet) #--> lệnh cho combobox thay đổi data theo sheet chọn

		self.combobox_find_title.bind("<<ComboboxSelected>>",self.value_for_combobox_title) #--> lệnh combobox tìm các data duy nhất của cột đang chọn

		
		self.combobox_find_elemen_title.bind("<<ComboboxSelected>>",self.find_value_unique) #--> tìm các hàng cóa giá trị của cb
		
		self.last_focus = None
		
		self.treeview.bind("<Motion>",self.mycallback)

		root.protocol("WM_DELETE_WINDOW", self.delete_databe_undo_redo_when_quit_app)


	



	def delete_databe_undo_redo_when_quit_app(self): # xóa hết db của undo redo khi tắt app

		
		self.root.destroy()

		Database_Processing.delete_redo_table_db(self.database_undo_redo,self.name_table_database_undo_redo)

		self.database_undo_redo.close()
		self.file_database.close()
		self.file_sub_database.close()

		self.file_database_language.close()


	def mycallback(self,event):
		_iid = self.treeview.identify_row(event.y)

		if _iid != self.last_focus:
			if self.last_focus:
				
				self.treeview.item(self.last_focus, tags=[])
			self.treeview.item(_iid, tags=['focus'])
			self.last_focus = _iid

	

	def up_file_data_gui_one(self): #----> chức năng nút upfile data

		self.treeview.bind("<Button-1>", self.edit_treeview) #----> cho phép treeview chọn các cột

		self.path_file_data = Function_File_And_Data.take_path_file(self.value_entry_path_data) #---> lấy đường dẫn file data

		if self.path_file_data !="":

			Database_Processing.delete_redo_table_db(self.database_undo_redo,self.name_table_database_undo_redo) # xóa hết db của undo redo khi tắt app

			self.button_undo.configure(ctk.DISABLED)

			try:

				self.workbook_data =Function_File_And_Data.take_information_file_excel(self.path_file_data,self.combobox_sheet_data) #----> lấy workbook của file data

				name_file_data = Balance.truncate_text(self.workbook_data.name,10) #--> tên file data

				self.name_table = self.combobox_sheet_data.get().replace(" ", "_")

				

				self.label_name_file_data.configure(text=name_file_data) #--> đổi text label thành tên file

				self.lis_sheet_data = self.workbook_data.sheets


				self.lis_sheet_range_data = Function_File_And_Data.take_data_address(self.lis_sheet_data) #--> lấy các vùng chỉ có data

				Function_File_And_Data.push_value_database(self.lis_sheet_range_data,self.file_database) # --> đưa data của excel và database của sqlite

				shutil.copy(self.path_database, self.path_sub_database) # tạo ngẵn stack db

				self.lis_row_databse_hiden = [[] for _ in range(len(self.lis_sheet_range_data))] #--> tạo danh sách chứa các hàng không lấy của từng sheets

				lis_row_hiden = [item for sublist in self.lis_row_databse_hiden[self.combobox_sheet_data.current()] for item in sublist] #--> tạo list các hàng cần ẩn từ danh sách ẩn


				self.rowid = Function_File_And_Data.push_value_treeveiw(self.treeview,self.name_table,

				self.combobox_find_title,self.combobox_find_elemen_title,self.value_spin_box_left,

				self.value_spin_box_right,self.value_spin_box_top,self.value_spin_box_bottom,lis_row_hiden,self.file_database,self.notification_gui_one13,self.text_all_of_combobox)

				self.row_title = self.rowid[0] #--> lấy vị trí rowid của tiêu đề có lúc cần đến

				self.sub_rowd = self.rowid # --> tạo bản sao của rowid có lúc cần đến

				text_heading_treeveiw_check = self.treeview.heading("#1")["text"] #--> lấy text của tiêu đề
				

				if len(self.rowid) <=1 and text_heading_treeveiw_check ==self.notification_gui_one13 or len(self.rowid) <=1 and text_heading_treeveiw_check == "Không Có Data":
					# đặt điều kiện nếu tiêu đề trùng với tiêu đề loại bỏ thì đóng nút enter
					self.button_enter.configure(state=ctk.DISABLED)
				else:
					self.button_enter.configure(state=ctk.NORMAL)

				if self.switch_var_notification.get() =="on":
				
					self.list_box_notification.insert(0,f"File data '{self.workbook_data.name}' {self.notification_gui_one14}")

				self.last_focus = None
				
				
			except pywintypes.com_error as e:

				if e.args[0] == -2146959355:

					self.list_box_notification.insert(0,self.notification_gui_one1)
			
		else:

			if self.switch_var_notification.get() =="on":
			
				self.list_box_notification.insert(0,self.notification_gui_one2)

		

		self.curent_present_title = self.combobox_find_title.current()

		self.curent_present_elemen_title = self.combobox_find_elemen_title.current()


	def up_file_of_gui_one(self): #----> chức năng nút upfile of

		path_file = Function_File_And_Data.take_path_file(self.value_entry_path_official)



		if path_file !="":

			try:

				Database_Processing.delete_redo_table_db(self.database_undo_redo,self.name_table_database_undo_redo) # xóa hết db của undo redo khi tắt app

				self.button_undo.configure(ctk.DISABLED)

				self.workbook_of =Function_File_And_Data.take_information_file_excel(path_file,self.combobox_sheet_official)

				name_file_of = Balance.truncate_text(self.workbook_of.name,10)

				self.label_name_file_of.configure(text=name_file_of)

				self.lis_sheet_of = self.workbook_of.sheets

				if self.switch_var_notification.get() =="on":

					self.list_box_notification.insert(0,f"File official '{self.workbook_of.name}' {self.notification_gui_one14}")
			except pywintypes.com_error as e:

				if e.args[0] == -2146959355:

					self.list_box_notification.insert(0,self.notification_gui_one1)
		else:

			if self.switch_var_notification.get() =="on":
			
				self.list_box_notification.insert(0,self.notification_gui_one3)


	def save_file_of(self):

		try:

			path_save =Function_File_And_Data.save_file_excel(self.workbook_of.name,self.value_entry_path_save,self.workbook_of)
			
		except AttributeError:

			if self.switch_var_notification.get() =="on":

				self.list_box_notification.insert(0,self.notification_gui_one4)

		except pywintypes.com_error as e:
			
			if e.args[0] == -2147417848:

					self.list_box_notification.insert(0,self.notification_gui_one6)
		
		
	def chane_data_sheet(self,event): #--> chức năng combobox đổi data của các sheets

		

		self.name_table = self.combobox_sheet_data.get().replace(" ", "_")
		

		lis_row_hiden = [item for sublist in self.lis_row_databse_hiden[self.combobox_sheet_data.current()] for item in sublist]

	
		
		self.rowid = Function_File_And_Data.push_value_treeveiw(self.treeview,self.name_table,

			self.combobox_find_title,self.combobox_find_elemen_title,self.value_spin_box_left,

			self.value_spin_box_right,self.value_spin_box_top,self.value_spin_box_bottom,lis_row_hiden,self.file_database,self.notification_gui_one13,self.text_all_of_combobox)

		text_heading_treeveiw_check = self.treeview.heading("#1")["text"]

		


		if len(self.rowid) <=1 and text_heading_treeveiw_check ==self.notification_gui_one13 or len(self.rowid) <=1 and text_heading_treeveiw_check == "Không Có Data":

			self.button_enter.configure(state=ctk.DISABLED)

			if self.switch_var_notification.get() =="on":

				self.list_box_notification.insert(0,self.notification_gui_one5)

		else:

			if len(self.rowid)>0:

				self.row_title = self.rowid[0]

			self.sub_rowd = self.rowid

			self.button_enter.configure(state=ctk.NORMAL)

		self.last_focus = None
		
			

	def value_for_combobox_title(self,event): #--> chức năng combobox tìm các value duy của cột để cho combobox elemen 


	
		lis_row_hiden = [item for sublist in self.lis_row_databse_hiden[self.combobox_sheet_data.current()] for item in sublist]


		

		if self.combobox_find_title.get() == self.text_all_of_combobox: #--> nếu tìm data là tất cả thì thay đổi rowid

			self.rowid = Function_Combobox.data_title(self.treeview,self.name_table,

			self.combobox_find_title,self.combobox_find_elemen_title,self.value_spin_box_left,

			self.value_spin_box_right,self.value_spin_box_top,self.value_spin_box_bottom,lis_row_hiden,self.file_database,self.notification_gui_one13,self.text_all_of_combobox)

			

			if len(self.rowid) >0:

				self.row_title = self.rowid[0]

				self.button_enter.configure(state=ctk.NORMAL)

			self.last_focus = None
			
		

		else: # --> còn lại khỏi đổi rowid chỉ thay đổi value của các giá trị tìm của combobox elemen

			Function_Combobox.data_title(self.treeview,self.name_table,

				self.combobox_find_title,self.combobox_find_elemen_title,self.value_spin_box_left,

				self.value_spin_box_right,self.value_spin_box_top,self.value_spin_box_bottom,lis_row_hiden,self.file_database,self.notification_gui_one13,self.text_all_of_combobox)

		self.curent_present_title = self.combobox_find_title.current()

		self.curent_present_elemen_title = self.combobox_find_elemen_title.current()

	def find_value_unique(self,event): # -->chức năng combobox tìm các hàng có value bằng giá trị của cb và hiên trên treeview

		if self.combobox_find_title.get() !="" and self.combobox_find_elemen_title.get() !="":

			lis_row_hiden = [item for sublist in self.lis_row_databse_hiden[self.combobox_sheet_data.current()] for item in sublist]
			
			
			self.rowid_search = Function_Combobox.find_data_title(self.treeview,self.name_table,

			self.combobox_find_title,self.combobox_find_elemen_title,int(self.value_spin_box_left),int(self.value_spin_box_right),

			self.file_database,lis_row_hiden)

			rowid_raw = [self.row_title] + self.rowid_search

			self.rowid = [value for value in self.sub_rowd if value in rowid_raw] # đổi rowid theo các hàng đẫ tìm thấy

			self.curent_present_elemen_title = self.combobox_find_elemen_title.current()
			
			self.last_focus = None
			

	def enter(self):

		
		
		try:


			sht_range_data = self.lis_sheet_range_data[self.combobox_sheet_data.current()] #--> sheets_range data các range


			sht_of = self.lis_sheet_of[self.combobox_sheet_official.get()] #--> sheets of thực hiện
		
		except AttributeError:

			print("Bạn chưa upfie")
		

		except pywintypes.com_error as e:

			if e.args[0] == -2146827864:

				if self.switch_var_notification.get() =="on":

					

					self.list_box_notification.insert(0,self.notification_gui_one6)


		selected_items = self.treeview.selection()

		if len(selected_items) ==0 and self.switch_var_take_title.get() == "off" :

			if self.switch_var_notification.get() =="on":

				if self.label_name_file_data.cget('text') != "Name File":

					self.list_box_notification.insert(0,self.notification_gui_one7)


				else:

					self.list_box_notification.insert(0,self.notification_gui_one8)

		else:

			
	
			Database_Processing.delete_redo_table_db(self.database_undo_redo,[self.name_table_database_undo_redo[1]]) # xóa bảng redo_table của db 


			self.button_redo.configure(state=ctk.DISABLED) # mở khóa undo và xóa bộ nhớ redo

			
			index_choose_of_treeveiw = [self.treeview.index(x) for x in self.treeview.selection()] # các index chọn trên treeview



			index_db = [self.rowid[1:][x] for x in index_choose_of_treeveiw] #--> các vị trí chọn trên treeview xử lý giống hàng của database


			# lis các range chọn đã được xử lý cắt
			lis_row_range = [(sht_range_data.rows[int(x) - 1][int(self.value_spin_box_left) :] if int(self.value_spin_box_right) == 0 else sht_range_data.rows[int(x) - 1][int(self.value_spin_box_left) : -int(self.value_spin_box_right)]) for x in index_db] #---> cắt range
			
			if self.switch_var_take_title.get() == "on": #--> ĐK lấy tiêu đề

				if int(self.value_spin_box_right) == 0: # xử lý cắt hàng tiêu đề giống trên

					title_value = sht_range_data.rows[int(self.value_spin_box_top)][int(self.value_spin_box_left) :]
				else:


					title_value = sht_range_data.rows[int(self.value_spin_box_top)][int(self.value_spin_box_left) : -int(self.value_spin_box_right)]			

				lis_row_range = [title_value] +lis_row_range #--> thêm vùng range của tiêu đề 

				index_db = [self.rowid[0]]+index_db #--> công thêm hàng tiêu đề cho db


			#----- thực hiện việc kiểm tra hàng
			columns_treeview = self.treeview["columns"] 

			headings = [self.treeview.heading(col)["text"] for col in columns_treeview] # ----> lấy value của tiêu đề để kiểm tra

			check_choose_columns = any(num[0] =="\u2611" for num in headings) # ------> kiểm tra các hàng đã được chọn

			if check_choose_columns:

			
				column_indices = [self.treeview["columns"].index(col_name)+1 for col_name in self.treeview["columns"] if self.treeview.heading(col_name, "text")[0] == "\u2611"]

			
				value = Database_Processing.take_data_database_columns(self.file_database,self.name_table,self.value_spin_box_left,index_db,column_indices)
			

			else:
				column_indices = []


				value = Database_Processing.take_data_database(self.file_database,self.name_table,self.value_spin_box_left,self.value_spin_box_right,index_db) #--> lấy value của excel từ các range đã chọn




			if self.switch_var_transpose.get() =="on": #--> ĐK transpose value

				value = np.transpose(value).tolist()

			try:
				lr_of = sht_of.range(self.value_column_excel+str(sht_of.cells.last_cell.row)).end('up').row + int(self.value_distance_excel) #--> vùng cuối có value của sheet of

				cloumn_last = get_column_letter(len(value[0])-1+self.number_cloumn_excel)+str(len(value)+lr_of-1) #--> column excel tương ứng độ dài của value

				range_of = sht_of.range(self.value_column_excel+str(lr_of)+":"+cloumn_last) #--> vùng range đền value hoàn chỉnh

				range_of.value = value # --> điền value vào file of

				range_of_address = range_of.address #---> lấy vùng range điền data để dùng cho undo redo
			
				range_choose = [] # --> tạo list rống để đưa vô undo redo không lỗi
				try:

					take_format = self.switch_var_format.get() #--> tạo 1 biến hứng điều kiện lấy định dạng

					if  take_format =="on": #--> ĐK lấy định dạng
						

						wx.apps.active.screen_updating = False # bỏ update màng hình file excel of

						if len(column_indices) <=1: # điều kiện chọn tôi đa 2 cột vì excel ko cho lấy định dạng nhiều cột

							if len(column_indices) ==0:

							

								range_choose = [x.get_address(row_absolute=False,column_absolute=False) for x in lis_row_range]

								

							else:

								range_choose = [x[y-1].get_address(row_absolute=False,column_absolute=False)  for x in lis_row_range for y in column_indices]
						

							try:

								Function_File_And_Data.color_the_data(range_choose,self.workbook_data.sheets[self.combobox_sheet_data.current()],
									self.switch_var_transpose.get(),sht_of,range_of)
							except pywintypes.com_error as e:
								self.list_box_notification.insert(0,e)
								if e.args[0] == -2147352567:
									
									messagebox.showinfo(self.notification_gui_one9[0],self.notification_gui_one9[1])	

							
							
							wx.apps.active.api.CutCopyMode = False

						else:
							messagebox.showinfo(self.notification_gui_one10[0],self.notification_gui_one10[1])	

							take_format = "off" #--> thay đổi lấy định dạng thành không nếu như chọn quá 3 cột

						wx.apps.active.screen_updating = True # mở lại update màng hình
					


					if self.switch_var_wrap_text.get() =='on':
						
						range_of.api.WrapText = True
						
						


				except pywintypes.com_error as e:

					if e.args[0] == -2146827864:

						messagebox.showinfo(self.notification_gui_one11[0],self.notification_gui_one11[1])	

						wx.apps.active.screen_updating = True

						self.switch_var_format = ctk.StringVar(value="off") #--> bỏ dịnh dạng



						if self.switch_var_notification.get() =="on":

							self.list_box_notification.insert(0,self.notification_gui_one12)

				data_title = []

				if self.value_copy_cut == "Cut":

					

					if check_choose_columns: #--> nếu có chọn các cột
						

						self.case = 1

						for item in selected_items:
							
							value_treeview = self.treeview.item(item, "values")

							values = [" " if i + 1 in column_indices else x for i, x in enumerate(value_treeview)]
							
							

							self.treeview.item(item,values=values)

						if self.switch_var_take_title.get() == "on":

							data_title = Database_Processing.query_title_db_case_1_3(self.file_database,self.name_table,self.value_spin_box_top,self.value_spin_box_left,self.value_spin_box_right)
						
							for x in column_indices:
							
								self.treeview.heading('#'+str(x), text=" ")

							text_heading = [self.treeview.heading(x)["text"] for x in self.treeview["columns"]]

							self.combobox_find_title.configure(values =[self.text_all_of_combobox] + text_heading)


						self.treeview.selection_remove(selected_items)

						lis_column_delete = [x+int(self.value_spin_box_left) for x in column_indices]

						Database_Processing.delete_row_columns_database(self.file_database,index_db,lis_column_delete,self.name_table)

					else: #--> nếu không chọn cột

						self.case = 2

						if int(self.value_spin_box_left) ==0  and int(self.value_spin_box_right) ==0: #---> trường hợp nếu bạn không ẩn cột nào thì xử lý cắt luôn chúng khỏi treeiew và ẩn khỏi db

							for item in selected_items:
								self.treeview.delete(item)
							self.lis_row_databse_hiden[self.combobox_sheet_data.current()].append(index_db)

							if self.switch_var_take_title.get() =="on": # nếu có lấy tiêu đề

								try:

									value_heading = self.treeview.item(self.treeview.get_children()[0])['values']
									
									
									for i, heading in enumerate(self.treeview["columns"]): # thay đổi lại value tiêu đề bằng hàng đầu tiên của treeveiw

										self.treeview.heading(heading, text=value_heading[i])

									self.treeview.delete(self.treeview.get_children()[0]) # soa đó xóa hàng đầu tiên để tạo hiệu ứng như thay hàng đầu tiên thành tiêu đề

									self.combobox_find_title.configure(values=[self.text_all_of_combobox] + value_heading) # thây đổi combobox tìm
								
								except IndexError: # trường hợp không còn value

									for i, heading in enumerate(self.treeview["columns"]):

										self.treeview.heading(heading, text=self.notification_gui_one13)					

						else:

							column_indices = [x for x in range(1,len(self.treeview["columns"])+1)]




							self.case = 3

							for item in selected_items:
							
								self.treeview.item(item, values=[''] * (len(self.treeview["columns"]) - 1))

							self.treeview.selection_remove(selected_items)

							having_column = [_ for _ in range(1,len(self.treeview["columns"])+1)]

							
						
							lis_column_hiden = [x+int(self.value_spin_box_left) for x in having_column]

						

							if self.switch_var_take_title.get() == "on":

								data_title = Database_Processing.query_title_db_case_1_3(self.file_database,self.name_table,self.value_spin_box_top,self.value_spin_box_left,self.value_spin_box_right)




								for x in having_column:
									
								
									self.treeview.heading('#'+str(x), text=" ")	

							self.treeview.selection_remove(selected_items)

							text_heading = [self.treeview.heading(x)["text"] for x in self.treeview["columns"]]

							self.combobox_find_title.configure(values =[self.text_all_of_combobox] + text_heading)

							Database_Processing.delete_row_columns_database(self.file_database,index_db,lis_column_hiden,self.name_table)


					self.last_focus = None
					

				this_value_undo = [self.combobox_sheet_data.current(),self.combobox_sheet_official.get(),self.case,self.value_copy_cut,

				index_db,range_of_address,take_format,check_choose_columns,range_choose,

				index_choose_of_treeveiw,self.switch_var_take_title.get(),self.combobox_find_elemen_title.get(),

				self.combobox_find_elemen_title.current(),self.combobox_find_title.get(),

				self.name_table,self.switch_var_transpose.get(),column_indices,

				self.value_spin_box_left,self.value_spin_box_right,self.value_spin_box_top,self.value_spin_box_bottom,data_title,self.switch_var_wrap_text.get()]
			
			
				str_this_value_undo = tuple([str(value) for value in this_value_undo])


				Database_Processing.push_value_database_undo_redo(self.database_undo_redo,str_this_value_undo,self.name_table_database_undo_redo) # đưa các giá trị undo redo và db


				lis_row_hiden = [item for sublist in self.lis_row_databse_hiden[self.combobox_sheet_data.current()] for item in sublist]

				
				self.rowid =  sorted(list(set(self.rowid) - set(lis_row_hiden)))

				
				if len(self.rowid) ==0:

					self.button_enter.configure(state=ctk.DISABLED)
				else:
					self.row_title = self.rowid[0]
				self.button_undo.configure(state= ctk.NORMAL)

				self.button_redo.configure(state=ctk.DISABLED)


				self.curent_present_title = self.combobox_find_title.current()

				self.curent_present_elemen_title = self.combobox_find_elemen_title.current()
			
				

				wx.apps.active.screen_updating = True
			except UnboundLocalError:

				print("đóng file chính nên không tìm đc biến sht_of")

				Database_Processing.delete_redo_table_db(self.database_undo_redo,self.name_table_database_undo_redo)

				self.list_box_notification.insert(0,'You have renamed the official sheet file or closed it, please re-upload the official file, or change the sheet name to the correct sheet name when uploading the file')

				self.button_undo.configure(state=ctk.DISABLED)
				self.button_redo.configure(state=ctk.DISABLED)

				return

	def command_undo(self):

		self.button_redo.configure(state=ctk.NORMAL)

		self.button_enter.configure(state=ctk.NORMAL)

		value_undo = Database_Processing.transfer_databas_back_and_forth(self.database_undo_redo,self.name_table_database_undo_redo[0],self.name_table_database_undo_redo[1])[0]

		try:

			sht_sheet_data = self.lis_sheet_data[int(value_undo[0])]
		except pywintypes.com_error as e:

			if e.args[0] == -2146827864:

				print("bạn đã đóng file data")

		except pywintypes.com_error as e:

			if e.args[0] == -2147352567:

				self.list_box_notification.insert(0,'You have renamed the official sheet file or closed it, please re-upload the official file, or change the sheet name to the correct sheet name when uploading the file')

				Database_Processing.delete_redo_table_db(self.database_undo_redo,self.name_table_database_undo_redo)
				self.button_undo.configure(state=ctk.DISABLED)
				self.button_redo.configure(state=ctk.DISABLED)


		sht_range_of = self.lis_sheet_of[value_undo[1]]

		case = value_undo[2]

		cut_and_copy = value_undo[3]

		index_db = value_undo[4]

		range_of_address = value_undo[5]

		take_format = value_undo[6]

		check_choose_columns = value_undo[7]

		range_choose = value_undo[8]

		index_choose_of_treeveiw = ast.literal_eval(value_undo[9])

		take_title = value_undo[10]

		value_elemen_title_combobox = value_undo[11]

		current_elemen_title_combobox = value_undo[12]

		value_title_combobox = value_undo[13]

		value_name_sheet_combobox = value_undo[14]

		check_transpose = value_undo[15]


		column_indices = ast.literal_eval(value_undo[16])

		box_left = int(value_undo[17])
		box_right = int(value_undo[18])
		box_top = int(value_undo[19])
		box_bottom = int(value_undo[20])

		data_title = ast.literal_eval(value_undo[21])

		warp_text = value_undo[22]



		if take_format =="on":


			sht_range_of.range(range_of_address).clear()

		else:

			sht_range_of.range(range_of_address).clear_contents()

			if warp_text =="on":
				


				sht_range_of.range(range_of_address).api.WrapText = False

		
		
		if cut_and_copy =="Cut":

			if case =="1" or case == "3":



				data_queryed = Database_Processing.update_again_database(self.file_database,self.file_sub_database,value_name_sheet_combobox,column_indices,index_db[1:-1],int(self.value_spin_box_left),int(self.value_spin_box_right),box_left)
		
				data_treeview = data_queryed


			
				if value_name_sheet_combobox == self.combobox_sheet_data.get().replace(" ", "_"):

					if value_elemen_title_combobox == self.combobox_find_elemen_title.get() and int(current_elemen_title_combobox) == self.combobox_find_elemen_title.current():

						if take_title =="on":



							data_treeview = data_queryed[1:]


							data_queryed_title = data_title


							for i, heading_id in enumerate(self.treeview['columns']):

								self.treeview.heading(heading_id, text=data_queryed_title[i])

							self.combobox_find_title.configure(values = [self.text_all_of_combobox]+list(data_queryed_title))



					
						lis_item = [self.treeview.get_children()[x] for x in index_choose_of_treeveiw]

					

						for i,item in enumerate(lis_item):
							

							self.treeview.item(item, values=(data_treeview[i]))

					else:

						if take_title =="on":

							data_queryed_title = data_queryed[0]

							for i, heading_id in enumerate(self.treeview['columns']):

								self.treeview.heading(heading_id, text=data_queryed_title[i])

							self.combobox_find_title.configure(values = [self.text_all_of_combobox]+list(data_queryed_title))

							self.combobox_find_title.current(self.curent_present_title)

						self.value_for_combobox_title(event=None)

						self.combobox_find_elemen_title.current(self.curent_present_elemen_title)

						if self.combobox_find_title.get() != self.text_all_of_combobox:
						
							self.find_value_unique(event=None)
						else:

							self.chane_data_sheet(event=None)
				
						
						
			

			if case =="2":
			
				self.lis_row_databse_hiden[int(value_undo[0])] = [x for x in self.lis_row_databse_hiden[int(value_undo[0])] if x != ast.literal_eval(index_db)] # loại bỏ các row

				self.rowid.extend(ast.literal_eval(index_db)) #----> trả lại các hàng đã ẩn
					
				self.rowid.sort()#----> sắp sếp lại các rowid

				self.row_title = self.rowid[0]	

				if value_name_sheet_combobox == self.combobox_sheet_data.get().replace(" ", "_"):
					data_queryed = Database_Processing.query_index_db(self.file_database,value_name_sheet_combobox,self.value_spin_box_left,self.value_spin_box_right,index_db[1:-1])

					data_queryed_title = data_queryed[0]
					if value_elemen_title_combobox == self.combobox_find_elemen_title.get() and int(current_elemen_title_combobox) == self.combobox_find_elemen_title.current():
						
						
						data_treeview = data_queryed
						if take_title =="on":

							
							
							data_treeview = data_queryed[1:]


							columns_treeview = self.treeview["columns"]

							headings = tuple([self.treeview.heading(col)["text"] for col in columns_treeview])

							if headings[0] != self.notification_gui_one13:

								self.treeview.insert("","0", values=headings)


							for x,y in enumerate(self.treeview["columns"]): #----> đổi tiêu đề lại

								self.treeview.heading(y,text=data_queryed_title[x])

							self.combobox_find_title.configure(values =[self.text_all_of_combobox] + list(data_queryed_title)) #----> thây đổi value của combobox tiêu đề
						
							self.combobox_find_title.current(self.curent_present_title)


						for index,value in zip(index_choose_of_treeveiw,data_treeview): #-----> chèn lại value trừ cái tiêu đề
						
							self.treeview.insert('', index, values=value)


						self.last_focus = None
								
					else:

						if take_title =="on":

							

							columns_treeview = self.treeview["columns"]

							headings = tuple([self.treeview.heading(col)["text"] for col in columns_treeview])

							if headings[0] != self.notification_gui_one13:

								self.treeview.insert("","0", values=headings)


							for x,y in enumerate(self.treeview["columns"]): #----> đổi tiêu đề lại

								self.treeview.heading(y,text=data_queryed_title[x])

							self.combobox_find_title.configure(values =[self.text_all_of_combobox] + list(data_queryed_title)) #----> thây đổi value của combobox tiêu đề
						

							self.combobox_find_title.current(self.curent_present_title)

						self.value_for_combobox_title(event=None)

						self.combobox_find_elemen_title.current(self.curent_present_elemen_title)

						if self.combobox_find_title.get() != self.text_all_of_combobox:
						
							self.find_value_unique(event=None)
						else:

							self.chane_data_sheet(event=None)
				

		len_data_table =  Database_Processing.check_data_existence(self.database_undo_redo,self.name_table_database_undo_redo[0])


		if len_data_table[0] ==0:

			self.button_undo.configure(state=ctk.DISABLED)

	def command_redo(self):

		self.button_undo.configure(state=ctk.NORMAL)

		value_redo = Database_Processing.transfer_databas_back_and_forth(self.database_undo_redo,self.name_table_database_undo_redo[1],self.name_table_database_undo_redo[0])[0]
		
		
		try:

			sht_range_of = self.lis_sheet_of[value_redo[1]]

			case = value_redo[2]

			cut_and_copy = value_redo[3]

			index_db = value_redo[4]

			range_of_address = value_redo[5]

			take_format = value_redo[6]

			check_choose_columns = value_redo[7]

			range_choose = value_redo[8]

			index_choose_of_treeveiw = ast.literal_eval(value_redo[9])

			take_title = value_redo[10]

			value_elemen_title_combobox = value_redo[11]

			current_elemen_title_combobox = value_redo[12]

			value_title_combobox = value_redo[13]

			value_name_sheet_combobox = value_redo[14]

			check_transpose = value_redo[15]


			column_indices = ast.literal_eval(value_redo[16])

			box_left = int(value_redo[17])
			box_right = int(value_redo[18])
			box_top = int(value_redo[19])
			box_bottom = int(value_redo[20])

			# data_title = ast.literal_eval(value_undo[21])

			warp_text = value_redo[22]



			if check_choose_columns == "True":

				indices = [int(x) for x in column_indices]

				values = Database_Processing.take_data_database_columns(self.file_database,value_name_sheet_combobox,box_left,ast.literal_eval(index_db),indices)
				
			else:
				values = Database_Processing.query_index_db(self.file_database,value_name_sheet_combobox,box_left,box_right,index_db[1:-1])

			if check_transpose =="on": #--> nếu có chọn transpose

				values = np.transpose(values).tolist()

			sht_range_of.range(range_of_address).value = values

			if warp_text == 'on':

				sht_range_of.range(range_of_address).api.WrapText = True

			try:

				if take_format == "on":

					sht_sheet_data = self.lis_sheet_data[int(value_redo[0])]

					wx.apps.active.screen_updating = False
				
					lisst_address = ast.literal_eval(range_choose) #--> chuyển đổi thành lits

					Function_File_And_Data.color_the_data(lisst_address,sht_sheet_data,check_transpose,sht_range_of,sht_range_of.range(range_of_address))		

					wx.apps.active.screen_updating = True

					wx.apps.active.api.CutCopyMode = False

			except pywintypes.com_error as e:

				if e.args[0] == -2146827864:

					print("bạn đã đóng file data")


			if cut_and_copy =="Cut":

				if case == "1" or case =="3":

					lis_column_delete = [x+int(box_left) for x in column_indices]

					

					Database_Processing.delete_row_columns_database(self.file_database,ast.literal_eval(index_db),lis_column_delete,value_name_sheet_combobox)

					if value_name_sheet_combobox == self.combobox_sheet_data.get().replace(" ", "_"):

						data_queryed =Database_Processing.query_index_db(self.file_database,value_name_sheet_combobox,int(self.value_spin_box_left),int(self.value_spin_box_right),index_db[1:-1])
						
						data_treeview = data_queryed
						
						if value_elemen_title_combobox == self.combobox_find_elemen_title.get() and int(current_elemen_title_combobox) == self.combobox_find_elemen_title.current():

							
							
							lis_item = [self.treeview.get_children()[x] for x in index_choose_of_treeveiw]
							
							if take_title =="on":

								for x in column_indices:

									
									item = self.treeview['columns'][x-1]


									self.treeview.heading(item, text=" ")

								value_combobox = [self.treeview.heading(item)['text'] for item in self.treeview["columns"]]

								self.combobox_find_title.configure(values = [self.text_all_of_combobox]+list(value_combobox))

								data_treeview = data_queryed[1:]

						
							lis_item = [self.treeview.get_children()[x] for x in index_choose_of_treeveiw]

						

							for i,item in enumerate(lis_item):
								

								self.treeview.item(item, values=(data_treeview[i]))

						else:

							if take_title =="on":

								for x in column_indices:

									
									item = self.treeview['columns'][x-1]

									self.treeview.heading(item, text=" ")

								value_combobox = [self.treeview.heading(item)['text'] for item in self.treeview["columns"]]

								self.combobox_find_title.configure(values = [self.text_all_of_combobox]+list(value_combobox))


								self.combobox_find_title.current(self.curent_present_title)

							self.value_for_combobox_title(event=None)

							self.combobox_find_elemen_title.current(self.curent_present_elemen_title)
							
							if self.combobox_find_title.get() != self.text_all_of_combobox:
							
								self.find_value_unique(event=None)
							else:

								self.chane_data_sheet(event=None)						

				if case =="2":				

					self.lis_row_databse_hiden[int(value_redo[0])].append(ast.literal_eval(index_db))


					lis_row_hiden = [item for sublist in self.lis_row_databse_hiden[int(value_redo[0])] for item in sublist]



					
					if value_name_sheet_combobox == self.combobox_sheet_data.get().replace(" ", "_"):
						
		
						if value_elemen_title_combobox == self.combobox_find_elemen_title.get() and int(current_elemen_title_combobox) == self.combobox_find_elemen_title.current():
							
							lis_item = [self.treeview.get_children()[x] for x in index_choose_of_treeveiw]

							for item in lis_item:

								
								self.treeview.delete(item)


							if take_title == "on":

								try:
								
									data_queryed_title = Database_Processing.query_title_db(self.file_database,value_name_sheet_combobox,lis_row_hiden,self.value_spin_box_left,self.value_spin_box_right)
									
									

									for i, heading in enumerate(self.treeview["columns"]):

										self.treeview.heading(heading, text=data_queryed_title[i]) # đổi tiêu đề thành giá trị của hàng đàu tiên

									self.treeview.delete(self.treeview.get_children()[0]) # xóa hàng đầu


									self.combobox_find_title.configure(values =[self.text_all_of_combobox] + list(data_queryed_title))

									self.combobox_find_title.current(self.curent_present_title)

								except IndexError:

									for i, heading in enumerate(self.treeview["columns"]):
										self.treeview.heading(heading, text=self.notification_gui_one13)

							self.last_focus = None
							
							
							
						else:

							if take_title =="on":

								data_queryed_title = Database_Processing.query_title_db(self.file_database,value_name_sheet_combobox,lis_row_hiden,self.value_spin_box_left,self.value_spin_box_right)

								for i, heading in enumerate(self.treeview["columns"]):

									self.treeview.heading(heading, text=data_queryed_title[i]) # đổi tiêu đề thành giá trị của hàng đàu tiên

								self.treeview.delete(self.treeview.get_children()[0]) # xóa hàng đầu


								self.combobox_find_title.configure(values =[self.text_all_of_combobox] + list(data_queryed_title))

								self.combobox_find_title.current(self.curent_present_title)

							self.value_for_combobox_title(event=None)

							self.combobox_find_elemen_title.current(self.curent_present_elemen_title)
							
							if self.combobox_find_title.get() != self.text_all_of_combobox:
							
								self.find_value_unique(event=None)
							else:

								self.chane_data_sheet(event=None)
								
					
					self.rowid =  sorted(list(set(self.rowid) - set(lis_row_hiden)))

					if len(self.rowid) ==0:

						self.button_enter.configure(state=ctk.DISABLED)

				
			len_data_table =  Database_Processing.check_data_existence(self.database_undo_redo,self.name_table_database_undo_redo[1])


			if len_data_table[0] ==0:

				self.button_redo.configure(state=ctk.DISABLED)


		except pywintypes.com_error as e:

			if e.args[0] == -2147352567:

				self.list_box_notification.insert(0,'You have renamed the official sheet file or closed it, please re-upload the official file, or change the sheet name to the correct sheet name when uploading the file')

				Database_Processing.delete_redo_table_db(self.database_undo_redo,self.name_table_database_undo_redo)
				self.button_undo.configure(state=ctk.DISABLED)
				self.button_redo.configure(state=ctk.DISABLED)
