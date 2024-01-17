from tkinter import *


from tkinter import filedialog

import xlwings as wx

# import sqlite3

import os

import numpy as np

from pathlib import Path


class Balance:


	def function_balance(main,lis,wei_,min_):

		for widget in lis:

			Grid.rowconfigure(main, widget, weight=wei_, minsize=min_)
			Grid.columnconfigure(main, widget, weight=wei_, minsize=min_)

	def function_balance2(main,widget):

		Grid.rowconfigure(main, widget, weight=1, minsize=0)
		Grid.columnconfigure(main, widget, weight=1, minsize=0)


	def center_window(window, width, height):
		screen_width = window.winfo_screenwidth()
		screen_height = window.winfo_screenheight()
		
		x = (screen_width - width) // 2
		y = (screen_height - height) // 2
		
		window.geometry(f"{width}x{height}+{x}+{y}")

		# tap.configure(width=width,height=height)


	def truncate_text(text, max_length):

		tail_file = Path(text).suffix

		if len(text) > max_length:
			return text[:max_length] + "_"+tail_file
		else:
			return text


	def take_width_and_height(geometry):

		plus_index = geometry.find('+')



		result = geometry[:plus_index].split("x")


		return result



		

class Setting_Entry:



	def default_data(entry,data):

		if entry.get() == "" or entry.get() !=data:

			entry.delete(0,END)

			entry.insert(0,data)
	

class Function_File_And_Data:

	def save_file_excel(name_file,path,worbook):

		file_path_save = filedialog.asksaveasfilename(filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsm")],title="Save File Excel",initialdir=path,initialfile=name_file)

		worbook.save(file_path_save)

	def take_path_file(path):

		file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsm")],initialdir=path,title="Choose File Excel")

		return file


	def value_treeveiw(treeview,all_data):

		treeview.delete(*treeview.get_children())
		


		columns = ["columns"+str(x) for x in range(0,len(all_data[0]))]

		

		treeview["columns"] = columns




		minsize = treeview.winfo_width()//len(treeview['columns'])

		
		for col_name,name in zip(treeview["columns"],all_data[0]):


			treeview.heading(col_name, text=name,anchor="w")

			treeview.column(col_name, stretch=False,minwidth=minsize)	

		for values in all_data[1:]:

			
					
			treeview.insert("", "end", values=values)


	def take_information_file_excel(path,combobox_sheet):

		wb = wx.Book(path)



		list_name_sheet = wb.sheet_names
		
		combobox_sheet.configure(values = list_name_sheet)

		combobox_sheet.current(0)

		return wb


	def take_data_address(lis_sheets):

		lis_range = []

		for sheet in lis_sheets:
				

			for x in sheet.used_range.rows:


				if x.value ==None:

					lis_range.append(sheet.range(x,sheet.used_range.columns[-1].rows[-1]))



				elif any(value is not None and value != " " for value in x.value):



				

					lis_range.append(sheet.range(x,sheet.used_range.columns[-1].rows[-1]))

					break
		
		return lis_range

	def create_value_database_language(file_database_language,name_table,lis_language):
		
		cursor = file_database_language.cursor()

		number_columns = len(lis_language[0])


		

		command_vreate_table = f''' CREATE TABLE IF NOT EXISTS {name_table[0]}( 

			
			{','.join(f'column{i} TEXT' for i in range(1,number_columns+1))});'''


		command_vreate_table2 = f''' CREATE TABLE IF NOT EXISTS {name_table[1]} (column1 TEXT);'''

		cursor.execute(command_vreate_table)

		cursor.execute(command_vreate_table2)

	def push_value_database_language(file_database_language,name_table,lis_language):

		cursor = file_database_language.cursor()

		number_columns = len(lis_language[0])


		matrix = tuple(tuple(str(item) if item is not None else " " for item in row)for row in lis_language)

		insert_value = f'INSERT INTO {name_table[0]} ({",".join(f"column{i}" for i in range(1,number_columns+1))}) VALUES ({", ".join("?" for _ in range(number_columns))})'

		cursor.executemany(insert_value,matrix)


		file_database_language.commit()

		insett_choose = f'INSERT INTO {name_table[1]} (column1) VALUES ("english")'

		cursor.execute(insett_choose)

		file_database_language.commit()


	def query_language_choose(file_database_language,name_table):

		cursor = file_database_language.cursor()

		query_choose = f'''SELECT * FROM {name_table}'''

		cursor.execute(query_choose)

		language_choose = cursor.fetchall()[0]

		return language_choose[0]

	def change_language(file_database_language,name_table,new_name):

		cursor = file_database_language.cursor()

		update_data_db = f'''UPDATE {name_table} SET column1 = '{new_name}' WHERE ROWID = 1'''

		cursor.execute(update_data_db)

		file_database_language.commit()

	def query_data_language(file_database_language,name_table,language_query):

		cursor = file_database_language.cursor()

		query_language_choose = f'''SELECT * FROM {name_table[0]} WHERE column1 ="{language_query}"'''

		cursor.execute(query_language_choose)

		language_set = cursor.fetchall()[0][1:]

		return language_set




	def push_value_database(lis_range,file_database):

		#_______________________________delete_all_table

		cursor = file_database.cursor()

		cursor.execute("SELECT name From sqlite_master WHERE type ='table';")

		tables = cursor.fetchall()

		for name_table in tables:

			name = name_table[0]

			command_delete = f' DROP TABLE IF EXISTS {name}'

			cursor.execute(command_delete)



		for range_  in lis_range:

			if range_.value == None:

				number_columns = 1

				matrix = (("Không Có Data",),)
			else:

				
				number_columns = len(range_.value[0])



				matrix = tuple(tuple(item if item is not None else " " for item in row)for row in range_.value)



		
			name_table = range_.sheet.name.replace(" ", "_")

			command_vreate_table = f''' CREATE TABLE IF NOT EXISTS {name_table}( 

			
			{','.join(f'column{i} TEXT' for i in range(1,number_columns+1))});'''


			cursor.execute(command_vreate_table)



			insert_value = f'INSERT INTO {name_table} ({",".join(f"column{i}" for i in range(1,number_columns+1))}) VALUES ({", ".join("?" for _ in range(number_columns))})'
			

			
			
			cursor.executemany(insert_value,matrix)

		file_database.commit()

	


	def query_data(treeview,name_table,combobox_title,combobox_elemen_title,left,right,top,bottom,lis_row,file_database,tex_hiden):

		

		cursor = file_database.cursor()

		cursor.execute(f"PRAGMA table_info({name_table});")

		columns_info = cursor.fetchall()

		all_column = [column[1] for column in columns_info]


		
		if int(right) == 0:
			columns = all_column[int(left):]
			

		else:
			columns = all_column[int(left):-int(right)]

		if len(columns)>0:

			select_query = f"SELECT {', '.join(columns)} FROM {name_table} WHERE ROWID NOT IN ({', '.join(str(y) for y in lis_row)});"
		

			cursor.execute(select_query)

			all_data = cursor.fetchall()



			if int(bottom) == 0:
				data = all_data[int(top):]
				

			else:
				data = all_data[int(top):-int(bottom)]

		else:

			data_end = [(tex_hiden)]

			Function_File_And_Data.value_treeveiw(treeview,data_end)

			combobox_title["values"] =[(tex_hiden)]

			combobox_title.current(0)

			combobox_elemen_title["values"] = ("?????",)
			
			combobox_elemen_title.current(0)

			combobox_elemen_title["state"]="disable"

			data = []


		return data



	def push_value_treeveiw(treeview,name_table,combobox_title,combobox_elemen_title,left,right,top,bottom,lis_row,file_database,tex_hiden,text_all):


		

		data_treeview = Function_File_And_Data.query_data(treeview,name_table,combobox_title,combobox_elemen_title,left,right,top,bottom,lis_row,file_database,tex_hiden)

		try:



			Function_File_And_Data.value_treeveiw(treeview,data_treeview)

		except IndexError:

			data_end = [[tex_hiden]]

			Function_File_And_Data.value_treeveiw(treeview,data_end)
		

		if len(data_treeview) > 0:

			data_title = data_treeview[0]

			combobox_title["values"] =(text_all,)+ data_title

	
		else:

			combobox_title["values"] =[(tex_hiden)]

		combobox_title.current(0)

		combobox_elemen_title["values"] = ("?????",)
		
		combobox_elemen_title.current(0)

		combobox_elemen_title["state"]="disable"

		# index = [x for x in range(1,len(data_treeview)+1)]

		cursor = file_database.cursor()



		cursor.execute(f"SELECT ROWID FROM {name_table};")

		row_id = cursor.fetchall()

		row_raw = [item[0] for item in row_id]

		if int(bottom) ==0:

			row_cut = row_raw[int(top):]
		else:

			row_cut = row_raw[int(top):-int(bottom)]

		row = list(set(row_cut) - set(lis_row))
		

		return row

	
	def color_the_data(divide_range,wb,check,sht_of,range_of):
		

		if check =="on":

			for index,rg in enumerate(divide_range):
				
				wb.range(rg).copy()

				# if check =="on":

				sht_of.range(range_of.columns[index]).paste('formats',transpose=True)

		else:

			for index,rg in enumerate(divide_range):
				
				wb.range(rg).copy()		

				sht_of.range(range_of.rows[index]).paste('formats')





		# 	if check =="on":
		# 		pass

		# 		# sht_of.range().paste('formats',transpose=True)
			
		# 	else:
		# 		sht_of.range(range_of.rows[index]).paste('formats')

		# 	start = start+len(lis_range)//divide_columns

		# # start = 0
		# end = 0

		# for lis_range in divide_range:
			

		# 	end = end + len(lis_range)-1
		
		# 	address_choose = ",".join(lis_range)


			
		# 	wb.range(address_choose).copy()

		

		# 	if check =="on":


		# 		sht_of.range(range_of.columns[start],range_of.columns[end//divide_columns]).paste('formats',transpose=True)
			
		# 	else:
		# 		sht_of.range(range_of.rows[start],range_of.rows[end//divide_columns]).paste('formats')

		# 	start = start+len(lis_range)//divide_columns

class Function_Combobox:

	
	def data_title(treeview,name_table,combobox_title,combobox_elemen_title,left,right,top,bottom,lis_row,file_database,tex_hiden,text_all):

		cursor = file_database.cursor()

		if combobox_title.get() == text_all and combobox_elemen_title.get() !="?????":
			
			

			combobox_elemen_title["values"] = ("?????",)
		
			combobox_elemen_title.current(0)

			combobox_elemen_title["state"]="disable"

			Function_File_And_Data.push_value_treeveiw(treeview,name_table,combobox_title,combobox_elemen_title,left,right,top,bottom,lis_row,file_database,tex_hiden,text_all)

			
		if combobox_title.get() != text_all and combobox_title.get() != tex_hiden:

			
			

			cursor.execute(f"SELECT column{combobox_title.current()+int(left)} FROM {name_table} WHERE ROWID NOT IN ({', '.join(str(y) for y in lis_row)})")

			columns_info = cursor.fetchall()
			
			result_list_oj = [item[0] for item in columns_info]

			if int(bottom) ==0:

				unique_values,indices = np.unique(result_list_oj[1+int(top):], return_index=True)



			else:

				unique_values,indices = np.unique(result_list_oj[1+int(top):-int(bottom)],return_index=True)

			result_list = unique_values[np.argsort(indices)]

			if len(result_list)>0:

				combobox_elemen_title["state"] ="readonly"

				combobox_elemen_title["values"] = result_list.tolist()

				combobox_elemen_title.current(0)
			else:

				combobox_elemen_title["values"] = ("Không có values",)
		
				combobox_elemen_title.current(0)

				combobox_elemen_title["state"]="disable"


		cursor.execute(f"SELECT ROWID FROM {name_table};")

		row_id = cursor.fetchall()

		row_raw = [item[0] for item in row_id]

		if int(bottom) ==0:

			row_cut = row_raw[int(top):]
		else:

			row_cut = row_raw[int(top):-int(bottom)]

		row = list(set(row_cut) - set(lis_row))
		

		return row

		
		

	def find_data_title(treeview,name_table,combobox_title,combobox_elemen_title,left,right,file_database,lis_row):

		cursor = file_database.cursor()

		cursor.execute(f"SELECT ROWID FROM {name_table} WHERE column{combobox_title.current()+int(left)} ='{combobox_elemen_title.get()}' AND ROWID NOT IN ({', '.join(str(y) for y in lis_row)});")


		row_id = cursor.fetchall()



		row = [item[0] for item in row_id]
		

		cursor.execute(f"SELECT * FROM {name_table} WHERE ROWID IN ({','.join(str(x) for x in row)});")

		


		data_searched = cursor.fetchall()


		for item in treeview.get_children():

			treeview.delete(item)
		

		result = [sublist[left:] if right == 0 else sublist[left:-right] for sublist in data_searched]

		
		
		for values in result:		
					
			treeview.insert("", "end", values=values)

		

		return row

		


class Database_Processing:

	def take_data_database_columns(file_database,name_table,left,index,columns):



		columns_db = [column+int(left) for column in columns]

		cursor = file_database.cursor()


		query = f'SELECT {",".join(f"column{i}" for i in columns_db)} FROM {name_table} WHERE rowid IN ({",".join(str(y) for y in index)});'



		cursor.execute(query)

		all_data = cursor.fetchall()

		return all_data


	def take_data_database(file_database,name_table,left,right,index):

		cursor = file_database.cursor()

		cursor.execute(f"PRAGMA table_info({name_table});")

		columns_info = cursor.fetchall()

		all_column = [column[1] for column in columns_info]


		
		if int(right) == 0:
			columns = all_column[int(left):]
			

		else:
			columns = all_column[int(left):-int(right)]

		if len(columns)>0:

			select_query = f"SELECT {', '.join(columns)} FROM {name_table} WHERE ROWID  IN ({', '.join(str(y) for y in index)});"


		cursor.execute(select_query)

		value = cursor.fetchall()

		return value



		


	def delete_row_columns_database(file_database,lis_row,lis_columns,name_table):

		cursor = file_database.cursor()

		
		query_delete = f'''UPDATE {name_table} SET {','.join(f'column{i} = " "' for i in lis_columns)} WHERE ROWID IN ({', '.join(str(y) for y in lis_row)});'''

		cursor.execute(query_delete)
		file_database.commit()

	def delete_redo_table_db(file_undo_redo_db,lis_name_table):

		cursor = file_undo_redo_db.cursor()


		for table_name in lis_name_table:
			cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
			existing_table = cursor.fetchone()

			if existing_table:
				# Xóa dữ liệu từ bảng
				cursor.execute(f"DELETE FROM {table_name}")
			

		# Lưu thay đổi và đóng kết nối
		file_undo_redo_db.commit()



	def check_data_existence(file_undo_redo_db,name_table):

		cursor = file_undo_redo_db.cursor()

		check_data = f'''SELECT COUNT(*) FROM {name_table};'''

		cursor.execute(check_data)

		len_data = cursor.fetchall()

		return len_data[0]



	def query_title_db(file_database,name_table,lis_hiden,left,right):

		cursor = file_database.cursor()



		query = f'''SELECT *FROM {name_table} WHERE ROWID NOT IN ({', '.join(str(x) for x in lis_hiden)}) LIMIT 1;'''

		cursor.execute(query)

		title_value = cursor.fetchall()[0]

		if int(right) == 0:
			title = title_value[int(left):]
			

		else:
			title = title_value[int(left):-int(right)]



		return title


	def query_title_db_case_1_3(file_database,name_table,top,left,right):

		cursor = file_database.cursor()

		query = f'''SELECT *FROM {name_table} WHERE ROWID ={1+int(top)}'''


		cursor.execute(query)

		data_raw = cursor.fetchall()[0]

		if int(right) == 0:
			data = data_raw[int(left):]
			

		else:

			data = data_raw[int(left):-int(right)]

		
		return data



	def push_value_database_undo_redo (file_undo_redo_db,value_undo_redo,name_table):

		

		cursor = file_undo_redo_db.cursor()


		for x in name_table:

			create_table = f''' CREATE TABLE IF NOT EXISTS {x}( 

				
				{','.join(f'column{i} TEXT' for i in range(1,len(value_undo_redo)+1))});'''

			cursor.execute(create_table)

		file_undo_redo_db.commit()

		# Lấy tên các cột trong bảng b
		cursor.execute(f"PRAGMA table_info({name_table[0]})")
		columns = [column[1] for column in cursor.fetchall()]


		# Tạo câu lệnh SQL INSERT tự động
		sql_insert = f"INSERT INTO {name_table[0]} ({', '.join(columns)}) VALUES ({', '.join(['?']*len(columns))})"

		# Thực thi câu lệnh INSERT
		cursor.execute(sql_insert, value_undo_redo)

		# Lưu thay đổi và đóng kết nối
		file_undo_redo_db.commit()



	def transfer_databas_back_and_forth(file_undo_redo_db,data_transfer,data_received):

		cursor = file_undo_redo_db.cursor()



		query_insert = f'''INSERT INTO {data_received} SELECT * FROM {data_transfer} ORDER BY ROWID DESC LIMIT 1;'''

		delete_value = f'''DELETE FROM {data_transfer} WHERE ROWID IN (SELECT ROWID FROM {data_transfer} ORDER BY ROWID DESC LIMIT 1);'''

		query_data = f'''SELECT *FROM {data_transfer} ORDER BY ROWID DESC LIMIT 1; '''

		
		cursor.execute(query_insert)

		cursor.execute(query_data)

		value = cursor.fetchall()

		

		

		cursor.execute(delete_value)

		file_undo_redo_db.commit()	

		return value
		

	def query_index_db(file_database,name_table,left,right,row):

		cursor = file_database.cursor()




		cursor.execute(f"PRAGMA table_info({name_table});")

		columns_info = cursor.fetchall()

		all_column = [column[1] for column in columns_info]


		
		if int(right) == 0:
			columns = all_column[int(left):]
			

		else:
			columns = all_column[int(left):-int(right)]

		query_data = f'''SELECT {', '.join(columns)} FROM {name_table} WHERE ROWID IN ({row});'''

		cursor.execute(query_data)


		all_data = cursor.fetchall()

		return all_data

	def update_again_database(file_database,file_sub_db,name_table,column_choose,lis_row,left,right,lefted):

		cursor_db = file_database.cursor()

		cursor_sub = file_sub_db.cursor()

		cursor_db.execute(f"PRAGMA table_info({name_table});")

		columns_info = cursor_db.fetchall()

		all_column = [column[1] for column in columns_info]

		column_used = [all_column[x-1+int(lefted)] for x in column_choose]



		query_sub_data = f'''SELECT {",".join(all_column)} FROM {name_table} WHERE ROWID IN ({lis_row});'''

		cursor_sub.execute(query_sub_data)

		data_sud = cursor_sub.fetchall()

		
		data_insert = [[data[x-1+int(lefted)] for x in column_choose] for data in data_sud]
		
	

		for row,data in zip(lis_row.split(","),data_insert):

			

			update_data_db = f'''UPDATE {name_table} SET {",".join(f'{column_used[x]} = "{data[x]}"'for x in range(0,len(column_used)))} WHERE ROWID = {row};'''

			cursor_db.execute(update_data_db)
		file_database.commit()

		

		result = [sublist[left:] if right == 0 else sublist[left:-right] for sublist in data_sud]

		return result


class Function_gui_two:


	def get_file_size_in_kb(file_path):

		
		size_in_bytes = os.path.getsize(file_path)

		size_in_kb = size_in_bytes / 1024  # Convert bytes to kilobytes
		return size_in_kb


	def compare_lists(list1, list2):
		set1 = set(list1)
		set2 = set(list2)
		
		missing_in_list1 = list(set1 - set2)
		 
		result = missing_in_list1
		
		return result