from function_gui_custom import Command_Gui_One
import xlwings as wx
from gui_custom import Toplevel_Setting
from tkinter import ttk,filedialog
from tkinter import *
from tkinter import messagebox
import customtkinter as ctk
from Function_Gui import Balance,Function_File_And_Data,Function_gui_two
import pandas as pd
import os

import pywintypes

class Command_Gui_Two(Command_Gui_One):

	def __init__(self,root):
		super().__init__(root)

		self.button_upfile_official_gui_two.configure(command = self.up_file_of_gui_two)

		self.button_addfile_data_gui_two.configure(command = self.add_file)

		self.button_remove_sheets.configure(command=self.delete_item)

		self.button_remove_all_sheets.configure(command=self.delete_all_item)

		self.button_enter_gui_two.configure(command=self.enter_gui2)

		self.button_save_gui_two.configure(command=self.save_file_of_gui2)

		self.lis_file_uped = []
	
		self.lis_sheets_of_file_uped = []


	def enter_gui2(self):

		try:

			sht_of =  self.lis_sheet_of[self.combobox_sheet_official_gui_two.get()]

			list_item = self.treeview_gui_two.get_children()

		except pywintypes.com_error as e:

			if e.args[0] == -2147023174 or e.args[0] == -2147352567:

				self.list_box_notification_gui_two.insert(0,'You have renamed the official sheet file or closed it, please re-upload the official file, or change the sheet name to the correct sheet name when uploading the file')

				return
		

		try:
			
			for item in list_item:

				index_fist_of_tree = self.treeview_gui_two.get_children(item)[0]
				
				name_file = self.treeview_gui_two.item(index_fist_of_tree, 'text')

				lis_name_sheet = [self.treeview_gui_two.item(elemen_item, 'values')[0] for elemen_item in self.treeview_gui_two.get_children(item)]

				path_file = self.treeview_gui_two.item(index_fist_of_tree, 'values')[1]

				file_merge_sheet = wx.Book(path_file)


				for sheet_name in lis_name_sheet:

					source_sheet = file_merge_sheet.sheets[sheet_name]

					wx.apps.active.screen_updating = False

					if self.value_combobox_regime == self.list_value_regime[0]:

						lr_of = sht_of.range(self.value_column_excel+str(sht_of.cells.last_cell.row)).end('up').row + int(self.value_distance_excel) #--> vùng cuối có value của sheet of

						
						if self.switch_var_write_name_sheet_gui2.get() == "on":

							sht_of.range(self.value_column_excel+str(lr_of)).value = "Sheets:   " +sheet_name +"     File " + name_file

							lr_of +=1


						source_sheet.api.UsedRange.Copy()

						if self.value_type_format_gui2 == self.lis_value_type_format_gui2[0]:

							sht_of.api.Range(self.value_column_excel+str(lr_of)).PasteSpecial(Paste= -4163)

							sht_of.api.Range(self.value_column_excel+str(lr_of)).PasteSpecial(Paste=-4122)
							

						if self.value_type_format_gui2 == self.lis_value_type_format_gui2[1]:

							sht_of.api.Range(self.value_column_excel+str(lr_of)).PasteSpecial(Paste= -4163)

						if self.value_type_format_gui2 == self.lis_value_type_format_gui2[2]:
							
							sht_of.api.Range(self.value_column_excel+str(lr_of)).PasteSpecial(Paste=-4122)

						if self.value_type_format_gui2 == self.lis_value_type_format_gui2[3]:

							sht_of.api.Range(self.value_column_excel+str(lr_of)).PasteSpecial(Paste=-4123)

						if self.value_type_format_gui2 == self.lis_value_type_format_gui2[4]:

							sht_of.api.Range(self.value_column_excel+str(lr_of)).PasteSpecial(Paste=-4104)
						
						wx.apps.active.api.CutCopyMode = False
						
					if self.value_combobox_regime == self.list_value_regime[1]:

						source_sheet.api.Copy(After=sht_of.api)
					wx.apps.active.screen_updating = True
					
				if self.switch_var_close_the_file_when_done.get() == "on":

					file_merge_sheet.close()
		except 	AttributeError:

			if self.switch_var_notification.get() =="on":
			
				self.list_box_notification_gui_two.insert(0,self.notification_gui_one4)

				if len(self.treeview_gui_two.get_children()) ==0:

					self.list_box_notification_gui_two.insert(0,self.notification_gui_one15)
	

			
	def up_file_of_gui_two(self):

		path_file = Function_File_And_Data.take_path_file(self.value_entry_path_official)

		if path_file !="":

			self.workbook_of_gui2 =Function_File_And_Data.take_information_file_excel(path_file,self.combobox_sheet_official_gui_two)

			

			self.lis_sheet_of = self.workbook_of_gui2.sheets

			if self.switch_var_notification.get() =="on":

				self.list_box_notification_gui_two.insert(0,f"{self.name_widget9} {self.workbook_of_gui2.name} {self.notification_gui_one14}")
		else:

			if self.switch_var_notification.get() =="on":
			
				self.list_box_notification_gui_two.insert(0,self.notification_gui_one6)

	

	def add_file(self):


		lis_path_file = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsm")],initialdir=self.value_entry_path_data,title="Choose File Excel")
		
		
		

		if lis_path_file !="":

			for path in lis_path_file:

			

				if path not in self.lis_file_uped:

					list_name_sheets_excel = pd.ExcelFile(path).sheet_names
					self.lis_file_uped.append(path)


					name_file = os.path.basename(path)

					self.lis_sheets_of_file_uped.append(list_name_sheets_excel)

					size = Function_gui_two.get_file_size_in_kb(path)

					parent = self.treeview_gui_two.insert("", "end",text=f"File----------->: {name_file}",tags=("color_file",))
					


					for name_sheet in list_name_sheets_excel:

						self.treeview_gui_two.insert(parent, "end",text=name_file, values=(name_sheet,path,f"{size:.2f}KB"))
					

					self.treeview_gui_two.item(parent, open=True)
				
				else:

					index_tree = self.lis_file_uped.index(path)

					index_tree_in_treeveiw = self.treeview_gui_two.get_children()[index_tree]

					

					branch_in_tree = self.treeview_gui_two.get_children(index_tree_in_treeveiw)

					list_sheet_choosing = self.lis_sheets_of_file_uped[index_tree]

					if len(list_sheet_choosing) > len(branch_in_tree):

						name_file = os.path.basename(path)

						size = Function_gui_two.get_file_size_in_kb(path)

						values_of_branch = [self.treeview_gui_two.item(item, 'values')[0] for item in branch_in_tree]

						sheet_additional = Function_gui_two.compare_lists(list_sheet_choosing,values_of_branch)

						for name_sheet in sheet_additional:

							index_sheet_additional = list_sheet_choosing.index(name_sheet)

							self.treeview_gui_two.insert(index_tree_in_treeveiw,str(index_sheet_additional),text=name_file, values=(name_sheet,path,f"{size:.2f}KB"),tags=("color_sheet_additional"))

						self.treeview_gui_two.tag_configure("color_sheet_additional",foreground="#8B7E66")

			quantity_sheets = len([item for sublist in [list(self.treeview_gui_two.get_children(x)) for x in self.treeview_gui_two.get_children()] for item in sublist])

			quantity_file = len(self.lis_file_uped)

			self.number_file_gui_two.configure(text=quantity_file)

			self.number_sheets_gui_two.configure(text=quantity_sheets)
		

			self.treeview_gui_two.tag_configure("color_file", background="#8B4500")

		else:

			if self.switch_var_notification.get() =="on":
			
				self.list_box_notification_gui_two.insert(0,self.notification_gui_one2)


	def save_file_of_gui2(self):

		try:

			path_save =Function_File_And_Data.save_file_excel(self.workbook_of_gui2.name,self.value_entry_path_save,self.workbook_of_gui2)
			
		
		except pywintypes.com_error as e:

			if e.args[0] == -2147417848:

					self.list_box_notification.insert(0,self.notification_gui_one6)

		except AttributeError:

			if self.switch_var_notification.get() =="on":

				self.list_box_notification_gui_two.insert(0,self.notification_gui_one4)

	def delete_item(self):

		selected_item = self.treeview_gui_two.selection()

		if selected_item:

			for item in selected_item:	

						
				
				if len(self.treeview_gui_two.get_children(item)) !=0:
					
					index_list = self.treeview_gui_two.index(item)

					self.lis_file_uped.pop(index_list)

					self.lis_sheets_of_file_uped.pop(index_list)

					self.treeview_gui_two.delete(item)

				else:

					item_parent = self.treeview_gui_two.parent(item)

					if len(self.treeview_gui_two.get_children(item_parent)) -1 ==0:


						index_list = self.treeview_gui_two.index(item)

						self.lis_file_uped.pop(index_list)

						self.lis_sheets_of_file_uped.pop(index_list)

						self.treeview_gui_two.delete(item_parent)
					else:
				
						self.treeview_gui_two.delete(item)
				

		quantity_sheets = len([item for sublist in [list(self.treeview_gui_two.get_children(x)) for x in self.treeview_gui_two.get_children()] for item in sublist])

		quantity_file = len(self.lis_file_uped)

		self.number_file_gui_two.configure(text=quantity_file)

		self.number_sheets_gui_two.configure(text=quantity_sheets)	

	

	def delete_all_item(self):

		selected_item = self.treeview_gui_two.get_children()

		if selected_item:

			result = messagebox.askokcancel(self.notification_gui_one16[0],self.notification_gui_one16[1])

			if result:

				for item in selected_item:

					self.treeview_gui_two.delete(item)


				self.lis_file_uped.clear()
				self.lis_sheets_of_file_uped.clear()


				quantity_sheets = len([item for sublist in [list(self.treeview_gui_two.get_children(x)) for x in self.treeview_gui_two.get_children()] for item in sublist])

				quantity_file = len(self.lis_file_uped)

				self.number_file_gui_two.configure(text=quantity_file)

				self.number_sheets_gui_two.configure(text=quantity_sheets)

				if self.switch_var_notification.get() =="on":

					self.list_box_notification_gui_two.insert(0,self.notification_gui_one17)
			else:
				pass
if __name__ == "__main__":

	root = ctk.CTk()		

	Command_Gui_Two(root)

	ctk.set_appearance_mode("Dark")

	ctk.set_default_color_theme("blue")

	Balance.center_window(root,860,520)

	root.mainloop()				
