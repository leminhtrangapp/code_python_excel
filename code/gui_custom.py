
from language import Image_Gui
from tkinter import *
from tkinter import ttk,filedialog
import customtkinter as ctk
from customtkinter import CTkToplevel


from Function_Gui import Balance,Setting_Entry,Function_File_And_Data
import ast

from openpyxl.utils import get_column_letter

import sqlite3


class Setup(Image_Gui):

	def __init__(self,root):

		super().__init__()

		self.root = root

		self.toplevel_window = None
			
		self.tapbar()

		self.frame()

		self.label()

		self.list_box()

		self.button()

		self.combobox()

		self.canvas()

		self.treeview()
		
		self.root.bind("<Double-Button-1>",self.clear_selection)

		self.root.bind("<Triple-Button-1>",self.clear_selection_heading_treeveiw)
	
		self.pack_widget()

		style = ttk.Style()

		style.theme_use("clam")
	

		self.root.option_add("*TCombobox*Listbox*Background", '#333333')

		self.root.option_add("*TCombobox*Listbox*foreground", 'white')

		style.map('TCombobox', fieldbackground=[('readonly','#003333'),('disabled','#2E8B57')],foreground=[('readonly', '#FFFFFF')],background = [('readonly', '#838B8B'),('disabled', '#8B5F65')],arrowcolor=[('readonly', '#8B1A1A')])



		style.map('TSpinbox', fieldbackground=[('readonly', '#333333')],foreground=[('readonly', '#FFFFFF')],background=[('readonly','white')],arrowcolor=[('readonly','#8B1A1A')])

		style.configure("Treeview", background="#003333",fieldbackground="#003333",font=("Tahoma", 10, "bold"),foreground="white")
		
		style.configure("Treeview.Heading",font=("Elephant", 15, "bold"),background="#528B8B",foreground="black")

		style.map('Treeview', background=[('selected', '#999900')], foreground=[('selected', 'black')])

		self.list_box_notification.configure(background="#003333",foreground="white")

		try:

			self.button_upfile_data.configure(image=self.image_up_file,compound="left")
			
			self.button_upfile_official.configure(image=self.image_up_file,compound="left")

			self.button_enter.configure(image=self.image_enter,compound="left")
			
			self.button_setting.configure(image=self.image_setting,compound="left")

			self.button_save.configure(image=self.image_save,compound="left")

			self.button_undo.configure(image=self.image_undo,compound="left")

			self.button_redo.configure(image=self.image_redo,compound="left")

		except AttributeError:

			notification_gui ="The icon for the button was not found (maybe you deleted the icon_button folder or the icons inside the folder), so now the interface will not be beautiful."

			self.list_box_notification.insert(0,notification_gui)


	def tapbar(self):

	

		self.mytab = ctk.CTkTabview(self.root,width=int(self.root.winfo_screenwidth()),

			height=int(self.root.winfo_screenheight()),segmented_button_selected_hover_color="blue")

		

		self.tap_1 = self.mytab.add(self.name_widget1)

		self.tap_2 = self.mytab.add(self.name_widget2)

		# self.tap_3 = self.mytab.add("Cách Hoạt Động")

	def frame(self):


		self.frame_upfile = ctk.CTkFrame(master=self.tap_1,fg_color="gray",height=100,corner_radius=15,width=250)
		self.frame_notification = ctk.CTkFrame(master=self.tap_1,fg_color="gray", height=180,corner_radius=15,width=250)	
		self.frame_main = ctk.CTkFrame(master=self.tap_1,fg_color="gray",corner_radius=10)
		self.frame_title = ctk.CTkFrame(master=self.frame_main,fg_color="#BBBBBB",height=90,corner_radius=10,width=250)
		self.frame_treeview = ctk.CTkFrame(master=self.frame_main,fg_color="#BBBBBB",corner_radius=10,width=250)
		self.frame_command = ctk.CTkFrame(master=self.tap_1,fg_color="gray",height=70,corner_radius=10,width=250)
		


	def label(self):

		self.label_name_file_data = ctk.CTkLabel(master=self.frame_upfile,fg_color="gray",text_color="black",text="Name File")

		self.label_name_sheet_data = ctk.CTkLabel(master=self.frame_upfile,fg_color="gray",text_color="black",text="Name Sheets")

		self.label_name_file_of = ctk.CTkLabel(master=self.frame_upfile,fg_color="gray",text_color="black",text="Name File")

		self.label_name_sheet_of = ctk.CTkLabel(master=self.frame_upfile,fg_color="gray",text_color="black",text="Name Sheets")		

		self.label_table = ctk.CTkLabel(self.frame_title,text=self.name_widget4,corner_radius=5,fg_color="#9C9C9C",text_color="black")

		self.label_find_title = ctk.CTkLabel(self.frame_title,text=self.name_widget5,fg_color="#BBBBBB",text_color="black",corner_radius=10)

		self.label_find_elemen_title = ctk.CTkLabel(self.frame_title,text=self.name_widget6,fg_color="#BBBBBB",text_color="black",corner_radius=10)

		self.label_notification = ctk.CTkLabel(self.frame_notification,text=self.name_widget7,corner_radius=5,fg_color="#9C9C9C",text_color="black")

	def list_box(self):

		self.list_box_notification = Listbox(self.frame_notification,font=("Arial", 10, "bold"),border=2,selectmode=MULTIPLE)

		self.xscrollbar_lis_box = ctk.CTkScrollbar(self.frame_notification, orientation='horizontal',border_spacing=5,corner_radius=10,button_hover_color="black",command=self.list_box_notification.xview)
		
		self.list_box_notification.configure(xscrollcommand=self.xscrollbar_lis_box.set)		

		self.yscrollbar_lis_box = ctk.CTkScrollbar(self.frame_notification, orientation='vertical',border_spacing=5,corner_radius=10,button_hover_color="black",command=self.list_box_notification.yview)

		self.list_box_notification.configure(yscrollcommand=self.yscrollbar_lis_box.set)	

	def button(self):

		self.button_upfile_data = ctk.CTkButton(self.frame_upfile, text=self.name_widget8,font=("Helvetica",11),corner_radius=50,fg_color="#2F4F4F",hover_color="#FF6600",border_width=0.7,border_color="black")
		
		self.button_upfile_official = ctk.CTkButton(self.frame_upfile, text=self.name_widget9,font=("Helvetica",11),corner_radius=50,fg_color="#2F4F4F",hover_color="#FF6600",border_width=0.7,border_color="black")

		self.button_enter = ctk.CTkButton(self.frame_command, text=self.name_widget10,border_width=1,border_color="black",fg_color="#2F4F4F",hover_color="#FF6600",height=60,width=230)
		

		self.button_setting = ctk.CTkButton(self.frame_command, text=self.name_widget11,border_width=0.7,border_color="black",fg_color="#2F4F4F",hover_color="#FF6600")

		self.button_undo = ctk.CTkButton(self.frame_command, text="UnDo",fg_color="#2F4F4F",hover_color="#FF6600",state=ctk.DISABLED,font =("Helvetica",11))

		self.button_redo = ctk.CTkButton(self.frame_command, text="ReDo",fg_color="#2F4F4F",hover_color="#FF6600",state=ctk.DISABLED,font =("Helvetica",11))	
	
		self.button_save = ctk.CTkButton(self.frame_command, text="Save",fg_color="#2F4F4F",hover_color="#FF6600",border_width=0.7,border_color="black",font =("Helvetica",11))

	def combobox(self):

		self.combobox_sheet_data = ttk.Combobox(self.frame_upfile,state="readonly",justify='center',width=20)
		
		self.combobox_sheet_official = ttk.Combobox(self.frame_upfile,state="readonly",justify='center',width=20)
	
		self.combobox_find_title = ttk.Combobox(self.frame_title,state="readonly",justify='center')

		self.combobox_find_elemen_title = ttk.Combobox(self.frame_title,values=[1],state="readonly",justify='center')

	def canvas(self):

		self.canvas_one = Canvas(self.frame_upfile, height=0.5, background='gray')

		self.canvas_two = Canvas(self.frame_title, height=1, background='gray')


	def treeview(self):

		self.treeview =ttk.Treeview(self.frame_treeview,show="headings",selectmode="none")

		self.xscrollbar = ctk.CTkScrollbar(self.frame_treeview, orientation='horizontal',border_spacing=5,corner_radius=10,button_hover_color="black",command=self.treeview.xview)
		
		self.treeview.configure(xscrollcommand=self.xscrollbar.set)		

		self.yscrollbar = ctk.CTkScrollbar(self.frame_treeview, orientation='vertical',border_spacing=5,corner_radius=10,button_hover_color="black",command=self.treeview.yview)

		self.treeview.configure(yscrollcommand=self.yscrollbar.set)

		self.treeview["columns"] =("Column 1", "Column 2")

		self.treeview.heading("Column 1", text="Column 1")
		self.treeview.heading("Column 2", text="Column 2")

		# for x in range(1,30):

		self.treeview.insert("", "0", "item0", text="Item 1", values=("Value 1.0", "Value 1.2"))
		self.treeview.insert("", "1", "item2", text="Item 2", values=("Value 2.1", "Value 2.2"))
		self.treeview.insert("", "end", "item3", text="Item 3", values=("Value 3.1", "Value 3.2"))
		
		self.treeview.bind("<Control-a>", self.select_all)	

		self.treeview.tag_configure('focus', background='#777777')

		

		

	#____________________________________________________________________________

	




	def select_all(self,event):

		for item in self.treeview.get_children():
			self.treeview.selection_add(item)

	def clear_selection_heading_treeveiw(self,event):

		for col_name in self.treeview["columns"]:

			if self.treeview.heading(col_name,"text")[0] =="\u2611":

				index = self.treeview["columns"].index(col_name)

				value = self.treeview.heading(index, 'text')

				value = value[1:]

				self.treeview.heading(index, text=value)

	def clear_selection(self, event):

		self.treeview.selection_remove(self.treeview.selection())

		
				
	def edit_treeview(self,event):
		
		if self.treeview.identify_region(event.x, event.y) == 'heading':


			self.treeview.unbind("<ButtonRelease-1>")
			
			column = self.treeview.identify_column(event.x)
	
			value = self.treeview.heading(column, 'text')


			if value[0] == "\u2611":

			 	value = value[1:]	 		 	

			elif value[0] != "\u2611":

				a =value

				value = "\u2611"+a			

			self.treeview.heading(column, text=value)

		if self.treeview.identify_region(event.x, event.y) =="cell":
			

			def select(event=None):
				self.treeview.selection_toggle(self.treeview.focus())

			self.treeview.bind("<ButtonRelease-1>", select)
	def pack_widget(self):

		self.mytab.grid(row=0,column=0)
		self.mytab._segmented_button.grid(sticky="w")
		self.mytab._segmented_button.configure(dynamic_resizing=True)
	
		#___________________________________________________________________________________ Frame
		self.frame_upfile.grid(row=0,column=0,stick="nsew",padx=5)
		self.frame_notification.grid(row=1,column=0,stick="nsew",padx=5,pady=5)
		self.frame_main.grid(row=0,column=1,stick="nsew",padx=2,rowspan=3,columnspan=2)
		self.frame_command.grid(row=2,column=0,stick="nsew",padx=5)
		self.frame_title.grid(row=0,column=0,stick="new",padx=5,pady=3)
		
		self.frame_upfile.grid_propagate(False)
		self.frame_notification.grid_propagate(False)
		self.frame_main.grid_propagate(False)
		self.frame_command.grid_propagate(False)
		

		#____________________________________________________________________________________Frame Upfile

		self.label_name_file_data.grid(row=0,column=0,padx=2,stick="s")
		self.label_name_sheet_data.grid(row=0,column=1,padx=2,stick="s")
		self.label_name_file_of.grid(row=3,column=0,padx=2,stick="s")
		self.label_name_sheet_of.grid(row=3,column=1,padx=2,stick="s")

		self.canvas_one.grid(row=2, column=0, columnspan=2,padx=10)

		self.button_upfile_data.grid(row=1,column=0,padx=4,pady=4,stick="nsew")
		self.button_upfile_official.grid(row=4,column=0,padx=4,pady=4,stick="nsew")

		self.combobox_sheet_data.grid(row=1,column=1,padx=4,stick="ew")
		self.combobox_sheet_official.grid(row=4,column=1,padx=4,stick="ew")

		#____________________________________________________________________________________Frame Main

		self.frame_treeview.grid(row=1,column=0,sticky="nsew",padx=5,pady=5)
		self.treeview.grid(row=1,column=0,sticky="nsew")

		
		self.yscrollbar.grid(row=1,column=1,sticky="sen")
		self.xscrollbar.grid(row=2,column=0,sticky="wse",padx=5,pady=5)

		#____________________________________________________________________________________Frame Title

		self.canvas_two.grid(row=1, column=0,columnspan=7,padx=30,pady=3,stick="ew")
		self.label_table.grid(row=2, column=0,columnspan=7,padx=5,stick="ew",pady=5)
		self.label_find_title.grid(row=0,column=0,pady=5,padx=5,stick="ew")
		self.label_find_elemen_title.grid(row=0,column=3,pady=5,padx=5,stick="ew")
		# 
		self.combobox_find_title.grid(row=0, column=1, pady=5,stick="ew")
		self.combobox_find_elemen_title.grid(row=0, column=4, pady=5,stick="ew")
		
		#_____________________________________________________________________________________Frame Command

		self.button_enter.grid(row=0, column=1,padx=3,pady=3,stick="nsew",rowspan=2)

		self.button_setting.grid(row=2, column=1,padx=3,pady=3,stick="nsew")

		self.button_undo.grid(row=0, column=0,padx=3,pady=3,stick="nsew")

		self.button_redo.grid(row=1, column=0,padx=3,pady=3,stick="nsew")

		self.button_save.grid(row=2, column=0,padx=3,pady=3,stick="nsew")

		#_______________________________________________________________________________________ Frame Notificatio

		self.label_notification.grid(row=0, column=0,padx=5,stick="ew",pady=5,columnspan=2)
		self.list_box_notification.grid(row=1, column=0,padx=3,stick="nsew")

		self.yscrollbar_lis_box.grid(row=1,column=1,sticky="sen")
		self.xscrollbar_lis_box.grid(row=2,column=0,sticky="wse")

class Balance_Widget(Setup):

	def __init__(self, root):

		super().__init__(root)

		self.balance_gui()
	
	def balance_gui(self):

		Balance.function_balance2(self.root,self.mytab)

		Balance.function_balance2(self.frame_main,self.frame_treeview)

		Balance.function_balance2(self.frame_treeview,self.treeview)

		
		Balance.function_balance2(self.tap_1,self.frame_main)
		Balance.function_balance2(self.frame_notification,self.list_box_notification)

		#_________________________________________________________________________________________________________Frame Upfile

		lis_frame_updile = [self.label_name_file_data,self.label_name_sheet_data,self.label_name_file_of,

		self.label_name_sheet_of,self.button_upfile_data,self.button_upfile_official,self.combobox_sheet_data,

		self.combobox_sheet_official]

		Balance.function_balance(self.frame_upfile,lis_frame_updile,1,0)

		#_____________________________________________________________________________________Frame Title

		lis_widget_title = [self.label_table,self.canvas_two,self.label_find_title,self.label_find_elemen_title,self.combobox_find_title
		,self.combobox_find_elemen_title]

		Balance.function_balance(self.frame_title,lis_widget_title,1,0)


		#______________________________________________________________________________________Frame Command

		lis_widget_command = [self.button_enter,self.button_setting,self.button_undo,self.button_redo,self.button_save]

		Balance.function_balance(self.frame_command,lis_widget_command,1,0)

		#______________________________________________________________________________________

class Toplevel_Setting(Balance_Widget):

	def __init__(self, root):

		super().__init__(root)

		

		self.spin_box_values = [i for i in range(0,100)]

		self.lis_cloumn_excel = [get_column_letter(x) for x in range(1,2638)]

		self.lis_distance_row = [str(x) for x in range(1,100)]

		self.list_value_theme = ttk.Style().theme_names()

		self.lis_vale_language = ['vietnames','english','japanese','chinese']

		self.value_entry_path_data = ""
		self.value_entry_path_official = ""
		self.value_entry_path_save = ""

		self.value_spin_box_left = 0
		self.value_spin_box_right = 0
		self.value_spin_box_top = 0
		self.value_spin_box_bottom = 0

		self.value_distance_excel = 1

		self.value_column_excel = "A"

		self.number_cloumn_excel = 1
		
		self.switch_var_wrap_text = ctk.StringVar(value="off")

		self.switch_var_take_title = ctk.StringVar(value="off")

		self.switch_var_transpose = ctk.StringVar(value="off")

		self.switch_var_format = ctk.StringVar(value="off")

		self.switch_var_notification = ctk.StringVar(value="on")

		self.value_copy_cut = "Copy"

		

		self.current_theme = 1


		self.select_mode_theme = 1

		
		self.button_setting.configure(command=self.open_toplevel)
		
	def open_toplevel(self):

		if self.toplevel_window is None:

			self.toplevel_window = CTkToplevel(self.root)  # create window if its None or destroyed

			self.toplevel_window.protocol("WM_DELETE_WINDOW", self.on_close_toplevel)


			Balance.center_window(self.toplevel_window,720,500)

			self.toplevel_window.after(100, self.toplevel_window.lift)

			self.widget_toplevel()

			self.pack_widget_toplevel()

			# self.toplevel_window.grab_set()

		
		else:

			self.toplevel_window.focus()  # if window exists focus it
			self.toplevel_window.lift(self.root)

	def widget_toplevel(self):

		self.frame_parameter = ctk.CTkFrame(master=self.toplevel_window,fg_color="dark gray",corner_radius=10)

		self.label_setting_toplevel_cut_value = ctk.CTkLabel(master=self.frame_parameter,font=("Helvetica",14),fg_color="dark gray",text_color="black",text=self.name_widget29)

		self.label_setting_toplevel_cut_value2 = ctk.CTkLabel(master=self.frame_parameter,font=("Helvetica",14),fg_color="dark gray",text_color="black",text=self.name_widget30)

		self.frame_align = ctk.CTkFrame(master=self.toplevel_window,fg_color="dark gray",corner_radius=10)
		
		self.label_setting_toplevel_align = ctk.CTkLabel(master=self.frame_align,font=("Helvetica",14),fg_color="dark gray",text_color="black",text=self.name_widget31)

#____________________________________________________________________________________________________________frame_cut_value	

		self.frame_cut_value = ctk.CTkFrame(master=self.frame_parameter,fg_color="gray",corner_radius=10)

		self.label_cut_top = ctk.CTkLabel(master=self.frame_cut_value,font=("Bold",12),fg_color="#00CD66",text_color="black",text=self.name_widget12,corner_radius=5)

		self.label_cut_left = ctk.CTkLabel(master=self.frame_cut_value,font=("Bold",12),fg_color="#00CD66",text_color="black",text=self.name_widget13,corner_radius=5)

		self.label_cut_right = ctk.CTkLabel(master=self.frame_cut_value,font=("Bold",12),fg_color="#00CD66",text_color="black",text=self.name_widget14,corner_radius=5)

		self.label_cut_bottom = ctk.CTkLabel(master=self.frame_cut_value,font=("Bold",12),fg_color="#00CD66",text_color="black",text=self.name_widget15,corner_radius=5)

		self.spin_box_top = ttk.Spinbox(self.frame_cut_value,state="readonly",justify="center",width=3,values=self.spin_box_values)

		self.spin_box_top.set(self.value_spin_box_top)

		self.spin_box_left = ttk.Spinbox(self.frame_cut_value,state="readonly",justify="center",width=3,values=self.spin_box_values)

		self.spin_box_left.set(self.value_spin_box_left)

		self.spin_box_right = ttk.Spinbox(self.frame_cut_value,state="readonly",justify="center",width=3,values=self.spin_box_values)

		self.spin_box_right.set(self.value_spin_box_right)

		self.spin_box_bottom = ttk.Spinbox(self.frame_cut_value,state="readonly",justify="center",width=3,values=self.spin_box_values)

		self.spin_box_bottom.set(self.value_spin_box_bottom)
#_____________________________________________________________________________________________________________Frame_align________________________________________________________________________________________________________frame_type_value

		self.frame_type_value = ctk.CTkFrame(master=self.frame_parameter,fg_color="gray",corner_radius=10,width=self.frame_parameter.winfo_screenwidth()//4)

		self.label_copy_cut = ctk.CTkLabel(master=self.frame_type_value,font=("Helvetica",13),fg_color="#00CD66",text_color="black",text="Copy Or Cut",corner_radius=5)

		self.label_border = ctk.CTkLabel(master=self.frame_type_value,font=("Helvetica",13),fg_color="#111111",text_color="black",text="",corner_radius=5)

		self.label_column_excel = ctk.CTkLabel(master=self.frame_type_value,font=("Helvetica",13),fg_color="#00CD66",text_color="black",text=self.name_widget16,corner_radius=5)

		self.label_distance_excel = ctk.CTkLabel(master=self.frame_type_value,font=("Helvetica",13),fg_color="#00CD66",text_color="black",text=self.name_widget17,corner_radius=5)

		self.combobox_copy_cut = ttk.Combobox(self.frame_type_value,state="readonly",justify='center',values=["Coppy","Cut"],width=5)
	
		self.combobox_copy_cut.set(self.value_copy_cut) 

		self.combobox_column_excel = ttk.Combobox(self.frame_type_value,state="readonly",justify='center',values=self.lis_cloumn_excel,width=5)

		self.combobox_column_excel.set(self.value_column_excel)


		self.combobox_distance_excel = ttk.Combobox(self.frame_type_value,state="readonly",justify='center',values=self.lis_distance_row,width=5)

		self.combobox_distance_excel.set(self.value_distance_excel)
	

		self.switch_title = ctk.CTkSwitch(master=self.frame_type_value, text=self.name_widget18,font=("Helvetica",13),text_color="black",variable=self.switch_var_take_title,onvalue="on", offvalue="off")

		self.switch_format = ctk.CTkSwitch(master=self.frame_type_value, text=self.name_widget19,font=("Helvetica",13),text_color="black",variable=self.switch_var_format,onvalue="on", offvalue="off")

		self.switch_transpose = ctk.CTkSwitch(master=self.frame_type_value, text="Transpose",font=("Helvetica",13),text_color="black",variable=self.switch_var_transpose,onvalue="on", offvalue="off")

#_____________________________________________________________________________________________________________Frame_align

		self.frame_system = ctk.CTkFrame(master=self.frame_align,fg_color="gray",corner_radius=10)

		self.switch_dark_mode = ctk.CTkSwitch(master=self.frame_system, text="Dark Mode",font=("Helvetica",13),text_color="black",command=self.chane_mode_theme)

		if self.select_mode_theme == 1:

			self.switch_dark_mode.select()

		else:

			self.switch_dark_mode.deselect()
	

		self.switch_notification = ctk.CTkSwitch(master=self.frame_system, text=self.name_widget20,font=("Helvetica",13),text_color="black",variable=self.switch_var_notification,onvalue="on", offvalue="off")
	
		self.switch_wrap_text = ctk.CTkSwitch(master=self.frame_system, text="Wrap Text",font=("Helvetica",13),text_color="black",variable=self.switch_var_wrap_text, onvalue="on", offvalue="off")

		self.label_language = ctk.CTkLabel(master=self.frame_system,font=("Bold",12),fg_color="#00CD66",text_color="black",text=self.name_widget38,corner_radius=5,width=5)

		self.combobox_language = ttk.Combobox(self.frame_system,state="readonly",justify='center',width=5,values=self.lis_vale_language)

		self.combobox_language.set(self.language_default)

		self.label_theme = ctk.CTkLabel(master=self.frame_system,font=("Bold",12),fg_color="#00CD66",text_color="black",text="Sample Themes",corner_radius=5,width=5)

		self.combobox_theme = ttk.Combobox(self.frame_system,state="readonly",justify='center',width=5,values=self.list_value_theme)

		self.combobox_theme.current(self.current_theme)

		self.combobox_theme.bind("<<ComboboxSelected>>",self.chane_theme)

		self.label_regime = ctk.CTkLabel(master=self.frame_system,font=("Bold",12),fg_color="#00CD66",text_color="black",text=self.name_widget21,corner_radius=5,width=5)

		self.button_delete_notification = ctk.CTkButton(self.frame_system, text=self.name_widget22,border_width=1,border_color="black",fg_color="#2F4F4F",hover_color="#FF6600",width=30,command=self.clear_notification)
		
		self.button_path_default = ctk.CTkButton(self.frame_system, text=self.name_widget23,border_width=1,border_color="black",fg_color="#2F4F4F",hover_color="#FF6600",width=30,command=self.default_data_gui_one)

		self.button_path_application = ctk.CTkButton(self.frame_system, text=self.name_widget24,border_width=1,border_color="black",fg_color="#2F4F4F",hover_color="#FF6600",width=30,command=self.setting_apply)

		self.button_path_close = ctk.CTkButton(self.frame_system, text=self.name_widget25,border_width=1,border_color="black",fg_color="#2F4F4F",hover_color="#FF6600",width=30,command=self.colse_toplevel)
		


		self.frame_path = ctk.CTkFrame(master=self.frame_system,fg_color="gray",corner_radius=10)
		
		
		self.entry_path_data = ctk.CTkEntry(master=self.frame_path,placeholder_text=self.name_widget26,corner_radius=10)
		
		if self.value_entry_path_data !="":

			self.entry_path_data.insert(0,self.value_entry_path_data)

		self.entry_path_official = ctk.CTkEntry(master=self.frame_path,placeholder_text=self.name_widget27,corner_radius=10)

		if self.value_entry_path_official != "":

			self.entry_path_official.insert(0,self.value_entry_path_official)

		self.entry_path_save = ctk.CTkEntry(master=self.frame_path,placeholder_text=self.name_widget28,corner_radius=10)

		if self.value_entry_path_save != "":

			self.entry_path_save.insert(0,self.value_entry_path_save)	

		self.button_path_data = ctk.CTkButton(self.frame_path, text="browse",border_width=1,border_color="black",fg_color="#2F4F4F",hover_color="#FF6600",width=50,command=self.path_for_entry_data)

		self.button_path_official = ctk.CTkButton(self.frame_path, text="browse",border_width=1,border_color="black",fg_color="#2F4F4F",hover_color="#FF6600",width=50,command=self.path_for_entry_official)

		self.button_path_save = ctk.CTkButton(self.frame_path, text="browse",border_width=1,border_color="black",fg_color="#2F4F4F",hover_color="#FF6600",width=50,command=self.path_for_entry_save)

	def path_for_entry_data(self):

		folder_path = filedialog.askdirectory(mustexist=True,title="Choose Fath")

		self.entry_path_data.delete(0,END)

		self.entry_path_data.insert(0,folder_path)

		self.value_entry_path_data = folder_path

		self.toplevel_window.after(10, self.toplevel_window.lift)

	def path_for_entry_official(self):

		folder_path = filedialog.askdirectory(mustexist=True,title="Choose Fath")

		self.entry_path_official .delete(0,END)

		self.entry_path_official.insert(0,folder_path)

		self.value_entry_path_official = folder_path

		self.toplevel_window.after(10, self.toplevel_window.lift)

	def path_for_entry_save(self):

		folder_path = filedialog.askdirectory(mustexist=True,title="Choose Fath")

		self.entry_path_save.delete(0,END)

		self.entry_path_save.insert(0,folder_path)

		self.value_entry_path_save = folder_path

		self.toplevel_window.after(10, self.toplevel_window.lift)

	def chane_theme(self,event):

		theme = self.combobox_theme.get()

		treestyle = ttk.Style()
		

		treestyle.theme_use(theme)

		self.current_theme = self.combobox_theme.current()

		if int(self.switch_dark_mode.get()) == 1:

			
			self.root.option_add("*TCombobox*Listbox*Background", '#333333')

			self.root.option_add("*TCombobox*Listbox*foreground", 'white')

			treestyle.map('TCombobox', fieldbackground=[('readonly','#003333'),('disabled','#2E8B57')],foreground=[('readonly', '#FFFFFF')],background = [('readonly', '#838B8B'),('disabled', '#8B5F65')],arrowcolor=[('readonly', '#8B1A1A')])



			treestyle.map('TSpinbox', fieldbackground=[('readonly', '#333333')],foreground=[('readonly', '#FFFFFF')],background=[('readonly','white')],arrowcolor=[('readonly','#8B1A1A')])

			treestyle.configure("Treeview", background="#003333",fieldbackground="#003333",font=("Tahoma", 10, "bold"),foreground="white")
			
			treestyle.configure("Treeview.Heading",font=("Elephant", 15, "bold"),background="#528B8B",foreground="black")

			treestyle.map('Treeview', background=[('selected', '#999900')], foreground=[('selected', 'black')])
			
			self.list_box_notification.configure(background="#003333",foreground="white")


		


			self.frame_title.configure(fg_color="#BBBBBB")

			self.label_find_title.configure(fg_color="#BBBBBB")
			self.label_find_elemen_title.configure(fg_color="#BBBBBB")
			

			self.frame_upfile.configure(fg_color="gray")
			self.frame_notification.configure(fg_color="gray")
			self.frame_main.configure(fg_color="gray")
			self.frame_command.configure(fg_color="gray")

			self.label_name_file_data.configure(fg_color="gray")
			self.label_name_sheet_data.configure(fg_color="gray")
			self.label_name_file_of.configure(fg_color="gray")
			self.label_name_sheet_of.configure(fg_color="gray")

			self.button_upfile_data.configure(fg_color="#2F4F4F")
			self.button_upfile_official.configure(fg_color="#2F4F4F")
			self.button_enter.configure(fg_color="#2F4F4F")
			self.button_setting.configure(fg_color="#2F4F4F")
			self.button_undo.configure(fg_color="#2F4F4F")
			self.button_redo.configure(fg_color="#2F4F4F")
			self.button_save.configure(fg_color="#2F4F4F")

		

			self.label_table.configure(fg_color="#9C9C9C")
			self.label_notification.configure(fg_color="#9C9C9C")


			self.button_upfile_official_gui_two.configure(fg_color="#2F4F4F")
			self.button_addfile_data_gui_two.configure(fg_color="#2F4F4F")
			self.button_enter_gui_two.configure(fg_color="#2F4F4F")
			self.button_setting_gui_two.configure(fg_color="#2F4F4F")
			self.button_save_gui_two.configure(fg_color="#2F4F4F")
			self.button_remove_sheets.configure(fg_color="#FF6633")
			self.button_remove_all_sheets.configure(fg_color="#C82E31")

			self.list_box_notification_gui_two.configure(background="#003333",foreground="white")

			self.label_title_gui_two.configure(fg_color="#9C9C9C")
			self.label_notification_gui_two.configure(fg_color="#9C9C9C")


			self.label_number_file_gui_two.configure(fg_color="#708090")
			self.label_number_sheets_gui_two.configure(fg_color="#708090")
			self.number_file_gui_two.configure(fg_color="#708090")
			self.number_sheets_gui_two.configure(fg_color="#708090")

			self.frame_setting_gui_two.configure(fg_color="#668B8B")

			self.frame_function_gui_two.configure(fg_color="#C0C0C0")
			self.frame_title_gui_two.configure(fg_color="#C0C0C0")
			self.frame_nofitication_gui2.configure(fg_color="#C0C0C0")

			self.frame_command_gui_two.configure(fg_color="#8B7B8B")

			self.treeview_gui_two.tag_configure("color_file", background="#8B4500")

			self.treeview_gui_two.tag_configure("color_sheet_additional",foreground="#8B7E66")
		
		else:			

			self.root.option_add("*TCombobox*Listbox*Background", '#DDDDDD')

			self.root.option_add("*TCombobox*Listbox*foreground", 'black')

			treestyle.map('TCombobox', fieldbackground=[('readonly','#DDDDDD'),('disabled','#698B69')],foreground=[('readonly', 'black')],background = [('readonly', '#548B54'),('disabled', '#8B5F65')],arrowcolor=[('readonly', '#556B2F')])

			treestyle.map('TSpinbox', fieldbackground=[('readonly', '#DDDDDD')],foreground=[('readonly', 'black')],background=[('readonly','black')],arrowcolor=[('readonly','#556B2F')])

			
			treestyle.configure("Treeview", background="#DDDDDD",fieldbackground="#DDDDDD",font=("Tahoma", 10, "bold"),foreground="#1C1C1C")
			
			treestyle.configure("Treeview.Heading",font=("Elephant", 15, "bold"),background="#4A708B",foreground="black")

			treestyle.map('Treeview', background=[('selected', '#3333FF')], foreground=[('selected', '#FFFAFA')])
			
			self.list_box_notification.configure(background="#DDDDDD",foreground="#1C1C1C")


			self.frame_title.configure(fg_color="#6CA6CD")

			self.label_find_title.configure(fg_color="#6CA6CD")
			self.label_find_elemen_title.configure(fg_color="#6CA6CD")
			

			self.frame_upfile.configure(fg_color="#528B8B")
			self.frame_notification.configure(fg_color="#528B8B")
			self.frame_main.configure(fg_color="#528B8B")
			self.frame_command.configure(fg_color="#528B8B")

			self.label_name_file_data.configure(fg_color="#528B8B")
			self.label_name_sheet_data.configure(fg_color="#528B8B")
			self.label_name_file_of.configure(fg_color="#528B8B")
			self.label_name_sheet_of.configure(fg_color="#528B8B")

			self.button_upfile_data.configure(fg_color="#006400")
			self.button_upfile_official.configure(fg_color="#006400")
			self.button_enter.configure(fg_color="#006400")
			self.button_setting.configure(fg_color="#006400")
			self.button_undo.configure(fg_color="#006400")
			self.button_redo.configure(fg_color="#006400")
			self.button_save.configure(fg_color="#006400")


			self.label_table.configure(fg_color="#008B45")
			self.label_notification.configure(fg_color="#008B45")


			self.button_upfile_official_gui_two.configure(fg_color="#006241")
			self.button_addfile_data_gui_two.configure(fg_color="#006241")
			self.button_enter_gui_two.configure(fg_color="#006241")
			self.button_setting_gui_two.configure(fg_color="#006241")
			self.button_save_gui_two.configure(fg_color="#006241")
			self.button_remove_sheets.configure(fg_color="#A0522D")
			self.button_remove_all_sheets.configure(fg_color="#B22222")

			self.list_box_notification_gui_two.configure(background="#DDDDDD",foreground="#1C1C1C")

			self.label_title_gui_two.configure(fg_color="#2E8B57")
			self.label_notification_gui_two.configure(fg_color="#2E8B57")


			self.label_number_file_gui_two.configure(fg_color="#5F9EA0")
			self.label_number_sheets_gui_two.configure(fg_color="#5F9EA0")
			self.number_file_gui_two.configure(fg_color="#5F9EA0")
			self.number_sheets_gui_two.configure(fg_color="#5F9EA0")


			self.frame_setting_gui_two.configure(fg_color="#008B8B")


			self.frame_function_gui_two.configure(fg_color="#528B8B")
			self.frame_title_gui_two.configure(fg_color="#528B8B")
			self.frame_nofitication_gui2.configure(fg_color="#528B8B")

			self.frame_command_gui_two.configure(fg_color="#99D1D3")
		
			
			self.treeview_gui_two.tag_configure("color_file", background="#556B2F")

			self.treeview_gui_two.tag_configure("color_sheet_additional",foreground="#00688B")

	def chane_mode_theme(self):
	
		if int(self.switch_dark_mode.get()) == 1:
	
			ctk.set_appearance_mode("Dark")

			self.treeview.tag_configure('focus', background='#777777')		

		else:

			ctk.set_appearance_mode("light")

			self.treeview.tag_configure('focus', background='#99CC66')
			
		self.select_mode_theme = int(self.switch_dark_mode.get())

		self.chane_theme(event=None)

		self.toplevel_window.after(10, self.toplevel_window.lift)
		
	def pack_widget_toplevel(self):

		self.label_setting_toplevel_cut_value.grid(row=0,column=0,stick="nw",padx=3,pady=3)

		self.label_setting_toplevel_cut_value2.grid(row=0,column=1,stick="nw",padx=3,pady=3)

		#______________________________________________________________________________________________________________frame_cut_value


		self.frame_parameter.grid(row=0,column=0,stick="nsew",padx=5,pady=5)
		
		self.frame_cut_value.grid(row=1,column=1,stick="nsew",padx=5,pady=5)
	
		self.label_cut_top.grid(row=1,column=1,stick="s")
	
		self.spin_box_top.grid(row=2,column=1,stick="n",padx=5)
		
		self.label_cut_left.grid(row=2,column=0,stick="e",padx=15)
		
		self.spin_box_left.grid(row=3,column=0,stick="e",padx=5)

		self.label_cut_right.grid(row=2,column=2,stick="w",padx=15)
		
		self.spin_box_right.grid(row=3,column=2,stick="w",padx=5)

		self.label_cut_bottom.grid(row=5,column=1,stick="n")
		
		self.spin_box_bottom.grid(row=4,column=1,stick="s",padx=5)

		lis_frame_toplevel_cut_value = [self.label_cut_top,self.spin_box_top,self.label_cut_left,

		self.spin_box_left,self.label_cut_right,self.spin_box_right,self.label_cut_bottom,self.spin_box_bottom]

		Balance.function_balance(self.frame_cut_value,lis_frame_toplevel_cut_value,1,0)
		#________________________________________________________________________________________________frame_type_value

		self.frame_type_value.grid(row=1,column=0,stick="nsew",padx=5,pady=5)


		self.label_copy_cut.grid(row=0,column=0,padx=3,pady=3,stick="ew")

		self.combobox_copy_cut.grid(row=0,column=1,padx=3,pady=3,stick="ew")

		self.label_border.grid(row=0,column=2,padx=3,pady=3,stick="ns",rowspan=3)

		self.label_column_excel.grid(row=1,column=0,padx=3,pady=30,stick="ew")

		self.combobox_column_excel.grid(row=1,column=1,padx=3,pady=30,stick="ew")

		self.label_distance_excel.grid(row=2,column=0,padx=3,pady=3,stick="ew")

		self.combobox_distance_excel.grid(row=2,column=1,padx=3,stick="ew")

		self.switch_title.grid(row=0,column=3,padx=3,pady=3,stick="ew")

		self.switch_format.grid(row=1,column=3,padx=3,pady=3,stick="ew")

		self.switch_transpose.grid(row=2,column=3,padx=3,pady=3,stick="ew")


		list_widget_frame_type_value = [self.label_copy_cut,self.combobox_copy_cut,self.label_column_excel,

		self.combobox_column_excel,self.label_distance_excel,self.combobox_distance_excel,self.switch_title,

		self.switch_format,self.switch_transpose]

		Balance.function_balance(self.frame_type_value,list_widget_frame_type_value,1,0)


		list_frame_parameter = [self.frame_type_value,self.frame_cut_value]

		Balance.function_balance(self.frame_parameter,list_frame_parameter,1,0)

		#__________________________________________________________________________________________Frame_align

		self.frame_align.grid(row=1,column=0,stick="nsew",padx=5,pady=5)

		self.label_setting_toplevel_align.grid(row=0,column=0,stick="nw",padx=3,pady=3)

		self.frame_system.grid(row=1,column=0,stick="nsew",padx=5,pady=5)
		
		self.switch_dark_mode.grid(row=0,column=0,stick="ew",padx=5,pady=5)
		
		self.switch_notification.grid(row=0,column=2,stick="ew",padx=5,pady=5)
		
		self.switch_wrap_text.grid(row=0,column=4,stick="ew",padx=5,pady=5)

		
		self.label_language.grid(row=1,column=0,stick="ew",padx=5,pady=20)
		self.combobox_language.grid(row=1,column=1,stick="ew",padx=5,pady=20)
		self.label_theme.grid(row=1,column=2,stick="ew",padx=5,pady=20)
		self.combobox_theme.grid(row=1,column=3,stick="ew",padx=5,pady=20)
		self.label_regime.grid(row=1,column=4,stick="ew",padx=5,pady=20)
		self.button_delete_notification.grid(row=1,column=5,stick="ew",padx=5,pady=20)

		self.frame_path.grid(row=2,column=0,stick="ew",padx=5,columnspan=6,pady=10)

		self.entry_path_data.grid(row=1,column=0,stick="ew",padx=5,pady=5)
		self.entry_path_official.grid(row=2,column=0,stick="ew",padx=5,pady=5)
		self.entry_path_save.grid(row=3,column=0,stick="ew",padx=5,pady=5)

		self.button_path_data.grid(row=1,column=1,stick="ew",padx=5,pady=5)
		self.button_path_official.grid(row=2,column=1,stick="ew",padx=5,pady=5)
		self.button_path_save.grid(row=3,column=1,stick="ew",padx=5,pady=5)


		self.button_path_default.grid(row=3,column=3,stick="ew",padx=5,pady=5)
		self.button_path_application.grid(row=3,column=4,stick="ew",padx=5,pady=5)
		self.button_path_close.grid(row=3,column=5,stick="ew",padx=5,pady=5)


		lis_widget_path = [self.entry_path_data,self.entry_path_official,self.entry_path_save]


		Balance.function_balance(self.frame_path,lis_widget_path,1,0)

		lis_frame_align = [self.label_language,self.combobox_language,self.label_theme,self.combobox_theme,

		self.label_regime,self.button_delete_notification,self.frame_path]

		Balance.function_balance(self.frame_system,lis_frame_align,1,0)
		

		Balance.function_balance2(self.frame_align,self.frame_system)

		lis_frame_toplevel = [self.frame_parameter,self.frame_align]

		Balance.function_balance(self.toplevel_window,lis_frame_toplevel,50,50)


	def clear_notification(self):

		if self.mytab.get() == self.name_widget1:

			self.list_box_notification.delete(0, 'end')

		else:

			self.list_box_notification_gui_two.delete(0, 'end')

	
	def default_data_gui_one(self):

		



		Setting_Entry.default_data(self.entry_path_data,self.path_app)

		Setting_Entry.default_data(self.entry_path_official,self.path_app)

		Setting_Entry.default_data(self.entry_path_save,self.path_app)

		self.combobox_copy_cut.set("Copy") 
		
		self.combobox_distance_excel.set(1)

		self.spin_box_top.set(0)
		self.spin_box_left.set(0)
		self.spin_box_right.set(0)
		self.spin_box_bottom.set(0)

		self.switch_title.configure(variable=ctk.StringVar(value="off"))

		self.switch_format.configure(variable=ctk.StringVar(value="off"))

		self.switch_transpose.configure(variable=ctk.StringVar(value="off"))

		self.switch_var_close_the_file_when_done = ctk.StringVar(value ="on")

	def colse_toplevel(self):

		self.toplevel_window.destroy()

		self.toplevel_window = None

	def on_close_toplevel(self):

		if self.toplevel_window:
			self.toplevel_window.destroy()
			self.toplevel_window = None


	def setting_apply(self):

		self.number_cloumn_excel = self.combobox_column_excel.current() +1

		self.value_column_excel	= self.combobox_column_excel.get()

		self.value_distance_excel =	self.combobox_distance_excel.get()

		if self.mytab.get() == self.name_widget1:


			self.value_path_data = self.entry_path_data.get()

			self.value_path_offical = self.entry_path_official.get()

			self.value_path_save = self.entry_path_save.get()

			self.value_spin_box_left = self.spin_box_left.get()

			self.value_spin_box_right = self.spin_box_right.get()

			self.value_spin_box_top = self.spin_box_top.get()

			self.value_spin_box_bottom = self.spin_box_bottom.get()		

			# self.switch_var_take_title = ctk.StringVar(value=self.switch_title.get())

			# self.switch_var_wrap_text = ctk.StringVar(value=self.switch_wrap_text.get())

			# self.switch_var_transpose = ctk.StringVar(value=self.switch_transpose.get())

			# self.switch_var_format = ctk.StringVar(value=self.switch_format.get())

			# self.switch_var_notification = ctk.StringVar(value = self.switch_notification.get())


			self.value_copy_cut = self.combobox_copy_cut.get()

			self.value_entry_path_data = self.entry_path_data.get()



			self.value_entry_path_official = self.entry_path_official.get()

			self.value_entry_path_save = self.entry_path_save.get()


			try:
				self.combobox_find_title.current(self.curent_present_title)

				self.value_for_combobox_title(event=None)

				self.combobox_find_elemen_title.current(self.curent_present_elemen_title)
				
				if self.combobox_find_title.get() != self.text_all_of_combobox:
				
					self.find_value_unique(event=None)

				else:

					self.chane_data_sheet(event=None)

				if self.switch_var_notification.get() =="on":

					self.list_box_notification.insert(0,self.notification_gui_one18)
			except AttributeError:
				print("chưa upfile")
		else:

			self.value_entry_path_data = self.entry_path_data.get()



			self.value_entry_path_official = self.entry_path_official.get()

			self.value_entry_path_save = self.entry_path_save.get()

			self.value_entry_path_data = self.entry_path_data.get()
	

			self.value_combobox_regime = self.combobox_regime.get()

			self.value_type_format_gui2 = self.combobox_type_format_gui2.get()
			

			# self.switch_var_write_name_sheet_gui2 = ctk.StringVar(value = self.switch_write_name_sheet_gui2.get())

			# self.switch_var_format_gui2 = ctk.StringVar(value = self.switch_format_gui2.get())

			# self.switch_var_close_the_file_when_done = ctk.StringVar(value = self.switch_close_the_file_when_done.get())

		if self.combobox_language.get() != self.language_default:

			self.file_database_language = sqlite3.connect(self.path_database_language)

			language_set = Function_File_And_Data.query_data_language(self.file_database_language,self.name_table_database_language,self.combobox_language.get())

			Function_File_And_Data.change_language(self.file_database_language,self.name_table_database_language[1],self.combobox_language.get())

			self.language_default = Function_File_And_Data.query_language_choose(self.file_database_language,self.name_table_database_language[1])

			self.text_all_of_combobox = language_set[0]

			if self.combobox_sheet_data.get() !="":

				values_of_combobox = self.combobox_find_title['values']

				self.combobox_find_title['values'] = [self.text_all_of_combobox] + list(values_of_combobox[1:])

				self.combobox_find_title.current(self.curent_present_title)

			
			self.notification_gui_one1 = language_set[1]

			self.notification_gui_one2 = language_set[2]

			self.notification_gui_one3 = language_set[3]

			self.notification_gui_one4 = language_set[4]

			self.notification_gui_one5 = language_set[5]

			self.notification_gui_one6 = language_set[6]

			self.notification_gui_one7 = language_set[7]

			self.notification_gui_one8 = language_set[8]

			self.notification_gui_one9 = ast.literal_eval(language_set[9])

			self.notification_gui_one10 = ast.literal_eval(language_set[10])



			self.notification_gui_one11 = ast.literal_eval(language_set[11])

			self.notification_gui_one12 = language_set[12]

			self.notification_gui_one13 = language_set[13]

			self.notification_gui_one14 = language_set[14]
				
			self.notification_gui_one15 = language_set[15]

			self.notification_gui_one16 = ast.literal_eval(language_set[16])

			self.notification_gui_one17 = language_set[17]

			self.notification_gui_one18 = language_set[18]			


			lis_name_tap = list(self.mytab._segmented_button._buttons_dict)
		

			self.mytab._segmented_button._buttons_dict[lis_name_tap[0]].configure(text=language_set[19])


			self.mytab._segmented_button._buttons_dict[lis_name_tap[1]].configure(text=language_set[20])

			self.label_table.configure(text= language_set[21])
			self.label_find_title.configure(text= language_set[22])
			self.label_find_elemen_title.configure(text= language_set[23])
			self.label_notification.configure(text= language_set[24])

			self.label_notification_gui_two.configure(text= language_set[24])

			self.button_upfile_data.configure(text= language_set[25])

			self.button_addfile_data_gui_two.configure(text= language_set[25])

			self.button_upfile_official.configure(text= language_set[26])

			self.button_upfile_official_gui_two.configure(text= language_set[26])

			self.button_enter.configure(text= language_set[27])

			self.button_enter_gui_two.configure(text= language_set[27])

			self.button_setting.configure(text= language_set[28])

			self.button_setting_gui_two.configure(text= language_set[28])

			self.label_cut_top.configure(text= language_set[29])
			
			self.label_cut_left.configure(text= language_set[30])
			
			self.label_cut_right.configure(text= language_set[31])
			
			self.label_cut_bottom.configure(text= language_set[32])

			self.label_column_excel.configure(text= language_set[33])
			
			self.label_distance_excel.configure(text= language_set[34])

			self.switch_title.configure(text= language_set[35])
			
			self.switch_format.configure(text= language_set[36])



			self.switch_notification.configure(text= language_set[37])
			
			self.label_regime.configure(text= language_set[38])

			self.button_delete_notification.configure(text= language_set[39])

			self.button_path_default.configure(text= language_set[40])

			self.button_path_application.configure(text= language_set[41])
			
			self.button_path_close.configure(text= language_set[42])

			self.entry_path_data.configure(placeholder_text = language_set[43])

			self.entry_path_official.configure(placeholder_text =language_set[44])
			
			self.entry_path_save.configure(placeholder_text =language_set[45])

			self.label_setting_toplevel_cut_value.configure(text= language_set[46])
		
			self.label_setting_toplevel_cut_value2.configure(text= language_set[47])
			
			self.label_setting_toplevel_align.configure(text= language_set[48])

			self.label_title_gui_two.configure(text= language_set[49])
			
			self.label_number_file_gui_two.configure(text= language_set[50])
			
			self.label_number_sheets_gui_two.configure(text= language_set[51])

			self.label_copy_cut.configure(text= language_set[52])

			self.label_language.configure(text= language_set[59])

			if self.mytab.get() == self.name_widget2:

				self.switch_write_name_sheet_gui2.configure(text= language_set[53])
				
				self.switch_close_the_file_when_done.configure(text= language_set[54])

				self.combobox_regime['values'] = ast.literal_eval(language_set[55])

				self.combobox_regime.set(language_set[56])

				self.combobox_type_format_gui2['values'] = ast.literal_eval(language_set[57])

				self.combobox_type_format_gui2.set(language_set[58])


			self.name_widget12 = language_set[29]

			self.name_widget13 = language_set[30]

			self.name_widget14 = language_set[31]

			self.name_widget15 = language_set[32]

			self.name_widget16 = language_set[33]

			self.name_widget17 = language_set[34]

			self.name_widget18 = language_set[35]

			self.name_widget19 = language_set[36]

			self.name_widget20 = language_set[37]
			
			self.name_widget21 = language_set[38]

			self.name_widget22 = language_set[39]

			self.name_widget23 = language_set[40]	

			self.name_widget24 = language_set[41]

			self.name_widget25 = language_set[42]

			self.name_widget26 = language_set[43]

			self.name_widget27 = language_set[44]

			self.name_widget28 = language_set[45]

			self.name_widget29 = language_set[46]

			self.name_widget30 = language_set[47]

			self.name_widget31 = language_set[48]

			self.name_widget32 = language_set[49]

			self.name_widget33 = language_set[50]

			self.name_widget34 = language_set[51]

			self.name_widget35 = language_set[52]

			self.name_widget36 = language_set[53]

			self.name_widget37 = language_set[54]



			self.list_value_regime = ast.literal_eval(language_set[55])

			self.value_combobox_regime = language_set[56]

			self.lis_value_type_format_gui2 = ast.literal_eval(language_set[57])

			self.value_type_format_gui2 = language_set[58]

			self.name_widget38 = language_set[59]

			self.file_database_language.close()