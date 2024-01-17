from gui_custom import Toplevel_Setting
from tkinter import ttk
from tkinter import *
import customtkinter as ctk
from Function_Gui import Balance,Function_File_And_Data


class Setup_Two(Toplevel_Setting):

	def __init__(self,root):
		super().__init__(root)

		self.frame_gui_two()

		self.label_gui_two()

		self.canvas_gui_two()

		self.treeview_of_gui_two()

		self.button_gui_two()

		self.combobox_gui_two()

		self.list_box_gui_two()

		self.pack_widget_gui_two()

		try:

			self.button_addfile_data_gui_two.configure(image=self.image_add_file,compound="top")

			self.button_upfile_official_gui_two.configure(image=self.image_up_file,compound="left")

			self.button_enter_gui_two.configure(image=self.image_enter,compound="bottom")
		
			self.button_setting_gui_two.configure(image=self.image_setting,compound="left")
			
			self.button_save_gui_two.configure(image=self.image_save,compound="left")

			self.button_remove_sheets.configure(image=self.image_delete,compound="left")
			
			self.button_remove_all_sheets.configure(image=self.image_clear,compound="left")

		except AttributeError:

			print("không thấy icon button")


		self.list_box_notification_gui_two.configure(background="#003333",foreground="white")
		
		self.switch_var_write_name_sheet_gui2 = ctk.StringVar(value="off")

		self.switch_var_close_the_file_when_done = ctk.StringVar(value="on")

		
	
	def frame_gui_two(self):


		self.frame_command_gui_two = ctk.CTkFrame(master=self.tap_2,fg_color="#8B7B8B",corner_radius=15)

		self.frame_function_gui_two = ctk.CTkFrame(master=self.frame_command_gui_two,fg_color="#C0C0C0",corner_radius=15,width=250)

		self.frame_title_gui_two = ctk.CTkFrame(master=self.frame_command_gui_two,fg_color="#C0C0C0",corner_radius=10)	

		self.frame_nofitication_gui2 = ctk.CTkFrame(master=self.frame_function_gui_two,fg_color="#C0C0C0",corner_radius=15)

		self.frame_setting_gui_two = ctk.CTkFrame(master=self.frame_function_gui_two,fg_color="#668B8B",corner_radius=15)

		

		self.frame_main_gui_two = ctk.CTkFrame(master=self.frame_title_gui_two,fg_color="#778899",corner_radius=15)


	def label_gui_two(self):

		self.label_title_gui_two = ctk.CTkLabel(self.frame_title_gui_two,text=self.name_widget32,corner_radius=10,fg_color="#9C9C9C",text_color="black")


		self.label_number_file_gui_two = ctk.CTkLabel(self.frame_title_gui_two,text=self.name_widget33,corner_radius=5,fg_color="#708090",text_color="black")

		self.label_number_sheets_gui_two = ctk.CTkLabel(self.frame_title_gui_two,text=self.name_widget34,corner_radius=5,fg_color="#708090",text_color="black")

		self.number_file_gui_two = ctk.CTkLabel(self.frame_title_gui_two,text="-",corner_radius=5,fg_color="#708090",text_color="black")

		self.number_sheets_gui_two = ctk.CTkLabel(self.frame_title_gui_two,text="-",corner_radius=5,fg_color="#708090",text_color="black")

		self.label_notification_gui_two = ctk.CTkLabel(self.frame_nofitication_gui2,text=self.name_widget7,corner_radius=10,fg_color="#9C9C9C",text_color="black")
	def canvas_gui_two(self):

		self.canvas_three = Canvas(self.frame_title_gui_two, height=1, background='gray')

		self.canvas_four = Canvas(self.frame_setting_gui_two, height=1, background='red')

		self.canvas_five = Canvas(self.frame_setting_gui_two, height=1, background='red')

	def treeview_of_gui_two(self):

		self.treeview_gui_two = ttk.Treeview(self.frame_main_gui_two)

		self.treeview_gui_two["columns"] = ("Name Sheet","Path","Size")

		list_name = ["Name Sheet","Path","Size"]

		minsize = 558//len(self.treeview_gui_two['columns'])

		self.treeview_gui_two.heading("#0", text="Name File")
		self.treeview_gui_two.column("#0",stretch=False,minwidth=minsize)


		for col_name,name in zip(self.treeview_gui_two["columns"],list_name):

			self.treeview_gui_two.heading(col_name, text=name,anchor="w")

			self.treeview_gui_two.column(col_name, stretch=False,minwidth=minsize)

		self.xscrollbar_gui_two = ctk.CTkScrollbar(self.frame_main_gui_two, orientation='horizontal',border_spacing=5,corner_radius=10,button_hover_color="black",command=self.treeview_gui_two.xview)
		
		self.treeview_gui_two.configure(xscrollcommand=self.xscrollbar_gui_two.set)		

		self.yscrollbar_gui_two = ctk.CTkScrollbar(self.frame_main_gui_two, orientation='vertical',border_spacing=5,corner_radius=10,button_hover_color="black",command=self.treeview_gui_two.yview)

		self.treeview_gui_two.configure(yscrollcommand=self.yscrollbar_gui_two.set)


	def button_gui_two(self):

		self.button_upfile_official_gui_two = ctk.CTkButton(self.frame_setting_gui_two, text=self.name_widget9,font=("Helvetica",9),corner_radius=50,fg_color="#2F4F4F",hover_color="#FF6600",border_width=0.7,border_color="black")

		self.button_addfile_data_gui_two = ctk.CTkButton(self.frame_setting_gui_two, text=self.name_widget8,corner_radius=20,fg_color="#2F4F4F",hover_color="#FF6600",border_width=0.7,border_color="black")

		self.button_enter_gui_two = ctk.CTkButton(self.frame_setting_gui_two, text=self.name_widget10,border_width=1,border_color="black",fg_color="#2F4F4F")

		self.button_setting_gui_two = ctk.CTkButton(self.frame_setting_gui_two, text=self.name_widget11,border_width=1,border_color="black",fg_color="#2F4F4F")

		self.button_save_gui_two = ctk.CTkButton(self.frame_setting_gui_two, text="Save",border_width=1,border_color="black",fg_color="#2F4F4F")

		self.button_remove_sheets = ctk.CTkButton(self.frame_setting_gui_two, text="Delete",border_width=1,border_color="black",fg_color="#FF6633")

		self.button_remove_all_sheets = ctk.CTkButton(self.frame_setting_gui_two, text="Clear",border_width=1.2,border_color="black",fg_color="#C82E31")

	def combobox_gui_two(self):


		self.combobox_sheet_official_gui_two =ttk.Combobox(self.frame_setting_gui_two,state="readonly",justify='center',width=15)

	def list_box_gui_two(self):

		self.list_box_notification_gui_two = Listbox(self.frame_nofitication_gui2,font=("Arial", 10, "bold"),border=2,selectmode=MULTIPLE)

		self.xscrollbar_lis_box_gui_two = ctk.CTkScrollbar(self.frame_nofitication_gui2, orientation='horizontal',border_spacing=5,corner_radius=10,button_hover_color="black",command=self.list_box_notification_gui_two.xview)
		
		self.list_box_notification_gui_two.configure(xscrollcommand=self.xscrollbar_lis_box_gui_two.set)		

		self.yscrollbar_lis_box_gui_two = ctk.CTkScrollbar(self.frame_nofitication_gui2, orientation='vertical',border_spacing=5,corner_radius=10,button_hover_color="black",command=self.list_box_notification_gui_two.yview)

		self.list_box_notification_gui_two.configure(yscrollcommand=self.yscrollbar_lis_box_gui_two.set)

	def pack_widget_gui_two(self):	

		self.frame_command_gui_two.grid(row=0,column=0,padx=5,pady=5,stick="nsew")
		
		self.frame_function_gui_two.grid(row=0,column=0,padx=5,pady=5,stick="nsew")
		
		self.frame_title_gui_two.grid(row=0,column=1,padx=5,pady=5,stick="nwes")

		self.frame_main_gui_two.grid(row=3,column=0,padx=5,pady=5,stick="nsew",columnspan=4)

		self.frame_command_gui_two.grid_propagate(False)
		self.frame_function_gui_two.grid_propagate(False)
		self.frame_title_gui_two.grid_propagate(False)

		#_______________________________________________________________________________________
		self.label_title_gui_two.grid(row=2,column=0,columnspan=4,stick="ew",padx=5,pady=5)

		self.canvas_three.grid(row=1,column=0,columnspan=4,stick="ew",padx=5,pady=5)

		self.label_number_file_gui_two.grid(row=0,column=0,padx=5,pady=5)
		self.label_number_sheets_gui_two.grid(row=0,column=2,padx=5,pady=5)
		self.number_file_gui_two.grid(row=0,column=1,padx=5,pady=5)
		self.number_sheets_gui_two.grid(row=0,column=3,padx=5,pady=5)

		#___________________________________________________________________________________________



		self.treeview_gui_two.grid(row=1,column=0,sticky="nsew")

		
		self.yscrollbar_gui_two.grid(row=1,column=1,sticky="sen")
		self.xscrollbar_gui_two.grid(row=1,column=0,sticky="wse")

		#_____________________________________________________________________________________________

		self.frame_nofitication_gui2.grid(row=0,column=0,sticky="nsew")
		self.frame_setting_gui_two.grid(row=1,column=0,sticky="nsew")

		#_____________________________________________________________________________________________

		

		self.button_upfile_official_gui_two.grid(row=1,column=0,padx=5,pady=5)

		self.combobox_sheet_official_gui_two.grid(row=2,column=0,pady=5)

		self.button_addfile_data_gui_two.grid(row=1,column=1,rowspan=2,stick="nsew",padx=5,pady=5)

		self.canvas_four.grid(row=3,column=0,padx=5,pady=5,columnspan=2)

		self.button_remove_all_sheets.grid(row=4,column=0,padx=5,stick="ns",pady=5)

		self.button_remove_sheets.grid(row=4,column=1,padx=5,stick="ns",pady=5)

		self.canvas_five.grid(row=5,column=0,padx=5,pady=5,columnspan=2)

		self.button_save_gui_two.grid(row=6,column=0,padx=5,stick="ns",pady=5)
		
		self.button_setting_gui_two.grid(row=7,column=0,padx=5,stick="ns",pady=5)

		self.button_enter_gui_two.grid(row=6,column=1,rowspan=2,stick="nsew",padx=5,pady=5)


		#_________________________________________________________

		self.label_notification_gui_two.grid(row=0, column=0,padx=10,stick="ew",pady=10,columnspan=2)

		self.list_box_notification_gui_two.grid(row=1, column=0,padx=3,stick="nsew")

		self.yscrollbar_lis_box_gui_two.grid(row=1,column=1,sticky="sen")
		self.xscrollbar_lis_box_gui_two.grid(row=2,column=0,sticky="wse")

class Balance_Widget_Gui_Two(Setup_Two):

	def __init__(self,root):
		super().__init__(root)


		self.balance_gui_two()

		self.button_setting_gui_two.configure(command=self.open_toplevel_gui_two)

	def open_toplevel_gui_two(self):

		self.open_toplevel()

		self.frame_cut_value.grid_forget()

		self.label_setting_toplevel_cut_value2.grid_forget()
		
		self.combobox_copy_cut.grid_forget()

		self.switch_title.grid_forget()
		self.switch_transpose.grid_forget()

		self.switch_format.grid_forget()

	
		self.widget_toplevel_gui_two()


	def widget_toplevel_gui_two(self):



		self.label_copy_cut.configure(text=self.name_widget35)

		self.combobox_regime = ttk.Combobox(self.frame_type_value,state="readonly",justify='center',values=self.list_value_regime,width=5)

		self.combobox_regime.set(self.value_combobox_regime)

		

		self.switch_write_name_sheet_gui2 = ctk.CTkSwitch(master=self.frame_type_value, text=self.name_widget36,font=("Helvetica",13),text_color="black",variable=self.switch_var_write_name_sheet_gui2,onvalue="on", offvalue="off")
		
		self.combobox_type_format_gui2 = ttk.Combobox(self.frame_type_value,state="readonly",justify='center',values=self.lis_value_type_format_gui2,width=30)

		self.combobox_type_format_gui2.set(self.value_type_format_gui2)

		self.switch_close_the_file_when_done = ctk.CTkSwitch(master=self.frame_type_value, text=self.name_widget37,font=("Helvetica",13),text_color="black",variable=self.switch_var_close_the_file_when_done,onvalue="on", offvalue="off")

		self.switch_write_name_sheet_gui2.grid(row=0,column=3,padx=3,pady=3,stick="ew",columnspan=2)

		self.combobox_type_format_gui2.grid(row=1,column=3,padx=3,pady=3,stick="w")	

		self.switch_close_the_file_when_done.grid(row=2,column=3,padx=3,pady=3,stick="ew",columnspan=2)

		self.combobox_regime.grid(row=0,column=1,padx=3,pady=3,stick="ew")

		self.frame_type_value.grid(row=1,column=0,stick="nsew",padx=5,pady=5,columnspan=2)

	def balance_gui_two(self):

		Balance.function_balance2(self.tap_2,self.frame_command_gui_two)

		Balance.function_balance2(self.frame_title_gui_two,self.frame_main_gui_two)

		Balance.function_balance2(self.frame_main_gui_two,self.treeview_gui_two)
	
		Balance.function_balance2(self.frame_command_gui_two,self.frame_title_gui_two)
		
		Balance.function_balance2(self.frame_function_gui_two,self.frame_nofitication_gui2)

		Balance.function_balance2(self.frame_nofitication_gui2,self.list_box_notification_gui_two)
		
		
		lis_widget_frame_setting = [self.button_addfile_data_gui_two,self.button_enter_gui_two]


		Balance.function_balance(self.frame_setting_gui_two,lis_widget_frame_setting,1,0)

