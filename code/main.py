from function_gui_two_custom import Command_Gui_Two
import customtkinter as ctk
from Function_Gui import Balance


class Main(Command_Gui_Two):
	
	def __init__(self,root):
			super().__init__(root)

if __name__ == "__main__":

	root = ctk.CTk()		

	Main(root)

	root.title("App Python Excel") 

	ctk.set_appearance_mode("Dark")

	ctk.set_default_color_theme("green")

	Balance.center_window(root,860,520)

	root.mainloop()