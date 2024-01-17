from PIL import Image
import customtkinter as ctk
import os
import sqlite3
from Function_Gui import Function_File_And_Data,Database_Processing

import ast
class database_language:

	def __init__(self):

		self.path_app = os.getcwd()

		# folder_database

		try:

			self.path_database_language = os.path.join(os.getcwd()+r'\folder_database','language_gui.db')

			self.path_database_undo_redo = os.path.join(os.getcwd()+r'\folder_database','undo_redo.db')
			
			self.path_sub_database = os.path.join(os.getcwd()+r'\folder_database','sub_mydatabase.db')
			
			self.path_database = os.path.join(os.getcwd()+r'\folder_database','mydatabase.db')

			self.file_database_language = sqlite3.connect(self.path_database_language)

			self.database_undo_redo = sqlite3.connect(self.path_database_undo_redo)	

			self.file_sub_database = sqlite3.connect(self.path_sub_database)

			self.file_database = sqlite3.connect(self.path_database)

		except sqlite3.OperationalError as d:

			create_folder = os.getcwd()+r'\folder_database'

			os.makedirs(create_folder)

			self.path_database_language = os.path.join(os.getcwd()+r'\folder_database','language_gui.db')

			self.path_database_undo_redo = os.path.join(os.getcwd()+r'\folder_database','undo_redo.db')
			
			self.path_sub_database = os.path.join(os.getcwd()+r'\folder_database','sub_mydatabase.db')
			
			self.path_database = os.path.join(os.getcwd()+r'\folder_database','mydatabase.db')

			self.file_database_language = sqlite3.connect(self.path_database_language)

			self.database_undo_redo = sqlite3.connect(self.path_database_undo_redo)	

			self.file_sub_database = sqlite3.connect(self.path_sub_database)

			self.file_database = sqlite3.connect(self.path_database)


		self.name_table_database_language = ["language_table","language_choose"]

		language_vienames = ["vietnames","Tất Cả","Kết Nối File Excel Thất Bại\nCó Thể Bạn Chưa Cài Excel\nHoặc Excel Không Tương Thích\n Hãy Thử 'RESTART' Lại Máy",

		"Bạn chưa chọn file data, hãy chọn lại file data","Bạn chưa chọn file official,hãy chọn lại file official","Bạn chưa mở file official",

		"sheet nầy đã không còn data vui lòng kiểm tra hoặc chọn sheet khác","Bạn đã đóng file official,vui lòng up lại file official",

		"Bạn chưa chọn data hãy chọn ít nhất 1 data và nhấn enter","Bạn vẫn chưa up data cho app",["Thông Báo", "Bạn không Thể Lấy Định Dạng\nNếu Vùng Chọn của bạn có merge\nhãy kiểm tra vùng bạn chọn hoặc bạn có lấy tiêu đề có merge không?"],

		["Thông Báo", "Do Excel Không Cho Lấy Nhiều Cell 1 Lần\nNên Bạn Sẽ Không Lấy Định Dạng Được Khi Chọn Trên 1 Cột\nNên Chỉ Lấy Được Data Thôi"],

		["Thông Báo", "Bạn Đẫ Đóng File data Nên Giờ Không Thể Lấy Được Định Dạng"],"Đã đóng lấy định dạng","Bạn Đã Ẩn Hết Data","đã được up",

		"Bạn chưa +add file",["Retry or Cancel", "Bạn Có Chắc Muốn Xóa Toàn Bộ File Đã Log?"],"Bạn vừa xóa toàn bộ file đã log","Cài Đặt Thành Công",

		"Xuất Data","Gộp Sheet","Bảng Data Thao Tác","Tìm Theo Tiêu Đề","Các Data Tiêu Đề","Bảng Thao Tác","File DaTa",

		"File Official","Xác Nhận","Cài Đặt","Trên","Trái","Phải","Dưới","Cột Xuất","Cách Khoảng Data","Lấy Tiêu Đề","Lấy Định Dạng\n(Chỉ dùng được khi\nfile data còn mở)"

		,"Thông Báo","Xóa Hết Thông Báo","Xóa","Mặc Định","Lưu Áp Dụng","Đóng","Đường Dẫn Up File Data","Đường Dẫn Up File Official","Đường Dẫn Up Save File",

		"Điều Chỉnh Thông Số","Ẩn Các Cột","Hệ Thống","Bảng Danh Sách File","Số File Đang Log","Số Sheets Đang Có","Chế Độ Gộp",

		"Điền Tên Sheets Và Tên File","Đóng File Khi Gộp Xong",["Gộp Data Vào Sheets","Gộp Sheets Vào File"],"Gộp Data Vào Sheets",

		["Lấy Data Và Định Dạng","Chỉ Lấy Data","Chỉ Lấy Định Dạng","Chỉ Lấy Công Thức","Lấy Toàn Bộ"],"Lấy Data Và Định Dạng","Ngôn Ngữ"]

		language_english = ["english", "All", "Failed to Connect Excel File\nYou may not have Excel installed\nOr Excel is not compatible\n Try 'RESTART' your machine",

		"You haven't selected a data file, please choose the data file again", "You haven't selected an official file, please choose the official file again", "You haven't opened the official file",

		"This sheet no longer has data, please check or select another sheet", "You have closed the official file, please upload the official file again",

		"You haven't selected data, please select at least 1 data and press enter", "You still haven't uploaded data for the app", ["Notification", "You Cannot Retrieve Format\nIf Your Selected Range is merged\nplease check the range you selected or check if you have merged headers?"],

		["Notification", "Excel Doesn't Allow Retrieving Multiple Cells at Once\nSo You Won't Get the Format When Selecting Across One Column\nYou Can Only Get Data"],

		["Notification", "You Have Closed the Data File So You Cannot Retrieve the Format Now"], "Closed to get the format", "You Have Hidden All Data", "has been uploaded",

		"You haven't +add a file", ["Retry or Cancel", "Are You Sure You Want to Delete All Logged Files?"], "You just deleted all logged files", "Installation Successful",

		"Export Data", "Merge Sheets", "Data Operation Table", "Search by Title", "Header Data", "Operation Table", "Data File",

		"Official File", "Confirm", "Settings", "On", "Left", "Right", "Under", "Export Column", "Data Range", "Get Header", "Get Format\n(Only applicable when\nthe data file is still open)"

		, "Notification", "Delete All Notifications", "Delete", "Default", "Save Apply", "Close", "Path to Upload Data File", "Path to Upload Official File", "Path to Upload Save File",

		"Adjust Parameters", "Hide Columns", "System", "List of Files", "Number of Files Logged", "Number of Sheets Currently Available", "Merge Mode",

		"Enter Sheet Names and File Names", "Close File After Merging", ["Merge Data Into Sheets", "Merge Sheets Into File"], "Merge Data Into Sheets",

		["Retrieve Data and Format", "Only Retrieve Data", "Only Retrieve Format", "Only Retrieve Formulas", "Retrieve All"] ,"Retrieve Data and Format","Language"]

		# 公式ファイル

		language_japanese = ["japanese","全て","Excelファイルの接続に失敗しました\nおそらくExcelがインストールされていないか、または互換性がありません\n 'RESTART' ボタンを押してみてください",

		"データファイルを選択していません、データファイルを再選択してください","公式ファイルを選択していません、公式ファイルを再選択してください","公式ファイルを開いていません",

		"このシートにはデータがありません。シートを確認するか、別のシートを選択してください","公式ファイルを閉じています。再度アップロードしてください",

		"データを選択していません。少なくとも1つのデータを選択し、Enterキーを押してください","アプリにデータをまだアップしていません",["お知らせ", "フォーマットを取得できません\n選択した領域にマージがある場合は、選択した領域またはマージされたヘッダーを確認してくださいか?"],

		["お知らせ", "Excelが1度に複数のセルを取得するのを許可していないため、1列で選択するとフォーマットが取得できません\nそのため、データだけを取得できます"],

		["お知らせ", "データファイルを既に閉じているため、現在はフォーマットを取得できません"],"フォーマットを取得しました","データをすべて非表示にしました","アップロードが完了しました",

		"ファイルを追加していません",["再試行またはキャンセル", "すべてのログファイルを削除してもよろしいですか?"],"すべてのログファイルを削除しました","設定が成功しました",

		"データの転送","シートの結合","データ操作テーブル","タイトルで検索","タイトルデータ","操作テーブル","データ",

		"公式 ","確認","設定","上","左","右","下","エクスポート列","データ間隔の取得","ヘッダーの取得","フォーマットの取得\n(データファイルがまだ開\nいている場合のみ使用可能"

		,"お知らせ","すべての通知を削除","削除","デフォルト","適用して保存","閉じる","データファイルのアップロードパス","公式ファイルのアップロードパス","保存ファイルのアップロードパス",

		"パラメータの調整","列を非表示にする","システム","ファイルリスト","ログに記録されているファイルの数","現在のシート数","結合モード",

		"シート名とファイル名を入力してください","結合が完了したらファイルを閉じる",["データをシートに結合","シートをファイルに結合"],"データをシートに結合",

		["データとフォーマットを取得","データのみ取得","フォーマットのみ取得","数式のみ取得","すべて取得"],"データとフォーマットを取得","言語"]


		language_chinese = ["chinese", "全部", "连接Excel文件失败\n可能您尚未安装Excel\n或Excel不兼容\n请尝试重新启动您的计算机",

		"您尚未选择数据文件，请重新选择数据文件", "您尚未选择官方文件，请重新选择官方文件", "您尚未打开官方文件",

		"此表没有数据，请检查或选择其他表", "您已关闭官方文件，请重新上传官方文件",

		"您尚未选择数据，请至少选择1个数据并按Enter键", "您仍未为应用程序上传数据", ["通知", "您无法获取格式\n如果您选择的区域已合并\n请检查您选择的区域或您是否选择了包含合并的标题？"],

		["通知", "由于Excel不允许一次提取多个单元格\n因此在一列上进行选择时，您将无法提取格式\n因此只能提取数据"],

		["通知", "您已关闭数据文件，因此现在无法获取格式"], "已关闭提取格式", "您已隐藏所有数据", "已上传",

		"您尚未+添加文件", ["重试或取消", "您确定要删除所有已记录的文件吗？"], "您刚刚删除了所有已记录的文件", "安装成功",

		"导出数据", "合并工作表", "操作数据表", "按标题查找", "标题数据", "操作表", "数据文件",

		"官方文件", "确认", "安装", "在上方", "在左边", "在右边", "在底部", "导出列", "数据间距", "获取标题", "获取格式\n（仅在\n数据文件仍在打开时有效）"

		, "通知", "清除所有通知", "删除", "默认", "保存应用", "关闭", "上传数据文件路径", "上传官方文件路径", "上传保存文件路径",

		"调整参数", "隐藏列", "系统", "文件列表", "正在记录的文件数", "正在使用的表数", "合并模式",

		"填写表格名称和文件名", "合并完成后关闭文件", ["合并数据到表格", "合并表格到文件"], "合并数据到表格",

		["提取数据和格式", "仅提取数据", "仅提取格式", "仅提取公式", "提取所有"], "提取数据和格式","語言"]


		lis_language = [language_vienames,language_english,language_japanese,language_chinese]

		
		Function_File_And_Data.create_value_database_language(self.file_database_language,self.name_table_database_language,lis_language)
		
		len_data_table =  Database_Processing.check_data_existence(self.file_database_language,self.name_table_database_language[0])

		if len_data_table[0] ==0:

			Function_File_And_Data.push_value_database_language(self.file_database_language,self.name_table_database_language,lis_language)


		del language_vienames

		del language_english

		del language_japanese

		del language_chinese

		del lis_language

class Language(database_language):

	def __init__(self):

		super().__init__()


		self.language_default = Function_File_And_Data.query_language_choose(self.file_database_language,self.name_table_database_language[1])
		
		language_set = Function_File_And_Data.query_data_language(self.file_database_language,self.name_table_database_language,self.language_default)

		self.text_all_of_combobox = language_set[0]

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

		self.name_widget1 = language_set[19]

		self.name_widget2 = language_set[20]

		# self.name_widget3 = "Bảng Data Thao Tác"

		self.name_widget4 = language_set[21]

		self.name_widget5 = language_set[22]

		self.name_widget6 = language_set[23]

		self.name_widget7 = language_set[24]

		self.name_widget8 = language_set[25]

		self.name_widget9 = language_set[26]

		self.name_widget10 = language_set[27]

		self.name_widget11 = language_set[28]

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

		self.name_widget38 = language_set[59]


		self.list_value_regime = ast.literal_eval(language_set[55])

		self.value_combobox_regime = language_set[56]

		self.lis_value_type_format_gui2 = ast.literal_eval(language_set[57])

		self.value_type_format_gui2 = language_set[58]


class Image_Gui(Language):

	def __init__(self):

		super().__init__()

		try:

			self.image_up_file = ctk.CTkImage(Image.open(self.path_app+r"\icon_button\up_file_image.png"))

			self.image_add_file = ctk.CTkImage(Image.open(self.path_app+r"\icon_button\add_file_image.png"))

			self.image_enter = ctk.CTkImage(Image.open(self.path_app+r"\icon_button\enter_image.png"))

			self.image_setting = ctk.CTkImage(Image.open(self.path_app+r"\icon_button\setting_image.png"))	

			self.image_save = ctk.CTkImage(Image.open(self.path_app+r"\icon_button\save_image.png"))	

			self.image_undo = ctk.CTkImage(Image.open(self.path_app+r"\icon_button\undo_image.png"))

			self.image_redo = ctk.CTkImage(Image.open(self.path_app+r"\icon_button\redo_image.png"))

			self.image_delete = ctk.CTkImage(Image.open(self.path_app+r"\icon_button\delete_image.png"))

			self.image_clear = ctk.CTkImage(Image.open(self.path_app+r"\icon_button\_clear_image.png"))

		except FileNotFoundError:

			create_folder = os.getcwd()+r'\icon_button'

			if not os.path.exists(create_folder):

				os.makedirs(create_folder)

