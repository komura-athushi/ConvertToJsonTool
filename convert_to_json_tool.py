#ウィンドウを表示するためのライブラリ
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
from tkinter import messagebox
#Excelファイルを開くためのライブラリ
#xlsxファイルには対応しなくなった
#import xlrd

#エクセルのデータを扱う。
import excel_data

#定数データ
import constant

#エクセルを読み込む。
data = excel_data.ExcelData()
data.load("file/data.xlsx")

#ウィンドウクラス。
class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        #なんか作って
        self.master = master
        #名前を設定
        self.master.title('Convert To Json Tool')
        window_size = str(constant.WINDOWS_WIDTH) + 'x' + str(constant.WINDOWS_HEIGHT)
        #とりあえずウィンドウはこれくらいで
        self.master.geometry(window_size)
        #これをしないとフレームがどうのこうので
        #placeしても表示されない
        self.pack(expand=1, fill=tk.BOTH)

        #ボタンの初期化
        self.init_button()
        #リストの初期化
        self.init_list_box()

        #エクエルデータを保持するリスト。
        self.books = []


    def init_button(self):
        #左フレームの作成と設置
        self.left_frame = ttk.Frame(self.master,width=constant.LEFT_FRAME_WIDTH,height=constant.LEFT_FRAME_HEIGHT)
        self.left_frame.place(relx=0,rely=0)

        #ボタンの作成
        self.load_button = tk.Button(self.left_frame, text="Excelファイル読み込み", font=("MSゴシック", "12", "bold"),
        command=self.load_excel)

        self.export_button = tk.Button(self.left_frame, text="Jsonファイル出力", font=("MSゴシック", "12", "bold"),
        command=self.export_json)

        #ボタンの設置
        self.load_button.place(relx=constant.LOAD_BUTTON_RELX,rely=constant.LOAD_BUTTON_RELY)
        self.export_button.place(relx=constant.EXPORT_BUTTON_RELX,rely=constant.EXPORT_BUTTON_RELY)

    def load_excel(self):
        return

    def export_json(self):
        return

    def init_list_box(self):

        #右フレームの作成と設定
        #self.right_frame = ttk.Frame(self.master,width=constant.RIGHT_FRAME_WIDTH,height=constant.RIGHT_FRAME_HEIGHT)
        self.right_frame = ttk.Frame(self.master)

        self.project_list = tk.Listbox(self.right_frame,
        listvariable=None,
        selectmode='single',
        width=100,
        height=200)

        self.right_frame.place(relx=0.5,rely=0)

        #フレームを親にスクロールバーを生成
        self.project_bar_y = tk.Scrollbar(self.right_frame,orient=tk.VERTICAL,command=self.project_list.yview)
        #スクロールバーを配置
        self.project_bar_y.pack(side=tk.RIGHT, fill=tk.Y)
        #キャンバスとスクロールバーを紐づける
        self.project_list.config(yscrollcommand=self.project_bar_y.set)
        #なんかよくわからんがスクロールバーが自動で調整される
        self.project_list.see('end')
        
         #リストを設置
        self.project_list.pack()
        
        return

    


#インスタンスを作成。
root = tk.Tk()
app = Application(master=root)
#ループ。
app.mainloop()



