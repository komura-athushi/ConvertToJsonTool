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
#data = excel_data.ExcelData()
#data.load("file/data.xlsx")

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

        #ヒエラルキービューの初期化
        self.init_hierarchy()

        #インスペクタービューの初期化
        self.init_inspector()

        #エクエルデータを保持するリスト。
        self.books = {}
        self.number = 1
        self.sum = self.number

    def init_button(self):
        #左フレームの作成と設置
        self.left_frame = ttk.Frame(self.master,width=constant.LEFT_FRAME_WIDTH,height=constant.LEFT_FRAME_HEIGHT)
        self.left_frame.place(relx=0,rely=0)

        self.inspector_sheet_name = tk.Label(self.left_frame,text='シート一覧')
        self.inspector_sheet_name.place(x=0.1,y=230.)

        #ボタンの作成
        self.load_button = tk.Button(self.left_frame, text="Excelファイル読み込み", font=("MSゴシック", "12", "bold"),
        command=self.load_excel)

        self.export_button = tk.Button(self.left_frame, text="Jsonファイル出力", font=("MSゴシック", "12", "bold"),
        command=self.export_json)

        #ボタンの設置
        self.load_button.place(relx=constant.LOAD_BUTTON_RELX,rely=constant.LOAD_BUTTON_RELY)
        self.export_button.place(relx=constant.EXPORT_BUTTON_RELX,rely=constant.EXPORT_BUTTON_RELY)

        

    def load_excel(self):
        #読み込むファイルの拡張子を指定    
        typ = [('Excelファイル','*.xlsx')]
        #ファイル選択ダイアログを表示
        fn = filedialog.askopenfilename(filetypes=typ)

        book = excel_data.ExcelData()
        #エクセルファイルを読み込む
        if book.load(fn,self.sum) == False:
            #ファイルオープンに失敗したら。
            return

        #リストに追加
        self.books[self.sum] =book
        self.sum += 1

        #リストボックスに名前を追加
        self.excel_list.insert(tk.END, book.file_name) 
        return

    def delete_excel(self):
        number = self.excel_list.curselection()
        #何も選択されてなかったら処理をしない
        if len(number) == 0:
            return
        #リストボックスから削除
        self.excel_list.delete(number)

        i = 0
        number2 = 0
        for i in self.books:
            if number[0] == number2:
                #リストから削除
                self.books.pop(i)
                break
            number2 += 1
        return

    def export_json(self):
        return

    def null(self):
        return

    def select_listbox(self,event):
        number = self.excel_list.curselection()
        #何も選択されてなかったら処理しない
        if len(number) == 0:
            return
        number2 = 0
        excel = None
        for i in self.books:
            if number[0] == number2:
                excel = self.books[i]
                break
            number2+=1

        #リストをクリア
        self.sheet_list.delete(0,tk.END)

        sheets = excel.get_sheets()
        for sheet in sheets:
            #リストボックスに名前を追加
            self.sheet_list.insert(tk.END, sheet.sheet_name)

    def init_hierarchy(self):

        #ヒエラルキーフレームの作成と設定
        #self.right_frame = ttk.Frame(self.master,width=constant.RIGHT_FRAME_WIDTH,height=constant.RIGHT_FRAME_HEIGHT)
        self.hierarchy_frame = ttk.Frame(self.master)

        self.excel_list = tk.Listbox(self.hierarchy_frame,
        listvariable=None,
        selectmode='single',
        width=37,
        height=29)

        self.excel_list.bind('<<ListboxSelect>>', self.select_listbox)
        self.hierarchy_frame.place(relx=0.5,rely=0)

        #フレームを親にスクロールバーを生成
        self.list_bar_y = tk.Scrollbar(self.hierarchy_frame,orient=tk.VERTICAL,command=self.excel_list.yview)
        #スクロールバーを配置
        self.list_bar_y.pack(side=tk.RIGHT, fill=tk.Y)
        #キャンバスとスクロールバーを紐づける
        self.excel_list.config(yscrollcommand=self.list_bar_y.set)
        #なんかよくわからんがスクロールバーが自動で調整される
        self.excel_list.see('end')
        
         #リストを設置
        self.excel_list.pack()

        #画像の削除ボタンを生成
        self.delete_button = ttk.Button(
        self.hierarchy_frame,
        text='削除',
        command=self.delete_excel)
        self.delete_button.pack(side=tk.LEFT)

    def init_inspector(self):

        self.inspector_frame = ttk.Frame(self.master)


        self.sheet_list = tk.Listbox(self.inspector_frame,
        listvariable=None,
        selectmode='single',
        width=33,
        height=15)

       

        self.inspector_frame.place(relx=0,rely=0.5)

             


        #フレームを親にスクロールバーを生成
        self.sheet_list_bar_y = tk.Scrollbar(self.inspector_frame,orient=tk.VERTICAL,command=self.sheet_list.yview)
        #スクロールバーを配置
        self.sheet_list_bar_y.pack(side=tk.RIGHT, fill=tk.Y)
        #キャンバスとスクロールバーを紐づける
        self.sheet_list.config(yscrollcommand=self.sheet_list_bar_y.set)
        #なんかよくわからんがスクロールバーが自動で調整される
        self.sheet_list.see('end')
        
         #リストを設置
        self.sheet_list.pack()


    
        return


#インスタンスを作成。
root = tk.Tk()
app = Application(master=root)
#ループ。
app.mainloop()



