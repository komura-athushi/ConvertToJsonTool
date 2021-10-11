#ウィンドウを表示するためのライブラリ
import tkinter as tk

#Excelファイルを開くためのライブラリ
#xlsxファイルには対応しなくなった
#import xlrd

#エクセルのデータを扱う。
import excel_data

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
        #とりあえずウィンドウはこれくらいで
        self.master.geometry("600x600")
        #これをしないとフレームがどうのこうので
        #placeしても表示されない
        self.pack(expand=1, fill=tk.BOTH)
        #メニューの初期化
        self.init_menu()
        #リストの初期化
        self.init_list_box()

        #エクエルデータを保持するリスト。
        self.books = []

    def load_excel(self):
        return

    def export_json(self):
        return

    def init_list_box(self):
        return

    def init_menu(self):
        #メニューを生成
        self.mbar = tk.Menu()
        #メニューコマンドを生成
        self.mcom = tk.Menu(self.mbar,tearoff=0)
        #コマンドを追加
        self.mcom.add_command(label='Excelファイル読み込み',command=self.load_excel)
        self.mcom.add_command(label='jsonファイル書き出し',command=self.export_json)
        self.mbar.add_cascade(label='ファイル',menu=self.mcom)
        self.master['menu'] = self.mbar


#インスタンスを作成。
root = tk.Tk()
app = Application(master=root)
#ループ。
app.mainloop()



