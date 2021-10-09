#ウィンドウを表示するためのライブラリ
import tkinter as tk

#Excelファイルを開くためのライブラリ
#xlsxファイルには対応しなくなった
import xlrd

import openpyxl


#エクセルの要素を読み込む
def lead_sheet(sheet):
    lines = []
    for line in sheet:
        #1行ずつ読み込んで、リストに追加する。
        lines.append(lead_col(line))
    return lines

#1行ごとの列を読み込む。
def lead_col(line):
    cells = []
    #1列ずつ読み込む。
    for cell in line:
        cells.append(cell.value)
    return cells

#Excelのファイルを読み込む。
book = openpyxl.load_workbook("file/data.xlsx")
#シート枚数を取得。
sheet_number = len(book.worksheets)
for number in range(sheet_number):
    #number番目のシートを取得。
    sheet = book.worksheets[number]
    lines = []
    #1行ずつ読み込む。
    for row in sheet.rows:
        line = []
        #1列ずつ読み込む。
        for cell in row:
            line.append(cell.value)
        #1行ずつリストに追加する。
        #lines.append(line)
    lines = lead_sheet(sheet)
    print(sheet)
    print(lines)





class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        #なんか作って
        self.master = master
        #名前を設定
        self.master.title('Convert To Json Tool')
        #とりあえずウィンドウはこれくらいで
        self.master.geometry("600x400")
        #これをしないとフレームがどうのこうので
        #placeしても表示されない
        self.pack(expand=1, fill=tk.BOTH)


root = tk.Tk()
app = Application(master=root)
app.mainloop()



