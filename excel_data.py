import openpyxl
            

#エクセルのシート毎のデータを保持
class Sheet():
    def __init__(self):
        #リストを初期化。
        self.rows = {}
        self.sheet_name = None


    #行ごとのセルを読み込む。
    def load_cell(self,row):
        cells = []
        #セルを1つずつ読み込む。
        i = 0
        for cell in row:
            i += 1
            #3つ以上は読み込まない。
            if i >=3:
                break
            cells.append(cell.value)
        return cells[0],cells[1]

    #エクセルのシートを読み込む。
    def load(self,sheet):
        self.sheet_name = sheet.title
        i = 0
        for line in sheet:
            #最初の1行目は読み込まない
            if i == 0:
                i += 1
                continue
            #1行ずつ読み込んで、リストに追加する。
            key,value = self.load_cell(line)
            if key == None or value == None:
                continue
            self.rows[key] = value

    #行番号と列番号を指定して、セルを取得。
    '''def get_cell(self,row_number,col_number):
        #リストは0から始まるが、エクセルは1から始まるので-1する。
        row_number -= 1
        col_number -= 1

        #指定された行番号が読み込んだ行を超えていたら。
        if row_number >= len(self.rows):
            return None

        row = self.rows[row_number]
        #指定された列番号が読み込んだ列を超えていたら。
        if col_number >= len(row.cells):
            return None
        return row.cells[col_number]

    #行番号を指定して、行要素を取得。
    def get_cells(self,row_number):
        row_number -= 1
        #指定された行番号が読み込んだ行を超えていたら。
        if row_number >= len(self.rows):
            return None

        return self.rows[row_number].cells'''

    #行要素の数を取得
    def get_row_number(self):
        return len(self.rows)

class ExcelData():
    def __init__(self):
        self.sheets = []
        self.sheet_number = None
        self.is_init = False
        self.id = 0
        self.file_name = None

    def load(self,file_path,id):
        try:
            #Excelのファイルを読み込む。
            book = openpyxl.load_workbook(file_path)
        except:
            #ファイルオープンに失敗したらFalseを返す。
            return False

        #シート枚数を取得。
        self.sheet_number = len(book.worksheets)
        for number in range(self.sheet_number):
            #number番目のシートを取得。
            sheet = book.worksheets[number]
            sh = Sheet()
            #シートを読み込む。
            sh.load(sheet)
            #リストに追加。
            self.sheets.append(sh)


        #ファイルの名前を抽出していく、/と.を除いていく
        slash_number = file_path.rfind('/')
        if slash_number == -1:
            self.file_name = file_path
        else:
            self.file_name = file_path[slash_number + 1:]
        self.id = id
        self.is_init = True
        return True

    #シートのデータを取得
    def get_sheet(self,sheet_number):
        if sheet_number >= self.sheet_number or self.is_init == False:
            return None
        return self.sheets[sheet_number]

    #シート一覧を取得
    def get_sheets(self):
        return self.sheets

