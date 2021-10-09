import tkinter as tk


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