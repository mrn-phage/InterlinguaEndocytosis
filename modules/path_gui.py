import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox


def gui():
    """
    guiウィンドウを表示し、選択したpathをdict形式で返す。
    path_word_dictの内、['content']のvalueはcancelボタンを押したか判断する。
    Trueだと中身在り
    Falseだとキャンセルを押しているため、中身無し
    """

    path_word_dict = dict()
    # --------------------------------------------------------------------------
    # ボタンクリック時の処理関数

    class FuncMaker():
        """
        ボタンクリック時の関数クラス
        関数は内部にinner関数を作っていない場合、tkinterのmainloopに入った瞬間
        実行されるので注意
        """

        def __init__(self):
            self.path = None

        # Excelファイルを選択する
        # コードを置いたディレクトリが初期位置となる
        def file_select(self):
            def inner():
                fTyp = [("Excelファイル", "*.xlsx"),
                        ("マクロ付きExcelファイル", "*.xlsm"),
                        ("csvファイル", "*.csv")]
                iDir = os.path.dirname(__file__)
                fpath = filedialog.askopenfilename(filetypes=fTyp,
                                                   initialdir=iDir)
                self.path.set(fpath)
            return inner

        # フォルダを選択する
        # デスクトップが初期位置となる

        def dire_select(self):
            def inner():
                desktop_path =\
                    os.getenv("HOMEDRIVE") + \
                    os.getenv("HOMEPATH") + "\\Desktop"
                fpath = filedialog.askdirectory(initialdir=desktop_path)
                self.path.set(fpath)
            return inner

    def apply():

        messagebox.showinfo(
            'Result', '構成情報↓↓\n' + const.path.get() + '\n'
            'に追加情報↓↓\n' + add.path.get() + '\n'
            'を追加しました'
        )

        root.destroy()

    def cancel():

        const.path.set('')
        add.path.set('')
        save.path.set('')
        # key.word.set(None)
        # value.word.set(None)

        root.destroy()

    # --------------------------------------------------------------------------
    # 入力されたワードクラス

    class WordStorage():
        def __init__(self):
            self.word = None

    # --------------------------------------------------------------------------
    # ボタン処理関数インスタンスとワードインスタンス生成

    const = FuncMaker()
    add = FuncMaker()
    save = FuncMaker()
    # key = WordStorage()
    # value = WordStorage()

    # --------------------------------------------------------------------------
    # rootの作成

    root = tk.Tk()
    root.title('unification Tool')
    root.resizable(False, False)

    # --------------------------------------------------------------------------
    # Frame1の作成
    frame1 = ttk.Frame(root, padding=10)
    frame1.grid()

    # ラベルの作成
    s1 = tk.StringVar()
    s1.set('構成情報>>')
    label1 = ttk.Label(frame1, textvariable=s1)
    label1.grid(row=0, column=0)

    # 入力ボックスの作成
    const.path = tk.StringVar()
    entry1 = ttk.Entry(frame1, textvariable=const.path, width=50)
    entry1.grid(row=0, column=1)

    # ボタンの作成
    button1 =\
        ttk.Button(frame1, text=u'選択', command=const.file_select())
    button1.grid(row=0, column=2)

    # ラベルの作成
    s2 = tk.StringVar()
    s2.set('追加情報>>')
    label2 = ttk.Label(frame1, textvariable=s2)
    label2.grid(row=1, column=0)

    # 入力ボックスの作成
    add.path = tk.StringVar()
    entry2 = ttk.Entry(frame1, textvariable=add.path, width=50)
    entry2.grid(row=1, column=1)

    # ボタンの作成
    button2 =\
        ttk.Button(frame1, text=u'選択', command=add.file_select())
    button2.grid(row=1, column=2)

    # ラベルの作成
    s3 = tk.StringVar()
    s3.set('出力先>>')
    label3 = ttk.Label(frame1, textvariable=s3)
    label3.grid(row=2, column=0)

    # 入力ボックスの作成
    save.path = tk.StringVar()
    entry3 = ttk.Entry(frame1, textvariable=save.path, width=50)
    # 初期値としてデスクトップをセット
    entry3.insert(tk.END,
                  os.getenv("HOMEDRIVE") + os.getenv("HOMEPATH") + "\\Desktop")
    entry3.grid(row=2, column=1)

    # ボタンの作成
    button3 =\
        ttk.Button(frame1, text=u'選択', command=save.dire_select())
    button3.grid(row=2, column=2)

    # --------------------------------------------------------------------------
    # Frame2の作成
    # frame2 = ttk.Frame(root)
    # frame2.grid()
    #
    # # ラベルの作成
    # s11 = tk.StringVar()
    # s11.set('共通する見出しの単語>>')
    # label11 = ttk.Label(frame2, textvariable=s11)
    # label11.grid(row=0, column=0)
    #
    # # 入力ボックスの作成
    # key.word = tk.StringVar()
    # entry11 = ttk.Entry(frame2, textvariable=key.word, width=29)
    # entry11.grid(row=0, column=1)
    #
    # # ラベルの作成
    # s22 = tk.StringVar()
    # s22.set('参照する値の単語>>')
    # label22 = ttk.Label(frame2, textvariable=s22)
    # label22.grid(row=1, column=0)
    #
    # # 入力ボックスの作成
    # value.word = tk.StringVar()
    # entry22 = ttk.Entry(frame2, textvariable=value.word, width=29)
    # entry22.grid(row=1, column=1)

    # --------------------------------------------------------------------------
    # Frame3の作成
    frame3 = ttk.Frame(root, padding=(0, 5))
    frame3.grid(row=2)

    # 実行ボタンの作成
    go_button = ttk.Button(frame3, text='実行', command=apply)
    go_button.pack(side=tk.LEFT)

    # キャンセルボタンの作成
    no_go_button = ttk.Button(frame3, text='キャンセル', command=cancel)
    no_go_button.pack(side=tk.LEFT)

    # --------------------------------------------------------------------------
    # 閉じるボタン等の無効化
    root.overrideredirect(True)

    # --------------------------------------------------------------------------
    # loop開始

    root.mainloop()

    path_word_dict['const'] = const.path.get()
    path_word_dict['add'] = add.path.get()
    path_word_dict['save'] = save.path.get() + r"\new.xlsx"
    # path_word_dict["key_word"] = key.word.get()
    # path_word_dict["val_word"] = value.word.get()
    # path_word_dictのkeyの数でcancel関数に入ったか確認。
    if path_word_dict['const'] == ''\
            or path_word_dict['add'] == ''\
            or path_word_dict['save'] == r"\new.xlsx":
        path_word_dict['content'] = False
    else:
        path_word_dict['content'] = True

    if __name__ == '__main__':
        print([(i, j) for i, j in path_word_dict.items()])

    return path_word_dict


if __name__ == '__main__':
    gui()
