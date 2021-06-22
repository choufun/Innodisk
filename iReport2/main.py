import tkinter as tk
from tkinter import filedialog

import os

from reader import *

class App:
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master).pack()
        self.button1 = tk.Button(self.frame, text="Select", command=self.browse, width=10, height=2).pack(side=tk.LEFT, padx=10)
        self.filename = tk.StringVar()
        self.filename.set("File")
        self.label = tk.Label(self.frame, textvariable=self.filename).pack(side=tk.LEFT, fill=tk.X, expand=tk.YES)
        self.button2 = tk.Button(self.frame, text="Compile", command=self.compile, width=10, height=2).pack(side=tk.LEFT, padx=10)

    def browse(self):
        filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select An Image").split('/')[-1]
        self.filename.set(filename)

    def compile(self):
        reader = Reader(self.filename.get())
        pass


def main():
    root = tk.Tk()
    root.title("Innodisk: iReport")
    display = App(root)
    root.geometry("600x50")
    root.mainloop()


if __name__ == "__main__":
    main()
