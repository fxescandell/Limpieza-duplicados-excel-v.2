# main.py
import tkinter as tk
from gui import DuplicatesRemoverApp

def main():
    root = tk.Tk()
    app = DuplicatesRemoverApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
