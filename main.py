import tkinter as tk
from tkinter import messagebox
from docrefine.gui.app import App
from docrefine.config import log_app

if __name__ == "__main__":
    try:
        log_app("Booting DocRefine Pro (Modular)...")
        root = tk.Tk()
        App(root)
        root.mainloop()
    except Exception as e:
        msg = f"Fatal Boot Error: {e}"
        print(msg)
        messagebox.showerror("Fatal Error", msg)