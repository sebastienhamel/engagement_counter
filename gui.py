import tkinter as tk
from tkinter import filedialog, messagebox
import json
import os
from docx import Document

CACHE_FILE = "cache.json"

# Load cached text
def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("text", "")
    return ""

# Save text to cache
def save_cache(text):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump({"text": text}, f)

# Submit handler
def on_submit():
    text = text_input.get("1.0", tk.END).strip()
    save_cache(text)
    output.config(state="normal")
    output.delete("1.0", tk.END)
    output.insert(tk.END, text)
    output.config(state="disabled")
    messagebox.showinfo("Saved", "Text saved to cache.")

# Open .docx file
def open_docx(file_path=None):
    if not file_path:
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        doc = Document(file_path)
        content = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        text_input.delete("1.0", tk.END)
        text_input.insert(tk.END, content)

# Drag-and-drop support (Windows only)
def drop(event):
    file_path = root.tk.splitlist(event.data)[0]
    if file_path.lower().endswith(".docx"):
        open_docx(file_path)

# GUI
root = tk.Tk()
root.title("Free Text + DOCX Editor")

text_input = tk.Text(root, height=10, width=60)
text_input.insert(tk.END, load_cache())
text_input.pack(padx=10, pady=10)

btn_frame = tk.Frame(root)
tk.Button(btn_frame, text="Submit", command=on_submit).pack(side="left", padx=5)
tk.Button(btn_frame, text="Open DOCX", command=open_docx).pack(side="left", padx=5)
btn_frame.pack()

tk.Label(root, text="Submitted Text:").pack()
output = tk.Text(root, height=5, width=60, state="disabled")
output.pack(padx=10, pady=(0, 10))

# Drag-and-drop setup
try:
    import tkinterdnd2 as tkdnd  # Install if needed
    from tkinterdnd2 import DND_FILES

    root.destroy()
    root = tkdnd.TkinterDnD.Tk()
    root.title("Free Text + DOCX Editor (DnD)")

    text_input = tk.Text(root, height=10, width=60)
    text_input.insert(tk.END, load_cache())
    text_input.pack(padx=10, pady=10)
    text_input.drop_target_register(DND_FILES)
    text_input.dnd_bind("<<Drop>>", drop)

    btn_frame = tk.Frame(root)
    tk.Button(btn_frame, text="Submit", command=on_submit).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Open DOCX", command=open_docx).pack(side="left", padx=5)
    btn_frame.pack()

    tk.Label(root, text="Submitted Text:").pack()
    output = tk.Text(root, height=5, width=60, state="disabled")
    output.pack(padx=10, pady=(0, 10))

except ImportError:
    print("tkinterdnd2 not installed — drag-and-drop disabled.")
    print("Install it with: pip install tkinterdnd2")

root.mainloop()
