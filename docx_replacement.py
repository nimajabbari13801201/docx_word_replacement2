from docx import Document
import tkinter as tk
import ttkbootstrap as ttk
from tkinter import filedialog
def set_text_preserving_text_formatting(paragraph, text):
    runs = paragraph.runs
    if not runs:
        return
    runs[0].text = text
    for run in runs[1:]:
        r = run._element
        r.getparent().remove(r)
file_path=''
def browse():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    file_path = file_path.replace('/', '\\\\')
    filepath.set(file_path)
def replace():
    old=old_word.get()
    new=new_word.get()
    doc = Document(filepath.get())
    for p in doc.paragraphs:
        if old in p.text:
            set_text_preserving_text_formatting(p, p.text.replace(old, new))
    doc.save(filepath.get())
#create window
window=ttk.Window(themename='darkly')
window.title('word replacement')
window.geometry('400x300')
#top label
top_label=ttk.Label(master=window,text='Browse the file you want '+"\n"+' type the old word and the new word '+"\n"+' then click replace button',font='Calibri 12 bold')
top_label.pack(side='top')
#file path
filepath_form=ttk.Frame(master=window)
filepath_form.pack()
filepath=tk.StringVar()
path_entry=ttk.Entry(master=filepath_form,textvariable=filepath)
path_entry.pack(side='left')
browse_button=ttk.Button(master=filepath_form,text='Browse',command=browse)
browse_button.pack(side='left')
#old word frame
oldword_form=ttk.Frame(master=window)
oldword_form.pack()
oldword_label=ttk.Label(master=oldword_form,text='old_word:',font='Calibri 10')
oldword_label.pack(side='left',padx=10)
old_word=tk.StringVar()
oldword_entry=ttk.Entry(master=oldword_form,textvariable=old_word)
oldword_entry.pack(side='left',pady=10)
#new word frame
newword_form=ttk.Frame(master=window)
newword_form.pack()
newword_label=ttk.Label(master=newword_form,text='new_word:',font='Calibri 10')
newword_label.pack(side='left',padx=10)
new_word=tk.StringVar()
newword_entry=ttk.Entry(master=newword_form,textvariable=new_word)
newword_entry.pack(side='left',pady=10)
#replace button
replace_button=ttk.Button(master=window,text='Replace',command=replace)
replace_button.pack(pady=10)
#run
window.mainloop()
