import openpyxl
from openpyxl import Workbook
import os
import tkinter as tk
from tkinter import simpledialog, messagebox, Frame, Label, ttk
import ttkbootstrap as tb
from tkcalendar import DateEntry
from enum import Enum, auto

class QuestionType(Enum):
    Date = auto(), # calendar / date picker
    ValuePicker = auto(), # Spin Box Widget
    TextBox = auto(), # text widget
    Numeric = auto(), # only Numeric allowed
    ComboBox = auto() # ComboBox / Preselect Option picker
    KeepingValues = auto() # value picker based on how well it holds in fridge / freezer
    


# Function to create a new Excel file with headers if it doesn't exist
def create_excel_file(file_path):
    try:
        # Try opening an existing file
        wb = openpyxl.load_workbook(file_path)
        print(f"Workbook '{file_path}' loaded.")
    except FileNotFoundError:
        # If file does not exist, create a new one
        wb = Workbook()
        sheet = wb.active
        # Define headers for the food tracking sheet
        sheet.append(["Date", "Recipe Link", "Dish Name", "Cooking Method", "Notes", "Enjoyment Rating", "Ease of Preparation", "How well did it hold / freeze", ])
        wb.save(file_path)
        print(f"Workbook '{file_path}' created.")
    
    return wb

def submit_answers(workbook, entries, file_path):
    sheet = workbook.active
    recipeData = []
    for key,value in entries.items():
        if isinstance(value,tk.Listbox):
            selected_indices = value.curselection()  # Get indices of selected items
            selected_items = [value.get(i) for i in selected_indices]  # Get items based on selected indices
            my_string = ",".join(selected_items)
            recipeData.append(my_string)
            value.select_clear(0, tk.END)
        else:
            recipeData.append(value.get())
            value.delete(0, tk.END)
        
    sheet.append(recipeData)
    workbook.save(file_path)
    
    
def validate_numeric_input(P):
    if P == "" or P.isdigit():
        return True
    else:
        return False
    
def question_prompt(workbook, file_path):
    questions = [["Date", QuestionType.Date], ["Recipe Link?", QuestionType.TextBox], ["Dish Name", QuestionType.TextBox], ["Cooking Method", QuestionType.ComboBox], ["Any Notes?", QuestionType.TextBox], ["Rate your Enjoyment 1-10", QuestionType.ValuePicker], ["Ease of Preperation", QuestionType.ValuePicker], ["Fridge/Freezer compatible?", QuestionType.KeepingValues]]
    entries = {}
    # Text color)   
    for i, question in enumerate(questions):
        match question[1]:
            case QuestionType.Date:
                tk.Label(questionContainer, text=question[0], anchor="w", bg="#606060", fg="white").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                date_entry = tb.DateEntry(questionContainer, width=30, bootstyle="success")
                date_entry.grid(row=i, column=1, pady=5, sticky="e")
                entries[question[0]] = date_entry.entry  # Store the date entry widget
                
            case QuestionType.TextBox:
                tk.Label(questionContainer, text=question[0], anchor="w", bg="#606060", fg="white").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                entry = tk.Entry(questionContainer, width=30, bg="#3F3F3F", fg="white")
                entry.grid(row=i, column=1, pady=5, sticky="e")
                entries[question[0]] = entry
                
            case QuestionType.Numeric:
                tk.Label(questionContainer, text=question[0], anchor="w", bg="#606060", fg="white").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                vcmd = (root.register(validate_numeric_input), '%P')
                entry = tk.Entry(questionContainer, validate="key", validatecommand=vcmd, bg="#3F3F3F", fg="white", width=30,)
                entry.grid(row=i, column=1, pady=5, sticky="e")
                entries[question[0]] = entry
                
            case QuestionType.ComboBox:
                tk.Label(questionContainer, text=question[0], anchor="w", bg="#606060", fg="white").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                entry = ttk.Combobox(questionContainer, values=["Sheet Pan", "Combo", "One Pot"], width=30)
                entry.grid(row=i, column=1, pady=5, sticky="e")
                entries[question[0]] = entry
                
            case QuestionType.ValuePicker:
                tk.Label(questionContainer, text=question[0], anchor="w", bg="#606060", fg="white").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                entry = ttk.Spinbox(questionContainer, from_=0, to=10, width=30)
                entry.grid(row=i, column=1, pady=5, sticky="e")
                #entry.configure(background="black")
                entries[question[0]] = entry
            
            case QuestionType.KeepingValues:
                tk.Label(questionContainer, text=question[0], anchor="w", bg="#606060", fg="white").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                listbox = tk.Listbox(questionContainer, width=30, selectmode=tk.MULTIPLE)
                listbox.grid(row=i, column=1, pady=5, sticky="e")
                items = ["Item 1", "Item 2", "Item 3", "Item 4"]
                for item in items:
                    listbox.insert(tk.END, item)
                entries[question[0]] = listbox
                
        
    submit_btn = ttk.Button(questionContainer, text="Submit", command=lambda: submit_answers(workbook, entries, file_path), bootstyle="success")
    submit_btn.grid(row=len(questions), column=0, columnspan=2, pady=10)


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filename = "food_log.xlsx"  # Excel file name
    file_path = os.path.join(script_dir, filename)
    wb = create_excel_file(file_path)

    ## GUI SET UP
    root = tb.Window(themename="superhero")
    root.geometry("1000x600") 
    root.title("Recipe Tracker")
        
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    # Add a frame that scales
    mainframe = tk.Frame(root, bg="#3D3D3D")
    mainframe.grid(row=0, column=0, sticky="nsew")
    mainframe.grid_rowconfigure(0, weight=1)
    mainframe.grid_columnconfigure(0, weight=1)
    
    # Create frame within center_frame
    questionContainer = tk.Frame(mainframe, bg="#606060")
    questionContainer.grid(row=0, column=0, padx=50, pady=50)
    questionContainer.grid_rowconfigure(0, weight=0)
    questionContainer.grid_columnconfigure(0, weight=0)
    
    question_prompt(wb, file_path)
    
    root.mainloop()