import openpyxl
from openpyxl import Workbook
import os
import tkinter as tk
from tkinter import simpledialog, messagebox, Frame, Label, ttk
import ttkbootstrap as tb
from tkcalendar import DateEntry
from enum import Enum, auto

#Enums to help define question Types and filter based on them. 
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
        sheet.append(["Date", "Recipe Link", "Dish Name", "Cooking Method", "Notes", "Enjoyment Rating", "Ease of Preparation", "Meal Prep Compatability", ])
        wb.save(file_path)
        print(f"Workbook '{file_path}' created.")
    
    return wb

#Submit Answers button parses the entries, clears the input fields and saves the answers to your workbook 
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
    
# Function ensures that you can only enter numeric input into the text field    
def validate_numeric_input(P):
    if P == "" or P.isdigit():
        return True
    else:
        return False
    
def question_prompt(workbook, file_path):
    questions = [["Date", QuestionType.Date], ["Recipe Link:", QuestionType.TextBox], ["Dish Name:", QuestionType.TextBox], ["Cooking Method:", QuestionType.ComboBox], ["Notes:", QuestionType.TextBox], ["Enjoyment", QuestionType.ValuePicker], ["Ease of Preperation", QuestionType.ValuePicker], ["Fridge/Freezer compatible?", QuestionType.KeepingValues]]
    entries = {}  
    for i, question in enumerate(questions):
        match question[1]:
            case QuestionType.Date: #Date Entry Widget
                tk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                date_entry = tb.DateEntry(questionContainer, width=30, bootstyle="success")
                date_entry.grid(row=i, column=1, pady=5, sticky="e")
                entries[question[0]] = date_entry.entry 
                
            case QuestionType.TextBox: #Text Box Widget
                tk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                entry = tk.Entry(questionContainer, width=30)
                entry.grid(row=i, column=1, pady=5, sticky="e")
                entries[question[0]] = entry
                
            case QuestionType.Numeric: #Text Box Widget with Numeric filter
                tk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                vcmd = (root.register(validate_numeric_input), '%P')
                entry = tk.Entry(questionContainer, validate="key", validatecommand=vcmd, width=30,)
                entry.grid(row=i, column=1, pady=5, sticky="e")
                entries[question[0]] = entry
                
            case QuestionType.ComboBox: #Combo Box
                tk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                entry = ttk.Combobox(questionContainer, values=["Sheet Pan", "Combo", "One Pot"], width=30)
                entry.grid(row=i, column=1, pady=5, sticky="e")
                entries[question[0]] = entry
                
            case QuestionType.ValuePicker: #1-10 Spin Box
                tk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                entry = ttk.Spinbox(questionContainer, from_=0, to=10, width=30)
                entry.grid(row=i, column=1, pady=5, sticky="e")
                entries[question[0]] = entry
            
            case QuestionType.KeepingValues: #Multi-select value picker (listbox)
                tk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                listbox = tk.Listbox(questionContainer, width=30, selectmode=tk.MULTIPLE)
                listbox.grid(row=i, column=1, pady=5, sticky="e")
                items = ["Freezes well", "Freezes Poorly", "Holds well in Fridge", "Does not hold well in Fridge"]
                for item in items:
                    listbox.insert(tk.END, item)
                entries[question[0]] = listbox
                
    #Submit Button    
    submit_btn = ttk.Button(questionContainer, text="Submit", command=lambda: submit_answers(workbook, entries, file_path), bootstyle="success")
    submit_btn.grid(row=len(questions), column=0, columnspan=2, pady=10)


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filename = "food_log.xlsx"  # Excel file name
    file_path = os.path.join(script_dir, filename)
    wb = create_excel_file(file_path)

    ### GUI SET UP ###
    root = tb.Window(themename="superhero")
    root.geometry("1000x600") 
    root.title("Recipe Tracker")
        
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    # Add a frame that scales
    mainframe = tk.Frame(root)
    mainframe.grid(row=0, column=0, sticky="nsew")
    mainframe.grid_rowconfigure(0, weight=1)
    mainframe.grid_columnconfigure(0, weight=1)
    
    # Create frame within center_frame
    questionContainer = tk.Frame(mainframe)
    questionContainer.grid(row=0, column=0, padx=50, pady=50)
    questionContainer.grid_rowconfigure(0, weight=0)
    questionContainer.grid_columnconfigure(0, weight=0)
    
    question_prompt(wb, file_path)
    
    root.mainloop()