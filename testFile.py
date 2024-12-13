import openpyxl
from openpyxl import Workbook
import os
import tkinter as tk
from tkinter import simpledialog, messagebox, Frame, Label, ttk
from tkcalendar import DateEntry
from enum import Enum, auto

class QuestionType(Enum):
    Date = auto(), # calendar / date picker
    ValuePicker = auto(), # Spin Box Widget
    TextBox = auto(), # text widget
    Numeric = auto(), # only Numeric allowed
    ComboBox = auto() # ComboBox / Preselect Option picker
    
    
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
        sheet.append(["Date", "Recipe Link", "Dish Name", "Cooking Method", "Notes", "Enjoyment Rating", "How well did it hold / freeze", ])
        wb.save(file_path)
        print(f"Workbook '{file_path}' created.")
    
    return wb

def submit_answers(workbook, entries, file_path):
    sheet = workbook.active
    recipeData = []
    for key,value in entries.items():
        recipeData.append(value.get())
        value.delete(0, tk.END)
        
    sheet.append(recipeData)
    workbook.save(file_path)
    print(recipeData)
    
    
def validate_numeric_input(P):
    if P == "" or P.isdigit():
        return True
    else:
        return False
    
def question_prompt(workbook, file_path):
    
    questions = [["Date", QuestionType.Date], ["Recipe Link?", QuestionType.TextBox], ["Dish Name", QuestionType.TextBox], ["Cooking Method", QuestionType.ComboBox], ["Any Notes?", QuestionType.TextBox], ["Rate your Enjoyment 1-10", QuestionType.ValuePicker]]
    entries = {}

    for i, question in enumerate(questions):
        match question[1]:
            case QuestionType.Date:
                ttk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                date_entry = DateEntry(questionContainer, width=30, date_pattern='yyyy-mm-dd')
                date_entry.grid(row=i, column=1, pady=5)
                entries[question[0]] = date_entry  # Store the date entry widget
                
            case QuestionType.TextBox:
                ttk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                entry = ttk.Entry(questionContainer, width=30)
                entry.grid(row=i, column=1, pady=5)
                entries[question[0]] = entry
                
            case QuestionType.Numeric:
                ttk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                vcmd = (root.register(validate_numeric_input), '%P')
                entry = tk.Entry(questionContainer, validate="key", validatecommand=vcmd)
                entry.grid(row=i, column=1, pady=5)
                entries[question[0]] = entry
                
            case QuestionType.ComboBox:
                ttk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                entry = ttk.Combobox(questionContainer, values=["Sheet Pan", "Combo", "One Pot"])
                entry.grid(row=i, column=1, pady=5)
                entries[question[0]] = entry
                
            case QuestionType.ValuePicker:
                print("trying to do spinbox")
                ttk.Label(questionContainer, text=question[0], anchor="w").grid(row=i, column=0, sticky="w", pady=5, padx=15)
                entry = ttk.Spinbox(questionContainer, from_=0, to=10)
                entry.grid(row=i, column=1, pady=5)
                entries[question[0]] = entry
                
        
    submit_btn = ttk.Button(questionContainer, text="Submit", command=lambda: submit_answers(workbook, entries, file_path))
    submit_btn.grid(row=len(questions), column=0, columnspan=2, pady=10)


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filename = "food_log.xlsx"  # Excel file name
    file_path = os.path.join(script_dir, filename)
    wb = create_excel_file(file_path)

    ## GUI SET UP
    root = tk.Tk()
    root.geometry("450x400") 
    root.title("Recipe Tracker")
    
    # Apply the clam theme
    style = ttk.Style(root)
    style.theme_use("xpnative")
    
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    # Add a frame that scales
    mainframe = Frame(root)
    mainframe.grid(row=0, column=0, sticky="nsew") 
    
    # Create frame within center_frame
    questionContainer = Frame(mainframe, bg="grey")
    questionContainer.pack(fill="both", expand=True, padx=20, pady=20)
    
    question_prompt(wb, file_path)
    
    root.mainloop()