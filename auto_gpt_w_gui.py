import tkinter as tk
from tkinter import *
from tkinter import filedialog
import os
import openai
import pandas as pd
from tkinter import messagebox
import numpy as np
import traceback


def browse_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        print("Selected file:", file_path)
        excel_file_var.set(file_path)
        selected_file_label.config(text="Selected file: " + os.path.basename(file_path))

window = tk.Tk()
window.title('GPT automaton')
window.minsize(height=500, width=500)

# Listbox
listbox = tk.Listbox(window, width=80, height=25)
listbox.grid(row=2, columnspan=3)

# Scrollbar
scrollbar = Scrollbar(window)
scrollbar.grid(row=2, column=3, sticky="ns")
listbox.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=listbox.yview)

# Selected file label
selected_file_label = Label(window, text="Choose question file")
selected_file_label.grid(row=3, column=1)

# Browse button
browse_button = Button(window, text= "Browse", command=browse_file)
browse_button.grid(row=3, column=2)

Button(window, text="Quit", command=window.quit).grid(row=11, column=2)

api_key_path = 'api_key.txt'
with open(api_key_path, "r") as file:
    api_key = file.read().strip()

openai.api_key = api_key

excel_file_var = StringVar()
question_col = 'Questions'
answer_col = 'Answers'


def process_file():
    try:
        df = pd.read_excel(excel_file_var.get(), sheet_name='Sheet1')
        col_index = 2
        row_index = 0
        is_empty = pd.isnull(df.at[row_index, answer_col])
        num_rows = len(df.index)
        listbox.insert(END, f'{num_rows} questions found.')
        listbox.see(END)  
        window.update()  

        for idx, row in df.iterrows():
            question = row[question_col]
            listbox.insert(END, f'Processing {idx+1}/{num_rows} questions.')              
            window.update()
            processing_message = f'Processing question: {question}'
            listbox.insert(END, processing_message)
            listbox.see(END)  # Scroll to the latest message
            window.update()  # Update the GUI to display the new message

            # Make an API call to generate the answer
            response = openai.Completion.create(
                engine='text-davinci-003',
                prompt=question,
                max_tokens=100
            )

            # Extract the generated answer from the API response
            generated_answer = response.choices[0].text.strip()

            # Convert the answer to string
            generated_answer = str(generated_answer)

            listbox.insert(END, 'Done')            
            listbox.see(END)
            window.update()

            if is_empty:
                df.at[idx, df.columns[col_index]] = generated_answer
            else:
                while not is_empty:
                    col_index += 1
                    try:
                        is_empty = pd.isnull(df.iloc[row_index, col_index])
                    except IndexError:
                        new_col_name = 'Answers' + str(col_index)
                        df[new_col_name] = np.nan
                        is_empty = True
                df.at[idx, df.columns[col_index]] = generated_answer

            # update the last written column index        
            df.to_excel(excel_file_var.get(), index=False)
        
        
        listbox.insert(END, f'All questions are written successfully to the {os.path.basename(file_path)}')
        listbox.see(END)
        window.update()

    except Exception as e:
        error_message = f"Error: {str(e)}"
        traceback_message = traceback.format_exc()
        listbox.insert(END, error_message)
        listbox.insert(END, traceback_message)
        listbox.see(END)
        
process_button = Button(window, text="Process File", command=process_file)
process_button.grid(row=4, column=1, columnspan=4)

window.mainloop()
