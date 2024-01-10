

import tkinter as tk
from tkinter import filedialog
import pywhatkit as pwk
import time
import pandas as pd
import numpy as np
import keyboard as k
import pyautogui


DEFAULT_MESSAGE = 'Hi, we represent a distinguished modest fashion company based in Istanbul, Turkey. We firmly believe that our product line is ideally suited for your customer portfolio. Furthermore, we are delighted to extend an exceptional discount opportunity for your first order. Please kindly access the following link to our wholesale website:\n\nhttps://ws.toucheprive.com'



# Function to send messages to multiple numbers
def send_message():
    excel_paths = excel_entry.get().split("|")
    use_default_message = default_checkbox_var.get()
    default_message = default_message_entry.get("1.0", tk.END).strip() or DEFAULT_MESSAGE
    custom_message = custom_message_entry.get("1.0", tk.END).strip()


    for excel_path in excel_paths:
        excel = pd.read_excel(excel_path)
        excel = excel[excel['International Phone'].notna()]
        numbers = np.array(excel["International Phone"].str.replace(' ', ''))
        formatted_numbers = ["+" + number[2:] for number in numbers]

       

        for number in formatted_numbers:
            if use_default_message:
                message = default_message
            else:
                message = custom_message

            pwk.sendwhatmsg_instantly(number, message, 20)  

            time.sleep(2)
            pyautogui.click(1737, 973)
            time.sleep(2)
            pyautogui.click(1845, 972)
            k.press_and_release('enter')
            time.sleep(1)
            k.press_and_release('ctrl+w')
            time.sleep(1)


            
            # print("1")
            # pyautogui.click(1050, 950)
            # time.sleep(10)
            # k.press_and_release('enter')
            # print('2')

    # Clear the input fields
    excel_entry.delete(0, tk.END)
    default_message_entry.delete("1.0", tk.END)
    custom_message_entry.delete("1.0", tk.END)
    default_checkbox_var.set(1)

   

# Function to browse for Excel files
def browse_excel_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx")])
    excel_entry.delete(0, tk.END)
    excel_entry.insert(tk.END, "|".join(file_paths))


# Create the main window
window = tk.Tk()
window.title("WhatsApp Message Sender")

# Excel path label and entry
excel_label = tk.Label(window, text="Excel Files:")
excel_label.pack()
excel_entry = tk.Entry(window)
excel_entry.pack()

# Browse button to open file dialog
browse_button = tk.Button(window, text="Browse", command=browse_excel_files)
browse_button.pack()


# Default message label and entry
default_message_label = tk.Label(window, text="Default Message:")
default_message_entry = tk.Text(window, height=5, width=50)
default_message_entry.pack()
default_message_entry.insert(tk.END, DEFAULT_MESSAGE)

# Checkbox for default or custom message
default_checkbox_var = tk.IntVar()
default_checkbox = tk.Checkbutton(window, text="Use Default Message", variable=default_checkbox_var)
default_checkbox.pack()



# Custom message label and entry
custom_message_label = tk.Label(window, text="Custom Message:")
custom_message_label.pack()
custom_message_entry = tk.Text(window, height=3, width=30)
custom_message_entry.pack()


# Send button
send_button = tk.Button(window, text="Send", command=send_message)
send_button.pack()

# Run the main event loop
window.mainloop()






