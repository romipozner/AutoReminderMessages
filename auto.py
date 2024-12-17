import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import json
import os
from urllib.parse import quote 
 

driver_path = './chromedriver'

# Define the configuration file path and required file name
CONFIG_FILE = "file_path.json"
REQUIRED_FILE_NAME = "tasks_template.xlsx"

# Load the Excel file path from the config file if it exists.
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
        return config.get("file_path")
    return None

# Save the Excel file path to the config file.
def save_config(file_path):
    with open(CONFIG_FILE, 'w') as f:
        json.dump({"file_path": file_path}, f)

# Prompt the user to select the correct Excel file path and save it if valid.
def get_excel_path():
    while True:
        # Show a prompt for the user to upload the file
        messagebox.showinfo("Upload Required", f"Please upload the '{REQUIRED_FILE_NAME}' file to continue.")

        root = tk.Tk()
        root.withdraw()  # Hide the root window
        file_path = filedialog.askopenfilename(
            title=f"Select the Excel File ({REQUIRED_FILE_NAME})",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        root.destroy()

        if not file_path:
            messagebox.showwarning("File Not Selected", "No file was selected. Please upload the required Excel file.")
            continue

        # Check if the selected file matches the required name
        if os.path.basename(file_path) != REQUIRED_FILE_NAME:
            messagebox.showerror("Incorrect File", f"The file must be named '{REQUIRED_FILE_NAME}'. Please try again.")
        else:
            save_config(file_path)  # Save the correct file path
            return file_path

# Load data from the Excel file, ensuring the correct file is selected.
def load_data():
    file_path = load_config()
    if not file_path:
        file_path = get_excel_path()
        if not file_path:
            return None, None  # Return empty DataFrames if no valid file selected
    
    try:
        df_tasks = pd.read_excel(file_path, sheet_name='Table', engine='openpyxl')
        df_names = pd.read_excel(file_path, sheet_name='Contacts', engine='openpyxl')
        return df_tasks, df_names
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file: {e}")
        return None, None


#create data frame for specific worker with his tasks
def find_tasks(name):
    df_tasks, df_names = load_data()
    
    # Strip and remove spaces from the name
    name = name.strip().replace(" ", "")
    
    # Create a mask for the filtering conditions
    role1 = df_tasks['role1'].str.replace(" ", "").str.strip() == name
    role2 = df_tasks['role2'].str.replace(" ", "").str.strip() == name
    role3 = df_tasks['role3'].str.replace(" ", "").str.strip() == name
    role4 = df_tasks['role4'].str.replace(" ", "").str.strip() == name
    
    # Use the lists to filter and keep the original order
    tasks = df_tasks[role1 | role2 | role3 | role4]
    
    return tasks


#create the weekly massage for spcific worker with all his tasks
def create_weekly_message(my_tasks,name,worker):
    # Define a mapping for the days of the week to numerical values
    day_order = {
        'Sunday': 1,
        'Monday': 2,
        'Tuesday': 3,
        'Wednesday': 4,
        'Thursday': 5,
        'Friday': 6,
        'Saturday': 7
    }
    # Strip any leading or trailing spaces from the day strings
    my_tasks.iloc[:, 0] = my_tasks.iloc[:, 0].str.strip()
    # Add a new column 'day_order' to the DataFrame based on the mapping
    my_tasks['day_order'] = my_tasks.iloc[:, 0].map(day_order)
    # Sort the DataFrame first by 'day_order' and then by 'date'
    my_tasks = my_tasks.sort_values(by=['day_order', my_tasks.columns[1]])

    # Start creating the message
    message = f"Hi, this is {name}.\n These are your tasks for the upcoming week:\n\n"
    
    for index, row in my_tasks.iterrows():
        name_of_task = row[2].replace(" ","")
        if row[8] == worker:
            name_of_task += "(role3)"
        elif row[9]== worker:
            name_of_task += "(role4)"
        day = row[0]
        date = row[1].strftime('%d-%m-%Y').replace("-","/")  # Now this is already in 'YYYY-MM-DD' format
        from_ = row[3].strftime('%H:%M')
        to_ = row[5].strftime('%H:%M')
        if pd.isna(row[10]):
            message += f"{day} {date} \n *{name_of_task}* {str(from_)}-{str(to_)} \n\n" 
        else:
            comment = row[10]
            message += f"{day} {date} \n *{name_of_task}* {str(from_)}-{str(to_)} ({comment})\n\n"     
    
    message += "Have a nice weekend."
    return message

def create_tomorrow_message(day, my_tasks, name, worker):
    message = f"Hi, this is {name}(:\nThese are your tasks for tomorrow:\n "
    tasks = my_tasks[my_tasks['day'].str.replace(" ","").str.strip() == day]
    for index, row in tasks.iterrows():
        name_of_task = row[2].replace(" ","")
        if row[8] == worker:
            name_of_task += "(role3)"
        elif row[9]== worker:
            name_of_task += "(role4)"
        from_ = row[3].strftime('%H:%M')
        to_ = row[5].strftime('%H:%M')
        brief = row[4].strftime('%H:%M')
        if pd.isna(row[10]):
            message += "\n" + "*" + name_of_task + "*" + f" {str(from_)}-{str(to_)}\n "
        else:
            comment = row[10]
            message += "\n" + "*" + name_of_task + "*" + f" {str(from_)}-{str(to_)}({comment}) \n " 
        message += f"brief {brief}\n\n"
    message +=  "Waiting for your approval."
   
    return message

def work_in_specific_day(day):
    df_tasks, df_names = load_data()
    # Filter the DataFrame to include only tasks on the specified day
    tasks_for_day = df_tasks[df_tasks['day'].str.replace(" ","").str.strip() == day]
    # Extract the workers' columns:
    workers = tasks_for_day[['role1', 'role2','role3','role4']]
    # Stack the worker columns into a single Series and drop NaNs (in case some tasks don't have all roles filled)
    workers = workers.stack().dropna()
    # Get unique worker names
    unique_workers = pd.Series(workers).str.strip().unique()
    return unique_workers

# Creates a custom popup window to confirm WhatsApp connection.
def ask_whatsapp_connected():
    # Create a new Tkinter window
    popup = tk.Toplevel()
    popup.title("whatsapp connection")
    
    # Set window size
    popup.geometry("300x100")

    # Add label to the window
    label = ttk.Label(popup, text="Were you able to successfully connect to WhatsApp Web?", wraplength=250, anchor="center", justify="center")
    label.pack(pady=10, fill="both")

    # Variable to store the user's answer
    user_answer = tk.BooleanVar()

    # Function to handle Yes button click
    def on_yes():
        user_answer.set(True)  # Set the answer to True (Yes)
        popup.destroy()  # Close the popup window

    # Function to handle No button click
    def on_no():
        user_answer.set(False)  # Set the answer to False (No)
        popup.destroy()  # Close the popup window

    # Create Yes and No buttons
    yes_button = ttk.Button(popup, text="yes", command=on_yes)
    yes_button.pack(side="left", padx=20, pady=10)

    no_button = ttk.Button(popup, text="no", command=on_no)
    no_button.pack(side="right", padx=20, pady=10)

    # Wait until the user answers (this will block until popup is closed)
    popup.wait_window()

    # Return the user's answer
    return user_answer.get()


def send_all_tasks(name):
    df_tasks, df_names = load_data()
    service = Service(executable_path=driver_path)
    options = webdriver.ChromeOptions()
    
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("http://web.whatsapp.com/")
    
    wait = WebDriverWait(driver, 100)
    
    # Ask the user if they have connected to WhatsApp
    if not ask_whatsapp_connected():
        driver.quit()
        return
    
    successful_sends = []
    failed_sends = []

    # Concatenate the 4 columns
    combined_series = pd.concat([
    df_tasks['role1'].astype(str).str.strip(),
    df_tasks['role2'].astype(str).str.strip(),
    df_tasks['role3'].astype(str).str.strip(),
    df_tasks['role4'].astype(str).str.strip()
])
    # Drop duplicates
    unique_names = combined_series.drop_duplicates()
    
    try:
        for worker in unique_names:
            row_index = df_names.index[df_names['Name'].astype(str).str.strip() == worker.strip()] 
            phone_number = df_names.loc[row_index[0], 'Phone_Number']
            my_tasks = find_tasks(worker)
            
            if not my_tasks.empty:
                message = create_weekly_message(my_tasks, name, worker)

                # Encode the message and format the phone number into the URL
                whatsapp_url = f"https://web.whatsapp.com/send?phone={phone_number}&text={quote(message)}"
                
                driver.get(whatsapp_url)
                
                try:
                    # Wait for the "Send" button to appear, then click it
                    element = WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'button[aria-label="שליחה"]'))
                    )
                    element.click()  # Click the send button
                    
                    time.sleep(2)  # Wait to ensure the message is sent
                    successful_sends.append(worker)
                except TimeoutException:
                    print(f"Failed to send message to {worker} ({phone_number})")
                    failed_sends.append(worker)
                    continue  # Skip to the next contact
                
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        time.sleep(30)
        driver.quit()
        
        # Create a new Excel file for the results
        df_results = pd.DataFrame({
            'Successful Sends': pd.Series(successful_sends),
            'Failed Sends': pd.Series(failed_sends)
        })

        result_file_path = './'
        df_results.to_excel(result_file_path, index=False)

    print("All tasks sent.")


def send_tasks_for_day(day, name):
    df_tasks, df_names = load_data()
    service = Service(executable_path=driver_path)
    options = webdriver.ChromeOptions()
    
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("http://web.whatsapp.com/")
    
    wait = WebDriverWait(driver, 100)
    
    # Ask the user if they have connected to WhatsApp
    if not ask_whatsapp_connected():
        driver.quit()
        return
    
    successful_sends = []
    failed_sends = []


    try:
        contacts = work_in_specific_day(day)
        
        for worker in contacts:
            worker = worker.replace(" ","")
            my_tasks = find_tasks(worker)

            # Assuming there's a 'Phone_Number' column in Excel
            phone_number = df_names[df_names['Name'] == worker]['Phone_Number'].values[0]
            phone_number = str(phone_number)
            
            if not my_tasks.empty:
                message = create_tomorrow_message(day, my_tasks, name, worker)

                # Encode the message and format the phone number into the URL
                whatsapp_url = f"https://web.whatsapp.com/send?phone={phone_number}&text={quote(message)}"
                
                driver.get(whatsapp_url)
                
                try:
                    # Wait for the "Send" button to appear, then click it
                    element = WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'button[aria-label="שליחה"]'))
                    )
                    element.click()  # Click the send button
                    
                    time.sleep(5)  # Wait for the message to be sent
                    successful_sends.append(worker)
                except TimeoutException:
                    print(f"Failed to send message to {worker} ({phone_number})")
                    failed_sends.append(worker)
                    continue  # Skip to the next contact
                
    finally:
        time.sleep(10)
        driver.quit()
        
        # Create a new Excel file for the results
        df_results= pd.DataFrame({
            'Successful Sends': pd.Series(successful_sends),
            'Failed Sends': pd.Series(failed_sends)
        })

        result_file_path = './'
        df_results.to_excel(result_file_path, index=False)
    
    print(f"Tasks for {day} sent.")


def create_gui():
    """Create the main GUI window with buttons for various actions."""
    root = tk.Tk()
    root.title("AutoReminderMessages")

    # Set up the GUI layout
    frame = ttk.Frame(root, padding="10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    # Variable to store the user's name
    saved_name = tk.StringVar()

    # Add a label for the Text box
    text_title = ttk.Label(frame, text="Enter youre name:")
    text_title.grid(row=0, column=0, pady=5, padx=5, sticky=tk.W)

    # Create a Text widget to enter the name
    name_entry = ttk.Entry(frame, width=20)
    name_entry.grid(row=0, column=1, pady=5, padx=5, sticky=(tk.W))

    # Function to save the entered name
    def save_name():
        entered_name = name_entry.get().strip()  # Get the text and remove any leading/trailing spaces
        if entered_name:
            saved_name.set(entered_name)
            print(f"Name saved: {entered_name}")
        else:
            print("Please enter a valid name.")

    # Add a "Save Name" button
    save_button = ttk.Button(frame, text="save", command=save_name)
    save_button.grid(row=1, column=0, columnspan=2, pady=5)

    # Function to send messages for all tasks
    def send_messages_for_all():
        if saved_name.get():
            send_all_tasks(saved_name.get())  # Pass the saved name
        else:
            print("Please save your name first.")

    # Add a button to send all messages
    send_all_button = ttk.Button(frame, text="send weekly reminder", command=send_messages_for_all)
    send_all_button.grid(row=2, column=0, columnspan=2, pady=5, sticky="ew")

    # Add a label and combobox for selecting the day
    day_label = ttk.Label(frame, text="Choose day:")
    day_label.grid(row=3, column=0, pady=5, sticky=tk.W)

    day_combobox = ttk.Combobox(frame, values=['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'])
    day_combobox.grid(row=3, column=1, pady=5, sticky=(tk.W))

    # Function to send messages for the selected day
    def send_messages_for_day():
        selected_day = day_combobox.get()
        if saved_name.get():
            send_tasks_for_day(selected_day, saved_name.get())  # Pass the saved name
        else:
            print("Please save your name first.")

    # Add a button to send messages for the selected day
    send_day_button = ttk.Button(frame, text="Send daily reminder", command=send_messages_for_day)
    send_day_button.grid(row=4, column=0, columnspan=2, pady=5)

    root.mainloop()

# Load data to ensure file path is correct before creating GUI
df_tasks, df_names = load_data()
if df_tasks is not None and df_names is not None:
    create_gui()
else:
    print("Program terminated: Excel file not loaded correctly.")




