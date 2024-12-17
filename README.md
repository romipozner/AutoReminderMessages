# AutoReminderMessages

## Project Overview
AutoReminderMessages is a software solution designed to help companies efficiently remind their employees about upcoming tasks for the next day or the upcoming week. By automatically generating and sending task reminders based on a pre-prepared Excel table, the software eliminates the need for manual follow-ups and minimizes errors, helping businesses streamline their operations.

---

## Features
- **Automated Task Reminders**: Generates and sends task reminders for employees based on input from an Excel file.
- **Customizable Schedule**: Users can configure the tool to send reminders for tasks due the **next day** or the **upcoming week**.
- **Error Reduction**: Automates the process to avoid human errors in sending reminders.
- **Streamlined Workflow**: Saves time and effort by reducing manual follow-ups.

---

## How to Use

Follow these steps to set up and run the AutoReminderMessages project:

1. **Prerequisites**
   - Make sure you have the following installed:
     - Python (>= 3.x)
     - Required Python libraries: pandas, openpyxl, selenium etc.  
      
2. **Prepare the Excel File**
   - Prepare an Excel file containing tasks, employees, times, due dates ect. The file structure attached in this folder under the name "tasks_template".
   - Save the file in the project's directory.

3. **Run the Software**
   - Run the script by executing the following command in your terminal:
     ```bash
     python auto.py
     ```

4. **Generate Reminder Messages**
   - The software will process the Excel file and generate task reminders for employees.
   - The reminders can be configured to focus on tasks for the next day or the upcoming week.

5. **Output**
   - Reminders will be generated and sent automatically (depending on the configuration).

---

## Technologies and Tools
The project uses the following technologies:
- **Python**: Core programming language for processing data and automation.
- **pandas**: Library for reading and manipulating Excel data.
- **openpyxl**: For working with Excel files.
- **Excel**: Input file format for task management.
- **Selenium and ChromeDriver** for automation.
- **Tkinter** for the user interface.

---

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---



