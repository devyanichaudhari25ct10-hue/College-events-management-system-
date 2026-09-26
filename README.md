# College-events-management-system-
📌 About the Project

The College Events Management System is a desktop-based GUI application developed using Python, Tkinter, and Excel.

This system is designed to manage college event information in an organized way. Users can add, view, search, update, and delete college events through a simple graphical interface.

Event data is stored in an Excel file, which is automatically created when the program is run for the first time.

🎯 Objectives

- To manage college event information digitally.
- To provide a simple graphical user interface.
- To store event details in an Excel file.
- To make searching and updating events easier.
- To practice Python GUI and file-handling concepts.

🛠️ Technologies Used

- Python
- Tkinter – for GUI development
- ttk Treeview – for displaying event records in table format
- OpenPyXL – for reading and writing Excel files
- Excel (.xlsx) – for storing event data

✨ Features

🔐 Login System

The application provides a basic admin login.

Username: "admin"
Password: "1234"

➕ Add Event

Users can add a new event by entering:

- Event ID
- Event Name
- Date
- Venue
- Organizer
- Category

The event information is saved in the Excel file.

👀 View Events

All stored events are displayed in a table using Tkinter's "Treeview".

🔎 Search Event

Users can search for an event using:

- Event ID
- Event Name

The matching event is displayed in the table.

✏️ Update Event

Users can enter an Event ID, find the existing event, modify its details, and save the updated information.

🗑️ Delete Event

Users can delete an event by entering its Event ID.

🚪 Logout

The user can log out and return to the login screen.

📊 Event Details Stored

Each event contains the following information:

Field| Description
Event ID| Unique ID of the event
Event Name| Name of the college event
Date| Date of the event
Venue| Location of the event
Organizer| Person/department organizing the event
Category| Type of event

📁 Project Structure

College-Events-Management-System/
│
├── college_events.py
├── college_events.xlsx
└── README.md

«"college_events.xlsx" is created automatically when the program is executed if the file does not already exist.»

▶️ How to Run the Project

1. Install Python

Make sure Python is installed on your computer.

2. Install OpenPyXL

Open the terminal and run:

pip install openpyxl

3. Run the Program

Run the Python file:

python college_events.py

The login screen will appear.

4. Login

Use:

Username: admin
Password: 1234

After successful login, the dashboard will be displayed.

🔄 Working Flow

Start
  ↓
Login
  ↓
Dashboard
  ↓
Choose an Operation
  ↓
Add / View / Search / Update / Delete
  ↓
Excel File
  ↓
Logout

📚 Python Concepts Used

This project uses the following concepts:

- Functions
- Variables
- Conditional statements
- Loops
- Nested functions
- "Entry"
- "Label"
- "Button"
- "Frame"
- "Treeview"
- "pack()"
- "grid()"
- "winfo_children()"
- File handling
- Excel file handling
- "if-else"
- "for" loop

📦 Python Modules Used

import tkinter as tk
from tkinter import ttk
from openpyxl import Workbook, load_workbook
import os

Module Purpose

- tkinter – Creates the graphical user interface.
- ttk – Provides the Treeview table.
- openpyxl – Creates, reads, updates, and saves Excel data.
- os – Checks whether the Excel file already exists.

💾 Data Storage

The project uses an Excel file named:

college_events.xlsx

The Excel file contains these columns:

Event ID | Event Name | Date | Venue | Organizer | Category

🔮 Future Scope

The project can be further improved by adding:

- Event registration
- Admin and student roles
- Event reminders
- Date-wise event filtering
- Better input validation
- Database connectivity
- Event notification system
- Export and reporting features

🎓 Academic Purpose

This project was developed as an academic project to understand the practical use of Python GUI programming, Tkinter, Excel file handling, and basic CRUD operations.

👩‍💻 Developer

Devyani Chaudhari

B.Tech – Computer Technology
