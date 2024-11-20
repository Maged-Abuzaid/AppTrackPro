import logging
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from datetime import datetime
import pandas as pd
import json
import webbrowser
import shutil
from PIL import Image, ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD  # Use tkinterdnd2 for drag-and-drop functionality
from config.settings_manager import CONFIG_JSON_PATH, default_config, save_theme, SPREADSHEET_ID, ENABLE_GOOGLE_SYNC, \
    ASSETS_DIR

# Import configuration settings from settings_manager.py
from config.settings_manager import (
    ICON_PATH,
    PERSONAL_INFO_FILE,
    DATA_FILE_PATH,
    base_path,
    CONFIG_JSON_PATH,
    SERVICE_ACCOUNT_FILE,
    theme
)

# Import utility functions for file I/O and Google Sheets synchronization
from src.utils.file_io import read_applications_from_excel, save_applications_to_excel
from src.utils.google_sheets import (
    read_from_google_sheets,
    write_to_google_sheets,
    delete_row_in_google_sheets,
)

# Import the centralized resource_path function from utils/utils.py
from src.utils.utils import resource_path

def load_personal_info():
    """
    Loads personal information from a JSON file.
    Returns default data if the file does not exist.
    """
    if os.path.exists(PERSONAL_INFO_FILE):
        # Load JSON data if the file exists
        with open(PERSONAL_INFO_FILE, "r") as file:
            return json.load(file)
    else:
        # Return default information if file is missing
        return {
            "First Name": "John",
            "Last Name": "Doe",
            "Email": "john.doe@example.com",
            "Password": "password123",
            "Phone Number": "+1 (555) 123-4567",
            "Address Line 1": "123 Main St",
            "City": "Anytown",
            "State": "CA",
            "Zip Code": "12345",
            "Full Address": "123 Main St, Anytown, CA 12345",
            "University": "State University",
            "Degree": "BS in Computer Science",
        }

class AppTrackPro(TkinterDnD.Tk):  # Inherit from TkinterDnD.Tk for drag-and-drop
    def __init__(self):
        super().__init__()
        self.APPLICATIONS_FILE_NAME = "Applications.xlsx"
        self.SERVICE_ACCOUNT_FILE_NAME = "service_account.json"  # Define the service account file name
        self.CONFIG_DIR = os.path.join(base_path, 'config')
        self.DATA_DIR = os.path.join(base_path, 'Data')
        self.bg_color = "#2E2E2E"
        self.fg_color = "#FFFFFF"
        self.sync_to_google = False  # Initialize sync_to_google to default
        self.status_combobox = None
        self.edit_entry = None  # Initialize edit_entry as None
        self.menu_visible = False  # Variable to track menu visibility

        # Configure the main window
        self.configure_window()

        # Initialize paths
        self.initialize_paths()

        # Load and apply theme
        self.load_and_apply_theme()

        # Initialize preferences
        self.initialize_preferences()

        # Create UI components
        self.create_ui_components()

        # Load assets (images)
        self.load_assets()

        # Initialize additional GUI components
        self.initialize_additional_gui()

        # Load application data
        self.load_application_data()

        # Setup the main layout
        self.setup_main_layout()

        # Schedule periodic tasks
        self.schedule_tasks()

        # Apply the current theme
        self.apply_theme()

        # Initialize sync task
        self.sync_task = None
        self.schedule_sync()

    def configure_window(self):
        # Use the native title bar by removing overrideredirect
        self.title("AppTrackPro")
        self.geometry("1300x600")
        try:
            icon_path = resource_path(os.path.join('assets', 'app_icon.png'))
            icon_image = tk.PhotoImage(file=icon_path)
            self.iconphoto(True, icon_image)  # Use .png file for the application icon
        except Exception as e:
            print(f"Error loading icon: {e}")
            logging.error(f"Error loading icon: {e}")

    def initialize_paths(self):
        self.BASE_PATH = base_path  # Use AppData base path directly
        self.CONFIG_DIR = os.path.join(self.BASE_PATH, 'config')
        self.CONFIG_JSON_PATH = CONFIG_JSON_PATH  # Already set to AppData path
        self.DATA_DIR = os.path.join(self.BASE_PATH, 'Data')
        self.DATA_FILE_PATH = DATA_FILE_PATH  # Directly use the AppData path for Applications.xlsx
        os.makedirs(self.DATA_DIR, exist_ok=True)
        os.makedirs(self.CONFIG_DIR, exist_ok=True)

    def load_and_apply_theme(self):
        self.is_dark_mode = self.load_theme_from_config()
        if self.is_dark_mode:
            self.set_dark_mode()
        else:
            self.set_light_mode()

    def initialize_preferences(self):
        """Initialize preferences like Google Sync based on the configuration."""
        self.sync_to_google = ENABLE_GOOGLE_SYNC
        self.google_sync_var = tk.BooleanVar(value=self.sync_to_google)

    def create_ui_components(self):
        self.create_custom_menu_bar()

    def load_assets(self):
        try:
            self.upload_xlsx_icon = ImageTk.PhotoImage(
                Image.open(resource_path(os.path.join('assets', 'upload_xlsx.png'))).resize((64, 64))
            )
            self.upload_json_icon = ImageTk.PhotoImage(
                Image.open(resource_path(os.path.join('assets', 'upload_json.png'))).resize((64, 64))
            )
            self.upload_sheets_id_icon = ImageTk.PhotoImage(
                Image.open(resource_path(os.path.join('assets', 'upload_sheets_id.png'))).resize((42, 42))
            )
            try:
                self.google_sync_icon = tk.PhotoImage(file=os.path.join(ASSETS_DIR, 'google_sync.png'))
            except Exception as e:
                print(f"Error loading google_sync.png: {e}")
                self.google_sync_icon = None

            try:
                self.applications_icon = tk.PhotoImage(file=os.path.join(ASSETS_DIR, 'applications.png'))
            except Exception as e:
                print(f"Error loading applications.png: {e}")
                self.applications_icon = None

        except Exception as e:
            print(f"Error loading assets: {e}")
            logging.error(f"Error loading assets: {e}")

    def initialize_additional_gui(self):
        self.selected_row = None
        self.selected_column = None
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.perform_search())
        self.applications_tree = None
        self.position_entry = None
        self.company_entry = None
        self.url_entry = None
        self.applications_df = pd.DataFrame()

    def load_application_data(self):
        """Loads application data from Applications.xlsx in AppData."""
        try:
            self.applications_df = self.read_applications_from_excel(self.DATA_FILE_PATH)
        except Exception as e:
            print(f"Error: Could not read the Excel file from AppData. {str(e)}")
            self.applications_df = pd.DataFrame()

    def setup_main_layout(self):
        main_paned_window = tk.PanedWindow(self, orient="horizontal")
        main_paned_window.pack(side='top', fill='both', expand=True)

        self.tab_control = ttk.Notebook(main_paned_window)
        self.add_application_tab = ttk.Frame(self.tab_control)
        self.view_edit_applications_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.add_application_tab, text="Add Application")
        self.tab_control.add(self.view_edit_applications_tab, text="View/Edit Applications")
        main_paned_window.add(self.tab_control, stretch="always")

        self.right_notebook = ttk.Notebook(main_paned_window)
        self.clipboard_tab = ttk.Frame(self.right_notebook)
        self.right_notebook.add(self.clipboard_tab, text="Clipboard")
        main_paned_window.add(self.right_notebook, stretch="always")

        main_paned_window.paneconfig(self.tab_control, minsize=300)
        main_paned_window.paneconfig(self.right_notebook, minsize=125)

        self.create_add_application_tab()
        self.create_view_edit_applications_tab()
        self.create_personal_info_tab()

    def schedule_tasks(self):
        if self.sync_to_google:
            self.after(60000, self.sync_to_google_sheets)
            self.schedule_sync()
            self.apply_theme()

#-----------------------------------------------
    def load_theme_from_config(self):
        """Load the saved theme setting from app_config.json in AppData, or create it with default settings if it doesn't exist."""
        # Check if app_config.json exists in AppData (CONFIG_JSON_PATH)
        if not os.path.exists(CONFIG_JSON_PATH):
            # Create app_config.json in AppData with default theme settings
            default_config = {"theme": "Light"}
            with open(CONFIG_JSON_PATH, "w") as config_file:
                json.dump(default_config, config_file, indent=4)
            return False  # Default to Light theme

        # Load theme from existing app_config.json in AppData
        try:
            with open(CONFIG_JSON_PATH, "r") as config_file:
                config = json.load(config_file)
                return config.get("theme", "Light") == "Dark"
        except json.JSONDecodeError:
            messagebox.showerror("Error", "app_config.json is corrupted. Reverting to default settings.")
            return False  # Default to Light theme

    def create_add_application_tab(self):
        """
        Sets up the 'Add Application' tab with entry fields for company, position,
        and URL, as well as a 'Submit' button to add the application.
        """
        # Font settings for labels and entries
        label_font = ("TkDefaultFont", 12)
        entry_font = ("TkDefaultFont", 12)

        # Form frame to hold input fields, making it easier to center content
        form_frame = tk.Frame(self.add_application_tab)
        form_frame.grid(row=0, column=0, padx=20, pady=20, sticky="n")

        # Input for 'Company' (first input field)
        tk.Label(form_frame, text="Company:", font=label_font).grid(
            row=0, column=0, padx=10, pady=(0, 5), sticky="w"
        )
        self.company_entry = tk.Entry(form_frame, font=entry_font, width=60)
        self.company_entry.grid(row=1, column=0, padx=10, pady=(0, 10))

        # Input for 'Position' (second input field)
        tk.Label(form_frame, text="Position:", font=label_font).grid(
            row=2, column=0, padx=10, pady=(0, 5), sticky="w"
        )
        self.position_entry = tk.Entry(form_frame, font=entry_font, width=60)
        self.position_entry.grid(row=3, column=0, padx=10, pady=(0, 10))

        # Input for 'Application Portal URL'
        tk.Label(form_frame, text="Application Portal URL:", font=label_font).grid(
            row=4, column=0, padx=10, pady=(0, 5), sticky="w"
        )
        self.url_entry = tk.Entry(form_frame, font=entry_font, width=60)
        self.url_entry.grid(row=5, column=0, padx=10, pady=(0, 20))

        # 'Submit' button for adding the application (updated to ttk.Button)
        add_button = ttk.Button(
            form_frame,
            text="Submit",
            command=self.save_application,
            style="Custom.TButton"
        )
        add_button.grid(row=6, column=0, pady=10)

        # Center form_frame within add_application_tab
        self.add_application_tab.grid_rowconfigure(0, weight=1)
        self.add_application_tab.grid_columnconfigure(0, weight=1)
        form_frame.grid_rowconfigure(0, weight=1)
        form_frame.grid_columnconfigure(0, weight=1)

    def create_view_edit_applications_tab(self):
        """
        Sets up the 'View/Edit Applications' tab with a search bar and a Treeview
        to display application data with a vertical scrollbar and column configuration.
        """
        # Attempt to load applications data
        try:
            self.applications_df = read_applications_from_excel(DATA_FILE_PATH)
        except Exception as e:
            print(f"Error: Could not read the Excel file. {str(e)}")
            self.applications_df = pd.DataFrame()

        # Frame for search bar
        search_frame = tk.Frame(self.view_edit_applications_tab)
        search_frame.grid(row=0, column=0, sticky="ew")

        # Search label and entry field
        tk.Label(search_frame, text="Search:").pack(side="left", padx=5)
        search_entry = tk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side="left", padx=5)

        # Frame for the main Treeview
        frame = tk.Frame(self.view_edit_applications_tab)
        frame.grid(row=1, column=0, sticky="nsew")
        self.view_edit_applications_tab.rowconfigure(1, weight=1)
        self.view_edit_applications_tab.columnconfigure(0, weight=1)

        # Define Treeview columns
        columns = ("No", "Company", "Position", "Application Portal URL", "Date Applied", "Status")
        self.applications_tree = ttk.Treeview(frame, columns=columns, show="headings", height=15)

        # Configure column headers and alignment
        for col in columns:
            self.applications_tree.heading(
                col,
                text=col,
                anchor="center" if col in ["No", "Date Applied", "Status"] else "w"
            )
            if col in ["Date Applied", "Status", "No"]:
                self.applications_tree.column(
                    col,
                    anchor="center",
                    stretch=True,
                    width=120 if col != "No" else 50
                )
            else:
                self.applications_tree.column(col, anchor="w", stretch=True, width=150)

        # Bind Treeview events for clicking and context menu
        self.applications_tree.bind("<Button-1>", self.on_treeview_click)
        self.applications_tree.bind("<Button-3>", self.show_context_menu)  # Right-click menu
        self.populate_treeview(self.applications_df)

        # Add vertical scrollbar for Treeview
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.applications_tree.yview)
        self.applications_tree.configure(yscrollcommand=vsb.set)

        # Position Treeview and scrollbar in grid
        self.applications_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # Configure grid weights
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def create_personal_info_tab(self):
        """
        Creates the 'Personal Information' tab with editable fields
        and clipboard copy functionality for each field.
        """
        row = 0
        self.personal_info_entries = {}

        # Load personal information from JSON file
        personal_info = load_personal_info()

        # Add each personal info item as a label-entry pair
        for label, value in personal_info.items():
            # Field label (clickable to copy value)
            field_label = tk.Label(self.clipboard_tab, text=label + ":", fg="blue", cursor="hand2")
            field_label.grid(row=row, column=0, padx=5, pady=5, sticky="e")

            # Bind click event to copy value to clipboard
            field_label.bind("<Button-1>", lambda e, val=value: self.copy_to_clipboard(val))

            # Entry widget for displaying (and editing) field value
            entry = tk.Entry(self.clipboard_tab, width=38)
            entry.insert(0, value)
            entry.grid(row=row, column=1, padx=5, pady=5, sticky="w")

            # Store the entry widget to update values later if needed
            self.personal_info_entries[label] = entry
            row += 1

        # Save button setup (updated to ttk.Button)
        button_frame = tk.Frame(self.clipboard_tab)
        button_frame.grid(row=row, column=0, columnspan=2, pady=10)
        save_button = ttk.Button(
            button_frame,
            text="Save",
            command=self.save_personal_info,
            style="Custom.TButton"
        )
        save_button.grid(row=0, column=0)
        button_frame.grid_columnconfigure(0, weight=1)

    # Treeview Setup and Interaction
    def populate_treeview(self, df):
        """
        Populates the applications Treeview with data from the DataFrame.
        """
        # Clear existing data in the Treeview
        for item in self.applications_tree.get_children():
            self.applications_tree.delete(item)

        # Insert new data into the Treeview
        for index, row in df.iterrows():
            values = [index + 1] + list(row)
            self.applications_tree.insert("", "end", iid=index, values=values)

    def refresh_treeview(self):
        """
        Clears and repopulates the Treeview with the most current DataFrame Data.
        """
        # Clear existing Treeview Data
        for item in self.applications_tree.get_children():
            self.applications_tree.delete(item)

        # Reload and display the current Data in the Treeview
        self.populate_treeview(self.applications_df)

    def on_treeview_click(self, event):
        """
        Handles single-click events on the Treeview cells.
        Supports dropdown selection for 'Status' column and URL opening for the 'Application Portal URL' column.
        """
        # Close any open status dropdown
        if self.status_combobox:
            self.status_combobox.destroy()
            self.status_combobox = None

        # Identify the selected row and column
        selected_item = self.applications_tree.selection()
        if not selected_item:
            return  # Exit if no item is selected

        selected_item = selected_item[0]
        column = self.applications_tree.identify_column(event.x)
        col_index = int(column[1:]) - 1  # Convert column to zero-based index

        # Show status dropdown if the "Status" column is clicked
        if col_index == 5:
            self.show_status_dropdown(selected_item, col_index)
            # Clear selection to avoid URL opening inadvertently
            self.selected_row = None
            self.selected_column = None
            return

        # Handle URL column double-click behavior
        if col_index == 3:  # 'Application Portal URL' column
            values = self.applications_tree.item(selected_item, "values")
            url = values[col_index] if len(values) > col_index else ""

            # Check if the URL cell is clicked twice consecutively
            if self.selected_row == selected_item and self.selected_column == col_index:
                # Open URL in the default web browser if valid
                if url and url.startswith("http"):
                    print("Opening URL:", url)
                    try:
                        webbrowser.open(url, new=2)  # Open in a new browser tab
                    except webbrowser.Error:
                        print("Error: Could not open the URL.")
                # Reset tracking after opening URL
                self.selected_row = None
                self.selected_column = None
            else:
                # Update selection on first click, without opening the URL
                self.selected_row = selected_item
                self.selected_column = col_index
                print("Row selected. Click again to open the URL if in the URL column.")
        else:
            # Reset selection if another column is clicked
            self.selected_row = None
            self.selected_column = None

    def on_treeview_double_click(self, event):
        """
        Enables in-place editing of Treeview cells on double-click.
        Displays an entry widget over the selected cell for editable columns.
        """
        selected_item = self.applications_tree.selection()
        if not selected_item:
            return  # Exit if no item is selected

        selected_item = selected_item[0]
        column = self.applications_tree.identify_column(event.x)
        col_index = int(column[1:]) - 1  # Zero-based column index

        # Get cell coordinates and current value
        x, y, width, height = self.applications_tree.bbox(selected_item, column)
        current_value = self.applications_tree.item(selected_item, "values")[col_index]

        # Create an entry widget over the cell for editing
        self.edit_entry = tk.Entry(self.applications_tree, width=width)
        self.edit_entry.insert(0, current_value)
        if x and y:
            self.edit_entry.place(x=x, y=y, width=width, height=height)

        # Bind actions to save or cancel editing
        self.edit_entry.bind("<Return>", lambda e: self.save_edit(selected_item, col_index))
        self.edit_entry.bind("<FocusOut>", lambda e: self.edit_entry.destroy())
        self.edit_entry.focus_set()

    def on_treeview_cell_edit(self, event):
        """
        Enables direct in-cell editing for non-URL, non-Status columns.
        Activates an editable entry when the relevant cell is selected.
        """
        selected_item = self.applications_tree.selection()
        if not selected_item:
            return  # Exit if no item is selected

        selected_item = selected_item[0]
        column = self.applications_tree.identify_column(event.x)
        col_index = int(column[1:]) - 1  # Zero-based index in applications_df

        # Allow editing for non-Status and non-URL columns only
        if self.applications_tree["columns"][col_index] not in ["Status", "Application Portal URL"]:
            self.applications_tree.bind("<KeyRelease>", lambda e: self.save_direct_edit(selected_item, col_index))

    # Editing and Saving
    def create_edit_entry(self, item_id, col_index):
        """
        Creates an Entry widget directly within the Treeview cell, allowing in-cell editing.
        Only available for non-URL columns.
        """
        # Get the bounding box coordinates of the cell for placing the Entry widget
        x, y, width, height = self.applications_tree.bbox(item_id, column="#" + str(col_index + 1))

        # Retrieve the current value of the cell for editing
        current_value = self.applications_tree.item(item_id, "values")[col_index]

        # Create and place an Entry widget at the cell's location with its current value
        self.edit_entry = tk.Entry(self.applications_tree, width=width)
        self.edit_entry.insert(0, current_value)
        self.edit_entry.place(x=x, y=y, width=width, height=height)

        # Bind events to save or cancel edit on Enter key or focus loss
        self.edit_entry.bind("<Return>", lambda e: self.save_edit(item_id, col_index))
        self.edit_entry.bind("<FocusOut>", lambda e: self.edit_entry.destroy())
        self.edit_entry.focus_set()  # Set focus to the entry widget for immediate editing

    def save_edit(self, item_id, col_index):
        """
        Saves the edited value from the Entry widget back to both the Treeview cell and the DataFrame.
        """
        # Check if edit_entry exists and retrieve the new value
        if self.edit_entry:
            new_value = self.edit_entry.get()

            # Update the Treeview cell with the new value
            values = list(self.applications_tree.item(item_id, "values"))
            values[col_index] = new_value
            self.applications_tree.item(item_id, values=values)

            # Update the DataFrame with the new value
            column_name = self.applications_tree["columns"][col_index]
            self.applications_df.at[int(item_id), column_name] = new_value

            # Save the DataFrame to Excel
            save_applications_to_excel(self.applications_df)

            # Conditionally sync updated data to Google Sheets if sync is enabled
            if self.sync_to_google:
                try:
                    write_to_google_sheets(self.applications_df)
                    print(f"Updated '{column_name}' synced to Google Sheets for row {item_id}.")
                except Exception as e:
                    print(f"Error: Could not sync with Google Sheets. {str(e)}")
            else:
                print("Google Sync is disabled. Changes were not synced to Google Sheets.")

            # Destroy the Entry widget after saving the edit
            self.edit_entry.destroy()
            self.edit_entry = None  # Reset edit_entry

            print(f"Saved edit: {new_value} in cell ({item_id}, {col_index}).")
        else:
            print("Edit entry does not exist to save.")

    def save_direct_edit(self, item_id, col_index):
        """
        Saves the edited value directly within the cell without needing an Entry widget.
        Only allows saving for non-URL and non-Status columns.
        """
        # Retrieve the directly edited value from the Treeview cell
        edited_value = self.applications_tree.item(item_id, "values")[col_index]

        # Update the DataFrame with the edited value for the specified column
        column_name = self.applications_tree["columns"][col_index]
        self.applications_df.at[int(item_id), column_name] = edited_value

        # Persist changes by saving the updated DataFrame to Excel
        save_applications_to_excel(self.applications_df)

        # Unbind the key release event after saving to prevent unintended edits
        self.applications_tree.unbind("<KeyRelease>")

    # Clipboard and Copying
    def copy_to_clipboard(self, value):
        """
        Copies the specified value to the system clipboard.

        Parameters:
        value (str): The text value to copy to the clipboard.
        """
        # Clear existing clipboard content
        self.clipboard_clear()

        # Append the specified value to the clipboard
        self.clipboard_append(value)

        # Log confirmation of the copied value
        print(f"Copied to clipboard: {value}")

    def copy_rows(self, row_ids):
        """
        Copies the values of the selected rows from the Treeview to the clipboard.
        Each row's values are tab-separated, and rows are separated by newlines.
        """
        if not row_ids:
            print("No rows selected for copying.")
            return

        copied_text = ""
        for row_id in row_ids:
            # Retrieve all cell values for the specified row
            row_data = self.applications_tree.item(row_id, "values")
            # Concatenate row values into a single tab-separated string
            row_text = "\t".join(str(item) for item in row_data)
            copied_text += row_text + "\n"

        # Copy the concatenated text to the clipboard
        self.clipboard_clear()
        self.clipboard_append(copied_text.strip())

        # Log confirmation of the copied rows
        print(f"Copied {len(row_ids)} rows to clipboard.")

    # Data Management and Synchronization
    def sync_from_google_sheets(self):
        """Fetch data from Google Sheets if Google Sync is enabled."""
        if not self.sync_to_google:
            print("Google Sync is disabled. Skipping sync from Google Sheets.")
            return

        try:
            # Retrieve the latest data from Google Sheets
            google_df = read_from_google_sheets()

            # Replace NaN values with empty strings
            google_df = google_df.fillna('')

            # Check for differences and update if necessary
            if not google_df.empty:
                if not google_df.equals(self.applications_df):
                    print("Detected changes in Google Sheets. Updating local data.")
                    self.applications_df = google_df

                    # Ensure the Treeview is initialized before updating it
                    if hasattr(self, 'applications_tree') and self.applications_tree:
                        self.populate_treeview(self.applications_df)
                    else:
                        print("Error: applications_tree is not initialized yet. Will populate later.")
                else:
                    print("No changes detected in Google Sheets.")
        except Exception as e:
            print(f"Error syncing data from Google Sheets: {e}")
            # Log the error but do not disable Google Sync
            logging.error(f"Error syncing data from Google Sheets: {e}")

    def sync_to_google_sheets(self):
        """Push local DataFrame data to Google Sheets if Google Sync is enabled."""
        if not self.sync_to_google:
            print("Google Sync is disabled. Skipping sync to Google Sheets.")
            return

        try:
            # Update Google Sheets with the current DataFrame data
            write_to_google_sheets(self.applications_df)
            print("Data synced to Google Sheets successfully.")
        except Exception as e:
            print(f"Error syncing data to Google Sheets: {e}")
            # Log the error but do not disable Google Sync
            logging.error(f"Error syncing data to Google Sheets: {e}")

    def schedule_sync(self):
        """Schedules periodic syncing from Google Sheets every 60 seconds."""
        # Schedule the next sync and store the task ID
        self.sync_task = self.after(60000, self.schedule_sync)

        if self.sync_to_google:
            # Perform synchronization with Google Sheets
            self.sync_from_google_sheets()
        else:
            print("Google Sync is disabled. Skipping sync from Google Sheets.")

    def save_application(self):
        """
        Captures Data from input fields, validates it, and saves it as a new application entry.
        Updates both the local DataFrame and Google Sheets, then refreshes the Treeview.
        """
        # Retrieve input Data and clean up extra spaces
        position = self.position_entry.get().strip()
        company = self.company_entry.get().strip()
        url = self.url_entry.get().strip()  # URL is optional
        date_applied = datetime.now().strftime("%Y-%m-%d")
        status = "Submitted"  # Default status for new applications

        # Validate required fields (company and position)
        if not company or not position:
            print("Please fill out the Company and Position fields before adding an application.")
            messagebox.showerror("Error", "Company and Position are required fields.")
            return  # Stop if required fields are missing

        # Ensure DataFrame has the correct columns if it's empty
        if self.applications_df.empty:
            self.applications_df = pd.DataFrame(
                columns=["Company", "Position", "Application Portal URL", "Date Applied", "Status"])

        # Log current DataFrame columns for debugging purposes
        print("Current DataFrame columns:", self.applications_df.columns)

        # Create a new row of Data in DataFrame format
        new_data = pd.DataFrame(
            [[company, position, url, date_applied, status]],
            columns=["Company", "Position", "Application Portal URL", "Date Applied", "Status"]
        )

        # Append the new Data to the applications DataFrame
        self.applications_df = pd.concat([self.applications_df, new_data], ignore_index=True)

        # Save the updated DataFrame to the local Excel file
        save_applications_to_excel(self.applications_df)
        print("Data saved locally to Excel.")

        # Sync updated Data to Google Sheets only if sync is enabled
        if self.sync_to_google:
            try:
                write_to_google_sheets(self.applications_df)
                print("Data synced to Google Sheets.")
            except FileNotFoundError as e:
                print(f"Google Sheets sync failed: {e}")
                messagebox.showerror("Error", f"Google Sheets sync failed: {e}")
            except Exception as e:
                print(f"[ERROR] Unexpected error during Google Sheets sync: {e}")
                messagebox.showerror("Error", f"Google Sheets sync failed: {e}")

        # Refresh the Treeview to display the new application
        self.populate_treeview(self.applications_df)

        # Clear the input fields after saving
        self.clear_input_fields()

    def clear_input_fields(self):
        """
        Clears the input fields in the 'Add Application' tab.
        """
        self.company_entry.delete(0, tk.END)
        self.position_entry.delete(0, tk.END)
        self.url_entry.delete(0, tk.END)

    # Search and Filter
    def perform_search(self):
        """
        Filters the Treeview to display only rows containing the search term.
        If no search term is entered, all rows are displayed.
        """
        # Retrieve and clean the search term (convert to lowercase for case-insensitive matching)
        search_term = self.search_var.get().strip().lower()

        # If the search term is empty, display all rows
        if not search_term:
            self.populate_treeview(self.applications_df)
            return

        # Filter the DataFrame: retain rows that contain the search term in any column
        filtered_df = self.applications_df[
            self.applications_df.apply(
                lambda row: search_term in row.astype(str).str.lower().to_string(), axis=1
            )
        ]

        # Refresh the Treeview to show only the rows in the filtered DataFrame
        self.populate_treeview(filtered_df)

    # Context Menu and Cell Interaction
    def show_context_menu(self, event):
        """
        Displays a context menu on right-click with options based on the selected cell's column.
        General options include 'Delete Row' and 'Copy Row', with column-specific edit options.
        """
        # Identify the row and column where the right-click occurred
        row_id = self.applications_tree.identify_row(event.y)
        column_id = self.applications_tree.identify_column(event.x)
        col_index = int(column_id[1:]) - 1  # Convert to zero-based index

        # Get all selected rows
        selected_rows = self.applications_tree.selection()

        # Only show the context menu if a row is clicked
        if not row_id:
            return

        # Create the context menu
        context_menu = tk.Menu(self, tearoff=0)

        if len(selected_rows) > 1:
            # If multiple rows are selected, provide the option to delete all
            context_menu.add_command(label="Delete Selected Rows", command=lambda: self.delete_rows(selected_rows))
            context_menu.add_command(label="Copy Selected Rows", command=lambda: self.copy_rows(selected_rows))
        else:
            # General options: Delete Row, Copy Row
            context_menu.add_command(label="Delete Row", command=lambda: self.delete_rows([row_id]))
            context_menu.add_command(label="Copy Row", command=lambda: self.copy_row(row_id))

            # Column-specific options based on the column index
            if col_index == 1:  # Company column
                context_menu.add_command(label="Edit Company",
                                         command=lambda: self.edit_cell(row_id, col_index, "Company"))
            elif col_index == 2:  # Position column
                context_menu.add_command(label="Edit Position",
                                         command=lambda: self.edit_cell(row_id, col_index, "Position"))
            elif col_index == 3:  # URL Portal column
                context_menu.add_command(label="Edit URL",
                                         command=lambda: self.edit_cell(row_id, col_index, "Application Portal URL"))
            elif col_index == 4:  # Date Applied column
                context_menu.add_command(label="Edit Date",
                                         command=lambda: self.edit_cell(row_id, col_index, "Date Applied"))
            elif col_index == 5:  # Status column
                context_menu.add_command(label="Edit Status",
                                         command=lambda: self.show_status_dropdown(row_id, col_index))

        # Display the context menu at the mouse cursor position
        context_menu.post(event.x_root, event.y_root)

    def delete_rows(self, row_ids):
        """
        Deletes the selected rows from the Treeview, DataFrame, and Google Sheets.
        Updates the local Excel file and Treeview to reflect the deletion.
        """
        if not row_ids:
            print("No rows selected for deletion.")
            return

        # Convert row_ids to integers and sort them in descending order to prevent index shifting
        row_indices = sorted([int(row_id) for row_id in row_ids], reverse=True)

        # Remove the rows from the DataFrame and Treeview
        for row_index in row_indices:
            if row_index in self.applications_df.index:
                # Remove the row from the DataFrame
                self.applications_df = self.applications_df.drop(row_index)
            else:
                print(f"Row ID {row_index} not found in DataFrame index.")

        # Reset the DataFrame index after deletions
        self.applications_df.reset_index(drop=True, inplace=True)

        # Update the Treeview
        self.populate_treeview(self.applications_df)

        # Sync with Google Sheets if enabled
        if self.sync_to_google:
            try:
                # Update Google Sheets with the current DataFrame data
                write_to_google_sheets(self.applications_df)
                print("Data synced to Google Sheets after deletion.")
            except Exception as e:
                print(f"Error syncing data to Google Sheets: {e}")
        else:
            print("Google Sync is disabled. Changes were not synced to Google Sheets.")

        # Save the updated DataFrame to Excel
        save_applications_to_excel(self.applications_df)
        print("Data saved locally to Excel after deletion.")

    def edit_cell(self, row_id, col_index, column_name):
        """
        Creates an Entry widget directly over the specified Treeview cell for inline editing.
        """
        # If an edit Entry already exists, destroy it to prevent multiple editors
        if self.edit_entry:
            self.edit_entry.destroy()

        # Retrieve the cell coordinates and current value
        x, y, width, height = self.applications_tree.bbox(row_id, column="#" + str(col_index + 1))
        current_value = self.applications_tree.item(row_id, "values")[col_index]

        # Create an Entry widget for editing
        self.edit_entry = tk.Entry(self.applications_tree, width=width)
        self.edit_entry.insert(0, current_value)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.focus_set()  # Ensure the Entry widget is focused

        # Bind actions to save on Enter key press or focus out
        self.edit_entry.bind("<Return>", lambda e: self.save_edit(row_id, col_index))
        self.edit_entry.bind("<FocusOut>", lambda e: self.save_edit(row_id, col_index))

    def show_status_dropdown(self, item_id, col_index):
        """
        Displays a dropdown menu for editing the 'Status' column in the Treeview.

        Parameters:
        - item_id (str): Identifier of the row containing the status to edit.
        - col_index (int): Index of the 'Status' column.
        """
        status_options = ["Submitted", "Rejected", "Interview", "Offer"]
        current_status = self.applications_tree.item(item_id, "values")[col_index]

        # Create a dropdown menu (Combobox) with status options
        self.status_combobox = ttk.Combobox(self.applications_tree, values=status_options, state="readonly")
        self.status_combobox.set(current_status)
        x, y, width, height = self.applications_tree.bbox(item_id, column="#" + str(col_index + 1))

        # Position the dropdown if coordinates are valid
        if x and y:
            self.status_combobox.place(x=x, y=y, width=width, height=height)
        else:
            print("Error: Unable to place the combobox due to invalid bounding box values.")

        # Focus on the dropdown and bind selection event for saving
        self.status_combobox.focus_set()
        self.status_combobox.bind("<<ComboboxSelected>>", lambda event: self.save_status(item_id, col_index))

    def save_status(self, item_id, col_index):
        """
        Saves the selected status from the dropdown to the Treeview, DataFrame, and Google Sheets.
        """
        # Retrieve the new status from the dropdown menu
        new_status = self.status_combobox.get()
        values = list(self.applications_tree.item(item_id, "values"))
        values[col_index] = new_status
        self.applications_tree.item(item_id, values=values)

        # Update the DataFrame with the new status
        column_name = self.applications_tree["columns"][col_index]
        self.applications_df.at[int(item_id), column_name] = new_status

        # Save changes to the Excel file locally
        try:
            save_applications_to_excel(self.applications_df, DATA_FILE_PATH)
            print(f"Status '{new_status}' saved for row {item_id} in Excel.")
        except Exception as e:
            print(f"Error: Could not save to the Excel file. {str(e)}")

        # Conditionally sync the updated status to Google Sheets if sync is enabled
        if self.sync_to_google:
            try:
                write_to_google_sheets(self.applications_df)
                print(f"Status '{new_status}' synced with Google Sheets for row {item_id}.")
            except Exception as e:
                print(f"Error: Could not sync with Google Sheets. {str(e)}")
        else:
            print("Google Sync is disabled. Changes were not synced to Google Sheets.")

        # Destroy the dropdown after saving
        self.status_combobox.destroy()
        self.status_combobox = None

    # Personal Info
    def save_personal_info(self):
        """Save personal information to a JSON file in AppData and update clipboard labels."""
        personal_info_data = {label: entry.get() for label, entry in self.personal_info_entries.items()}

        try:
            with open(PERSONAL_INFO_FILE, "w") as file:
                json.dump(personal_info_data, file, indent=4)
            print("Personal information saved successfully.")
        except Exception as e:
            print(f"Error saving personal information: {e}")
            messagebox.showerror("Error", f"Failed to save personal information: {e}")
            return

    #Setup Settings
    def start_move(self, event):
        self.xwin = self.winfo_x()
        self.ywin = self.winfo_y()
        self.startx = event.x_root
        self.starty = event.y_root

    def do_move(self, event):
        deltax = event.x_root - self.startx
        deltay = event.y_root - self.starty
        x = self.xwin + deltax
        y = self.ywin + deltay
        self.geometry(f"+{x}+{y}")

    def minimize_window(self):
        # Minimize window
        self.overrideredirect(False)
        self.iconify()

    def on_close(self):
        self.destroy()

    def create_custom_menu_bar(self):
        """Create a custom menu bar with theme-aware styling."""
        # Main menu bar frame
        self.menu_bar = tk.Frame(self, bg=self.menu_bg_color, height=25)
        self.menu_bar.pack(side='top', fill='x')

        # Settings button
        self.settings_button = tk.Menubutton(
            self.menu_bar,
            text='⚙️',
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color,
            relief='flat',
            padx=10
        )
        self.settings_button.pack(side='left')

        # Settings menu with 'Configuration' and 'Toggle Theme'
        self.settings_menu = tk.Menu(
            self.settings_button,
            tearoff=0,
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color
        )
        self.settings_menu.add_command(label="Applications File", command=self.open_applications_config_dialog)
        self.settings_menu.add_command(label="Google Sync", command=self.open_settings_dialog)
        self.settings_menu.add_command(label='Switch Theme', command=self.toggle_theme)
        self.settings_button.config(menu=self.settings_menu)

        # Google Sync Toggle Checkbutton next to the settings button
        self.google_sync_var = tk.BooleanVar(value=self.sync_to_google)
        self.google_sync_checkbutton = tk.Checkbutton(
            self.menu_bar,
            text="Enable Google Sync",
            variable=self.google_sync_var,
            command=self.toggle_sync,
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color,
            indicatoron=True,
            selectcolor=self.menu_active_bg,
            relief="flat",
            padx=10
        )
        self.google_sync_checkbutton.pack(side='left')

    def toggle_settings_menu(self, event=None):
        """Toggle the visibility of the settings menu."""
        if self.menu_visible:
            self.settings_menu.unpost()  # Hide the menu if it is already open
        else:
            self.settings_menu.post(self.settings_button.winfo_rootx(),
                                    self.settings_button.winfo_rooty() + self.settings_button.winfo_height())
        self.menu_visible = not self.menu_visible  # Toggle the visibility state

    def toggle_sync(self):
        """Toggle the Google Sync setting and update app_config.json accordingly."""
        self.sync_to_google = self.google_sync_var.get()
        print(f"Sync to Google Sheets: {'Enabled' if self.sync_to_google else 'Disabled'}")

        # Update configuration
        self.update_config(ENABLE_GOOGLE_SYNC=self.sync_to_google, theme="Dark" if self.is_dark_mode else "Light")

        # Cancel any existing scheduled sync
        if self.sync_task is not None:
            self.after_cancel(self.sync_task)
            self.sync_task = None

        # Schedule the sync task
        self.schedule_sync()

        # If enabling sync, perform an immediate sync
        if self.sync_to_google:
            self.sync_to_google_sheets()
        else:
            print("Google Sync is disabled. Skipping initial sync.")

    def apply_theme(self):
        """Apply the selected theme and update the menu bar."""
        if self.is_dark_mode:
            self.set_dark_mode()
        else:
            self.set_light_mode()

        # Update all widgets with the new theme
        self.update_all_widgets_theme(self)
        # Update the menu bar specifically
        self.update_menu_bar_theme()

    def set_dark_mode(self):
        self.bg_color = "#2E2E2E"
        self.fg_color = "#FFFFFF"
        self.entry_bg_color = "#3A3A3A"
        self.entry_fg_color = "#FFFFFF"
        self.button_bg_color = "#3E3E3E"
        self.menu_bg_color = "#3E3E3E"
        self.menu_fg_color = "#FFFFFF"
        self.menu_active_bg = "#5E5E5E"

        style = ttk.Style()
        style.theme_use("alt")  # Use 'alt' theme for better customization

        # Configure styles for ttk widgets
        style.configure("TLabel", background=self.bg_color, foreground=self.fg_color)
        style.configure("TFrame", background=self.bg_color)
        style.configure("TButton", background=self.button_bg_color, foreground=self.fg_color)
        style.configure("TEntry", fieldbackground=self.entry_bg_color, foreground=self.entry_fg_color)
        style.configure("Treeview", background=self.entry_bg_color, foreground=self.entry_fg_color,
                        fieldbackground=self.entry_bg_color)
        style.map('Treeview', background=[('selected', '#6A6A6A')], foreground=[('selected', '#FFFFFF')])
        style.configure("Treeview.Heading", background=self.button_bg_color, foreground=self.fg_color)
        style.configure("TNotebook", background=self.bg_color)
        style.configure("TNotebook.Tab", background=self.bg_color, foreground=self.fg_color)
        style.map("TNotebook.Tab", background=[('selected', self.entry_bg_color)])
        style.configure("TFrame", background=self.bg_color)
        style.configure("TCombobox", fieldbackground=self.entry_bg_color, background=self.entry_bg_color,
                        foreground=self.entry_fg_color)
        style.map('TCombobox', fieldbackground=[('readonly', self.entry_bg_color)],
                  background=[('readonly', self.entry_bg_color)],
                  foreground=[('readonly', self.entry_fg_color)])

        # Custom style for Save buttons
        style.configure(
            "Custom.TButton",
            background=self.button_bg_color,
            foreground=self.fg_color,
            borderwidth=1,
            focusthickness=3,
            focuscolor='none',
            font=('TkDefaultFont', 12),
            padding=(10, 5)
        )
        style.map(
            "Custom.TButton",
            background=[('active', self.menu_active_bg)],
            foreground=[('active', self.fg_color)]
        )

    def set_light_mode(self):
        self.bg_color = "#F0F0F0"
        self.fg_color = "#000000"
        self.entry_bg_color = "#FFFFFF"
        self.entry_fg_color = "#000000"
        self.button_bg_color = "#E0E0E0"
        self.menu_bg_color = "#E0E0E0"
        self.menu_fg_color = "#000000"
        self.menu_active_bg = "#C0C0C0"

        style = ttk.Style()
        style.theme_use("alt")  # Use 'alt' theme for better customization

        # Configure styles for ttk widgets
        style.configure("TLabel", background=self.bg_color, foreground=self.fg_color)
        style.configure("TFrame", background=self.bg_color)
        style.configure("TButton", background=self.button_bg_color, foreground=self.fg_color)
        style.configure("TEntry", fieldbackground=self.entry_bg_color, foreground=self.entry_fg_color)
        style.configure("Treeview", background=self.entry_bg_color, foreground=self.entry_fg_color,
                        fieldbackground=self.entry_bg_color)
        style.map('Treeview', background=[('selected', '#D9D9D9')], foreground=[('selected', '#000000')])
        style.configure("Treeview.Heading", background=self.button_bg_color, foreground=self.fg_color)
        style.configure("TNotebook", background=self.bg_color)
        style.configure("TNotebook.Tab", background=self.bg_color, foreground=self.fg_color)
        style.map("TNotebook.Tab", background=[('selected', self.entry_bg_color)])
        style.configure("TFrame", background=self.bg_color)
        style.configure("TCombobox", fieldbackground=self.entry_bg_color, background=self.entry_bg_color,
                        foreground=self.entry_fg_color)
        style.map('TCombobox', fieldbackground=[('readonly', self.entry_bg_color)],
                  background=[('readonly', self.entry_bg_color)],
                  foreground=[('readonly', self.entry_fg_color)])

        # Custom style for Save buttons
        style.configure(
            "Custom.TButton",
            background=self.button_bg_color,
            foreground=self.fg_color,
            borderwidth=1,
            focusthickness=3,
            focuscolor='none',
            font=('TkDefaultFont', 12),
            padding=(10, 5)
        )
        style.map(
            "Custom.TButton",
            background=[('active', self.menu_active_bg)],
            foreground=[('active', self.fg_color)]
        )

    def update_all_widgets_theme(self, widget):
        """Recursively update the theme for all widgets."""
        for child in widget.winfo_children():
            # Check if the widget is a ttk widget
            if isinstance(child, ttk.Widget):
                pass  # Skip ttk widgets; they're styled via ttk.Style
            else:
                # Get the list of options the widget supports
                options = child.keys()
                # Set 'bg' or 'background' if supported
                if 'bg' in options or 'background' in options:
                    child.config(bg=self.bg_color)
                # Set 'fg' or 'foreground' if supported
                if 'fg' in options or 'foreground' in options:
                    child.config(fg=self.fg_color)

                # Specific adjustments for certain widget types
                if isinstance(child, tk.Entry):
                    if 'bg' in options or 'background' in options:
                        child.config(bg=self.entry_bg_color)
                    if 'fg' in options or 'foreground' in options:
                        child.config(fg=self.entry_fg_color)
                elif isinstance(child, tk.Button):
                    if 'activebackground' in options:
                        child.config(activebackground=self.button_bg_color)
                    if 'activeforeground' in options:
                        child.config(activeforeground=self.fg_color)
                elif isinstance(child, tk.Text):
                    if 'bg' in options or 'background' in options:
                        child.config(bg=self.entry_bg_color)
                    if 'fg' in options or 'foreground' in options:
                        child.config(fg=self.entry_fg_color)
                # For frames and toplevels, only set 'bg'
                elif isinstance(child, (tk.Frame, tk.Toplevel)):
                    if 'bg' in options or 'background' in options:
                        child.config(bg=self.bg_color)
            # Recursively update child widgets
            self.update_all_widgets_theme(child)

    def update_menu_bar_theme(self):
        """Update theme for the menu bar components (Settings and Google Sync Checkbutton)."""
        # Update menu bar background color
        self.menu_bar.config(bg=self.menu_bg_color)

        # Update Settings button
        self.settings_button.config(
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color
        )

        # Update Settings menu items to match theme
        self.settings_menu.config(
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color
        )
        for index in range(self.settings_menu.index('end') + 1):
            self.settings_menu.entryconfig(
                index,
                background=self.menu_bg_color,
                foreground=self.menu_fg_color,
                activebackground=self.menu_active_bg,
                activeforeground=self.menu_fg_color
            )

        # Update Google Sync Checkbutton
        self.google_sync_checkbutton.config(
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color,
            selectcolor=self.menu_active_bg  # Match checkbox color with theme
        )

    def toggle_theme(self):
        """Toggle between Dark and Light themes and save to app_config.json in AppData."""
        self.is_dark_mode = not self.is_dark_mode
        theme = "Dark" if self.is_dark_mode else "Light"
        self.apply_theme()

        # Call save_theme from settings_manager to save the theme to the AppData config
        save_theme(theme)  # This function now manages saving only in AppData

    def bind_events_to_children(self, parent_widget, click_handler, drop_handler=None):
        """
        Bind click and drag-and-drop events to all child widgets within the parent_widget,
        excluding interactive widgets.

        Parameters:
        - parent_widget: The frame whose child widgets will have events bound.
        - click_handler: The method to handle click events.
        - drop_handler: The method to handle drop events (optional).
        """
        widgets = parent_widget.winfo_children()
        for widget in widgets:
            # Exclude interactive widgets to allow normal user interaction
            if isinstance(widget, (tk.Entry, tk.Text, ttk.Entry, ttk.Combobox)):
                continue

            # Bind the click event
            widget.bind("<Button-1>", click_handler)

            # Register drag-and-drop if a drop_handler is provided
            if drop_handler:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', drop_handler)

            # Recursively bind events to child widgets
            self.bind_events_to_children(widget, click_handler, drop_handler)

    def select_app_file(self, event=None):
        """Prompt user to select Applications.xlsx, and copy it to AppData."""
        file_path = filedialog.askopenfilename(title="Select Applications.xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                shutil.copy(file_path, DATA_FILE_PATH)
                self.app_file_path_var.set(DATA_FILE_PATH)
                print(f"[DEBUG] Applications.xlsx copied to {DATA_FILE_PATH}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy Applications.xlsx: {e}")

    def select_service_account_file(self, event=None):
        """Prompt user to select Service Account JSON, and copy it to AppData."""
        file_path = filedialog.askopenfilename(title="Select Service Account JSON",
                                               filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                shutil.copy(file_path, SERVICE_ACCOUNT_FILE)
                self.service_account_file_path_var.set(SERVICE_ACCOUNT_FILE)
                print(f"[DEBUG] Service Account JSON copied to {SERVICE_ACCOUNT_FILE}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy Service Account JSON: {e}")

    def service_account_file_drop(self, event):
        """Handle the drop event for the Service Account JSON file."""
        print("[DEBUG] Service Account JSON file drop detected.")
        file_path = event.data
        file_list = self.tk.splitlist(file_path)
        if file_list:
            file_path = file_list[0]
            # Validate the file type
            if file_path.lower().endswith('.json'):
                try:
                    shutil.copy(file_path, SERVICE_ACCOUNT_FILE)
                    self.service_account_file_path_var.set(SERVICE_ACCOUNT_FILE)
                    print(f"[DEBUG] Service Account JSON copied to {SERVICE_ACCOUNT_FILE}")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to copy Service Account JSON: {e}")
            else:
                messagebox.showerror("Invalid File", "Please drop a valid JSON file.")

    def app_file_drop(self, event):
        """Handle the drop event for the Applications.xlsx file."""
        print("[DEBUG] Applications.xlsx file drop detected.")
        file_path = event.data
        file_list = self.tk.splitlist(file_path)
        if file_list:
            file_path = file_list[0]
            # Validate the file type
            if file_path.lower().endswith('.xlsx'):
                try:
                    shutil.copy(file_path, DATA_FILE_PATH)
                    self.app_file_path_var.set(DATA_FILE_PATH)
                    print(f"[DEBUG] Applications.xlsx copied to {DATA_FILE_PATH}")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to copy Applications.xlsx: {e}")
            else:
                messagebox.showerror("Invalid File", "Please drop a valid Excel (.xlsx) file.")

    def open_settings_dialog(self):
        """Open a dialog to configure the Service Account JSON and Spreadsheet ID."""
        dialog = tk.Toplevel(self)
        dialog.title("Google Sync Configuration")
        dialog.geometry("600x300")  # Adjusted size for larger content

        if hasattr(self, 'google_sync_icon') and self.google_sync_icon:
            dialog.iconphoto(False, self.google_sync_icon)
        # Keep settings dialog in front
        dialog.transient(self)
        dialog.grab_set()
        dialog.focus()

        # Use current theme colors
        bg = self.bg_color
        fg = self.fg_color
        entry_bg = self.entry_bg_color
        entry_fg = self.entry_fg_color
        button_bg = self.button_bg_color

        dialog.config(bg=bg)

        # --- Service Account JSON file ---
        self.service_account_file_path_var = tk.StringVar(value=self.get_current_service_account_file_path())
        self.service_file_button = tk.Frame(dialog, bg=button_bg, relief='raised', bd=2)
        self.service_file_button.pack(pady=20, fill='x', padx=50)

        # Bind the entire frame for the service account button
        self.service_file_button.bind("<Enter>", lambda e: self.service_file_button.config(relief='groove'))
        self.service_file_button.bind("<Leave>", lambda e: self.service_file_button.config(relief='raised'))
        self.service_file_button.bind("<Button-1>", lambda e: self.select_service_account_file())

        # Use tkinterdnd2 methods for drag-and-drop
        self.service_file_button.drop_target_register(DND_FILES)
        self.service_file_button.dnd_bind('<<Drop>>', self.service_account_file_drop)

        # Place the icon and labels inside the frame for service account file
        self.service_icon_label = tk.Label(self.service_file_button, image=self.upload_json_icon, bg=button_bg)
        self.service_icon_label.pack(side='left', padx=(10, 5), pady=10)

        # Text label and path display for the service account file
        self.service_text_frame = tk.Frame(self.service_file_button, bg=button_bg)
        self.service_text_frame.pack(side='left', fill='x', expand=True)

        self.service_text_label = tk.Label(
            self.service_text_frame,
            text="Upload or Drop Service Account JSON File Here",
            bg=button_bg,
            fg=fg,
            font=('TkDefaultFont', 12)
        )
        self.service_text_label.pack(anchor='w', pady=(5, 0))

        self.service_file_label = tk.Label(
            self.service_text_frame,
            textvariable=self.service_account_file_path_var,
            bg=button_bg,
            fg=fg,
            wraplength=500,
            justify='left',
            font=('TkDefaultFont', 8)
        )
        self.service_file_label.pack(anchor='w', pady=(0, 5))

        # Bind events to all child widgets within service_file_button
        self.bind_events_to_children(
            self.service_file_button,
            self.select_service_account_file,
            self.service_account_file_drop
        )

        # --- Google Sheets Spreadsheet ID ---
        self.sheets_id_var = tk.StringVar(value=self.get_current_spreadsheet_id())

        self.sheets_id_button = tk.Frame(
            dialog,
            bg=button_bg,
            relief='raised',
            bd=2
        )
        self.sheets_id_button.pack(pady=10, fill='x', padx=50)

        self.sheets_id_button.bind("<Enter>", lambda e: self.sheets_id_button.config(relief='groove'))
        self.sheets_id_button.bind("<Leave>", lambda e: self.sheets_id_button.config(relief='raised'))
        self.sheets_id_button.bind("<Button-1>", lambda e: self.sheets_id_entry.focus_set())

        # Place the icon inside the button frame
        self.sheets_id_icon_label = tk.Label(
            self.sheets_id_button,
            image=self.upload_sheets_id_icon,
            bg=button_bg
        )
        self.sheets_id_icon_label.pack(side='left', padx=(20, 5), pady=15)

        # Create a frame for the label and entry widget
        self.sheets_text_frame = tk.Frame(
            self.sheets_id_button,
            bg=button_bg
        )
        self.sheets_text_frame.pack(side='left', pady=10)

        # Place the label and entry widget side by side
        self.sheets_id_text_label = tk.Label(
            self.sheets_text_frame,
            text="Spreadsheet ID:",
            bg=button_bg,
            fg=fg,
            font=('TkDefaultFont', 12)
        )
        self.sheets_id_text_label.pack(side='left', padx=(0, 5))

        self.sheets_id_entry = tk.Entry(
            self.sheets_text_frame,
            textvariable=self.sheets_id_var,
            bg=entry_bg,
            fg=entry_fg,
            font=('TkDefaultFont', 8),
            bd=1,
            highlightthickness=2.5,
            relief='sunken',
            width=45
        )
        self.sheets_id_entry.pack(side='left')

        # Bind events to all child widgets within sheets_id_button
        self.bind_events_to_children(
            self.sheets_id_button,
            lambda e: self.sheets_id_entry.focus_set(),
            None  # Assuming no drag-and-drop for Spreadsheet ID
        )

        # Additionally, bind events directly to the frame to cover any gaps
        self.sheets_id_button.bind("<Button-1>", lambda e: self.sheets_id_entry.focus_set())
        # If you have a drop handler for spreadsheet ID, bind it here
        # self.sheets_id_button.drop_target_register(DND_FILES)
        # self.sheets_id_button.dnd_bind('<<Drop>>', self.sheets_id_drop_handler)

        # Save Changes button using ttk.Button with custom style
        ttk.Button(
            dialog,
            text="Save Changes",
            command=lambda: self.save_settings(dialog),
            style="Custom.TButton"
        ).pack(pady=20)

    def open_applications_config_dialog(self):
        """Open a dialog to configure the Applications.xlsx file."""
        dialog = tk.Toplevel(self)
        dialog.title("Applications File Configuration")
        dialog.geometry("600x200")  # Adjusted size for content

        if hasattr(self, 'applications_icon') and self.applications_icon:
            dialog.iconphoto(False, self.applications_icon)

        # Keep settings dialog in front
        dialog.transient(self)
        dialog.grab_set()
        dialog.focus()

        # Set dialog theme based on current mode
        bg = self.bg_color
        fg = self.fg_color
        entry_bg = self.entry_bg_color
        entry_fg = self.entry_fg_color
        button_bg = self.button_bg_color

        dialog.config(bg=bg)

        # --- Applications.xlsx file ---
        self.app_file_path_var = tk.StringVar(value=self.get_current_applications_file_path())
        self.app_file_button = tk.Frame(dialog, bg=button_bg, relief='raised', bd=2)
        self.app_file_button.pack(pady=20, fill='x', padx=50)

        # Bind the frame directly to trigger file selection and drag-and-drop
        self.app_file_button.bind("<Enter>", lambda e: self.app_file_button.config(relief='groove'))
        self.app_file_button.bind("<Leave>", lambda e: self.app_file_button.config(relief='raised'))
        self.app_file_button.bind("<Button-1>", lambda e: self.select_app_file())

        # Use tkinterdnd2 methods for drag-and-drop
        self.app_file_button.drop_target_register(DND_FILES)
        self.app_file_button.dnd_bind('<<Drop>>', self.app_file_drop)

        # Place the icon and labels inside the frame
        self.app_icon_label = tk.Label(self.app_file_button, image=self.upload_xlsx_icon, bg=button_bg)
        self.app_icon_label.pack(side='left', padx=(10, 5), pady=10)

        # Create a frame for the text labels
        self.app_text_frame = tk.Frame(self.app_file_button, bg=button_bg)
        self.app_text_frame.pack(side='left', fill='x', expand=True)

        # Text label for instructions
        self.app_text_label = tk.Label(
            self.app_text_frame,
            text="Upload or Drop Applications.xlsx File Here",
            bg=button_bg,
            fg=fg,
            font=('TkDefaultFont', 12)
        )
        self.app_text_label.pack(anchor='w', pady=(5, 0))

        # Path label
        self.app_file_label = tk.Label(
            self.app_text_frame,
            textvariable=self.app_file_path_var,
            bg=button_bg,
            fg=fg,
            wraplength=500,
            justify='left',
            font=('TkDefaultFont', 8)
        )
        self.app_file_label.pack(anchor='w', pady=(0, 5))

        # Bind events to all child widgets within app_file_button
        self.bind_events_to_children(
            self.app_file_button,
            self.select_app_file,
            self.app_file_drop
        )

        # Save Changes button using ttk.Button with custom style
        ttk.Button(
            dialog,
            text="Save Changes",
            command=lambda: self.save_applications_settings(dialog),
            style="Custom.TButton"
        ).pack(pady=20)

    def save_settings(self, dialog):
        """Save settings related to Google Sync and close the dialog."""
        # Retrieve values from the UI
        service_account_path = self.service_account_file_path_var.get()
        spreadsheet_id = self.sheets_id_var.get()

        # Validate inputs
        if not os.path.isfile(service_account_path):
            messagebox.showerror("Error", "Service Account JSON file does not exist.")
            return
        if not spreadsheet_id.strip():
            messagebox.showerror("Error", "Spreadsheet ID cannot be empty.")
            return

        # Update configuration (Assuming you have a method or mechanism to handle configurations)
        try:
            # Example: Update a configuration dictionary or write to a config file
            config = {
                "SERVICE_ACCOUNT_FILE": service_account_path,
                "SPREADSHEET_ID": spreadsheet_id
            }
            with open('config.json', 'w') as config_file:
                json.dump(config, config_file, indent=4)
            print("[DEBUG] Settings saved successfully.")
            messagebox.showinfo("Success", "Settings have been saved successfully.")
            dialog.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {e}")

    def save_applications_settings(self, dialog):
        """Save settings related to Applications.xlsx and close the dialog."""
        # Retrieve values from the UI
        applications_path = self.app_file_path_var.get()

        # Validate inputs
        if not os.path.isfile(applications_path):
            messagebox.showerror("Error", "Applications.xlsx file does not exist.")
            return

        # Update configuration (Assuming you have a method or mechanism to handle configurations)
        try:
            # Example: Update a configuration dictionary or write to a config file
            config = {
                "DATA_FILE_PATH": applications_path
            }
            with open('config.json', 'w') as config_file:
                json.dump(config, config_file, indent=4)
            print("[DEBUG] Applications settings saved successfully.")
            messagebox.showinfo("Success", "Applications settings have been saved successfully.")
            dialog.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Applications settings: {e}")

    def get_current_google_sync_setting(self):
        """Retrieve the current ENABLE_GOOGLE_SYNC setting from settings_manager.py."""
        try:
            return ENABLE_GOOGLE_SYNC
        except ImportError:
            return False  # Default to False if not set

    # Additional methods for file selection, Data handling, and layout setup

    def get_current_applications_file_path(self):
        """Retrieve the current Applications.xlsx file path from settings_manager.py."""
        try:
            from config.settings_manager import DATA_FILE_PATH
            if os.path.isfile(DATA_FILE_PATH):
                return os.path.abspath(DATA_FILE_PATH)
            else:
                print("DATA_FILE_PATH does not point to an existing file.")
                return "No file selected"
        except Exception as e:
            print(f"Error retrieving DATA_FILE_PATH: {e}")
            return "No file selected"

    def get_current_service_account_file_path(self):
        """Retrieve the current Service Account JSON file path from settings_manager.py or show 'No file selected'."""
        try:
            # Check if the path actually exists
            if os.path.isfile(SERVICE_ACCOUNT_FILE):
                return os.path.abspath(SERVICE_ACCOUNT_FILE)
            else:
                print("SERVICE_ACCOUNT_FILE does not point to an existing file.")
                return "No file selected"
        except ImportError:
            return "No file selected"
        except Exception as e:
            print(f"Error retrieving SERVICE_ACCOUNT_FILE: {e}")
            return "No file selected"

    def get_current_spreadsheet_id(self):
        """Retrieve the current Spreadsheet ID from config.settings."""
        try:
            return SPREADSHEET_ID
        except ImportError:
            print("Error: Could not import SPREADSHEET_ID from config.settings.")
            return ""
        except Exception as e:
            print(f"Error retrieving SPREADSHEET_ID: {e}")
            return ""

    def update_config(self, **kwargs):
        """Update configuration settings in app_config.json."""
        try:
            # Load the current configuration
            with open(CONFIG_JSON_PATH, "r") as config_file:
                config = json.load(config_file)
        except (FileNotFoundError, json.JSONDecodeError):
            # If the file does not exist or is corrupted, start with default configuration
            config = default_config

        # Update the configuration with new values
        config.update(kwargs)

        # Save the updated configuration back to app_config.json
        with open(CONFIG_JSON_PATH, "w") as config_file:
            json.dump(config, config_file, indent=4)
        print("[DEBUG] Configuration updated in app_config.json.")

    def reload_configurations(self):
        """Reload configurations from app_config.json."""
        config_json_path = os.path.join(base_path, "config", "app_config.json")
        try:
            with open(config_json_path, "r") as config_file:
                config = json.load(config_file)

            # Update variables
            self.sync_to_google = config.get("ENABLE_GOOGLE_SYNC", False)
            self.DATA_FILE_PATH = config.get("DATA_FILE_PATH", os.path.join(base_path, "Data", "Applications.xlsx"))
            self.SERVICE_ACCOUNT_FILE = config.get("SERVICE_ACCOUNT_FILE",
                                                   os.path.join(base_path, "config", "service_account.json"))
            self.SPREADSHEET_ID = config.get("SPREADSHEET_ID", "")
            theme = config.get("theme", "Light")

            print(f"[DEBUG] Reloading configurations. Theme found: {theme}")

            # Update theme
            self.is_dark_mode = True if theme.lower() == "dark" else False
            print(f"[DEBUG] is_dark_mode set to: {self.is_dark_mode}")
            self.apply_theme()

            # Re-read the Excel file with the updated path
            try:
                self.applications_df = read_applications_from_excel(self.DATA_FILE_PATH)
                self.populate_treeview(self.applications_df)
                print("[DEBUG] Applications Data reloaded successfully.")
            except Exception as e:
                print(f"[ERROR] Could not read the Excel file after reloading configurations: {e}")
                self.applications_df = pd.DataFrame()
                self.populate_treeview(self.applications_df)

            # Re-establish Google Sync if enabled
            if self.sync_to_google:
                self.sync_to_google_sheets()
                self.schedule_sync()

            print("[DEBUG] Configurations reloaded successfully.")
        except Exception as e:
            print(f"[ERROR] Failed to reload configurations: {e}")
            messagebox.showerror("Error", f"Failed to reload configurations: {e}")

    def update_google_sync_setting(self, enable_google_sync):
        """
        Update ENABLE_GOOGLE_SYNC setting in app_config.json.
        """
        try:
            # Load current config
            with open(self.config_file_path, "r") as file:
                config = json.load(file)

            # Update ENABLE_GOOGLE_SYNC value
            config["ENABLE_GOOGLE_SYNC"] = enable_google_sync

            # Write updated config back to file
            with open(self.config_file_path, "w") as file:
                json.dump(config, file, indent=4)

            print(f"Google Sync setting updated to: {enable_google_sync}")
        except Exception as e:
            print(f"Error updating Google Sync setting: {e}")


if __name__ == "__main__":
    app = AppTrackPro()
    app.mainloop()


