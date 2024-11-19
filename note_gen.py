from tkinter import filedialog
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import csv
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox

# Function to generate and display the note in real time
def update_note(*args):
    notes_entry.delete("1.0", tk.END)

    work_order = work_order_entry.get().strip()
    site_contact = site_contact_entry.get().strip().title()  # Ensuring title case
    new_priority = priority_var.get().strip()
    contractor = contractor_entry.get().strip().title()      # Ensuring title case
    quick_note = quick_notes_var.get()
    additional_note = additional_note_entry.get("1.0", tk.END).strip()

    generated_note = ""
    
    if work_order:
        generated_note += f"Work Order:     {work_order}\n"
    if new_priority:
        generated_note += f"New Priority:   {new_priority}\n"
    if site_contact:
        generated_note += f"Site Contact:   {site_contact}\n"
    if contractor:
        generated_note += f"Contractor:     {contractor}\n"

    actions = {
        "Contacted": contacted_var.get(),
        "Dispatched": dispatched_var.get(),
        "Emailed": emailed_var.get(),
        "Confirmed": confirmed_var.get(),
        "Cancelled": cancelled_var.get(),
        "POC Callback": poc_callback_var.get(),
    }

    actions_added = False
    for action, is_checked in actions.items():
        if is_checked:
            if not actions_added:
                generated_note += "\nAction:\n"
                actions_added = True
            generated_note += f"   - {action}:\n"

    if quick_note:
        generated_note += f"\nNotes:\n  - {quick_note}\n"

    if additional_note:
        if not quick_note:
            generated_note += "\nNotes:\n"
        generated_note += f"{additional_note}\n"

    notes_entry.insert("1.0", generated_note)

    # Enable save button when any changes are made
    save_button.config(state=tk.NORMAL)

# Function to copy the note text to clipboard
def copy_to_clipboard():
    root.clipboard_clear()
    root.clipboard_append(notes_entry.get("1.0", tk.END).strip())
    root.update()

# Function to clear all fields
def clear_all():
    work_order_entry.delete(0, tk.END)
    site_contact_entry.delete(0, tk.END)
    contractor_entry.delete(0, tk.END)
    priority_var.set("")
    quick_notes_var.set("")
    additional_note_entry.delete("1.0", tk.END)
    notes_entry.delete("1.0", tk.END)
    for var in [contacted_var, dispatched_var, emailed_var, confirmed_var, cancelled_var, poc_callback_var]:
        var.set(False)
    hide_success_label()  # Hide the success label when clearing all fields
    save_button.config(state=tk.DISABLED)  # Disable save button initially

def validate_inputs():
    work_order = work_order_entry.get().strip()
    if not work_order:  # Check if the Work Order field is empty
        messagebox.showerror("Input Error", "The 'Work Order' field is mandatory.")
        return False
    return True

# Function to display or hide the priority change reason label
def toggle_priority_reason_label(*args):
    current_priority = priority_var.get()
    if current_priority in ["P1 Emergency", "P2 Immediate", "P3 Urgent", "P3.5 Escalated Routine", "P4 Routine", "P5 Specific Date"]:
        priority_reason_label.place(x=120, y=65)  # Adjust position to suit layout
    elif current_priority == "":  # If the dropdown is cleared or empty
        priority_reason_label.place_forget()  # Hide the label


# Function to save note and display confirmation
def save_note(event=None):  # Accept an optional event parameter for hotkey binding
    if not validate_inputs():
        return  # Exit the function if validation fails

    # Extract data for CSV
    work_order = work_order_entry.get().strip()
    site_contact = site_contact_entry.get().strip().title()
    new_priority = priority_var.get().strip()
    contractor = contractor_entry.get().strip().title()
    quick_note = quick_notes_var.get()
    additional_note = additional_note_entry.get("1.0", tk.END).strip()
    actions = []
    if contacted_var.get(): actions.append("Contacted")
    if dispatched_var.get(): actions.append("Dispatched")
    if emailed_var.get(): actions.append("Emailed")
    if confirmed_var.get(): actions.append("Confirmed")
    if cancelled_var.get(): actions.append("Cancelled")
    if poc_callback_var.get(): actions.append("POC Callback")
    actions_str = ", ".join(actions)

    # Save data to CSV
    file_exists = os.path.isfile("notegenhistory.csv")
    with open("notegenhistory.csv", "a", newline='') as csvfile:
        fieldnames = ["Date Time", "Work Order", "Site Contact", "New Priority", "Contractor", "Actions", "Notes",]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerow({
            "Date Time": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Work Order": work_order,
            "Site Contact": site_contact,
            "New Priority": new_priority,
            "Contractor": contractor,
            "Actions": actions_str,
            "Notes": f"{quick_note} {additional_note}".strip()
        })

    messagebox.showinfo("Save Confirmation", "Note saved to notegenhistory.csv successfully.")
    show_success_label()
    save_button.config(state=tk.DISABLED)  # Disable save button until changes are made

# Function to show "Successfully Saved" label
def show_success_label():
    timestamp = datetime.now().strftime("   -   %d-%m-%Y %H:%M")
    success_label.config(text=f"Successfully Saved{timestamp}")
    success_label.place(relx=1.0, rely=1.0, anchor="se")
    root.after(20000, hide_success_label)  # Hide after 3 seconds

# Function to hide the "Successfully Saved" label
def hide_success_label():
    success_label.place_forget()

# Function to load history data into the Treeview
def load_history():
    global history_tree, selected_item_details_text, history_count_label

    # Clear existing items in the Treeview
    for item in history_tree.get_children():
        history_tree.delete(item)

    # Set Treeview column headings without Actions and Notes
    if os.path.exists("notegenhistory.csv"):
        with open("notegenhistory.csv", "r") as csvfile:
            csvreader = csv.DictReader(csvfile)
            headers = csvreader.fieldnames

            if headers:
                # Exclude Actions and Notes from the columns to display
                display_headers = [header for header in headers if header not in ["Actions", "Notes"]]

                # Set Treeview columns
                history_tree["columns"] = display_headers
                for header in display_headers:
                    history_tree.heading(header, text=header)
                    history_tree.column(header, anchor='center', width=100)

                # Read data into a list and sort it by date in descending order
                rows = list(csvreader)
                rows = sorted(rows, key=lambda x: datetime.strptime(x["Date Time"], "%d/%m/%Y %H:%M"), reverse=True)

                # Insert rows into the Treeview without Actions and Notes
                for row in rows:
                    values = [row[header] for header in display_headers]
                    history_tree.insert("", "end", values=values)

    # Update the history count label
    history_count_label.config(text=f"Total Records: {len(history_tree.get_children())}")

# Function to display details of the selected item with exact format
def show_selected_item_details(event):
    selected_item = history_tree.focus()  # Get the selected item in the Treeview
    if selected_item:
        item_values = history_tree.item(selected_item, "values")
        headers = history_tree["columns"]

        # Extract the values that are displayed in the Treeview
        work_order = item_values[headers.index("Work Order")].strip() if "Work Order" in headers else ""
        new_priority = item_values[headers.index("New Priority")].strip() if "New Priority" in headers else ""
        site_contact = item_values[headers.index("Site Contact")].strip() if "Site Contact" in headers else ""
        contractor = item_values[headers.index("Contractor")].strip() if "Contractor" in headers else ""

        # Load Actions and Notes separately from the CSV since they are not displayed in Treeview
        actions = ""
        notes = ""
        if work_order:  # Proceed only if there's a valid work order to match
            with open("notegenhistory.csv", "r") as csvfile:
                csvreader = csv.DictReader(csvfile)
                for row in csvreader:
                    if row["Work Order"].strip() == work_order:
                        actions = row.get("Actions", "").strip()
                        notes = row.get("Notes", "").strip()
                        break

        # Start formatting details, only include non-empty values
        formatted_details = ""

        if work_order:
            formatted_details += f"Work Order:     {work_order}\n"
        if new_priority:
            formatted_details += f"New Priority:   {new_priority}\n"
        if site_contact:
            formatted_details += f"Site Contact:   {site_contact}\n"
        if contractor:
            formatted_details += f"Contractor:     {contractor}\n"

        if actions:
            action_items = actions.split(", ")
            formatted_details += "\nAction:\n"
            for action in ["Contacted", "Dispatched", "Emailed", "Confirmed", "Cancelled", "POC Callback"]:
                if action in action_items:
                    formatted_details += f"   - {action}:\n"

        if notes:
            formatted_details += f"\nNotes:\n  - {notes}\n"

        # Clear existing content and insert formatted details
        selected_item_details_text.delete("1.0", tk.END)
        selected_item_details_text.insert("1.0", formatted_details)

# Variables to track search results and index
search_results = []
current_result_index = -1

# Function to remove previous highlights
def remove_highlights():
    # Iterate through all the rows and remove the highlight tag
    for item in history_tree.get_children():
        history_tree.item(item, tags=())  # Removing all tags

# Modified focus function to highlight and scroll to search results
def focus_on_result():
    # Remove previous highlights
    remove_highlights()

    # Check if there are valid search results to highlight
    if search_results and 0 <= current_result_index < len(search_results):
        # Focus on the current result
        history_tree.selection_remove(history_tree.selection())
        result_item = search_results[current_result_index]
        history_tree.selection_add(result_item)
        history_tree.see(result_item)  # Scroll to the item

        # Highlight the result by adding a tag to it
        history_tree.item(result_item, tags=("highlight",))
        history_tree.tag_configure("highlight", background="yellow")

# Updated search_history() to reset and prepare for highlighting

def search_history():
    global search_var, history_tree, search_results, current_result_index, history_count_label

    # Get the search term from the search entry box
    search_term = search_var.get().strip().lower()

    # Clear existing items in the Treeview
    for item in history_tree.get_children():
        history_tree.delete(item)

    # Initialize the search results
    search_results = []
    current_result_index = -1

    # Reload the data based on search input
    if os.path.exists("notegenhistory.csv"):
        with open("notegenhistory.csv", "r") as csvfile:
            csvreader = csv.DictReader(csvfile)
            headers = csvreader.fieldnames

            if headers:
                # Exclude Actions and Notes from the columns to display
                display_headers = [header for header in headers if header not in ["Actions", "Notes"]]

                # Set Treeview columns
                history_tree["columns"] = display_headers
                for header in display_headers:
                    history_tree.heading(header, text=header)
                    history_tree.column(header, anchor='center', width=100)

                # Insert rows into the Treeview that match the search term
                for row in csvreader:
                    if any(search_term in row[header].lower() for header in headers if row[header]):
                        values = [row[header] for header in display_headers]
                        item_id = history_tree.insert("", "end", values=values)
                        search_results.append(item_id)  # Track matching rows

    # Update the history count label
    history_count_label.config(text=f"Total Records: {len(history_tree.get_children())}")

# Function to move to the next search result
def next_result():
    global current_result_index
    if search_results:
        current_result_index = (current_result_index + 1) % len(search_results)
        focus_on_result()

# Function to confirm exit
def confirm_exit():
    if messagebox.askokcancel("Quit",  "    Any unsaved work will be lost \n\n                     Continue?"):
        root.destroy()  # Quit the application if the user confirms

# Root setup
root = tk.Tk()
root.title("NoteGen")
root.geometry("520x780")
root.eval('tk::PlaceWindow . center')
root.resizable(False, False)


# Create menu
menu = tk.Menu(root)
root.config(menu=menu)

fmenu = tk.Menu(root)
root.config(menu=menu)

file_menu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="File", menu=file_menu)

options_menu = tk.Menu(file_menu, tearoff=0)
file_menu.add_cascade(label="Options", menu=options_menu)
options_menu.add_command(label="History", command=lambda: open_history_window())
file_menu.add_separator()
file_menu.add_command(label="Exit", command=confirm_exit)  # Use confirm_exit function

# Bind the close button (Windows X button) to the confirm_exit function
root.protocol("WM_DELETE_WINDOW", confirm_exit)


def open_history_window(event=None):  # Accept an optional event parameter for hotkey binding
    global search_var, history_tree, selected_item_details_text, history_count_label
    history_window = tk.Toplevel(root)
    history_window.title("History")
    history_window.geometry("675x750")
    history_window.transient(root)

    # Create a search entry and button in the history window
    search_frame = ttk.Frame(history_window)
    search_frame.pack(pady=5)

    search_label = ttk.Label(search_frame, text="Search:")
    search_label.pack(side="left", padx=5)

    # Use a StringVar to track the search input in real time
    search_var = tk.StringVar()
    search_var.trace("w", lambda name, index, mode: search_history())

    search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
    search_entry.pack(side="left", padx=5)

    next_button = ttk.Button(search_frame, text="Next", command=next_result)
    next_button.pack(side="left", padx=5)

    # Scrollbars for the history tree widget
    history_tree_frame = ttk.Frame(history_window)
    history_tree_frame.pack(expand=True, fill="x")

    history_scrollbar_y = ttk.Scrollbar(history_tree_frame, orient="vertical")
    history_scrollbar_y.pack(side="right", fill="y")

    history_scrollbar_x = ttk.Scrollbar(history_tree_frame, orient="horizontal")
    history_scrollbar_x.pack(side="bottom", fill="x")

    # Create Treeview widget
    history_tree = ttk.Treeview(history_tree_frame, height=15, yscrollcommand=history_scrollbar_y.set, xscrollcommand=history_scrollbar_x.set)

    # Configure the scrollbars
    history_scrollbar_y.config(command=history_tree.yview)
    history_scrollbar_x.config(command=history_tree.xview)

    history_tree["show"] = "headings"
    history_tree.pack(expand=True, fill="both")

    # Bind selection event to show details
    history_tree.bind("<<TreeviewSelect>>", show_selected_item_details)

    # Create a Text widget to display the details of the selected item
    selected_item_details_frame = ttk.Frame(history_window)
    selected_item_details_frame.pack(expand=True, fill="both", padx=25, pady=15)

    selected_item_details_text = tk.Text(selected_item_details_frame, wrap="word", height=15, width=80)
    selected_item_details_text.pack(expand=True, fill="x")

    # Create a frame to hold the Export, Load buttons, and History Count label
    button_frame = ttk.Frame(history_window)
    button_frame.pack(pady=10)

    # Create an "Export to CSV" button and add it to the button frame
    export_button = ttk.Button(button_frame, text="Export to CSV", command=export_csv)
    export_button.pack(side="left", padx=10, pady=10)

    # Create a "Load History" button and add it to the button frame
    load_button = ttk.Button(button_frame, text="Reload History", command=load_history)
    load_button.pack(side="left", padx=10, pady=10)

    # Create a label to display the count of rows in the history
    history_count_label = ttk.Label(button_frame, text="Total Records: 0")
    history_count_label.pack(side="right", padx=10, pady=10)

    # Load history data as soon as the window opens
    load_history_on_open(history_window)

# Function to load history data when window opens
def load_history_on_open(history_window):
    if os.path.exists("notegenhistory.csv"):
        load_history()  # Load history if CSV exists
    else:
        # If CSV does not exist, notify user and prompt to select a file
        response = messagebox.askyesno("File Not Found", "The history CSV file does not exist. Would you like to choose a location to load a CSV?")
        if response:
            file_path = filedialog.askopenfilename(defaultextension=".csv",
                                                   filetypes=[("CSV files", "*.csv")],
                                                   title="Select CSV File")
            if file_path:
                try:
                    with open(file_path, "r") as csvfile:
                        reader = csv.reader(csvfile)
                        with open("notegenhistory.csv", "w", newline='') as outfile:
                            writer = csv.writer(outfile)
                            for row in reader:
                                writer.writerow(row)
                    load_history()  # Load the CSV after it has been copied

                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred while loading the CSV file: {e}")
            else:
                messagebox.showinfo("Cancelled", "No file selected. History will not be loaded.")
        else:
            messagebox.showinfo("Cancelled", "History will not be loaded.")


# Function to export and save CSV file to selected location
def export_csv():
    # Open a file save dialog to select save location
    file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV files", "*.csv")],
                                             title="Save CSV File")
    if file_path:
        try:
            # Copy the CSV data from the current file to the selected location
            with open("notegenhistory.csv", "r") as infile, open(file_path, "w", newline='') as outfile:
                reader = csv.reader(infile)
                writer = csv.writer(outfile)

                # Write each row from infile to outfile
                for row in reader:
                    writer.writerow(row)

            messagebox.showinfo("Export Successful", f"File has been saved successfully to {file_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while exporting the CSV file: {e}")




# short cuts
# Bind Ctrl+S to save the note
root.bind("<Control-s>", save_note)
# Bind Ctrl+H to open the history window
root.bind("<Control-h>", open_history_window)



# NoteGen Tab
input_frame = ttk.Frame(root, relief="groove", borderwidth=2)
input_frame.place(x=10, y=20, width=500, height=125)

input_frame.grid_columnconfigure(0, weight=1)
input_frame.grid_columnconfigure(1, weight=1)

ttk.Label(input_frame, text="Work Order:", anchor="center").grid(row=0, column=0, sticky="sw", padx=20, pady=5)
work_order_entry = ttk.Entry(input_frame, width=35)
work_order_entry.grid(row=1, column=0, padx=10, pady=2)
work_order_entry.bind("<KeyRelease>", update_note)

ttk.Label(input_frame, text="Site Contact:", anchor="center").grid(row=2, column=0, sticky="sw", padx=20, pady=5)
site_contact_entry = ttk.Entry(input_frame, width=35)
site_contact_entry.grid(row=3, column=0, padx=10, pady=2)
site_contact_entry.bind("<KeyRelease>", update_note)

ttk.Label(input_frame, text="Contractor:", anchor="center").grid(row=2, column=1, sticky="sw", padx=20, pady=5)
contractor_entry = ttk.Entry(input_frame, width=35)
contractor_entry.grid(row=3, column=1, padx=10, pady=2)
contractor_entry.bind("<KeyRelease>", update_note)

ttk.Label(input_frame, text="New Priority?", anchor="center").grid(row=0, column=1, sticky="sw", padx=20, pady=5)
priority_var = tk.StringVar()
priority_var.trace("w", toggle_priority_reason_label)
priority_options = ["", "P1 Emergency", "P2 Immediate", "P3 Urgent", "P3.5 Escalated Routine", "P4 Routine", "P5 Specific Date"]
priority_menu = ttk.Combobox(input_frame, textvariable=priority_var, values=priority_options, width=33, state="readonly")
priority_menu.grid(row=1, column=1, padx=10, pady=2)

side_frame = ttk.Frame(root)
side_frame.place(x=10, y=160, width=560, height=210)

actions_frame = ttk.Frame(side_frame, relief="groove", borderwidth=2)
actions_frame.place(x=0, y=0, width=130, height=210)
ttk.Label(actions_frame, text="Action:", anchor="center").pack(anchor="nw", padx=10, pady=5)

def create_action_row(label_text, var):
    checkbox = ttk.Checkbutton(actions_frame, text=label_text, variable=var, command=update_note)
    checkbox.pack(anchor="sw", padx=10, pady=2)

contacted_var = tk.BooleanVar()
dispatched_var = tk.BooleanVar()
emailed_var = tk.BooleanVar()
confirmed_var = tk.BooleanVar()
cancelled_var = tk.BooleanVar()
poc_callback_var = tk.BooleanVar()

create_action_row("Contacted", contacted_var)
create_action_row("Dispatched", dispatched_var)
create_action_row("Emailed", emailed_var)
create_action_row("Confirmed", confirmed_var)
create_action_row("Cancelled", cancelled_var)
create_action_row("POC Callback", poc_callback_var)

notes_frame = ttk.Frame(side_frame, relief="groove", borderwidth=2)
notes_frame.place(x=140, y=0, width=360, height=210)

quick_notes_var = tk.StringVar()
quick_notes_var.trace_add("write", update_note)
quick_notes_options = [
    "", 
    "Cancellation Reason: No Longer Required",
    "Cancellation Reason: Site Requested",
    "Cancellation Reason: Transport availability",
    "Priority Change Reason: Escalated",
    "Priority Change Reason: Non-urgent Request",
    "Site Contact to Contact Contractor",
    "Site Contact to Dispatch Workorder",
    "Transport NEPT Request Confirmed",
    "Transport Self Drive Request Confirmed",
]

ttk.Label(notes_frame, text="Quick Note:", anchor="w").grid(row=0, column=0, sticky="ew", padx=5, pady=5)
quick_notes_menu = ttk.Combobox(notes_frame, textvariable=quick_notes_var, values=quick_notes_options, width=52, state="readonly")
quick_notes_menu.grid(row=1, column=0, padx=5, pady=5)

ttk.Label(notes_frame, text="Additional Note:", anchor="w").grid(row=2, column=0, sticky="ew", padx=5, pady=5)
additional_note_entry = tk.Text(notes_frame, width=40, height=6, wrap=tk.WORD)
additional_note_entry.grid(row=3, column=0, padx=5, pady=2)
additional_note_entry.bind("<KeyRelease>", update_note)

ttk.Label(root, text="Notes:", anchor="center").place(x=50, y=510)
notes_entry = tk.Text(root, width=68, height=10, wrap=tk.WORD)
notes_entry.place(x=10, y=380, width=500, height=275)

copy_button = ttk.Button(root, text="Copy", command=copy_to_clipboard)
copy_button.place(x=140, y=670)
copy_button_tooltip = ttk.Label(root, text="Copy the generated note to clipboard", background="lightyellow", relief="solid", borderwidth=1)

def show_copy_tooltip(event):
    copy_button_tooltip.place(x=event.x_root - root.winfo_x() - 40, y=event.y_root - root.winfo_y() + 20)

def hide_copy_tooltip(event):
    copy_button_tooltip.place_forget()

copy_button.bind("<Enter>", show_copy_tooltip)
copy_button.bind("<Leave>", hide_copy_tooltip)

clear_button = ttk.Button(root, text="Clear All", command=clear_all)
clear_button.place(x=225, y=670)
clear_button_tooltip = ttk.Label(root, text="Clear all input fields", background="lightyellow", relief="solid", borderwidth=1)

def show_clear_tooltip(event):
    clear_button_tooltip.place(x=event.x_root - root.winfo_x() - 40, y=event.y_root - root.winfo_y() + 20)

def hide_clear_tooltip(event):
    clear_button_tooltip.place_forget()

clear_button.bind("<Enter>", show_clear_tooltip)
clear_button.bind("<Leave>", hide_clear_tooltip)

save_button = ttk.Button(root, text="Save", command=save_note)
save_button.place(x=310, y=670)
save_button.config(state=tk.DISABLED)
save_button_tooltip = ttk.Label(root, text="Save the current note to CSV file", background="lightyellow", relief="solid", borderwidth=1)

def show_save_tooltip(event):
    save_button_tooltip.place(x=event.x_root - root.winfo_x() - 40, y=event.y_root - root.winfo_y() + 20)

def hide_save_tooltip(event):
    save_button_tooltip.place_forget()

save_button.bind("<Enter>", show_save_tooltip)
save_button.bind("<Leave>", hide_save_tooltip)

# Add the priority reason label (initially hidden)
priority_reason_label = ttk.Label(notes_frame, text="*Reason For Priority Change", foreground="dark red")
priority_reason_label.place_forget()  # Hide the label initially

# Success label setup
success_label = ttk.Label(root, text="Successfully Saved", foreground="green")

root.mainloop()

# pyinstaller -w --onefile main.py
