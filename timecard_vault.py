''' Imports '''
import pandas as pd
import pyodbc
import os
import platform
import traceback
import re
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

''' Configurations '''
ACCESS_DB = r"file path"
TABLE_NAME = "Entries"
active_treeviews = []
PDF_contract = ""
LAST_TOTAL_HOURS = 0  # will hold the latest calculated value for PDF use



''' Connect to Access '''
def get_conn():
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={ACCESS_DB};"
    )
    return pyodbc.connect(conn_str)


''' Clean and Upload Data '''
def import_timecard(file_path, sheet_name=None, progress_callback=None):
    # Skip sheets named 'example' (case-insensitive)
    if sheet_name and sheet_name.strip().lower() == "example":
        print(f"Skipping sheet named 'example' in {file_path}")
        return

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)  # Correct: uses specified sheet
        file_name = os.path.basename(file_path)
        file_name = file_name.replace(" ", "")
        
        # before upload check for duplicate file
        if duplicate_file(file_name, sheet_name):
            print(f"Duplicate: {file_name} | {sheet_name} — skipping")
            show_duplicate_warning(f"Duplicate: {file_name} ({sheet_name})")
            return
        
        if not duplicate_file(file_name, sheet_name):
            try:
                # clean column names
                df.columns = df.columns.str.strip()
                df["Totals"] = pd.to_numeric(df["Totals"], errors="coerce")
                df = df.dropna(subset=["Totals"])

                # select and rename relevant columns
                df = df[["Name", "Month", "Year", "Contract Name", "Project Manager", "Totals"]]
                df.columns = ["Name", "Month", "Year", "Contract_Name", "Project_Manager", "Hours"]

                # remove spaces from data
                df["Month"] = df["Month"].astype(str).str.replace(" ", "", regex=False)
                df["Contract_Name"] = df["Contract_Name"].astype(str).str.replace(" ", "", regex=False)
                df["Project_Manager"] = df["Project_Manager"].astype(str).str.replace(" ", "", regex=False)
                df["Hours"] = df["Hours"].astype(str).str.replace(" ", "", regex=False)

                # fill down Name, Month, and Year from the first non-empty cell
                # check if values exist
                try:
                    df["Name"] = df["Name"].dropna().iloc[0]
                    df["Month"] = df["Month"].dropna().iloc[0]
                    df["Year"] = df["Year"].dropna().iloc[0]
                except IndexError:
                    print(f"Skipping sheet '{sheet_name}' in '{file_path}' — missing Name/Month/Year")

                    # popup UI warning
                    popup = Toplevel()
                    popup.title("Missing Data Warning")
                    popup.geometry("400x100+150+150")
                    Label(popup, text=f"'{sheet_name}' in {os.path.basename(file_path)}\nis missing Name, Month, or Year.\nSheet skipped.", justify="center", wraplength=380).pack(pady=10)
    
                    return
                
                # remove 0s from total hours
                df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce")
                df = df[df["Hours"] > 0]

                # Remove rows where Contract_Name is NaN
                df = df.dropna(subset=["Contract_Name"])

                # Remove rows where Contract_Name is blank or just whitespace
                df["Contract_Name"] = df["Contract_Name"].astype(str).str.strip()
                df = df[df["Contract_Name"] != ""]


                # add filename source
                path = file_path.replace(" ", "")
                df["Source_File"] = os.path.basename(path)
                df["Sheet_Name"] = sheet_name

                # upload to Access
                conn = get_conn()
                cursor = conn.cursor()
                row_count = len(df)

                for i, (_, row) in enumerate(df.iterrows()):
                    val = str(row.Contract_Name).strip().lower()
                    if pd.isna(row.Contract_Name) or val == "" or val == "nan":
                        print(f"Skipping row with missing or invalid contract in sheet '{sheet_name}'")
                        continue

                    cursor.execute(
                        f"""
                        INSERT INTO {TABLE_NAME}
                        (Name, Month, Year, Contract_Name, Project_Manager, Hours, Source_File, Sheet_Name)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)

                        """,
                        row.Name, row.Month, row.Year,
                        row.Contract_Name, row.Project_Manager,
                        row.Hours, row.Source_File, row.Sheet_Name

                    )

                    # update progress
                    if progress_callback:
                        percent = ((i + 1) / row_count) * 100
                        progress_callback(percent)

                conn.commit()
                conn.close()
                print(f"Imported: {file_path} | Sheet: {sheet_name}")

            except Exception as e:
                print(f"Error processing {file_path} | Sheet: {sheet_name}: {e}")
                traceback.print_exc()
        else:
            show_duplicate_warning()
            return

        if progress_callback:
            progress_callback(100)

    except Exception as e:
        print(f"Failed to import {file_path} (sheet: {sheet_name}): {e}")


''' TKINTER UI '''
# create a window and set size
root = Tk()
root.title("Timecard Vault")
root.geometry("800x600+100+50") # Width x Height + X_offset + Y_offset

# frames
# top
fram = Frame(root)
fram.pack(pady=30)
fram_14 = Frame(root)
fram_14.pack()
fram1_25 = Frame(root)
fram1_25.pack(pady=30)
# middle
fram1 = Frame(root)
fram1.pack(pady=20)
# second middle
fram1_5 = Frame(root)
fram1_5.pack(pady=10)
# bottom
fram2 = Frame(root)
fram2.pack(pady=10)


# uploading
# upload progress bar
progress = Progressbar(fram_14, orient = HORIZONTAL,
              length = 100, mode = 'determinate')
# upload not success label
duplicate_label = Label(fram_14, text='Error: Duplicate file, upload unsuccessful.')
print_error_label = Label(fram_14, text='Print error.')
# upload button
ubutt = Button(fram, text='Upload Timecard', padding=10)
ubutt.pack(side=LEFT)
# delete timecard button
dbutt = Button(root, text='Delete Timecard Entry')
dbutt.place(x=0, y=0) 


# searching
# label for searching
Label(fram1, text='Contract Name:').pack(side=LEFT)
# adding of single line text box
edit = Entry(fram1) 
# position of text box
edit.pack(side=LEFT, fill=BOTH, expand=1) 
#set focus
edit.focus_set()
# clear button
cbutt = Button(fram2, text='Clear')  
cbutt.pack(side=RIGHT) 
# search button
sbutt = Button(fram1, text='Search')  
sbutt.pack(side=LEFT) 
# advanced search
asbutt = Button(fram1, text='Advanced Search')
asbutt.pack(side=RIGHT)
# total hours display
hours_label = Label(fram1_5, text="Total Hours: " + str(LAST_TOTAL_HOURS))
hours_label.pack(side=LEFT, padx=10)


# search results treeview 
# to display results
results_tree = Treeview(root, columns=("Entry_ID","Name", "Month", "Year", "Contract_Name", "Project_Manager", "Hours", "Source_File"), show="headings")
for col in results_tree["columns"]:
    results_tree.heading(col, text=col)
    results_tree.column(col, width=100)
results_tree.pack(fill=BOTH, expand=True, pady=20)
active_treeviews.append(results_tree)
# print treeview
pbutt= Button(root, text='Print')
pbutt.place(relx = 1.0, 
                  rely = 0.0,
                  anchor ='ne')


''' Functions to Make Buttons Work '''
# treeview functions
# loading data into tree
def load_all_data():
    try:
        conn = get_conn()
        cursor = conn.cursor()
        cursor.execute(
            f"""
            SELECT
                Entry_ID, [Name], [Month], [Year], Contract_Name, Project_Manager, Hours, Source_File
            FROM
                {TABLE_NAME}
            """
        )
        rows = cursor.fetchall()
        conn.close()
        return [tuple(row) for row in rows]
    except Exception as e:
        print("Failed to load data from Access.")
        traceback.print_exc()
        return []
    
# create the tree
def populate_tree(tree, data):
    # clear the tree first
    for row in tree.get_children():
        tree.delete(row)
    # insert new data
    for row_data in data:
        tree.insert('', 'end', values=row_data)

# refresh data
def refresh_all_trees():
    global main_tree_data
    main_tree_data = load_all_data()

    for tree in active_treeviews[:]:  # copy to avoid mutation issues
        try:
            if str(tree) in tree.tk.call("winfo", "children", "."):
                populate_tree(tree, main_tree_data)
            else:
                active_treeviews.remove(tree)  # clean up invalid tree
        except Exception as e:
            print(f"Treeview refresh error: {e}")
            active_treeviews.remove(tree)

# search functions
# search contract name in Access
def search(search_text, tree, header, full_data):
    cleaned_search_term = search_text.strip().replace(" ", "").lower()

    # clear the tree
    for row in tree.get_children():
        tree.delete(row)

    if cleaned_search_term == "":
        populate_tree(tree, full_data)
        return

    filtered_data = [
        row for row in full_data
        if cleaned_search_term in str(row).lower()
    ]
    populate_tree(tree, filtered_data)

# calculate the total hours
def calculate_hours(search_term):
    global LAST_TOTAL_HOURS
    # remove spaces from search term
    search_term = edit.get()
    cleaned_search_term = search_term.replace(" ", "")

    if not cleaned_search_term:
        total_hours = 0
        hours_label.configure(text="Total Hours: " + str(total_hours))
        return

    try:
        total_hours = 0
        conn = get_conn()
        cursor = conn.cursor()
        cursor.execute(
            f"""
            SELECT
                Hours
            FROM
                {TABLE_NAME}
            WHERE
                Contract_Name LIKE ?;
            """,
            f"%{cleaned_search_term}%",
        )
        rows = cursor.fetchall()
        # calculation
        for (hours,) in rows:
            total_hours += hours
        conn.close()

        LAST_TOTAL_HOURS = total_hours
        hours_label.configure(text=f"Total Hours: " + str(total_hours))
        return total_hours
        

    except Exception as e:
        hours_label.configure(text="Total Hours: Error")
        traceback.print_exc()

# combines search and calculate functions
def search_and_calculate(event=None):
    search(edit.get(), results_tree, "Contract_Name", main_tree_data)
    calculate_hours(edit.get())

# calculate hours from the advanced search
def advanced_calculate_hours(search_term, month=None, year=None):
    global LAST_TOTAL_HOURS

    # clean contract name input
    if not search_term:
        total_hours = 0
        hours_label.configure(text="Total Hours: " + str(total_hours))
        return

    cleaned_search_term = search_term.replace(" ", "")
    total_hours = 0

    try:
        total_hours = 0
        conn = get_conn()
        cursor = conn.cursor()

        # build WHERE conditions and parameter list dynamically
        where_clauses = ["Contract_Name LIKE ?"]
        params = [f"%{cleaned_search_term}%"]

        if month:
            where_clauses.append("Month = ?")
            params.append(month)

        if year:
            where_clauses.append("[Year] = ?")
            params.append(year)

        query = f"""
            SELECT [Hours]
            FROM {TABLE_NAME}
            WHERE {" AND ".join(where_clauses)};
        """

        cursor.execute(query, params)
        rows = cursor.fetchall()
        for (hours,) in rows:
            total_hours += hours

        conn.close()
        LAST_TOTAL_HOURS = total_hours
        hours_label.configure(text=f"Total Hours: " + str(total_hours))
        return total_hours

    except Exception as e:
        hours_label.configure(text="Total Hours: Error")
        traceback.print_exc()

# advanced search helper
def on_advanced_search(contract, month, year):
    assbutt_press_and_return_contract_name(contract, month, year)
    advanced_calculate_hours(contract, month, year)

# clear search results
def clear(tree):
    if tree and tree.get_children():
        for item in tree.get_children():
            tree.delete(item)
    edit.delete(0, END)  # clear text in entry
    total_hours = 0
    hours_label.configure(text="Total Hours: " + str(total_hours)) # clear the calculated hours
    edit.focus_set()
cbutt.config(command=lambda: clear(results_tree))

# advanced search process
def advanced_search():
    # create the popup
    popup = Toplevel()
    popup.title("Advanced Search")
    popup.geometry("600x150+100+50") # Width x Height + X_offset + Y_offset

    # frames
    pfram1 = Frame(popup)
    pfram1.pack(pady=40)
    pfram2 = Frame(popup)
    pfram2.pack(pady=10)
    pfram3 = Frame(popup)
    pfram3.pack(pady=10)

    # searching labels and entries
    # contract
    Label(pfram1, text='Contract Name:').pack(side=LEFT)
    contract = Entry(pfram1) 
    contract.pack(side=LEFT, fill=BOTH, expand=1) 
    #set focus
    contract.focus_set()
    contract.bind("<Return>", lambda event: on_advanced_search(contract.get(), month.get(), year.get()))
    # month
    Label(pfram1, text='Month:').pack(side=LEFT)
    month = Entry(pfram1) 
    month.pack(side=LEFT, fill=BOTH, expand=1) 
    month.bind("<Return>", lambda event: on_advanced_search(contract.get(), month.get(), year.get()))
    # year
    Label(pfram1, text='Year:').pack(side=LEFT)
    year = Entry(pfram1) 
    year.pack(side=LEFT, fill=BOTH, expand=1) 
    year.bind("<Return>", lambda event: on_advanced_search(contract.get(), month.get(), year.get()))
    # search button
    assbutt = Button(pfram2, text='Search')  
    assbutt.pack(side=BOTTOM) 
    assbutt.config(command=lambda: on_advanced_search(contract.get(), month.get(), year.get()))

    # clear past search in regular search bar
    edit.delete(0, END)  # clear text in entry

    # grab inputs from UI
    advanced_calculate_hours(edit.get(), month.get(), year.get())

# detect when advanced search is pressed
def advanced_search_button_press(contract, month, year):
    cleaned_contract = contract.strip().replace(" ", "").lower()
    cleaned_month = month.strip().replace(" ", "").lower()
    cleaned_year = year.strip().replace(" ", "").lower()

    filtered_data = []

    for row in main_tree_data:
        row_str = str(row).lower().replace(" ", "")
        if cleaned_contract and cleaned_month and cleaned_year:
            if cleaned_contract in row_str and cleaned_month in row_str and cleaned_year in row_str:
                filtered_data.append(row)
        elif cleaned_contract and cleaned_month:
            if cleaned_contract in row_str and cleaned_month in row_str:
                filtered_data.append(row)
        elif cleaned_contract and cleaned_year:
            if cleaned_contract in row_str and cleaned_year in row_str:
                filtered_data.append(row)
        elif cleaned_contract:
            if cleaned_contract in row_str:
                filtered_data.append(row)

    populate_tree(results_tree, filtered_data)

# combine advanced_search_button_press and return_contract functions
def assbutt_press_and_return_contract_name(contract,month, year):
    advanced_search_button_press(contract, month, year)
    return_contract(contract)

# get contract for PDF name
def return_contract(contract):
    global PDF_contract
    PDF_contract = contract


# upload an excel file
# progress bar update
def update_progress(percent):
    progress['value'] = percent
    root.update_idletasks()

# import files
def import_multiple_files():
    duplicate_label.pack_forget()

    # open file explorer browse
    file_paths = filedialog.askopenfilenames(
        title="Select Excel Timecards",
        filetypes=[("Excel Files", "*.xls*")]
    )

    # error
    if not file_paths:
        return
    
    # import file with progress update
    def run_import():
        for file_path in file_paths:
            file_name = os.path.basename(file_path)

            progress.pack(padx=20, pady=10, side=LEFT) 

            try:
                xls = pd.ExcelFile(file_path)
                for sheet_name in xls.sheet_names:
                    progress['value'] = 0
                    root.update_idletasks()

                    import_timecard(file_path, sheet_name=sheet_name, progress_callback=update_progress)

                    progress['value'] = 100
                    root.update_idletasks()
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
        progress.pack_forget()
        refresh_all_trees()

    # shows the progress bar for an extra second so user is able to see it
    root.after(100, run_import)

# duplicate file check
def duplicate_file(file_name, sheet_name):
    conn = get_conn()
    cursor = conn.cursor()
    cursor.execute(
        f"""
        SELECT COUNT(*) FROM {TABLE_NAME}
        WHERE Source_File = ? AND Sheet_Name = ?
        """,
        (file_name, sheet_name),
    )
    count = cursor.fetchone()[0]
    conn.close()
    return count > 0

# show warning for duplicate file
def show_duplicate_warning(message=None):
    if message:
        duplicate_label.config(text=message)
    duplicate_label.pack(side=RIGHT)
    root.update_idletasks()



# delete functions
# remove from tree
def delete_by_treeview():
    # create the popup
    popup = Toplevel()
    popup.title("Delete Timecard")
    popup.geometry("800x400+100+50") # Width x Height + X_offset + Y_offset

    # frames
    pfram1 = Frame(popup)
    pfram1.pack()
    pfram2 = Frame(popup)
    pfram2.pack(pady=10)
    pfram3 = Frame(popup)
    pfram3.pack(pady=10)

    # existing entries treeview 
    # to display results
    entries_tree = Treeview(popup, columns=("Entry_ID", "Name", "Month", "Year", "Contract_Name", "Project_Manager", "Hours", "Source_File"), show="headings")
    for col in entries_tree["columns"]:
        entries_tree.heading(col, text=col)
        entries_tree.column(col, width=100)
    entries_tree.pack(fill=BOTH, expand=True, pady=20)
    active_treeviews.append(entries_tree)

    # searching
    # label for searching
    Label(pfram1, text='File Name:').pack(side=LEFT)
    # adding of single line text box
    edit = Entry(pfram1) 
    edit.bind("<Return>", lambda event: search(edit.get(), entries_tree, "Source_File", popup_tree_data))


    # position of text box
    edit.pack(side=LEFT, fill=BOTH, expand=1) 
    #set focus
    edit.focus_set() 
    # search button
    sbutt = Button(pfram1, text='Search')  
    sbutt.pack(side=RIGHT) 
    sbutt.config(command=lambda: search(edit.get(), entries_tree, "Source_File", popup_tree_data))

    # delete button
    delete = Button(pfram3, text="Delete", padding=10)
    delete.pack()
    delete.config(command=lambda:get_selected_items(entries_tree))
    #popup.bind("<BackSpace>", lambda event: get_selected_items(entries_tree))

    popup_tree_data = load_all_data()  # get data just for this popup
    populate_tree(entries_tree, popup_tree_data)

# return selected items
def get_selected_items(tree):
    keys = []
    selected_iids = tree.selection()
    if selected_iids:
        for iid in selected_iids:
            item_data = tree.item(iid)
            raw_value = item_data['values'][0]
            try:
                entry_id = int(str(raw_value).strip("(), "))  # Strip tuple-like chars
                keys.append(entry_id)
            except ValueError:
                print(f"Skipping invalid entry ID: {raw_value}")

        
    else:
        print("No items selected.")

    # create the popup
    popup = Toplevel()
    popup.title("WARNING!")
    popup.geometry("200x150+100+50") # Width x Height + X_offset + Y_offset

    Label(popup, text="Warning! \nAre you sure you would \nlike to delete these files? \nYou cannot undo this action.",justify="center").pack()

    yesbutt = Button(popup, text="Yes")
    yesbutt.pack()
    yesbutt.config(command=lambda: on_yes_button_pressed(keys, popup, tree))


    nobutt = Button(popup, text="No")
    nobutt.pack()
    nobutt.config(command=lambda: close_window(popup))

# combines function for when yes button is clicked
def on_yes_button_pressed(keys, window, tree):
    delete_entries(keys)
    close_window(window)
    # reload data for popup treeview
    updated_popup_data = load_all_data()
    populate_tree(tree, updated_popup_data)
    refresh_all_trees()

# close chosen window
def close_window(window):
    window.destroy()

# delete the entries
def delete_entries(keys):
    try:
        conn = get_conn()
        cursor = conn.cursor()

        print("Attempting to delete Entry_IDs:", keys)  # debug

        for key in keys:
            print("Deleting Entry_ID:", key)  # debug
            cursor.execute(f"DELETE FROM {TABLE_NAME} WHERE Entry_ID = ?", (key,))

        conn.commit()
        conn.close()

        # success popup
        popup = Toplevel()
        popup.title("Success!")
        popup.geometry("250x50+100+50")
        Label(popup, text="Success! Your file(s) were deleted.").pack(side=TOP)

    except Exception as e:
        traceback.print_exc()
        popup = Toplevel()
        popup.title("Error!")
        popup.geometry("250x50+100+50")
        Label(popup, text="Error. Your file(s) were NOT deleted.").pack(side=TOP)

# helper function
def delete_by_file_and_sheet(file_name, sheet_name):
    try:
        conn = get_conn()
        cursor = conn.cursor()
        cursor.execute(
            f"DELETE FROM {TABLE_NAME} WHERE Source_File = ? AND Sheet_Name = ?",
            (file_name, sheet_name),
        )
        conn.commit()
        conn.close()
        print(f"Deleted all entries for {file_name} | {sheet_name}")
    except Exception as e:
        print(f"Error deleting sheet {sheet_name} from {file_name}: {e}")
        traceback.print_exc()

#check for file
def file_check(file):
    contains_file = False
    count = 0
    try:
        conn = get_conn()
        cursor = conn.cursor()
        cursor.execute(
            f"""
            SELECT
                Source_File
            FROM
                {TABLE_NAME}
            WHERE
                Source_File LIKE ?;
            """,
            f"%{file}%",
        )
        rows = cursor.fetchall()
        for (files,) in rows:
            count += 1
        conn.close()

        if (count > 0):
            contains_file = True

        return contains_file

    except Exception as e:
        print(f"Error: {e}")


# printing
# open PDF
def print_pdf(filename):
    if platform.system() == "Windows":
        os.startfile(filename)  # open PDF
    else:
        print_error_label.pack(side=RIGHT)  # show the print error

# put tree into PDF format
def export_treeview_to_pdf(tree):
    duplicate_label.pack_forget()
    raw_name = ""

    if edit.get() != "":
        raw_name = edit.get().strip()
    else:
        raw_name = PDF_contract.strip()

    contract_name = re.sub(r'[^\w\-]', '_', raw_name)
    default_filename = f"timecard_report_{contract_name}.pdf"
    filename = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        initialfile=default_filename,
        title="Save Timecard Report As"
    )

    # exit if no filename selected
    if not filename:
        print("PDF generation cancelled by user.")
        return

    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter

    x_offset = 50
    y_offset = height - 50
    row_height = 20

    c.setFont("Helvetica-Bold", 16)
    c.drawString(x_offset, y_offset, "Timecard Report                                                           Total Hours Spent:")

    c.setFont("Helvetica", 12)
    y_offset -= 25
    c.drawString(x_offset, y_offset, "Generated by Timecard Vault                                                                       " + str(LAST_TOTAL_HOURS))
    
    y_offset -= 15
    c.drawString(x_offset, y_offset, " ")

    y_offset -= 30  # add some space before table

    # get headings
    columns = tree["columns"]
    headings = [tree.heading(col)["text"] for col in columns]

    # draw headings
    c.setFont("Helvetica-Bold", 8)
    for i, heading in enumerate(headings):
        c.drawString(x_offset + i * 68, y_offset, heading)

    y_offset -= row_height

    # get rows
    c.setFont("Helvetica-Bold", 7)
    for row_id in tree.get_children():
        row = tree.item(row_id)['values']
        for i, value in enumerate(row):
            c.drawString(x_offset + i * 68, y_offset, str(value))
        y_offset -= row_height
        if y_offset < 50:
            c.showPage()
            y_offset = height - 50

    c.save()
    print(f"PDF saved as {filename}")

    if os.path.exists(filename):
        print_pdf(filename)
    else:
        print(f"Error: File not found at {filename}")



''' Button Assignment '''
# search
sbutt.config(command=search_and_calculate)
edit.bind("<Return>", search_and_calculate)
# upload timecard
ubutt.config(command=import_multiple_files)
# print 
pbutt.config(command=lambda: export_treeview_to_pdf(results_tree))
# delete
dbutt.config(command=delete_by_treeview)
# advanced search
asbutt.config(command=advanced_search)


''' Window Loop '''
main_tree_data = load_all_data()
populate_tree(results_tree, main_tree_data)
root.mainloop()



