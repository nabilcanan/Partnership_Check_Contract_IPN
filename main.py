import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
import pandas as pd


def compare_excel_files(file1_path, file2_path, header_name):
    try:
        # Read data from both Excel files into DataFrames
        df1 = pd.read_excel(file1_path, engine='openpyxl')
        df2 = pd.read_excel(file2_path, engine='openpyxl')

        # Extract the IPNs from both DataFrames
        ipns_in_last_week = set(df1[header_name])
        ipns_in_this_week = set(df2[header_name])

        # Find the IPNs that are in last week's file but not in this week's file
        missing_ipns = ipns_in_last_week - ipns_in_this_week

        if not missing_ipns:
            result_label.config(text="There are no missing IPNs between the two Excel files.")
        else:
            # Create a new DataFrame containing the rows corresponding to the missing IPNs
            diff_df = df1[df1[header_name].isin(missing_ipns)]

            # Create a Treeview table to display the missing data
            table = ttk.Treeview(result_frame)
            table["columns"] = list(diff_df.columns)
            table.heading("#0", text="")
            for col in diff_df.columns:
                table.heading(col, text=col)
                table.column(col, width=100)

            for index, row in diff_df.iterrows():
                table.insert("", "end", values=list(row))

            table.pack(pady=10)

            # Add horizontal scrollbar
            h_scrollbar = ttk.Scrollbar(result_frame, orient="horizontal", command=table.xview)
            h_scrollbar.pack(fill="x")
            table.configure(xscrollcommand=h_scrollbar.set)

            # Add vertical scrollbar
            v_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=table.yview)
            v_scrollbar.pack(side="right", fill="y")
            table.configure(yscrollcommand=v_scrollbar.set)

            # Ask the user to save the new Excel file
            save_path = filedialog.asksaveasfilename(title="Save New Excel File",
                                                     filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
            if save_path:
                # Export the rows corresponding to missing IPNs to a new Excel file with .xlsx extension
                diff_df.to_excel(save_path, index=False, engine='openpyxl')
                result_label.config(text=result_label.cget("text") + f"\nNew file saved: {save_path}.xlsx")

    except Exception as e:
        result_label.config(text=f"Error occurred: {str(e)}")


def select_file():
    file_path = filedialog.askopenfilename(title="Select Excel file",
                                           filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")))
    if file_path:
        return file_path
    else:
        return None


def prompt_for_header_name():
    header_name = simpledialog.askstring("Header Name", "Enter the header name by which you want to compare:")
    return header_name


def compare_files_on_button_click():
    file1_path = select_file()
    file2_path = select_file()

    if file1_path and file2_path:
        header_name = prompt_for_header_name()
        if header_name:
            compare_excel_files(file1_path, file2_path, header_name)
        else:
            messagebox.showinfo("Invalid Header Name", "Please enter a valid header name.")


def exit_program():
    window.destroy()


window = tk.Tk()
window.title("Excel File Comparator")
window.geometry("800x500")


title_label = tk.Label(window, text="Excel File Comparator", font=("Times New Roman", 24, "bold"))
title_label.pack(pady=20)

file_label_1 = tk.Label(window, text="Select last week's contract, then select this week's contract file", bg="white")
file_label_1.pack(pady=10)

compare_button = tk.Button(window, text="Compare Files", command=compare_files_on_button_click,
                           font=("Times New Roman", 16, "bold"), bg="green", fg="white", width=20, height=2)
compare_button.pack(pady=10)

result_frame = tk.Frame(window)
result_frame.pack(pady=20)

result_label = tk.Label(window, text="", bg="white", wraplength=600)
result_label.pack(pady=10)

exit_button = tk.Button(window, text="Exit", command=exit_program, font=("Times New Roman", 16, "bold"),
                        bg="red", fg="white", width=20, height=2)
exit_button.pack(pady=10)

window.mainloop()
