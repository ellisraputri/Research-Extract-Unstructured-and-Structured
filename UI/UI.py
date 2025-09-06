from tkinter import Tk, StringVar, Frame, Label, Entry, Button, Listbox, DISABLED, NORMAL, filedialog, messagebox
import extraction_machine

# Initialize main window
root = Tk()
root.title("DOCX Data Extractor")
root.geometry("600x600")

selected_files = []
processed_files = {}  # Store extracted data in memory instead of saving immediately
pending_filename = StringVar()

# UI Components
exit_frame = Frame(root)
exit_frame.pack(side="top", anchor="ne", padx=10, pady=10)
Button(exit_frame, text="Exit", command=root.quit, width=10, bg="red", fg="white").pack()

frame = Frame(root, padx=20, pady=10)
frame.pack()
Label(frame, text="Enter File Name:", font=("Arial", 10)).pack()
Entry(frame, textvariable=pending_filename, width=30).pack(pady=5)

Button(root, text="Select Files", command=lambda: select_files_ui(), width=20).pack(pady=5)

Label(root, text="Selected Files:", font=("Arial", 10, "bold")).pack(pady=5)
selected_files_listbox = Listbox(root, width=50, height=6)
selected_files_listbox.pack(pady=5)

extract_button = Button(root, text="Extract Data", command=lambda: process_files_ui(), width=20, state=DISABLED, bg="lightgray", fg="gray")
extract_button.pack(pady=5)

Label(root, text="Extracted Files:", font=("Arial", 10, "bold")).pack(pady=5)
extracted_files_listbox = Listbox(root, width=50, height=6)
extracted_files_listbox.pack(pady=5)

Button(root, text="Download Selected File", command=lambda: download_file_ui(), bg="green", fg="white").pack(pady=5)

def select_files_ui():
    global selected_files
    selected_files = extraction_machine.select_files()
    
    selected_files_listbox.delete(0, "end")
    for file in selected_files:
        selected_files_listbox.insert("end", file.split("/")[-1])
    
    extract_button.config(state=NORMAL, bg="blue", fg="white")

def process_files_ui():
    file_name = pending_filename.get() + ".xlsx" 
    extracted_data = extraction_machine.process_selected_files(selected_files)

    if extracted_data is not None and not extracted_data.empty:
        processed_files[file_name] = extracted_data  
        extracted_files_listbox.insert("end", file_name) 
        messagebox.showinfo("Success", "Data extraction complete! Click 'Download Selected File' to save.")
    else:
        messagebox.showwarning("Extraction Failed", "No data was extracted.")

def download_file_ui():
    """Downloads the selected extracted file properly."""
    selected_index = extracted_files_listbox.curselection()
    if selected_index:
        selected_file = extracted_files_listbox.get(selected_index[0])

        if selected_file in processed_files:  # Ensure data exists
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", 
                filetypes=[("Excel Files", "*.xlsx")], 
                title="Save Extracted Data", 
                initialfile=selected_file
            )
            if save_path:
                processed_files[selected_file].to_excel(save_path, index=False)  # Save only when clicked
                messagebox.showinfo("Success", f"File saved to {save_path}")
        else:
            messagebox.showwarning("Error", "File not found in memory.")

root.mainloop()
