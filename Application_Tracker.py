import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkcalendar import DateEntry
from tkinter import messagebox
from tkinter import filedialog
from openpyxl import Workbook
from datetime import datetime


class JobApplicationTracker:
    def __init__(self, root):
        self.root = root
        self.root.title("Job Application Tracker")
        
        # Create DataFrame to store job applications
        self.df = pd.DataFrame(columns=["Company", "Website", "Date Applied", "Interview Scheduled", "Interview Date", "Job Title"])
        
        # Create GUI elements
        self.company_label = ttk.Label(root, text="Company:")
        self.company_entry = ttk.Entry(root)
        
        self.website_label = ttk.Label(root, text="Website:")
        self.website_entry_linkedin_var = tk.BooleanVar()
        self.website_entry_indeed_var = tk.BooleanVar()
        self.website_entry_linkedin = ttk.Checkbutton(root, text="LinkedIn", variable=self.website_entry_linkedin_var)
        self.website_entry_indeed = ttk.Checkbutton(root, text="Indeed", variable=self.website_entry_indeed_var)

        self.date_label = ttk.Label(root, text="Date Applied:")
        self.date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
        
        self.description_label = ttk.Label(root, text="Job Title: ")
        self.description_entry = ttk.Entry(root)
        #self.browse_button = ttk.Button(root, text="Browse", command=self.browse_file)

        self.interview_var = tk.BooleanVar()
        self.interview_checkbutton = ttk.Checkbutton(root, text="Interview Scheduled", variable=self.interview_var)
        self.interview_var.trace_add("write", self.toggle_interview_date_entry)
        self.interview_date_label = ttk.Label(root, text="Interview Date:")
        self.interview_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
        
        #self.interview_date_label = ttk.Label(root, text="Interview Date:")
        #self.interview_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
        
        self.export_button = ttk.Button(root, text="Export to Excel", command=self.add_application)
        
        # Grid layout
        self.company_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.company_entry.grid(row=0, column=1, padx=5, pady=5)
        
        self.website_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.website_entry_linkedin.grid(row=1, column=1, padx=5, pady=5) 
        self.website_entry_indeed.grid(row=1, column=2, padx=5, pady=5) 

        self.date_label.grid(row=2, column=0,padx=5, pady=5, sticky="e")
        self.date_entry.grid(row=2, column=1,padx=5, pady=5)
        
        self.description_label.grid(row=3, column=0,padx=5, pady=5, sticky="e")
        self.description_entry.grid(row=3, column=1, padx=5, pady=5)
        #self.browse_button.grid(row=3, column=2, padx=5, pady=5)
        self.interview_checkbutton.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        self.interview_date_label.grid(row=5, column=0, padx=5, pady=5, sticky="e")
        self.interview_date_entry.grid(row=5, column=1, padx=5, pady=5)

        self.export_button.grid(row=6, column=1, columnspan=1, padx=5, pady=5, sticky="we")
    
    def toggle_interview_date_entry(self, *args):
        if self.interview_var.get():
            self.interview_date_entry.configure(state="normal")
        else:
            self.interview_date_entry.configure(state="disabled")
    
    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.description_entry.delete(0, tk.END)
            self.description_entry.insert(0, file_path)

    def add_application(self):
        company = self.company_entry.get()
        website = ""
        #if self.interview_var.get():
        if self.website_entry_linkedin.instate(['selected']):
            website += "LinkedIn"
        if self.website_entry_indeed.instate(['selected']):
            website += ", Indeed"
        website = website.rstrip(", ")
        date_applied = self.date_entry.get()
        interview_scheduled = self.interview_var.get()
        description = self.description_entry.get()        
       
        if interview_scheduled:
            interview_date = self.interview_date_entry.get_date()
        else:
            interview_date = None
        
        new_row = {"Company": company, "Website": website, "Date Applied": date_applied, 
                "Interview Scheduled": interview_scheduled, "Interview Date": interview_date,
                "Job Description": description}
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.export_to_excel()
        #self.add_application()

    def export_to_excel(self):
        try:
            file_name = "job_applications.xlsx"
            try:
                existing_df = pd.read_excel(file_name)
                combined_df = pd.concat([existing_df, self.df], ignore_index=True)
            except FileNotFoundError:
                combined_df = self.df
            combined_df.to_excel(file_name, index=False)
            print(f"Data exported to {file_name}")
            messagebox.showinfo("Export Successful", f"Data exported to {file_name}")
        except Exception as e:
            print(f"Error occurred during export: {e}")
            messagebox.showerror("Export Error", f"An error occurred during export: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = JobApplicationTracker(root)
    root.mainloop()
