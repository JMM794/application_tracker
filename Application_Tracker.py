import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkcalendar import DateEntry
from openpyxl import Workbook
from datetime import datetime


class JobApplicationTracker:
    def __init__(self, root):
        self.root = root
        self.root.title("Job Application Tracker")
        
        # Create DataFrame to store job applications
        self.df = pd.DataFrame(columns=["Company", "Website", "Date Applied", "Interview Scheduled", "Interview Date"])
        
        # Create GUI elements
        self.company_label = ttk.Label(root, text="Company:")
        self.company_entry = ttk.Entry(root)
        
        self.website_label = ttk.Label(root, text="Website:")
        self.website_entry_linkedin_var = tk.BooleanVar()
        self.website_entry_indeed_var = tk.BooleanVar()
        self.website_entry_linkedin = ttk.Checkbutton(root, text="LinkedIn", variable=self.website_entry_linkedin_var)
        self.website_entry_indeed = ttk.Checkbutton(root, text="Indeed", variable=self.website_entry_indeed_var)

        self.date_label = ttk.Label(root, text="Date Applied:")
        print("Creating DataEntry widget....")
        self.date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
        print("DataEntry widget created.")

        self.interview_var = tk.BooleanVar()
        self.interview_checkbutton = ttk.Checkbutton(root, text="Interview Scheduled", variable=self.interview_var)
        self.interview_date_label = ttk.Label(root, text="Interview Date:")
        self.interview_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
        
        #self.interview_date_label = ttk.Label(root, text="Interview Date:")
        #self.interview_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
        
        self.export_button = ttk.Button(root, text="Export to Excel", command=self.export_to_excel)
        
        # Grid layout
        self.company_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.company_entry.grid(row=0, column=1, padx=5, pady=5)
        
        self.website_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.website_entry_linkedin.grid(row=1, column=1, padx=5, pady=5) 
        self.website_entry_indeed.grid(row=1, column=2, padx=5, pady=5) 

        self.date_label.grid(row=2, column=0,padx=5, pady=5, sticky="e")
        self.date_entry.grid(row=2, column=1,padx=5, pady=5)
        print("Data Entry widget created")

        self.interview_checkbutton.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        self.interview_date_label.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.interview_date_entry.grid(row=4, column=1, padx=5, pady=5)

        self.export_button.grid(row=5, column=1, columnspan=1, padx=5, pady=5, sticky="we")
    
    def add_application(self):
        company = self.company_entry.get()
        website = ""
        if self.interview_var.get():
            if self.website_entry_linkedin.instate(['selected']):
                website += "LinkedIn"
            if self.website_entry_indeed.instate(['selected']):
                website += ", Indeed"
            website = website.rstrip(", ")
        date_applied = self.date_entry.get()
        interview_scheduled = self.interview_var.get()
        date_applied = self.date_entry.get()
        interview_scheduled = self.interview_var.get()
        
       
        if interview_scheduled:
            interview_date = self.interview_date_entry.get_date()
        else:
            interview_date = None
        
        new_row = {"Company": company, "Website": website, "Date Applied": date_applied, 
                "Interview Scheduled": interview_scheduled, "Interview Date": interview_date}
        
        
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)

    def export_to_excel(self):
        file_name = "job_applications.xlsx"
        self.df.to_excel(file_name, index=False)
        print(f"Data exported to {file_name}")

if __name__ == "__main__":
    root = tk.Tk()
    app = JobApplicationTracker(root)
    root.mainloop()
