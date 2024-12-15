import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import webbrowser
import imgkit

def load_datesheet(file_path):
    sheet_data = pd.read_excel(file_path, sheet_name='Complete')
    cleaned_data = sheet_data.iloc[2:, [0, 1, 2, 3, 4, 5]]
    cleaned_data.columns = ['Day', 'Date', 'Code 1', 'Course Name 1', 'Code 2', 'Course Name 2']
    morning_data = cleaned_data[['Day', 'Date', 'Code 1', 'Course Name 1']].rename(
        columns={'Code 1': 'Code', 'Course Name 1': 'Course Name'}
    )
    morning_data['Time'] = '09:00 - 12:00'
    afternoon_data = cleaned_data[['Day', 'Date', 'Code 2', 'Course Name 2']].rename(
        columns={'Code 2': 'Code', 'Course Name 2': 'Course Name'}
    )
    afternoon_data['Time'] = '1:00 - 4:00'
    combined_data = pd.concat([morning_data, afternoon_data], ignore_index=True)
    combined_data = combined_data.dropna(subset=['Code', 'Course Name'])
    return combined_data

class DateSheetApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Date Sheet Generator")
        self.root.geometry("900x700")
        self.root.config(bg="#1e1e2f")
        self.data = None
        self.selected_courses = []

        upload_frame = ttk.Frame(root, style="TFrame")
        upload_frame.pack(pady=20)
        self.upload_btn = ttk.Button(upload_frame, text="Upload Date Sheet", style="TButton", command=self.upload_datesheet)
        self.upload_btn.pack()

        search_frame = ttk.Frame(root, style="TFrame")
        search_frame.pack(pady=20, fill=tk.X)
        ttk.Label(search_frame, text="Search Courses:", style="TLabel").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_bar = ttk.Entry(search_frame, textvariable=self.search_var, style="TEntry")
        self.search_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.search_bar.bind("<KeyRelease>", self.update_suggestions)

        self.suggestions_list = tk.Listbox(root, height=5, selectmode=tk.SINGLE, bg="#2b2b3f", fg="white", bd=0, font=("Arial", 12))
        self.suggestions_list.pack(pady=10, fill=tk.X)
        self.suggestions_list.bind("<Double-1>", self.add_course)
        style = ttk.Style()
        style.configure("TLabelframe", background="#2C3E50", foreground="white", font=("Helvetica", 12))
        selected_frame = ttk.LabelFrame(root, text="Selected Courses", style="TLabelframe")
        selected_frame.pack(pady=20, fill=tk.BOTH, expand=True)
        self.selected_list = tk.Listbox(selected_frame, height=10, bg="#2b2b3f", fg="white", bd=0, font=("Arial", 12))
        self.selected_list.pack(pady=10, fill=tk.BOTH, expand=True)
        self.selected_list.bind("<Double-1>", self.remove_course)

        self.generate_btn = ttk.Button(root, text="Generate Date Sheet", style="TButton", command=self.generate_datesheet)
        self.generate_btn.pack(pady=10)

        self.save_btn = ttk.Button(root, text="Save Date Sheet", style="TButton", command=self.save_datesheet, state=tk.DISABLED)
        self.save_btn.pack(pady=10)

        self.style_config()

    def style_config(self):
        style = ttk.Style()
        style.configure("TButton", font=("Arial", 12), background="#4CAF50", padding=10)
        style.map("TButton", background=[('active', '#45a049')])
        style.configure("TLabel", font=("Arial", 12), foreground="#ffffff", background="#1e1e2f")
        style.configure("TEntry", font=("Arial", 12), padding=5)
        style.configure("TFrame", background="#1e1e2f")
        style.configure("TLabelFrame", font=("Arial", 12), foreground="#ffffff", background="#1e1e2f")

    def upload_datesheet(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.data = load_datesheet(file_path)
            self.update_suggestions()

    def update_suggestions(self, event=None):
        if self.data is not None:
            query = self.search_var.get().lower()
            suggestions = self.data[self.data['Code'].str.contains(query, case=False) |
                                    self.data['Course Name'].str.contains(query, case=False)]
            self.suggestions_list.delete(0, tk.END)
            for _, row in suggestions.iterrows():
                self.suggestions_list.insert(tk.END, f"{row['Code']} - {row['Course Name']}")

    def add_course(self, event):
        selected = self.suggestions_list.get(tk.ACTIVE)
        if selected and selected not in self.selected_courses:
            self.selected_courses.append(selected)
            self.selected_list.insert(tk.END, selected)

    def remove_course(self, event):
        selected = self.selected_list.get(tk.ACTIVE)
        if selected in self.selected_courses:
            self.selected_courses.remove(selected)
            self.selected_list.delete(tk.ACTIVE)

    def generate_datesheet(self):
        if not self.selected_courses:
            messagebox.showerror("Error", "No courses selected!")
            return

        html_content = """<html><head><title>Generated Date Sheet</title></head><body>
        <h1>Generated Date Sheet</h1>
        <table border="1" style="width:100%; border-collapse: collapse; text-align: left;">
        <tr><th>Day</th><th>Date</th><th>Time</th><th>Code</th><th>Course Name</th></tr>
        """

        for course in self.selected_courses:
            course_details = self.data[self.data['Code'].str.contains(course.split(' - ')[0], case=False)].iloc[0]
            html_content += f"<tr><td>{course_details['Day']}</td><td>{course_details['Date']}</td><td>{course_details['Time']}</td><td>{course_details['Code']}</td><td>{course_details['Course Name']}</td></tr>"

        html_content += "</table></body></html>"

        with open("generated_datesheet.html", "w") as file:
            file.write(html_content)

        webbrowser.open("generated_datesheet.html")
        self.save_btn.config(state=tk.NORMAL)

    def save_datesheet(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML Files", "*.html")])
        if save_path:
            with open("generated_datesheet.html", "r") as temp_file:
                html_content = temp_file.read()
            with open(save_path, "w") as save_file:
                save_file.write(html_content)
            messagebox.showinfo("Success", "Date sheet saved successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = DateSheetApp(root)
    root.mainloop()
