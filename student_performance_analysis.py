import os
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import messagebox, ttk
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from statistics import mode, StatisticsError

# === Seaborn Theme ===
sns.set(style="darkgrid",palette="pastel")

# === Professional Theme Colors ===
BG_WINDOW = "#ECEFF1"
BG_FRAME = "#CFD8DC"
ACCENT = "#4DB6AC"
BTN_SECONDARY = "#90A4AE"
TABLE_BG = "#98E2EA"
TABLE_FG = "#1E1E2F"

# === Configuration ===
EXCEL_FILE = r"C:\Users\PALLAVI PAWAR\OneDrive\Desktop\Excel_File\student_performance_analyzer.xlsx"
SHEET_NAME = "Students"

# === Student Class ===
class Student:
    def __init__(self, name, roll, marks, study_hours):
        self.name = str(name)
        self.roll = str(roll)
        self.marks = float(marks)
        self.study_hours = float(study_hours)

# === Analyzer Class ===
class Analyzer:
    def __init__(self, filename=EXCEL_FILE):
        self.filename = filename
        self.students = []
        self.ensure_excel_exists()

    def ensure_excel_exists(self):
        if not os.path.exists(self.filename):
            messagebox.showerror("Error", f"Excel file not found:\n{self.filename}")
            raise FileNotFoundError("Excel file missing")

    def load_data(self):
        self.students = []
        try:
            df = pd.read_excel(self.filename, sheet_name=SHEET_NAME)
            df.columns = [col.strip().lower() for col in df.columns]
            for _,row in df.iterrows():
                self.students.append(Student(
                    str(row.get("name", "")).strip(),
                    str(row.get("roll", "")).strip(),
                    float(row.get("marks", 0)),
                    float(row.get("study_hours", 0))
                ))
            self.students.sort(key=lambda s: int(s.roll) if s.roll.isdigit() else s.roll)
        except Exception as e:
            messagebox.showerror("Load Error", f"Error loading Excel:\n{e}")

# === GUI ===
def run_gui():
    analyzer = Analyzer()
    root = tk.Tk()
    root.title("🎓 Student Performance Analyzer")
    root.state('zoomed')  # Maximized window
    root.configure(bg=BG_WINDOW)
    root.resizable(True, True)

    # Header
    tk.Label(root, text="STUDENT PERFORMANCE ANALYTICS",
             bg=BG_WINDOW, fg=ACCENT, font=("Segoe UI", 20, "bold")).pack(pady=10)

    # Buttons frame
    btn_frame = tk.Frame(root, bg=BG_FRAME)
    btn_frame.pack(pady=5, padx=10)

    def create_btn(text, cmd, color, row, col):
        tk.Button(btn_frame, text=text, command=cmd, bg=color, fg="White",
                  font=("Segoe UI", 12, "bold"), relief="groove", width=18, pady=5).grid(row=row, column=col, padx=5, pady=5)

    # --- Table frame ---
    table_frame = tk.Frame(root, bg=BG_WINDOW)
    table_frame.pack(padx=10, pady=10, fill="both", expand=True)
    columns = ("roll", "name", "marks", "study_hours")
    tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=18)
    for col in columns:
        tree.heading(col, text=col.capitalize())
        tree.column(col, anchor="center", width=200)
        tree.pack(fill="both", expand=True)

    style = ttk.Style()
    style.configure("Treeview", background=TABLE_BG,foreground=TABLE_FG,fieldbackground=TABLE_BG,font=("Segoe UI", 10))
    style.configure("Treeview.Heading",font=("Segoe UI", 11, "bold"),foreground=ACCENT, background=BG_WINDOW)
    style.map("Treeview",background=[("selected", ACCENT)],foreground=[("selected", "#FFFFFF")])

    # --- Core functions ---
    def refresh():
        analyzer.load_data()
        tree.delete(*tree.get_children())
        if not analyzer.students:
            messagebox.showinfo("No Data", "No student data found in Excel.")
            return
        for s in analyzer.students:
            tree.insert("", tk.END, values=(s.roll, s.name, s.marks, s.study_hours))

    def analyze_stats():
        analyzer.load_data()
        if not analyzer.students:
            messagebox.showinfo("No Data", "No student data found in Excel.")
            return
        df = pd.DataFrame([vars(s) for s in analyzer.students])
        stats_window = tk.Toplevel(root)
        stats_window.title("Class Statistics")
        stats_window.geometry("500x500")
        stats_window.configure(bg=BG_WINDOW)
        tk.Label(stats_window, text="📊 Class Statistics", bg=BG_WINDOW, fg=ACCENT,
                 font=("Segoe UI", 16, "bold")).pack(pady=10)
        stats_text = f"Total Students: {len(analyzer.students)}\n\n"
        for col in ["marks", "study_hours"]:
            col_mean = round(df[col].mean(), 2)
            col_median = round(df[col].median(), 2)
            try:
                col_mode = mode(df[col])
            except StatisticsError:
                col_mode = "No unique mode"
            col_max = df[col].max()
            col_min = df[col].min()
            stats_text += f"{col.capitalize()}:\n"
            stats_text += f"  Mean = {col_mean}\n"
            stats_text += f"  Median = {col_median}\n"
            stats_text += f"  Mode = {col_mode}\n"
            stats_text += f"  Max = {col_max}\n"
            stats_text += f"  Min = {col_min}\n\n"
        text_widget = tk.Text(stats_window, bg=BG_WINDOW, fg=TABLE_FG, font=("Segoe UI", 11), wrap="word")
        text_widget.pack(fill="both", expand=True, padx=10, pady=10)
        text_widget.insert("1.0", stats_text)
        text_widget.config(state="disabled")

    def generate_report():
        if not tree.get_children():
            messagebox.showerror("Error", "No data to generate report!")
            return
        report_file = "student_report.pdf"
        c = canvas.Canvas(report_file, pagesize=A4)
        width,height = A4
        c.drawString(180, height-50, "Student Performance Report")
        c.setFont("Helvetica", 12)
        y = height - 100
        if analyzer.students:
            df = pd.DataFrame([vars(s) for s in analyzer.students])
            stats_text = ""
            for col in ["marks", "study_hours"]:
                col_mean = round(df[col].mean(), 2)
                col_median = round(df[col].median(), 2)
                try:
                    col_mode = mode(df[col])
                except StatisticsError:
                    col_mode = "No unique mode"
                col_max = df[col].max()
                col_min = df[col].min()
                stats_text += f"{col.capitalize()}: Mean={col_mean}, Median={col_median}, Mode={col_mode}, Max={col_max}, Min={col_min}\n"
            for line in stats_text.strip().split("\n"):
                c.drawString(50, y, line)
                y -= 20
            y -= 10
        c.drawString(50, y, "Roll No.")
        c.drawString(150, y, "Name")
        c.drawString(400, y, "Marks")
        c.drawString(500, y, "Study Hours")
        y -= 20
        for row_id in tree.get_children():
            row = tree.item(row_id)["values"]
            c.drawString(50, y, str(row[0]))
            c.drawString(150, y, str(row[1]))
            c.drawString(400, y, str(row[2]))
            c.drawString(500, y, str(row[3]))
            y -= 20
            if y < 50:
                c.showPage()
                y = height - 50
        c.save()
        messagebox.showinfo("Report Generated",f"PDF report saved as {report_file}")

    # --- Plot Functions ---
    def scatter_plot():
        analyzer.load_data()
        if not analyzer.students: messagebox.showerror("No Data", "Load students first."); return
        df = pd.DataFrame([vars(s) for s in analyzer.students])
        plt.figure(figsize=(8,6))
        sns.regplot(data=df, x="study_hours", y="marks")
        plt.title("Study Hours vs Marks (Regression)")
        plt.show()

    def bar_all():
        analyzer.load_data()
        if not analyzer.students: messagebox.showerror("No Data", "Load students first."); return
        df = pd.DataFrame([vars(s) for s in analyzer.students])
        plt.figure(figsize=(10,6))
        sns.barplot(data=df, y="name", x="marks", orient="h")
        plt.title("Marks of All Students")
        plt.show()

    def bar_top5():
        analyzer.load_data()
        if not analyzer.students: messagebox.showerror("No Data", "Load students first."); return
        df = pd.DataFrame([vars(s) for s in analyzer.students]).sort_values("marks", ascending=False).head(5)
        plt.figure(figsize=(8,4))
        sns.barplot(data=df, y="marks", x="name", orient="v")
        plt.title("Top 5 Performers")
        plt.show()

    def pie_chart():
        analyzer.load_data()
        if not analyzer.students: messagebox.showerror("No Data", "Load students first."); return
        categories = {">=85":0, "70-84":0, "50-69":0, "<50":0}
        for s in analyzer.students:
            if s.marks >= 85: 
                categories[">=85"] += 1
            elif s.marks >= 70: 
                categories["70-84"] += 1
            elif s.marks >= 50: 
                categories["50-69"] += 1
            else: categories["<50"] += 1
        plt.figure(figsize=(6,6))
        plt.pie(categories.values(), labels=categories.keys(), autopct="%1.1f%%", startangle=140,shadow=True,explode=[0,0.2,0,0])
        plt.title("Performance Categories")
        plt.show()

    def histogram_marks():
        analyzer.load_data()
        if not analyzer.students: messagebox.showerror("No Data", "Load students first."); return
        df = pd.DataFrame([vars(s) for s in analyzer.students])
        plt.figure(figsize=(8,6))
        sns.histplot(df['marks'], bins=10, kde=True, color=ACCENT)
        plt.title("Distribution of Marks")
        plt.xlabel("Marks")
        plt.ylabel("Number of Students")
        plt.show()

    def boxplot_study_hours():
        analyzer.load_data()
        if not analyzer.students: messagebox.showerror("No Data", "Load students first."); return
        df = pd.DataFrame([vars(s) for s in analyzer.students])
        plt.figure(figsize=(6,6))
        sns.boxplot(y=df['study_hours'], color=ACCENT)
        plt.title("Boxplot of Study Hours")
        plt.ylabel("Study Hours")
        plt.show()

    def correlation_heatmap():
        analyzer.load_data()
        if not analyzer.students: messagebox.showerror("No Data", "Load students first."); return
        df = pd.DataFrame([vars(s) for s in analyzer.students])
        corr = df[['marks', 'study_hours']].corr()
        plt.figure(figsize=(6,5))
        sns.heatmap(corr, annot=True, cmap="coolwarm", linewidths=0.5)
        plt.title("Correlation Heatmap")
        plt.show()

    def stacked_bar_categories():
        analyzer.load_data()
        if not analyzer.students: messagebox.showerror("No Data", "Load students first."); return
        df = pd.DataFrame([vars(s) for s in analyzer.students])
        categories = {'>=85': 0, '70-84': 0, '50-69': 0, '<50': 0}
        for mark in df['marks']:
            if mark >= 85: 
                categories['>=85'] += 1
            elif mark >= 70:
                categories['70-84'] += 1
            elif mark >= 50:
                categories['50-69'] += 1
            else: categories['<50'] += 1
        plt.figure(figsize=(8,5))
        plt.bar(categories.keys(), categories.values(), color=ACCENT)
        plt.title("Number of Students by Performance Category")
        plt.xlabel("Category")
        plt.ylabel("Number of Students")
        plt.show()

    # --- Buttons layout ---
    create_btn("Load Data", refresh, ACCENT, 0, 0)
    create_btn("Analyze", analyze_stats, ACCENT, 0, 1)
    create_btn("Generate Report", generate_report, ACCENT, 0, 2)
    create_btn("Clear Table", lambda: tree.delete(*tree.get_children()), ACCENT, 0, 3)

    create_btn("Scatter Plot", scatter_plot, ACCENT, 1, 0)
    create_btn("Bar Chart", bar_all, ACCENT, 1, 1)
    create_btn("Top 5", bar_top5, ACCENT, 1, 2)
    create_btn("Pie Chart", pie_chart, ACCENT, 1, 3)

    create_btn("Hist Marks", histogram_marks, ACCENT, 2, 0)
    create_btn("Box Study", boxplot_study_hours, ACCENT, 2, 1)
    create_btn("Heatmap", correlation_heatmap, ACCENT, 2, 2)
    create_btn("Stacked Bar", stacked_bar_categories, ACCENT, 2, 3)

    root.mainloop()

if __name__ == "__main__":
    run_gui()
