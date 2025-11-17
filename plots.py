import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# ---------- Theme ----------
BG_COLOR = "#D3C8EE"
PANEL_COLOR = "#C2B8E6"
ACCENT = "#6C5CE7"
TEXT = "#111827"


class ExcelPlotter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Plotter - Professional")
        self.root.geometry("1300x760")
        self.root.configure(bg=BG_COLOR)

        self.df = None
        self.fig = None
        self.ax = None
        self.canvas = None
        self.bottom_title = None

        self.build_ui()

    # ---------------- UI ----------------
    def build_ui(self):
        header = tk.Frame(self.root, bg=BG_COLOR, pady=8)
        header.pack(fill="x")
        tk.Label(header, text="📊 Excel Plotter", bg=BG_COLOR, fg=TEXT,
                 font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=15)

        control_frame = tk.Frame(self.root, bg=PANEL_COLOR)
        control_frame.pack(fill="x", padx=15, pady=10)

        # Left side controls
        left = tk.Frame(control_frame, bg=PANEL_COLOR)
        left.pack(side="left", fill="y", padx=(10, 10), pady=12)

        tk.Button(left, text="Load Excel File", bg=ACCENT, fg="white",
                  font=("Segoe UI", 11, "bold"), command=self.load_excel).pack(pady=6)

        tk.Label(left, text="Select Columns to Plot:", bg=PANEL_COLOR, fg=TEXT).pack(anchor="w", pady=(12, 2))

        self.column_listbox = tk.Listbox(left, selectmode="multiple", width=28, height=10)
        self.column_listbox.pack()

        # Title + Labels
        tk.Label(left, text="Title:", bg=PANEL_COLOR).pack(anchor="w", pady=(10, 0))
        self.title_entry = tk.Entry(left, width=25)
        self.title_entry.pack()

        tk.Label(left, text="X Label (optional):", bg=PANEL_COLOR).pack(anchor="w", pady=(10, 0))
        self.xlabel_entry = tk.Entry(left, width=25)
        self.xlabel_entry.pack()

        tk.Label(left, text="Y Label (optional):", bg=PANEL_COLOR).pack(anchor="w", pady=(10, 0))
        self.ylabel_entry = tk.Entry(left, width=25)
        self.ylabel_entry.pack()

        # Title Position
        tk.Label(left, text="Title Position:", bg=PANEL_COLOR).pack(anchor="w", pady=(10, 0))
        self.pos_var = tk.StringVar(value="top")
        ttk.Radiobutton(left, text="Top", value="top", variable=self.pos_var).pack(anchor="w")
        ttk.Radiobutton(left, text="Bottom", value="bottom", variable=self.pos_var).pack(anchor="w")
        ttk.Radiobutton(left, text="Blank", value="blank", variable=self.pos_var).pack(anchor="w")

        # Title Color + picker
        color_row = tk.Frame(left, bg=PANEL_COLOR)
        color_row.pack(anchor="w", pady=(10, 0))

        tk.Label(color_row, text="Title Color:", bg=PANEL_COLOR).pack(side="left")
        self.color_entry = tk.Entry(color_row, width=10)
        self.color_entry.insert(0, "#000000")
        self.color_entry.pack(side="left", padx=5)

        tk.Button(color_row, text="🎨", command=self.pick_color,
                  bg="white", relief="raised", width=3).pack(side="left")

        # Title Size
        tk.Label(left, text="Title Size:", bg=PANEL_COLOR).pack(anchor="w", pady=(10, 0))
        self.size_entry = tk.Entry(left, width=10)
        self.size_entry.insert(0, "14")
        self.size_entry.pack(anchor="w")

        tk.Button(left, text="Plot Graph", bg="#2ecc71", fg="white",
                  font=("Segoe UI", 11, "bold"), command=self.plot_graph).pack(pady=10)

        # Right area (plot container)
        right = tk.Frame(control_frame, bg=PANEL_COLOR)
        right.pack(side="left", fill="both", expand=True, padx=12, pady=12)

        self.graph_container = tk.Frame(right, bg="white", bd=1, relief="solid")
        self.graph_container.pack(fill="both", expand=True)

    # ---------- Color Picker ----------
    def pick_color(self):
        color = colorchooser.askcolor()[1]
        if color:
            self.color_entry.delete(0, tk.END)
            self.color_entry.insert(0, color)

    # ---------------- Load Excel ----------------
    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not path:
            return

        try:
            self.df = pd.read_excel(path)
        except Exception as e:
            messagebox.showerror("Error", f"Unable to load file:\n{e}")
            return

        self.column_listbox.delete(0, tk.END)
        for col in self.df.columns:
            self.column_listbox.insert(tk.END, col)

        messagebox.showinfo("Loaded", "Excel file loaded successfully!")

    # ---------------- Plot ----------------
    def plot_graph(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Load an Excel file first.")
            return

        indices = self.column_listbox.curselection()
        if len(indices) < 2:
            messagebox.showwarning("Select Columns", "Select at least TWO columns.")
            return

        x_column = self.column_listbox.get(indices[0])
        y_columns = [self.column_listbox.get(i) for i in indices[1:]]

        # Clean previous plot
        if self.canvas:
            self.canvas.get_tk_widget().destroy()

        if hasattr(self, "edit_frame"):
            self.edit_frame.destroy()

        if self.bottom_title:
            self.bottom_title.remove()
            self.bottom_title = None

        self.fig, self.ax = plt.subplots(figsize=(12, 8))  # Larger height

        for col in y_columns:
            try:
                self.ax.plot(self.df[x_column], self.df[col], marker="o", label=col)
            except:
                self.ax.plot(pd.to_numeric(self.df[x_column], errors="coerce"),
                             pd.to_numeric(self.df[col], errors="coerce"),
                             marker="o", label=col)

        title = self.title_entry.get().strip()
        xlabel = self.xlabel_entry.get().strip()
        ylabel = self.ylabel_entry.get().strip()
        color = self.color_entry.get().strip()

        try:
            size = int(self.size_entry.get())
        except:
            size = 14

        pos = self.pos_var.get()

        # ---------- Title Logic ----------
        if self.bottom_title:
            self.bottom_title.remove()

        if pos == "top":
            self.ax.set_title(title, fontsize=size, color=color)
            self.ax.title.set_y(1.02)

        elif pos == "bottom":
            self.ax.set_title("")  # remove top title
            self.bottom_title = self.fig.text(
                0.5, 0.02, title,
                ha='center',
                fontsize=size,
                color=color
            )

        else:  # blank
            self.ax.set_title("")

        self.ax.set_xlabel(xlabel)
        self.ax.set_ylabel(ylabel)

        self.ax.grid(True)
        self.ax.legend()

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_container)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # Edit button
        self.edit_frame = tk.Frame(self.graph_container, bg="white")
        self.edit_frame.pack(fill="x")

        tk.Button(
            self.edit_frame, text="🖉 Edit", command=self.open_edit_dialog,
            bg="white", relief="flat", font=("Segoe UI", 11)
        ).pack(side="left", padx=5, pady=5)

    # ---------------- Edit Dialog ----------------
    def open_edit_dialog(self):
        if not self.ax:
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Graph")
        dialog.geometry("350x380")
        dialog.grab_set()

        cur_title = self.title_entry.get()
        cur_xlabel = self.xlabel_entry.get()
        cur_ylabel = self.ylabel_entry.get()

        # Title
        tk.Label(dialog, text="Title:").pack(anchor="w", padx=12, pady=(10, 2))
        title_e = tk.Entry(dialog, width=35)
        title_e.pack(padx=12)
        title_e.insert(0, cur_title)

        # Title Size
        tk.Label(dialog, text="Title Font Size:").pack(anchor="w", padx=12, pady=(8, 2))
        size_e = tk.Entry(dialog, width=12)
        size_e.pack(padx=12, anchor="w")
        size_e.insert(0, self.size_entry.get())

        # Title Color
        tk.Label(dialog, text="Title Font Color (#hex):").pack(anchor="w", padx=12, pady=(8, 2))
        color_e = tk.Entry(dialog, width=20)
        color_e.pack(padx=12)
        color_e.insert(0, self.color_entry.get())

        # Title Position
        tk.Label(dialog, text="Title Position:").pack(anchor="w", padx=12, pady=(10, 2))
        pos_var = tk.StringVar(value=self.pos_var.get())

        pos_frame = tk.Frame(dialog)
        pos_frame.pack(anchor="w", padx=12)

        tk.Radiobutton(pos_frame, text="Top", value="top", variable=pos_var).pack(side="left")
        tk.Radiobutton(pos_frame, text="Bottom", value="bottom", variable=pos_var).pack(side="left")
        tk.Radiobutton(pos_frame, text="Blank", value="blank", variable=pos_var).pack(side="left")

        # X label
        tk.Label(dialog, text="X Axis:").pack(anchor="w", padx=12, pady=(10, 2))
        x_e = tk.Entry(dialog, width=35)
        x_e.pack(padx=12)
        x_e.insert(0, cur_xlabel)

        # Y label
        tk.Label(dialog, text="Y Axis:").pack(anchor="w", padx=12, pady=(10, 2))
        y_e = tk.Entry(dialog, width=35)
        y_e.pack(padx=12)
        y_e.insert(0, cur_ylabel)

        def apply_changes():
            t = title_e.get().strip()
            x = x_e.get().strip()
            y = y_e.get().strip()
            color = color_e.get().strip()
            pos = pos_var.get()

            try:
                size = int(size_e.get())
            except:
                size = 14

            self.title_entry.delete(0, tk.END)
            self.title_entry.insert(0, t)

            self.xlabel_entry.delete(0, tk.END)
            self.xlabel_entry.insert(0, x)

            self.ylabel_entry.delete(0, tk.END)
            self.ylabel_entry.insert(0, y)

            self.color_entry.delete(0, tk.END)
            self.color_entry.insert(0, color)

            self.size_entry.delete(0, tk.END)
            self.size_entry.insert(0, str(size))

            self.pos_var.set(pos)

            dialog.destroy()
            self.plot_graph()

        tk.Button(dialog, text="Apply", bg="#2ecc71", fg="white",
                  width=12, command=apply_changes).pack(pady=15)


# ---------- Run ----------
root = tk.Tk()
app = ExcelPlotter(root)
root.mainloop()
