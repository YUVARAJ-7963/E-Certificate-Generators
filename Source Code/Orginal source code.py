import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageTk
import tkinter as tk
from tkinter import filedialog, messagebox

class Certificate_Generator:
    def __init__(self, root):
        self.root = root
        self.root.title("Certificate Generator")
        self.root.geometry("1366x768")
        self.font_path = 'arial.ttf'
        self.font_size = 20
        self.positions = {}
        self.template_image = None
        self.file_path = None
        self.dragging = False
        self.dragged_heading = None
        self.dragged_label = None
        self.headings_labels = {}
        self.headings_labels2 = {}
        self.zoom_level = 1.0

        # scroll
        self.canvas = tk.Canvas(root, bg='#80ecbf')
        self.scroll_y = tk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scroll_x = tk.Scrollbar(root, orient="horizontal", command=self.canvas.xview)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.configure(yscrollcommand=self.scroll_y.set, xscrollcommand=self.scroll_x.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # main frame
        self.main_frame = tk.Frame(self.canvas, bg='#80ecbf')
        self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")

        # top heading (1st row)
        heading_label = tk.Label(self.main_frame, text="Certificate Generator", font=("Arial", 24), fg='black', bg='#80ecbf')
        heading_label.grid(row=0, column=0, columnspan=3, pady=10)

        # Buttons (2nd row)
        self.buttons_frame = tk.Frame(self.main_frame, bg='#80ecbf')
        self.buttons_frame.grid(row=1, column=0, padx=10, pady=10, columnspan=3)

        # Template Button
        self.load_template_button = tk.Button(self.buttons_frame, text="Load Template", command=self.load_template, bg='orange', fg='white', font=("Arial", 12), pady=5, padx=10)
        self.load_template_button.pack(side=tk.LEFT, padx=5)

        # Excel Button
        self.load_excel_button = tk.Button(self.buttons_frame, text="Load Excel File", command=self.load_excel_file, bg='green', fg='white', font=("Arial", 12), pady=5, padx=10)
        self.load_excel_button.pack(side=tk.LEFT, padx=5)

        # Generate Button
        self.convert_button = tk.Button(self.buttons_frame, text="Generate Certificte", command=self.set_headings, bg='blue', fg='white', font=("Arial", 12), pady=5, padx=10)
        self.convert_button.pack(side=tk.LEFT, padx=5)

        # Reset Button
        self.reset_button = tk.Button(self.buttons_frame, text="Reset", command=self.reset_headings, bg='red', fg='white', font=("Arial", 12), pady=5, padx=10)
        self.reset_button.pack(side=tk.LEFT, padx=5)

        # file and template name
        self.excel_name_label = tk.Label(self.buttons_frame, text="Excel File: Not Selected", font=("Arial", 12), fg='black', bg='#80ecbf')
        self.excel_name_label.pack(side=tk.LEFT, padx=5)
        self.template_name_label = tk.Label(self.buttons_frame, text="Template: Not Loaded", font=("Arial", 12), fg='black', bg='#80ecbf')
        self.template_name_label.pack(side=tk.LEFT, padx=5)

        # Excel Headings (3rd row)
        self.headings_frame = tk.Frame(self.main_frame, bg='#80ecbf')
        self.headings_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=10, columnspan=3)

        self.headings_canvas = tk.Canvas(self.headings_frame, height=45, bg='#80ecbf')
        self.headings_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.headings_inner_frame = tk.Frame(self.headings_canvas, bg='#80ecbf')
        self.headings_canvas.create_window((0, 0), window=self.headings_inner_frame, anchor="nw")

        # Template (4th row)
        self.template_frame = tk.Frame(self.main_frame, bg='#80ecbf')
        self.template_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=10, columnspan=3)

        self.template_canvas_frame = tk.Frame(self.template_frame, bg='#80ecbf')
        self.template_canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.template_canvas = tk.Canvas(self.template_canvas_frame, bg='#80ecbf')
        self.template_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.template_label = tk.Label(self.template_canvas, text="No Template Loaded", font=("Arial", 20), fg='Black', bg='#80ecbf')
        self.template_label.pack(expand=True)

        # Developer Name (5th row)
        developer_label = tk.Label(self.main_frame, text="@Developers: Yuvaraj", font=("Arial", 10), fg='black', bg='#80ecbf')
        developer_label.grid(row=4, column=0, pady=10, columnspan=3)


        self.canvas.bind("<Configure>", self.on_canvas_configure)

    def on_canvas_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def load_template(self):
        template_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpeg;*.jpg;*.png")])
        if template_path:
            self.template_image = Image.open(template_path)
            self.template_photo = ImageTk.PhotoImage(self.template_image)
            self.template_name_label.config(text=f"Template: {os.path.basename(template_path)}")
            self.template_label.config(image=self.template_photo, text="")
            self.update_canvas()

    def load_excel_file(self):
        if not self.template_image:
            messagebox.showerror("Error", "Please load a template first.")
            return
        
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file_path:
            self.excel_name_label.config(text=f"Excel File: {os.path.basename(self.file_path)}")
            df = pd.read_excel(self.file_path)
            headings = list(df.columns)
            for widget in self.headings_inner_frame.winfo_children():
                widget.destroy()

            # load the excel headings
            for heading in headings:
                label = tk.Label(self.headings_inner_frame, text=heading, bg='yellow', font=("Arial", 12), relief="raised", padx=5, pady=5)
                label.pack(side=tk.LEFT, padx=5, pady=5)
                label.bind("<ButtonPress-1>", self.start_drag)

            self.headings_canvas.update_idletasks()
            width = self.headings_inner_frame.winfo_reqwidth()
            self.headings_canvas.config(scrollregion=self.headings_canvas.bbox("all"))
            self.headings_inner_frame.place(x=(self.headings_canvas.winfo_width() - width) / 2 , y=5)


    def start_drag(self, event):
        widget = event.widget
        self.dragged_heading = widget.cget("text")
        self.dragging = True

        if self.dragged_heading in self.headings_labels2:
            self.dragged_label = self.headings_labels2[self.dragged_heading]
        else:
            self.dragged_label = tk.Label(self.template_canvas, text=self.dragged_heading, bg='yellow', font=("Arial", 12), relief="raised")
            self.headings_labels2[self.dragged_heading] = self.dragged_label        #load a selected excel heading on template
            self.dragged_label.place(x=event.x+125, y=event.y)

        if self.dragged_label.winfo_exists():
            self.dragged_label.lift()
            self.dragged_label.bind("<B1-Motion>", self.drag)
            self.dragged_label.bind("<ButtonRelease-1>", self.drop)

        self.drag_data = {"x": event.x, "y": event.y}

    def drag(self, event):
        if self.dragging and self.dragged_label:
            new_x = event.x_root - self.template_canvas.winfo_rootx()
            new_y = event.y_root - self.template_canvas.winfo_rooty()
            new_x = max(0, min(new_x, self.template_canvas.winfo_width() - self.dragged_label.winfo_width()))
            new_y = max(0, min(new_y, self.template_canvas.winfo_height() - self.dragged_label.winfo_height()))
            self.dragged_label.place(x=new_x, y=new_y)

    def drop(self, event):
        if self.dragging and self.dragged_label:
            x = self.dragged_label.winfo_x()
            y = self.dragged_label.winfo_y()
            self.positions[self.dragged_heading] = (x, y)
            self.dragging = False
            self.dragged_label.bind("<ButtonPress-1>", self.start_drag)
            self.dragged_label.bind("<B1-Motion>", self.drag)
            self.dragged_label.bind("<ButtonRelease-1>", self.drop)
            self.dragged_label = None

    def reset_headings(self):
        for widget in self.template_canvas.winfo_children():
            if widget != self.template_label:
                widget.destroy()
        self.positions.clear()
        self.headings_labels2.clear()
        self.dragged_label = None

    def update_canvas(self):
        self.template_canvas.update_idletasks()
        self.template_canvas.config(scrollregion=self.template_canvas.bbox("all"))
        for heading, (x, y) in self.positions.items():
            if heading in self.headings_labels:
                label = self.headings_labels[heading]
                label.place(x=x / self.zoom_level, y=y / self.zoom_level)

    def set_headings(self):
        if not self.file_path or not self.template_image:
            messagebox.showerror("Error", "Please load a template and Excel file first.")
            return

        df = pd.read_excel(self.file_path)
        output_folder = os.path.splitext(os.path.basename(self.file_path))[0]
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        for index, row in df.iterrows():
            certificate = self.template_image.copy()
            draw = ImageDraw.Draw(certificate)
            try:
                font = ImageFont.truetype(self.font_path, self.font_size)
                for heading, position in self.positions.items():
                    adjusted_position = (int(position[0] * self.zoom_level), int(position[1] * self.zoom_level))
                    draw.text(adjusted_position, str(row[heading]), fill='black', font=font)
                certificate_path = os.path.join(output_folder, f"certificate_{index + 1}.png")
                certificate.save(certificate_path)
            except Exception as e:
                messagebox.showerror("Error", f"Error generating certificate: {e}")
                return

        messagebox.showinfo("Success", f"Certificates generated successfully in {output_folder}")

def main():
    root = tk.Tk()
    app = Certificate_Generator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
