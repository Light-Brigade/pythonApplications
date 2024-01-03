import tkinter as tk
from tkinter import filedialog, messagebox, ttk 
from weasyprint import HTML
from docx import Document
from threading import Thread
import os
import time

class PDFConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("PDF Converter")

        # Calculate the center position of the screen
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        x_position = (screen_width - 600) // 2  # 600 is the width of the window
        y_position = (screen_height - 400) // 2  # 400 is the height of the window

        # Set the geometry to position the window at the center
        master.geometry(f"600x400+{x_position}+{y_position}")

        self.label_input = tk.Label(master, text="Input File:", font=("Helvetica", 12))
        self.label_input.pack(pady=5)  # Add padding between the label and the input field

        self.entry_input = tk.Entry(master, width=50, font=("Helvetica", 12))
        self.entry_input.pack(pady=5)

        self.button_browse_input = tk.Button(master, text="Browse", command=self.browse_input, font=("Helvetica", 12, "bold"))
        self.button_browse_input.pack(pady=5)

        self.label_output = tk.Label(master, text="Output File:", font=("Helvetica", 12))
        self.label_output.pack(pady=5)  # Add padding between the label and the output field

        self.entry_output = tk.Entry(master, width=50, font=("Helvetica", 12))
        self.entry_output.pack(pady=5)

        self.button_browse_output = tk.Button(master, text="Browse", command=self.browse_output, font=("Helvetica", 12, "bold"))
        self.button_browse_output.pack(pady=5)

        self.button_convert = tk.Button(master, text="Convert", command=self.convert, font=("Helvetica", 16, "bold"))
        self.button_convert.pack(pady=20)  # Add padding below the Convert button

        self.progress_label = tk.Label(master, text="", font=("Helvetica", 14, "bold"), foreground="blue")

        # Additional components to be lazily loaded
        self.lazy_loaded_component = None

    def browse_input(self):
        file_path = filedialog.askopenfilename(filetypes=[("Supported Files", "*.htm;*.txt;*.docx")])
        self.entry_input.delete(0, tk.END)
        self.entry_input.insert(0, file_path)

        # Lazily load additional components if not loaded
        if self.lazy_loaded_component is None:
            self.lazy_loaded_component = self.load_lazy_component()

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        self.entry_output.delete(0, tk.END)
        self.entry_output.insert(0, file_path)

        # Lazily load additional components if not loaded
        if self.lazy_loaded_component is None:
            self.lazy_loaded_component = self.load_lazy_component()

    def convert(self):
        input_file = self.entry_input.get()
        output_file = self.entry_output.get()

        if not input_file or not output_file:
            messagebox.showwarning("Error", "Please select input and output files.")
            return

        self.button_convert.config(state=tk.DISABLED)
        self.progress_label.config(text="Converting File to PDF...", foreground="blue")
        self.progress_label.pack()

        # Start a separate thread to perform the conversion
        conversion_thread = Thread(target=self.perform_conversion, args=(input_file, output_file))
        conversion_thread.start()

    def load_lazy_component(self):
        # Use threading to load the component
        thread = Thread(target=self.load_lazy_component_thread)
        thread.start()

    def load_lazy_component_thread(self):
        # Simulate a time-consuming process (replace this with actual logic)
        time.sleep(5)
        print("Lazy component loaded.")

    def perform_conversion(self, input_file, output_file):
        try:
            if os.path.exists(output_file):
                overwrite = messagebox.askyesno("File Exists", "The output file already exists. Do you want to overwrite it?")
                if not overwrite:
                    self.master.after(0, lambda: self.on_conversion_error("Conversion canceled. Output file already exists."))
                    return

            convert_to_pdf(input_file, output_file)
            self.master.after(0, self.on_conversion_complete)
        except FileNotFoundError as e:
            self.master.after(0, lambda: self.on_conversion_error(f"File not found: {e}"))
        except Exception as e:
            self.master.after(0, lambda: self.on_conversion_error(f"Error: {e}"))

    def on_conversion_complete(self):
        self.progress_label.config(text="Conversion successful. PDF saved.", foreground="green")
        self.button_convert.config(state=tk.NORMAL)
        self.progress_label.pack_forget()  # Hide the label

        messagebox.showinfo("Conversion Complete", "Conversion successful. PDF saved.")

    def on_conversion_error(self, error):
        self.progress_label.config(text=f"Error: {error}", foreground="red")
        self.button_convert.config(state=tk.NORMAL)
        self.progress_label.pack_forget()  # Hide the label

        messagebox.showerror("Conversion Error", f"Error: {error}")

def convert_to_pdf(input_file, output_file):
    if input_file.lower().endswith('.html') or input_file.lower().endswith('.htm'):
        convert_html_to_pdf(input_file, output_file)
    elif input_file.lower().endswith('.docx'):
        convert_docx_to_pdf(input_file, output_file)
    else:
        raise ValueError("Unsupported file format")

def convert_html_to_pdf(input_file, output_file):
    # Specify an empty list for stylesheets to avoid the creation of additional folders
    HTML(input_file).write_pdf(output_file, stylesheets=[])


def convert_docx_to_pdf(input_file, output_file):
    doc = Document(input_file)
    with open(output_file, 'wb') as f:
        doc.save(f, format='pdf')

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()
