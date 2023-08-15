import os
import re
import string
import textwrap
from collections import Counter

import PyPDF2
import tkinter as tk
import tkinter.ttk as ttk
from fpdf import FPDF
from tkinter import filedialog

# Center a window on the screen
def center_window(window):
    window.update_idletasks()

    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    window_width = window.winfo_width()
    window_height = window.winfo_height()

    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)

    window.geometry(f"{window_width}x{window_height}+{x}+{y}")


# Extract text from PDF
def extract_text_from_pdf(file_path, start_page=0, end_page=None, callback=None):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ''
        if end_page is None or end_page > len(reader.pages):
            end_page = len(reader.pages)
        for page in range(start_page, end_page + 1):
            text += reader.pages[page].extract_text()
            if callback:
                callback()
    return text

def count_words(text):
    words = re.findall(r'\b(?<![0-9])[a-zA-Z]+(?![0-9])\b', text.lower())
    unwanted_words = ["w", "j", "x", "q", "r", "s", "t", "v", "z", "k", "l", "m", "n", "p", "b", "c", "d", "f", "g", "h"]
    cleaned_words = [word for word in words if word not in unwanted_words and len(word) > 1]
    cleaned_words = [word for word in cleaned_words if not any(char.isdigit() for char in word)]
    cleaned_words = [word.strip(string.punctuation) for word in cleaned_words]
    return Counter(cleaned_words)

# Sort words by count
def sort_words_by_count(word_counts):
    return sorted(word_counts.items(), key=lambda x: x[1], reverse=True)

# Split a string into multiple lines based on a given maximum width
def split_string_into_lines(s, max_width):
    if len(s) <= max_width:
        return s
    
    lines = []
    while len(s) > max_width:
        # Find the last space within the max_width
        idx = s.rfind(' ', 0, max_width)
        if idx == -1:
            # If there's no space, break at the max_width
            idx = max_width
        lines.append(s[:idx])
        s = s[idx:].strip()
    if s:
        lines.append(s)
    return '\n'.join(lines)

def generate_pdf(sorted_word_counts, output_file):
    class PDF(FPDF):
        def header(self):
            self.set_font("Arial", "B", 10)
            self.cell(0, 10, "Words used in your book or selected pages", 0, 1, "C")

        def footer(self):
            self.set_y(-15)
            self.set_font("Arial", "I", 8)
            self.cell(0, 10, f"{self.page_no()}", 0, 0, "C")

    pdf = PDF()
    pdf.add_page()

    pdf.ln(10)
    pdf.set_fill_color(200, 220, 255)

    col_width_word = 60
    col_width_count = 15
    col_width_translation = 190 - (col_width_word + col_width_count)

    for word, count in sorted_word_counts:
        pdf.set_font("Arial", size=12)
        pdf.cell(col_width_word, 10, word, 1, 0, 'L', 1)
        pdf.set_font("Arial", size=8)
        pdf.cell(col_width_count, 10, f"({count})", 1, 0, 'L', 1)
        pdf.cell(col_width_translation, 10, "", 1, 1, 'L')

    pdf.output(output_file)

# Get page range from user
def get_page_range(total_pages):
    dialog = PageRangeDialog(app, total_pages)
    app.wait_window(dialog)
    return dialog.result

class PageRangeDialog(tk.Toplevel):
    def __init__(self, parent, total_pages):
        super().__init__(parent)
        self.title("Page Range")
        self.resizable(False, False)
        self.total_pages = total_pages
        self.result = None

        ttk.Label(self, text=f"Total pages in your file: {total_pages}.\n"
                             "You can choose to analyze a specific range of pages.\n"
                             "Enter start and end page separated by a comma or dash (e.g., 1-5) or 'all' for all pages.\n"
                             "Click 'OK' and then choose where to save the file.").pack(padx=12, pady=12)

        self.entry = ttk.Entry(self)
        self.entry.pack(padx=10, pady=10)

        btn_frame = ttk.Frame(self, style="TFrame")
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="OK", command=self.on_ok, takefocus=False).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.on_cancel, takefocus=False).pack(side="left", padx=5)

        center_window(self)

    def on_ok(self):
        pages_input = self.entry.get()
        if pages_input.lower() == 'all':
            self.result = (0, self.total_pages - 1)
        else:
            try:
                start_page, end_page = map(int, re.split(r'[-,]', pages_input))
                start_page -= 1
                end_page -= 1
                if 0 <= start_page < self.total_pages and 0 <= end_page < self.total_pages and start_page <= end_page:
                    self.result = (start_page, end_page)
                else:
                    ttk.messagebox.showwarning("Invalid Range", f"Please enter a valid page range between 1 and {self.total_pages}.")
                    return
            except:
                ttk.messagebox.showwarning("Invalid Format", "Invalid page range format. Please enter in the format 1-5 or 1,5.")
                return
        self.destroy()

    def on_cancel(self):
        self.destroy()


class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        default_font = ("Poppins", 10)
        style = ttk.Style()

        style.theme_use('clam')  # Available themes are 'clam', 'alt', 'default', 'classic'

        # Setting default widget styles (including font)
        style.configure('TFrame', background=self['background'])
        style.configure("TButton", font=default_font)
        style.configure("TLabel", font=default_font, padding=(5, 5))
        style.configure("TProgressbar", thickness=10, background="green")

        self.title("Book Word Analyzer")
        self.resizable(False, False) # prevent window from being resized        

        # Create the main menu
        self.menu = tk.Menu(self)

        # Add the 'About' menu item directly to the main menu
        self.menu.add_command(label="About", command=self.show_about)

        # Attach the menu to the main window
        self.config(menu=self.menu)


        # Browse PDF Button
        self.browse_button = ttk.Button(self, text="Browse PDF", command=self.browse_pdf, takefocus=False)
        self.browse_button.pack(pady=20)
        
        # Label to display chosen file
        self.file_label = ttk.Label(self, text="")
        self.file_label.pack(pady=5)
        
        # Button to start processing
        self.process_button = ttk.Button(self, text="Generate the list", command=self.process_pdf, takefocus=False, state=tk.DISABLED)
        self.process_button.pack(pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate", style="TProgressbar")

        self.progress.pack(pady=10, padx=10)
        style.configure("TProgressbar", thickness=10, background="green")  # Set the height/thickness to 20 pixels
        

    def show_about(self):
        # Disabling the 'About' menu item to prevent multiple instances
        self.menu.entryconfig("About", state="disabled")

        about_window = tk.Toplevel(self)
        about_window.title("About")
        about_window.resizable(False, False)

        # Re-enable the 'About' menu item once the 'About' window is closed
        about_window.protocol("WM_DELETE_WINDOW", lambda: [self.menu.entryconfig("About", state="normal"), about_window.destroy()])

        # Custom style to ensure consistent background color
        style = ttk.Style()
        style.configure("Clean.TLabel", background=about_window.cget('bg'))

        description = ("Book Word Analyzer is a project I initiated out   of personal interest.\n"
              "While reading a book in a foreign language, and grappling with its meaning, I was struck by an idea.\n"
              "Why not extract all the words and create a frequency list, along with a space for translations?\n"
              "With this list in hand, I could then craft flashcards, prioritizing the most frequent words.")


        # Ensure description doesn't exceed 50 characters per line for readability
        wrapped_description = textwrap.fill(description, width=50)

        author = "Author: Dmitri Markélov"
        
        # Using padx and pady for margins and Clean.TLabel style for a cleaner background
        ttk.Label(about_window, text=wrapped_description, style="Clean.TLabel").pack(padx=10, pady=10)
        ttk.Label(about_window, text=author, style="Clean.TLabel").pack(side="left", anchor="sw", padx=10)

        # Create hyperlinks for GitHub and LinkedIn inside a frame to group them together
        link_frame = ttk.Frame(about_window)
        link_frame.pack(anchor="e", padx=10)  # anchor to the right side

        github_link = "https://github.com/di-marko"
        linkedin_link = "https://www.linkedin.com/in/dmitri-mark%C3%A9lov/"

        # Using Clean.TLabel style for cleaner backgrounds
        github_label = ttk.Label(link_frame, text="GitHub", foreground="blue", cursor="hand2", style="Clean.TLabel")
        github_label.grid(row=0, column=0, padx=(0, 5))
        github_label.bind("<Button-1>", lambda e: self.open_link(github_link))

        linkedin_label = ttk.Label(link_frame, text="LinkedIn", foreground="blue", cursor="hand2", style="Clean.TLabel")
        linkedin_label.grid(row=0, column=1, padx=(5, 0))
        linkedin_label.bind("<Button-1>", lambda e: self.open_link(linkedin_link))

        center_window(about_window)  # center the about window


    def open_link(self, link):
        os.system(f"start \"\" \"{link}\"")

    def browse_pdf(self):
        self.input_path = filedialog.askopenfilename(title="Select the PDF you want to analyze", filetypes=[("PDF files", "*.pdf")])
        if self.input_path:
            processed_file_name = split_string_into_lines(os.path.basename(self.input_path), 25)
            self.file_label.config(text=processed_file_name)
            self.process_button.config(state=tk.NORMAL)


    def process_pdf(self):
        # Get total pages
        with open(self.input_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            total_pages = len(reader.pages)
        
        # Get the page range only once
        page_range = get_page_range(total_pages)
        if page_range is None:  # User canceled the dialog
            return  # or any other appropriate action
        start_page, end_page = page_range

        self.output_path = filedialog.asksaveasfilename(title="Select where to save the word list", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not self.output_path:
            return
            
        # Set progress bar maximum to the number of pages
        self.progress['maximum'] = end_page - start_page + 1
        
        try:
            # Use a lambda function to step the progress after each page is processed
            def step_progress():
                self.progress.step(1)
                self.update_idletasks()
                self.after(10)  # Wait for 10 milliseconds

            text = extract_text_from_pdf(self.input_path, start_page, end_page, callback=step_progress)

            word_counts = count_words(text)
            sorted_word_counts = sort_words_by_count(word_counts)
            generate_pdf(sorted_word_counts, self.output_path)
            
            tk.messagebox.showinfo("Success", f"PDF generated and saved to {self.output_path}!")
            open_pdf = tk.messagebox.askyesno("Open PDF", "Would you like to open the generated PDF?")
            if open_pdf:
                os.system(f"start \"\" \"{self.output_path}\"")
                
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred: {str(e)}")
            return

if __name__ == "__main__":
    app = Application()
    center_window(app)
    app.mainloop()
