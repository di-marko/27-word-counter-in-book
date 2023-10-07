import PyPDF2
import re
import string
from fpdf import FPDF
from collections import Counter


# Extract text from PDF
def extract_text_from_pdf(file_path, start_page=0, end_page=None, callback=None):
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        if end_page is None or end_page > len(reader.pages):
            end_page = len(reader.pages)
        for page in range(start_page, end_page + 1):
            text += reader.pages[page].extract_text()
            if callback:
                callback()
    return text


def count_words(text):
    """
    Count the frequency of each word in the text, excluding unwanted words.
    Parameters:
    - text (str): The input text to analyze.
    Returns:
    Counter: A Counter object with words as keys and their corresponding frequency as values.
    """
    # Define words that are to be excluded from the count.
    unwanted_words = set("w j x q r s t v z k l m n p b c d f g h".split())

    # Use regex to extract all words, converting them to lowercase.
    words = re.findall(r"\b(?<![0-9])[a-zA-Z]+(?![0-9])\b", text.lower())

    # Remove unwanted words and words with length 1.
    cleaned_words = [
        word for word in words if word not in unwanted_words and len(word) > 1
    ]

    # Remove any digits and punctuation from the words.
    cleaned_words = [word.strip(string.punctuation) for word in cleaned_words]

    return Counter(cleaned_words)


# Sort words by count
def sort_words_by_count(word_counts):
    return sorted(word_counts.items(), key=lambda x: x[1], reverse=True)


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
        pdf.cell(col_width_word, 10, word, 1, 0, "L", 1)
        pdf.set_font("Arial", size=8)
        pdf.cell(col_width_count, 10, f"({count})", 1, 0, "L", 1)
        pdf.cell(col_width_translation, 10, "", 1, 1, "L")

    pdf.output(output_file)


# Split a string into multiple lines based on a given maximum width
def split_string_into_lines(s, max_width):
    if len(s) <= max_width:
        return s

    lines = []
    while len(s) > max_width:
        # Find the last space within the max_width
        idx = s.rfind(" ", 0, max_width)
        if idx == -1:
            # If there's no space, break at the max_width
            idx = max_width
        lines.append(s[:idx])
        s = s[idx:].strip()
    if s:
        lines.append(s)
    return "\n".join(lines)
