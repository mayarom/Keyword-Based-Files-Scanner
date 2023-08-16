import os
from docx import Document
import tkinter as tk
from tkinter import ttk, filedialog
from ttkthemes import ThemedTk


def extract_text_from_docx(file_path, keywords):
    doc = Document(file_path)
    full_text = [paragraph.text for paragraph in doc.paragraphs]
    text = '\n'.join(full_text)
    found_keywords = []

    for keyword in keywords:
        if keyword in text:
            found_keywords.append(keyword)

    return found_keywords, text


def wrap_text(text, max_line_length):
    words = text.split()
    lines = []
    space_padding = "                    "  # This is a string with spaces
    current_line = space_padding  # Start with the padding

    for word in words:
        # Check if adding the word to the current line would exceed the max line length
        if len(current_line + word + " ") <= max_line_length:
            current_line += word + " "
        else:
            lines.append(current_line.strip())  # Append the current line to the list of lines
            current_line = space_padding + word + " "  # Start a new line with the padding, followed by the word

    if current_line.strip():  # If there's any content left in current_line
        lines.append(current_line.strip())

    return '\n'.join(lines)


def select_directory_and_scan(text_widget, status_label):
    print("Function called!")  # Debugging print

    keywords = keywords_entry.get().split(',')
    directory = filedialog.askdirectory()

    print(f"Directory selected: {directory}")  # Debugging print

    keyword_counts = {keyword: 0 for keyword in keywords}

    if directory:
        text_widget.delete(1.0, tk.END)

        file_contents = []  # Temporary list to hold file contents

        for filename in os.listdir(directory):
            print(f"Processing file: {filename}")  # Debugging print

            if filename.endswith(".docx"):
                full_path = os.path.join(directory, filename)
                try:
                    found_keywords, extracted_text = extract_text_from_docx(full_path, keywords)

                    # Only process the file if any keywords are found
                    if found_keywords:
                        for keyword in found_keywords:
                            keyword_counts[keyword] += 1

                        wrapped_text = wrap_text(extracted_text, max_line_length=130)

                        # Prepare file content to append
                        file_data = f"---------------------------------------------------------------------------" \
                                    f"------------------------------------------------------------------------\n{filename}:\n" \
                                    f"Keywords found: {', '.join(found_keywords)}\n\n{wrapped_text}\n\n"
                        lines = file_data.split('\n')
                        for index, line in enumerate(lines):
                            if index == 0:  # Color the filename red
                                result_text.insert(tk.END, "                    " + line + '\n', ("red", "center"))
                            else:  # Rest of the content
                                result_text.insert(tk.END, "                    " + line + '\n', ("right", "center"))

                except Exception as e:
                    print(f"Error processing file {filename}. Error: {e}")

        summary = "Summary:\n"
        for keyword, count in keyword_counts.items():
            summary += f"We found the keyword '{keyword}' in {count} files.\n"

        # First insert the summary
        text_widget.insert(tk.END, summary + "\n\n", ("bold", "center"))


# Set up the main window
root = ThemedTk(theme="equilux")
root.title("Keyword-Based Files Scanner")
root.geometry("1300x900")

# Define styles and colors
bg_color = "#2E3440"  # Dark background
font_color = "#ECEFF4"  # Light font color
font_style = ("Ubuntu", 14)
header_font_style = ("Ubuntu", 20, "bold")
button_font_style = ("Ubuntu", 14)

# Styling frames and labels for a consistent look
style = ttk.Style()
style.configure('TFrame', background=bg_color)
style.configure('TLabel', background=bg_color, foreground=font_color)

# Main content frame
main_frame = ttk.Frame(root, padding="10")
main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)  # Reduced padding for top balance

# Display header
header_label = ttk.Label(main_frame, text="Keyword-Based Files Scanner", font=header_font_style)
header_label.pack(pady=10)  # Reduced padding for balance

# User input for keywords
keywords_entry = ttk.Entry(main_frame, font=font_style, width=100)
keywords_entry.insert(0, "Enter keywords separated by commas")
keywords_entry.pack(pady=20)

# Button to initiate file scanning
button = tk.Button(main_frame, text="Select a Folder", font=button_font_style,
                   bg="#5E81AC", fg=font_color, activebackground="#88C0D0", relief=tk.FLAT, padx=20, pady=10,
                   command=lambda: select_directory_and_scan(result_text, status_label))
button.pack(pady=30)

result_text = tk.Text(main_frame, wrap=tk.WORD, height=20, width=110, font=font_style, bg="#3B4252", fg=font_color,
                      spacing2=9)
result_text.pack(padx=10, pady=10, side=tk.LEFT, fill=tk.BOTH, expand=True)

# Scrollbar for result area
scroll = tk.Scrollbar(main_frame, command=result_text.yview, bg="#4C566A")
result_text.configure(yscrollcommand=scroll.set, relief=tk.FLAT)
scroll.pack(side=tk.RIGHT, fill=tk.Y)

# Configure tags for specific text styles in result_text
result_text.tag_configure("blue_header", foreground="#81A1C1", justify="left")
result_text.tag_configure("bold", font=("Ubuntu", 14, "bold"))
result_text.tag_configure("red", foreground="#BF616A")
result_text.tag_configure("right", justify="right")
result_text.tag_configure("center", justify="center")
result_text.tag_configure("spacing", spacing1=10, spacing3=10)  # Add space above and below each line

# Label to display the status of scanning
status_label = ttk.Label(main_frame, text="", font=font_style)
status_label.pack(pady=20)

# Credit label
credit_label = ttk.Label(root, text="Developed by Maya Rom", font=("Ubuntu", 10, "italic"))
credit_label.pack(side=tk.BOTTOM, pady=5)

root.mainloop()
