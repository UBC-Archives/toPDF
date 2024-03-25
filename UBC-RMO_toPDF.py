#toPDF
#Version: 0.1
#https://recordsmanagement.ubc.ca
#https://www.gnu.org/licenses/gpl-3.0.en.html

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import time
from datetime import datetime
import os
import threading
import sys
from docx import Document
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# Global flag to signal thread to stop
stop_thread_flag = False

def browse_button_callback():
    folder_path = filedialog.askdirectory()
    input_path_entry.delete(0, tk.END)
    input_path_entry.insert(0, folder_path)

def run_conversion():
    global conversion_thread
    if conversion_thread and conversion_thread.is_alive():
        messagebox.showinfo("Info", "Conversion is already in progress.")
    else:
        # Disable input field, browse button, and convert button
        input_path_entry.config(state='disabled')
        browse_button.config(state='disabled')
        convert_button.config(state='disabled')

        input_folder = input_path_entry.get()
        output_subfolder = "Access"

        # Create and start the conversion thread
        conversion_thread = threading.Thread(target=perform_conversion, args=(input_folder, output_subfolder))
        conversion_thread.daemon = True
        conversion_thread.start()

        # Start the spinner on the main window
        start_spinner()

def start_spinner():
    # Create and start the spinner
    spinner_frame.pack_forget()  # Remove any existing spinner frame
    spinner_frame.pack(pady=10)
    spinner.start()

def stop_spinner():
    # Stop and hide the spinner
    spinner.stop()
    spinner_frame.pack_forget()

    # Enable input field, browse button, and convert button after conversion is complete
    input_path_entry.config(state='normal')
    browse_button.config(state='normal')
    convert_button.config(state='normal')

def perform_conversion(input_folder, output_subfolder):
    try:
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        start_time = time.time()  # Record the start time
        success, message, docx_processed = convert_and_combine(input_folder, output_subfolder, timestamp)
        end_time = time.time()  # Record the end time
        conversion_time = end_time - start_time  # Calculate the conversion time

        if success:
            combined_pdf_path = os.path.join(input_folder, f'_Combined_{timestamp}.pdf')
            access_folder_path = os.path.join(input_folder, output_subfolder)

            # Generate word cloud PDF only if there were .docx files converted
            if docx_processed:  # Check if .docx files were processed
                generate_word_cloud_pdf(input_folder, output_subfolder, timestamp)
                wordcloud_pdf_path = os.path.join(input_folder, f'_Word-Cloud_{timestamp}.pdf')

            root.after(0, show_completion_message, message, conversion_time, combined_pdf_path, access_folder_path, wordcloud_pdf_path if docx_processed else None)
            cleanup_text_txt_file(input_folder, output_subfolder)
        else:
            root.after(0, messagebox.showerror, "Error", "Conversion failed.\n{}".format(message))
    except Exception as e:
        root.after(0, messagebox.showerror, "Error", "An error occurred during conversion:\n{}".format(str(e)))
    finally:
        # Stop the spinner after the conversion is complete or encounters an error
        root.after(0, stop_spinner)

def show_completion_message(message, conversion_time, combined_pdf_path, access_folder_path, wordcloud_pdf_path=None):
    # Create a new Toplevel window
    completion_window = tk.Toplevel(root)
    completion_window.title("Conversion Completed!")

    # Format the elapsed time using strftime without decimals
    conversion_time_str = time.strftime("%H:%M:%S", time.gmtime(conversion_time))

    # Label to display the completion message including formatted time
    completion_label_text = f"Conversion completed in {conversion_time_str}\n\nCombined PDF saved as {message}."
    if wordcloud_pdf_path:
        completion_label_text += f"\n\nWord cloud saved as {wordcloud_pdf_path}."
    completion_label = tk.Label(completion_window, text=completion_label_text)
    completion_label.pack(padx=20, pady=5)

    # Frame to hold the buttons
    button_frame = tk.Frame(completion_window)
    button_frame.pack(pady=10)

    # Button to open the combined PDF
    open_pdf_button = tk.Button(button_frame, text="Open Combined PDF", command=lambda: open_combined_pdf(combined_pdf_path))
    open_pdf_button.pack(side='left', padx=10)

    # Button to open the "Access" folder
    open_folder_button = tk.Button(button_frame, text="Open PDFs Folder", command=lambda: open_folder(access_folder_path))
    open_folder_button.pack(side='left', padx=10)

    # Conditionally add the "Open Word Cloud" button
    if wordcloud_pdf_path:
        open_word_cloud_button = tk.Button(button_frame, text="Open Word Cloud", command=lambda: open_combined_pdf(wordcloud_pdf_path))
        open_word_cloud_button.pack(side='left', padx=10)

def cleanup_text_txt_file(input_folder, output_subfolder):
    text_txt_path = os.path.join(input_folder, output_subfolder, '_text.txt')
    try:
        os.remove(text_txt_path)
    except OSError as e:
        print(f"Error deleting _text.txt file: {e.strerror}")

def open_combined_pdf(pdf_path):
    try:
        os.startfile(pdf_path)  # Opens the file with the default associated application
    except Exception as e:
        messagebox.showerror("Error", "Error opening the combined PDF:\n{}".format(str(e)))

def open_folder(folder_path):
    try:
        os.startfile(folder_path)  # Opens the folder with the default associated application
    except Exception as e:
        messagebox.showerror("Error", "Error opening the folder:\n{}".format(str(e)))

def exit_app():
    global stop_thread_flag

    if conversion_thread and conversion_thread.is_alive():
        confirm = tk.messagebox.askyesno("Exit?", "A process is running. Are you sure you want to exit?")
        if not confirm:
            return
        # Signal the thread to stop
        stop_thread_flag = True

    root.destroy()

#Conversion Logic
import os
from PIL import Image, UnidentifiedImageError
from PyPDF2 import PdfReader, PdfWriter
from docx2pdf import convert

def create_output_folder(folder_path, subfolder_name):
    output_folder = os.path.join(folder_path, subfolder_name)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    return output_folder

def is_image(filename, input_folder):
    file_path = os.path.join(input_folder, filename)
    try:
        with Image.open(file_path) as img:
            return True
    except (UnidentifiedImageError, FileNotFoundError, OSError):
        return False

def is_docx(filename, input_folder=None):
    return filename.lower().endswith('.docx')

def convert_image_to_pdf(input_image, output_pdf):
    image = Image.open(input_image)
    image.save(output_pdf, 'PDF', resolution=100.0)

def convert_docx_to_pdf(input_docx, output_pdf):
    convert(input_docx, output_pdf)

def convert_files_to_pdf(input_folder, filter_function, output_subfolder):
    global stop_thread_flag
    output_folder = create_output_folder(input_folder, output_subfolder)
    pdf_paths = []

    # Open or create the _text.txt file in the output folder
    text_txt_path = os.path.join(output_folder, '_text.txt')
    with open(text_txt_path, 'w', encoding='utf-8') as text_file:
        for filename in os.listdir(input_folder):
            if stop_thread_flag:
                break
            if filter_function(filename, input_folder):  # Pass input_folder as argument
                file_path = os.path.join(input_folder, filename)
                pdf_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '_' + os.path.splitext(filename)[1][1:] + ".pdf")

                if is_image(filename, input_folder):  # Pass input_folder as argument
                    convert_image_to_pdf(file_path, pdf_path)
                elif is_docx(filename):
                    convert_docx_to_pdf(file_path, pdf_path)
                    # Extract and write text to _text.txt
                    for line in extract_text_from_docx(file_path):
                        text_file.write(line)
                        text_file.write('\n')  # Ensure spacing between documents

                pdf_paths.append(pdf_path)
    
    return pdf_paths

def combine_pdfs(pdf_paths, output_file):
    pdf_merger = PdfWriter()

    pdf_paths.sort()

    for pdf_path in pdf_paths:
        pdf_reader = PdfReader(pdf_path)
        for page_num in range(len(pdf_reader.pages)):
            pdf_merger.add_page(pdf_reader.pages[page_num])

    with open(output_file, 'wb') as output_pdf:
        pdf_merger.write(output_pdf)

def convert_and_combine(input_folder, output_subfolder, timestamp):
    pdf_paths_images = convert_files_to_pdf(input_folder, is_image, output_subfolder)
    pdf_paths_docx = convert_files_to_pdf(input_folder, is_docx, output_subfolder)

    pdf_paths = pdf_paths_images + pdf_paths_docx

    docx_processed = len(pdf_paths_docx) > 0  # True if any .docx files were processed

    if pdf_paths:
        combined_pdf_path = os.path.join(input_folder, f'_Combined_{timestamp}.pdf')
        combine_pdfs(pdf_paths, combined_pdf_path)
        return True, combined_pdf_path, docx_processed
    else:
        return False, "No files to convert. Combined file not created.", docx_processed

# Word Cloud
def extract_text_from_docx(docx_path):
    try:
        doc = Document(docx_path)
        for para in doc.paragraphs:
            yield para.text + '\n'
    except Exception as e:
        print(f"Error reading {docx_path}: {str(e)}")
        yield ''

def generate_word_cloud_pdf(input_folder, output_subfolder, timestamp):
    text_txt_path = os.path.join(input_folder, output_subfolder, '_text.txt')
    wordcloud_pdf_path = os.path.join(input_folder, f'_Word-Cloud_{timestamp}.pdf')

    # Read the whole text from file
    with open(text_txt_path, 'r', encoding='utf-8') as file:
        text = file.read()

    # Generate a word cloud image
    wordcloud = WordCloud(width=800, height=800, background_color='white', min_font_size=10).generate(text)

    # Display the generated image
    plt.figure(figsize=(8, 8), facecolor=None)
    plt.imshow(wordcloud, interpolation="bilinear")
    plt.axis("off")
    plt.tight_layout(pad=0)

    # Save the image in PDF format
    plt.savefig(wordcloud_pdf_path, format="pdf")
    plt.close()  # Close the figure to release memory

# File menu functions
def exit_app():
    if threading.active_count() > 1:
        confirm = tk.messagebox.askyesno("Exit?", "A process is running. Are you sure you want to exit?")
        if not confirm:
            return
    root.destroy()

def clear_fields():
    # Enable the Path entry
    input_path_entry.config(state=tk.NORMAL)

    # Clear the fields
    input_path_entry.delete(0, tk.END)

def show_help():
    help_message = "This program converts files in a directory to PDF format.\n\n" \
                   "Use 'Browse' to select the directory path containing the files you want to convert, " \
                   "then click 'Convert' to start the conversion process.\n\n" \
                   "The converted PDF files will be saved in a folder named 'Access' within the selected directory.\n\n" \
                   "You can also generate a word cloud PDF from the text extracted from .docx files.\n\n" \
                   "Once the conversion is complete, you can open the combined PDF and the 'Access' folder " \
                   "containing the converted PDFs.\n\n" \
                   "Note: Ensure the files you want to convert are accessible and supported by the program."
    tk.messagebox.showinfo("Help", help_message)

def show_about():
    about_message = "toPDF\nVersion 0.1"
   
    about_window = tk.Toplevel(root)
    about_window.title("About")
    about_window.resizable(False, False)

    about_label = tk.Label(about_window, text=about_message)
    about_label.pack(padx=20, pady=10)

    # Frame for UBC RMO
    ubc_frame = tk.Frame(about_window)
    ubc_frame.pack(padx=20, pady=20)

    ubc_label = tk.Label(ubc_frame, text="Developed by\nRecords Management Office\nThe University of British Columbia")
    ubc_label.pack()

    ubc_link_label = tk.Label(ubc_frame, text="https://recordsmanagement.ubc.ca", fg="blue", cursor="hand2")
    ubc_link_label.pack()

    def open_ubc_link(event):
        import webbrowser
        webbrowser.open("https://recordsmanagement.ubc.ca")

    ubc_link_label.bind("<Button-1>", open_ubc_link)

    # Frame for license
    license_frame = tk.Frame(about_window)
    license_frame.pack(padx=20, pady=10)

    license_label = tk.Label(license_frame, text="License: ")
    license_label.pack(side='left')

    license_link_label = tk.Label(license_frame, text="GPL-3.0", fg="blue", cursor="hand2")
    license_link_label.pack(side='left')

    def open_license_link(event):
        import webbrowser
        webbrowser.open("https://www.gnu.org/licenses/gpl-3.0.en.html")

    license_link_label.bind("<Button-1>", open_license_link)

root = tk.Tk()
root.title("toPDF")

# Set window size to 600x150 pixels and make it non-resizable
root.geometry("600x150")
root.resizable(False, False)

# Menu bar
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# File menu
file_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)

# File menu options
file_menu.add_command(label="Clear Fields", command=clear_fields)
file_menu.add_command(label="Exit", command=exit_app)

# Help menu
help_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Help", menu=help_menu)

# Help menu options
help_menu.add_command(label="Help", command=show_help)
help_menu.add_command(label="About", command=show_about)

# Frame to hold the path input field and the "Browse" button
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Input field for the path
input_label = tk.Label(frame, text="Directory Path:")
input_label.pack(side='left')

input_path_entry = tk.Entry(frame, width=60)
input_path_entry.pack(side='left', padx=10)

# "Browse" button
browse_button = tk.Button(frame, text="Browse", command=browse_button_callback)
browse_button.pack(side='left', padx=10)

# "Convert" button
convert_button = tk.Button(root, text="Convert", command=run_conversion)
convert_button.pack(pady=10)

# Initialize the conversion thread
conversion_thread = None

# Spinner
spinner_frame = tk.Frame(root)
spinner_frame.pack_forget()  # Initially hide the spinner frame
spinner = ttk.Progressbar(spinner_frame, mode='indeterminate', length=500)
spinner.pack()

# Set the window close event handler
root.protocol("WM_DELETE_WINDOW", exit_app)

# Run the Tkinter main loop
root.mainloop()
