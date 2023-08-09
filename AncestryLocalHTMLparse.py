import tkinter as tk
from tkinter import filedialog, messagebox
from bs4 import BeautifulSoup
import threading
import csv
import tkinter.ttk as ttk
import subprocess

# List of required packages
required_packages = ["beautifulsoup4", "tkinter", "csv"]

# Function to check and install missing packages
def install_missing_packages(packages):
    for package in packages:
        try:
            __import__(package)
        except ImportError:
            print(f"{package} is not installed. Installing...")
            subprocess.check_call(["pip", "install", package])

# Global variable to store the selected HTML file path
html_file = None

# Function to open an HTML file
def open_file():
    global html_file
    
    # Temporarily set the root window to be on top
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename()
    root.attributes("-topmost", False)

    if file_path:
        html_file = file_path
        selected_file_label.config(text="Selected file: " + html_file)

# Function to start parsing the HTML file
def start_parsing():
    if html_file:
        start_button.config(state=tk.DISABLED)
        processing_label.config(text="Processing...", foreground="blue")
        parse_thread = threading.Thread(target=parse_and_save_data, args=(html_file,))
        parse_thread.start()

# Function to update processing label
def update_processing_label():
    processing_label.config(text="Completed!", foreground="green")
    start_button.config(state=tk.NORMAL)

# Function to parse and save data from the HTML file
def parse_and_save_data(html_file):
    try:
        with open(html_file, 'r') as file:
            html_content = file.read()

        soup = BeautifulSoup(html_content, 'html.parser')

        # Find all match-entry elements in the HTML
        match_entry_elements = soup.find_all('match-entry')

        # Extract the test name from the HTML content
        testname_element = soup.find('span', class_='userCardContent hideVisually768 navRestrictedName').text
        testname = testname_element.strip()

        csv_file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if not csv_file:
            messagebox.showinfo("Information", "Data export canceled.", icon="info")
            return

        # Data extraction and processing logic
        extracted_data = []
        for match_entry in match_entry_elements:
            name_element = match_entry.select_one('h3 a.userCardTitle')
            name = name_element.text.strip() if name_element else None

            link_element = match_entry.select_one('a.userCardImg')
            link_href = link_element['href'] if link_element else None
            
            # Extract 'id_value' and 'testid2' from the 'href' attribute
            if link_href:
                link_parts = link_href.split('/')
                testid = link_parts[-3] if len(link_parts) >= 3 else None
                id_value = link_parts[-1] if link_parts else None
            else:
                testid, id_value = None, None

            shared_dna_element = match_entry.select_one('button.sharedDnaText')
            shared_dna_text = shared_dna_element.text.strip().split(' | ')[0] if shared_dna_element else None
            shared_cm = shared_dna_text.replace(" cM", "")
            
            side_element = match_entry.select_one('span.parentLineText')
            side = side_element.text.strip() if side_element else None

            tree_size_element = match_entry.select_one('div.treeSizeText')
            tree_size = tree_size_element.text.strip() if tree_size_element else None

            # Append extracted data as a tuple to the list
            extracted_data.append((testid, testname, id_value, name, shared_cm, side, tree_size))

        if csv_file:
            # Write extracted data to CSV file
            with open(csv_file, 'w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(["Test_ID", "Test_Name", "Match_ID", "Match_Name", "Shared_CM", "Side", "Tree_Size"])
                writer.writerows(extracted_data)

            # Show success message
            messagebox.showinfo("Information", "Data has been successfully saved to " + csv_file, icon="info")

    except Exception as e:
        # Handle errors during processing
        messagebox.showerror("Error", "An error occurred: " + str(e), icon="error")

    finally:
        # Update processing label after a delay
        processing_label.after(1000, update_processing_label)

# Function to create the GUI
def create_gui():
    global root, selected_file_label, processing_label, start_button

    root = tk.Tk()
    root.title("Ancestry HTML Parser")
    root.geometry("600x400")

    style = ttk.Style()
    style.configure("TButton", font=("Helvetica", 12), padding=10)
    style.configure("TLabel", font=("Helvetica", 8), padding=5)

    open_button = ttk.Button(root, text="Open HTML File", command=open_file)
    open_button.pack(pady=10)

    selected_file_label = ttk.Label(root, text="Selected HTML file:")
    selected_file_label.pack(pady=10)

    start_button = ttk.Button(root, text="Start", command=start_parsing)
    start_button.pack(pady=10)

    processing_label = ttk.Label(root, text="", foreground="blue", font=("Helvetica", 14, "bold"))
    processing_label.pack(pady=10)

    root.mainloop()

# Entry point of the script
if __name__ == "__main__":
    # Check and install missing packages before creating the GUI
    install_missing_packages(required_packages)

    # Now create and run the GUI
    create_gui()