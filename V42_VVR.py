# Install the required libraries before running:
# pip install pandas openpyxl PyMuPDF pydrive
# V4 GUI cleaning and Code cleanup

import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import time
from pathlib import Path
from datetime import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import fitz  # PyMuPDF
from io import BytesIO
from PIL import Image
from tqdm import tqdm  # Import the tqdm library
import sys

start_time = time.time()
# Replace with your actual Gemini API key
GOOGLE_API_KEY = "-"
print(">>>>> VV RECEIPTS <<<<<")
print("Program has Started...")
print("Please wait for file selection prompt...")

# Configure the Generative AI client
import google.generativeai as genai
genai.configure(api_key=GOOGLE_API_KEY)

# Gemini model configuration
MODEL_CONFIG = {
    "temperature": 0.2,
    "top_p": 1,
    "top_k": 32,
    #"max_output_tokens": 4096,
}

# Safety settings
safety_settings = [
    {
        "category": "HARM_CATEGORY_HARASSMENT",
        "threshold": "BLOCK_MEDIUM_AND_ABOVE",
    },
    {
        "category": "HARM_CATEGORY_HATE_SPEECH",
        "threshold": "BLOCK_MEDIUM_AND_ABOVE",
    },
    {
        "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
        "threshold": "BLOCK_MEDIUM_AND_ABOVE",
    },
    {
        "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
        "threshold": "BLOCK_MEDIUM_AND_ABOVE",
    },
]
def get_client_secrets_path():
    """
    Returns the correct path to client_secrets.json based on whether the script is running as a bundled executable or not.
    """
    if getattr(sys, 'frozen', False):
        # Running as a bundled executable (PyInstaller/PyArmor)
        base_dir = sys._MEIPASS  # Temporary extraction directory
    else:
        # Running as a regular script
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_dir, 'client_secrets.json')

# Load Gemini model
model = genai.GenerativeModel(
    model_name="gemini-2.0-flash",
    generation_config=MODEL_CONFIG,
    safety_settings=safety_settings,
)

# Authenticate and create the PyDrive client
gauth = GoogleAuth()
client_secrets_path = get_client_secrets_path()  # Get the correct path
gauth.LoadClientConfigFile(client_secrets_path)  # Load the config file
gauth.LocalWebserverAuth()  # Creates local webserver and auto handles authentication.
drive = GoogleDrive(gauth)

def clear_console():
    """Clears the console screen."""
    time.sleep(0.2)
    os.system('cls' if os.name == 'nt' else 'clear')

# Function to list all folders in Google Drive
def list_folders(drive, folder_id=None):
    query = "mimeType='application/vnd.google-apps.folder' and trashed=false"
    if folder_id:
        query += f" and '{folder_id}' in parents"
    folder_list = drive.ListFile({'q': query}).GetList()
    return folder_list

# Function to list PDFs in a specific folder
def list_pdfs_in_folder(drive, folder_id):
    query = f"'{folder_id}' in parents and trashed=false and mimeType='application/pdf'"
    file_list = drive.ListFile({'q': query}).GetList()
    return file_list

# Function to navigate folders and select PDFs
def navigate_and_select_pdfs(drive):
    selected_pdfs = []
    folder_stack = []  # Stack to keep track of folder navigation
    current_folder_id = None  # Start at the root folder

    while True:
        # List folders and PDFs in the current folder
        folders = list_folders(drive, current_folder_id)
        pdf_files = list_pdfs_in_folder(drive, current_folder_id) if current_folder_id else []

        clear_console()
        print(">>>>> VV RECEIPTS <<<<<")
        
        print("\nCurrent Folder Contents:")
        if folder_stack:  # If not in the root folder
            print("0. Go back to previous folder")
        else:
            print("(Root Folder)")

        # Display folders
        for i, folder in enumerate(folders):
            print(f"{i + 1}. {folder['title']} (Folder)")

        # Display PDFs
        for i, pdf in enumerate(pdf_files):
            print(f"{len(folders) + i + 1}. {pdf['title']} (PDF)")

        # Get user input and display total number of PDFs chosen
        print(f"Total Number of PDFs chosen: {len(selected_pdfs)}")
        selection = input("Enter the number of the item you want to select, 'X' to go back, or 'C' to confirm: ").strip().lower()

        # Handle 'X' input (go back)
        if selection == 'x':
            if not folder_stack:  # If already at the root folder
                print("Already at the root folder.")
            else:
                # Go back to the previous folder
                current_folder_id = folder_stack.pop()
                print("Proceeding to previous folder...")
            continue

        # Handle '0' to go back if inside a subfolder
        if selection == '0' and folder_stack:
            current_folder_id = folder_stack.pop()
            print("Proceeding to previous folder...")
            continue

        # Handle 'C' input (confirm selection)
        if selection == 'c':
            if not selected_pdfs:
                print("No PDFs selected. Please select PDFs first.")
                continue

            # Display selected PDFs
            print("\nSelected PDFs:")
            for i, pdf in enumerate(selected_pdfs):
                print(f"{i + 1}. {pdf['title']}")

            # Ask for confirmation to proceed
            proceed = input("Do you want to proceed with the conversion? (yes/no): ").strip().lower()
            if proceed == "yes":
                return selected_pdfs  # Return selected PDFs for processing
            else:
                # Reset selected PDFs and return to the root folder
                selected_pdfs = []  # Reset selected PDFs
                folder_stack = []  # Clear folder stack
                current_folder_id = None  # Return to root folder
                print("Returning to root folder...")
                continue  # Reloop to allow new selections

        try:
            # Parse user input into indices
            selected_indices = [int(i.strip()) - 1 for i in selection.split(",")]

            # Track selected folders and PDFs
            selected_folders = []
            new_selected_pdfs = []  # Temporary list to store newly selected PDFs
            for index in selected_indices:
                if 0 <= index < len(folders):
                    selected_folders.append(folders[index])
                elif len(folders) <= index < len(folders) + len(pdf_files):
                    new_selected_pdfs.append(pdf_files[index - len(folders)])
                else:
                    print(f"Invalid selection: {index + 1}. Skipping.")

            # Handle folder selection
            if len(selected_folders) > 1:
                print("Error: You can only select one folder at a time.")
            elif len(selected_folders) == 1:
                # Navigate into the selected folder
                selected_folder = selected_folders[0]
                print(f"Navigating to folder: {selected_folder['title']}")
                folder_stack.append(current_folder_id)  # Save current folder to stack
                current_folder_id = selected_folder['id']  # Update current folder

            # Add newly selected PDFs to the main list
            if new_selected_pdfs:
                selected_pdfs.extend(new_selected_pdfs)
                print(f"Number of PDFs chosen: {len(new_selected_pdfs)}")

        except ValueError:
            print("Invalid input. Please enter numbers separated by commas, 'X' to go back, or 'C' to confirm.")
            continue

    return selected_pdfs

# Start navigating from the root folder
selected_pdfs = navigate_and_select_pdfs(drive)

if not selected_pdfs:
    print("No PDFs selected. Exiting.")
    exit()

# Proceed with conversion for selected PDFs
print("Proceeding with conversion...")

# Function to format the image input to excel ------------------------------------------------------

def standardize_date(date_str):
    # List of possible date formats in the input
    possible_formats = [ # Changed way of typing from V3.4
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d %m %Y", "%d/%m/%y", "%d-%m-%y", "%d.%m.%y", "%d %m %y",
        "%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y", "%m %d %Y", "%m/%d/%y", "%m-%d-%y", "%m.%d.%y", "%m %d %y",
        "%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d", "%Y %m %d", "%Y%m%d", "%y/%m/%d", "%y-%m-%d", "%y.%m.%d", "%y %m %d",
        "%d %B %Y", "%B %d, %Y", "%d-%b-%Y", "%d %b %Y", "%b %d, %Y", "%Y, %B %d", "%Y %b %d", "%d %B %y", "%B %d, %y",
        "%d-%b-%y", "%d %b %y", "%b %d, %y", "%d%m%Y", "%m%d%Y", "%Y%m%d", "%d%m%y", "%m%d%y", "%y%m%d",
        "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%d.%m.%Y %H:%M:%S", "%d %m %Y %H:%M:%S", "%d/%m/%y %H:%M:%S",
        "%d-%m-%y %H:%M:%S", "%d.%m.%y %H:%M:%S", "%d %m %y %H:%M:%S", "%m/%d/%Y %H:%M:%S", "%m-%d-%Y %H:%M:%S",
        "%m.%d.%Y %H:%M:%S", "%m %d %Y %H:%M:%S", "%m/%d/%y %H:%M:%S", "%m-%d-%y %H:%M:%S", "%m.%d.%y %H:%M:%S",
        "%m %d %y %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y.%m.%d %H:%M:%S", "%Y %m %d %H:%M:%S",
        "%y/%m/%d %H:%M:%S", "%y-%m-%d %H:%M:%S", "%y.%m.%d %H:%M:%S", "%y %m %d %H:%M:%S", "%d/%m/%Y %I:%M:%S %p",
        "%d-%m-%Y %I:%M:%S %p", "%d.%m.%Y %I:%M:%S %p", "%d %m %Y %I:%M:%S %p", "%d/%m/%y %I:%M:%S %p", "%d-%m-%y %I:%M:%S %p",
        "%d.%m.%y %I:%M:%S %p", "%d %m %y %I:%M:%S %p", "%m/%d/%Y %I:%M:%S %p", "%m-%d-%Y %I:%M:%S %p", "%m.%d.%Y %I:%M:%S %p",
        "%m %d %Y %I:%M:%S %p", "%m/%d/%y %I:%M:%S %p", "%m-%d-%y %I:%M:%S %p", "%m.%d.%y %I:%M:%S %p", "%m %d %y %I:%M:%S %p",
        "%Y/%m/%d %I:%M:%S %p", "%Y-%m-%d %I:%M:%S %p", "%Y.%m.%d %I:%M:%S %p", "%Y %m %d %I:%M:%S %p", "%y/%m/%d %I:%M:%S %p",
        "%y-%m-%d %I:%M:%S %p", "%y.%m.%d %I:%M:%S %p", "%y %m %d %I:%M:%S %p", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%SZ",
        "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%dT%H:%M:%S.%f%z", "%dth %B %Y",
        "%B %dth, %Y", "%dth-%m-%Y", "%dth/%m/%Y", "%Y年%m月%d日", "%d/%m/%Y (%A)", "%Y-%m-%dT%H:%M:%S%z", "%dth %B %Y",
        "%B %dth, %Y", "%dth-%m-%Y", "%dth/%m/%Y"
    ]
    
    for fmt in possible_formats:
        try:
            # Attempt to parse the date
            return datetime.strptime(date_str, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass  # Continue trying other formats

    # If no format matches, return the original string or handle the error
    return "Invalid date format"

data = []
data.append({"Date": None, "Store Name": None, "Item": None, "Cost": None, "Total": None}) # V3.8

# Loop through the files, download PDFs, and convert them to images in memory | From V3.5
for file in tqdm(selected_pdfs, desc="Processing PDFs", unit="file"):
    print(f"\nProcessing file: {file['title']}")
    
    # Download the file content to a BytesIO object
    file_content = BytesIO()
    file.GetContentFile("temp.pdf")  # Download the file locally
    with open("temp.pdf", "rb") as f:
        file_content.write(f.read())  # Read it as bytes and store in BytesIO
    file_content.seek(0)  # Reset pointer to the beginning
    os.remove("temp.pdf")

    # Convert PDF to images using PyMuPDF
    try:
        pdf_document = fitz.open(stream=file_content.read(), filetype="pdf")  # Open the PDF file from memory
        for page_number in range(len(pdf_document)):
            page = pdf_document.load_page(page_number)  # Load the page
            pix = page.get_pixmap()  # Render page to an image
            img_bytes = pix.tobytes()  # Get the image bytes
            img = Image.open(BytesIO(img_bytes))  # Open the image using PIL
            
            # Process the image directly without saving
            def image_format(img):
                img_byte_arr = BytesIO()
                img.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                return [
                    {
                        "mime_type": "image/png",
                        "data": img_byte_arr,
                    }
                ]

            # Function to process Gemini output
            def gemini_output(image_path, system_prompt, user_prompt):
                image_info = image_format(image_path)
                input_prompt = [system_prompt, image_info[0], user_prompt]
                response = model.generate_content(input_prompt)
                return response.text

            # System and user prompts
            system_prompt = """
            You are a specialist in comprehending receipts.
            Input images in the form of receipts will be provided to you,
            and your task is to respond to questions based on the content of the input image.
            """

            user_prompt = """
            Display the date, the title, the items with final cost, total, rounding total (if exist)
            Do it in this format:
            Date: 
            Title: 
            Items with Final Cost:
            * Item Name: Cost
            * Item Name: Cost
            Total: 
            Rounding:
            """

            # Run the model and get unstructured text
            output_text = gemini_output(img, system_prompt, user_prompt)
            # if Date or Store exist, Value(0) + 1
            # after for loop, Value2 + Value
            # Parse the unstructured text
            Date = ""
            Store = ""
            TTL = "0.0"
            Rnd = "0.0"
            item = ""
            cost = ""
            Value1 = 0 # For Date and Store | V3.7
            for line in output_text.strip().split("\n"):
                if "Date" in line:
                    Date = line.split(":", 1)[1]
                    Date = Date.replace("**", "").strip()
                    if "/" in Date and len(Date) == 15 and Date.count("/") != 2:
                        Date = Date.split("/")[0]
                    if all(exclude not in Date for exclude in ["Not available", "N/A", "Not found on the receipt"]): # V3.7
                        Value1 += 1
                    Date = standardize_date(Date)      
                if "Store" in line or "Title" in line or "Restaurant" in line:
                    Store = line.split(":", 1)[1]
                    Store = Store.replace("**", "").strip()
                    if all(exclude not in Store for exclude in ["Not available", "N/A", "Not found on the receipt", "(Not explicitly stated, but based on the items, it's likely a grocery store receipt)"]): # V3.7
                        Value1 += 1
            if Value1 == 2:
                data.append({"Date": Date, "Store Name": Store})

            for line in output_text.strip().split("\n"):
                if ": " in line and all(exclude not in line for exclude in ["Total", "Rounding", "Title", "Date"]):  
                    item, cost = line.split(": ", 1)  
                    item = item.replace("*", "")  
                    cost = cost.replace("*", "")  
                    if ":" in cost:
                        parts = cost.split(":")
                        item += f": {parts[0].strip()}"  # Append the first half to item
                        cost = parts[1].strip()  # Keep the second half as cost
                    data.append({"Item": item.strip(), "Cost": cost.strip()})  

            # Method 2 delete and remake entry
            for line in reversed(output_text.strip().split("\n")):
                if "Total" in line:
                    TTL = line.split(":", 1)[1]
                    TTL = TTL.replace("**", "").strip()
                if "Rounding" in line:
                    Rnd = line.split(":", 1)[1]
                    Rnd = Rnd.replace("**", "").strip()
            try: #OMAIGOD IT TOOK ME 2+ HOURS TO UNDERSTAND
                Rnd = f"{float(Rnd):.2f}"
            except ValueError: # If Rnd is not a number 
                Rnd = 0.0  
            try: #OMAIGOD IT TOOK ME 2+ HOURS TO UNDERSTAND
                TTL = f"{float(TTL):.2f}" 
            except ValueError: # If TTL is not a number 
                TTL = 0.0
            if float(Rnd) <= 0.10:
                Rnd = f"{float(TTL) + float(Rnd):.2f}"
            if TTL != 0.0 and Rnd != 0.0: # V3.7
                TrueTotal = Rnd # V3.8
                data.append({"Total": TrueTotal})

        pdf_document.close()  # Close the PDF document | new from V3.5
    except Exception as e:
        print(f"Error converting {file['title']}: {e}")

# print("PDF to image conversion complete!")

# Convert the structured data into a DataFrame
df = pd.DataFrame(data)

# Save to an Excel file
print("=======================================================")
ExcelName = input("Select a name for the excel file: ").strip()
output_excel_path = ExcelName + ".xlsx"
df.to_excel(output_excel_path, index=False, engine="openpyxl")

print(df)
print(f"Data has been written to {output_excel_path}")
print("--- Process took %s seconds ---" % round((time.time() - start_time), 3))