import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import pandas as pd
import re
import os
import numpy 
import numpy as np
import shutil

# Update this to your Tesseract installation path
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"


def convert_to_csv_safe(img_folder, output_folder='CSV', retries=3):
    # Ensure the main output 'csv' folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Loop through each folder inside the IMG directory
    for folder_name in os.listdir(img_folder):
        folder_path = os.path.join(img_folder, folder_name)

        # Check if it's a directory (not a file)
        if os.path.isdir(folder_path):
            # Create a corresponding output subfolder in the 'csv' folder
            csv_subfolder = os.path.join(output_folder, folder_name)
            if not os.path.exists(csv_subfolder):
                os.makedirs(csv_subfolder)

            # Loop through PNG files in the current folder
            for file_name in os.listdir(folder_path):
                if file_name.endswith('.png'):
                    file_path = os.path.join(folder_path, file_name)

                   

def save_images(pdf_folder, output_folder='IMG'):
    # Ensure the main output 'img' folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Loop through all files in the PDF folder
    for filename in os.listdir(pdf_folder):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(pdf_folder, filename)
            
            # Create a subfolder for each PDF inside the 'img' folder
            pdf_name = os.path.splitext(filename)[0]
            pdf_output_folder = os.path.join(output_folder, pdf_name)
            
            if not os.path.exists(pdf_output_folder):
                os.makedirs(pdf_output_folder)
            
            # Convert PDF to images
            images = convert_from_path(pdf_path)
            
            # Save each page as an image in the PDF-named folder
            for i, image in enumerate(images):
                image_name = f"page_{i+1}.png"
                image.save(os.path.join(pdf_output_folder, image_name), 'PNG')

   
def clean_flight_code(flight_code):
    pattern = re.compile(r'^[A-Za-z0-9]{2}\s\d+$')
    # Ensure flight_code is a string and remove leading/trailing spaces
    flight_code = str(flight_code).strip()
    # If it matches the pattern with a space, return as is
    if pattern.match(flight_code):
        return flight_code

    # If it's a string with two alphanumeric characters directly followed by numbers, fix it
    elif re.match(r'^[A-Za-z0-9]{2}\d+$', flight_code):
        return f"{flight_code[:2]} {flight_code[2:]}"

    # Otherwise, return None
    else:
        return numpy.nan


def check_3_letter_code(value):
    # Ensure the value is a string and check if it's exactly 3 letters
    if isinstance(value, str) and len(value) == 3 and value.isalpha():
        return value.upper()  # Return in uppercase for consistency
    return numpy.nan

def clean_time_format(value):
    # Convert the value to string and strip spaces
    value_str = str(value).strip()

    # Handle cases like '09:45 | ETA' - extract valid time part
    if '|' in value_str:
        value_str = value_str.split('|')[0].strip()

    # Check for valid time format (HH:MM or H:MM)
    match = re.match(r'^(\d{1,2}):(\d{2})$', value_str)

    if match:
        hours, minutes = match.groups()
        # Check if hours are between 0 and 23 and minutes between 0 and 59
        if 0 <= int(hours) <= 23 and 0 <= int(minutes) <= 59:
            return value_str
    return numpy.nan

def extract_8_digits(prn_value):
    """
    Extracts the first 8 digits from a PRN value.
    :param prn_value: The PRN value as a string (e.g., '15029674/A').
    :return: The 8-digit number as a string or NaN if invalid.
    """
    if pd.isna(prn_value) or not isinstance(prn_value, str):
        return np.nan
    match = re.match(r'^(\d{8})', prn_value.strip())
    return match.group(1) if match else np.nan

def clean_name(value):
    value_str = str(value).strip('.')  # Remove leading and trailing dots
    # Check if it's a valid name (non-numeric and only alphabetic characters)
    if value_str.isalpha():
      if value_str=="Form" or value_str=="nan" or value_str=="From" or value_str=="Type" or value_str=="Data":
        return numpy.nan
      else:
        return value_str
    else:
        return numpy.nan
    


def clean_name_columns(df, columns_to_clean):

    def clean_name(raw_name):
        if pd.isna(raw_name):
            return raw_name
        # Remove leading dots and spaces
        cleaned_name = str(raw_name).lstrip(". ").strip()
        # Remove prefixes like 'J ..'
        cleaned_name = re.sub(r"^[A-Za-z]\s*\.\.", "", cleaned_name).strip()
        return cleaned_name

    # Apply cleaning to each specified column
    for column in columns_to_clean:
        if column in df.columns:
            df[column] = df[column].apply(clean_name)
        else:
            print(f"Warning: Column '{column}' not found in DataFrame.")
    
    return df

def clean_date_column(df):
    """
    Cleans the 'Date' column in a DataFrame.
    Converts dates from DD.MM.YYYY to MM/DD/YYYY format and skips rows where the date is already in MM/DD/YYYY format.
    Replaces invalid dates with the last valid date.
    
    Args:
        df (pd.DataFrame): Input DataFrame containing a 'Date' column.
    
    Returns:
        pd.DataFrame: DataFrame with the cleaned 'Date' column.
    """
    previous_date = None  # To track the last valid date

    def clean_date(value):
        nonlocal previous_date

        # Skip if value is already in MM/DD/YYYY format
        if re.match(r'^\d{2}/\d{2}/\d{4}$', str(value)):
            previous_date = value  # Update previous_date since it's valid
            return value

        if pd.isna(value) or value in ['Not found', '', None]:
            return previous_date  # Use the last valid date if available

        # Ensure the value is a string and strip spaces
        value = str(value).strip()

        # Match the format DD.MM.YYYY
        match = re.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{4})$', value)
        if match:
            # Extract day, month, year and rearrange to MM/DD/YYYY
            day, month, year = match.groups()
            formatted_date = f"{month.zfill(2)}/{day.zfill(2)}/{year}"
            previous_date = formatted_date  # Update the last valid date
            return formatted_date

        # If parsing fails, return the previous valid date
        return previous_date if previous_date else np.nan

    # Apply the cleaning function to the 'Date' column
    df['Date'] = df['Date'].apply(clean_date)
    return df



def extract_text_from_image(image_path):
    """
    Extract text from an image using pytesseract.
    :param image_path: Path to the image file.
    :return: Extracted text or an error message.
    """
    try:
        # Open the image
        img = Image.open(image_path)
        # Use pytesseract to extract text
        text = pytesseract.image_to_string(img)
        return text
    except Exception as e:
        return f"Error: {e}"

def format_extracted_text(raw_text):
    """
    Format extracted text by removing empty lines and unnecessary spaces.
    :param raw_text: Raw text extracted from the image.
    :return: Formatted text as a single string.
    """
    lines = raw_text.splitlines()
    formatted_output = [line.strip() for line in lines if line.strip()]  # Remove empty lines
    return "\n".join(formatted_output)

def extract_data_using_patterns(formatted_text, patterns):
    """
    Extract specific data points from the formatted text using regex patterns.
    :param formatted_text: The formatted text.
    :param patterns: A dictionary of field names and their regex patterns.
    :return: A dictionary of extracted data.
    """
    extracted_data = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, formatted_text)
        extracted_data[key] = match.group(1) if match else "Not found"
    return extracted_data

def main(img_folder):
    """
    Main function to process the image and extract data points.
    :param image_path: Path to the image file.
    """

    # Loop through each folder inside the IMG directory
    for folder_name in os.listdir(img_folder):
        folder_path = os.path.join(img_folder, folder_name)

        # Check if it's a directory (not a file)
        if os.path.isdir(folder_path):
            # Loop through PNG files in the current folder
            for file_name in os.listdir(folder_path):
                if file_name.endswith('.png'):
                    file_path = os.path.join(folder_path, file_name)
                    raw_text = extract_text_from_image(file_path)
                    # Format the extracted text
                    formatted_text = format_extracted_text(raw_text)
                    formatted_text = formatted_text.replace("|", "")
                    print(formatted_text)
                    # Define regex patterns for data extraction
                    patterns = {
                        "Date": r"Date:\s*(\d{2}\.\d{2}\.\d{4})",
                        # "Flight Arrival": r"Flight Arrival:\s*([A-Z0-9]{2,3} \d+)",  # Support for D3, SV, etc.
                        # "Flight Departure": r"Flight Departure:\s*([A-Z0-9]{2,3} \d+)",  # Adjusted for D3 169
                        # "AC Type:": r"AC Type::\s*([^\s]+)",
                        # "From": r"From:\s*([^\s]+)",
                        # "To": r"To:\s*([^\s]+)",
                        "STA": r"STA:\s*(\d{2}:\d{2})",
                        "ETA": r"ETA:\s*(\d{2}:\d{2})",
                        "ATA": r"ATA:\s*(\d{2}:\d{2})",
                        "STD": r"STD:\s*(\d{2}:\d{2})",
                        "ETD": r"ETD:\s*(\d{2}:\d{2})?",
                        "ATD": r"ATD:\s*(\d{2}:\d{2})?",
                        # "ARR PRN": r"(?:COORDINATION SHEET / TIME CHART|TC/TOC:)\s*(\d{8}/[A-Z])",  # Match 8 digits with letter
                        # "DEP PRN": r"(?:COORDINATION SHEET / TIME CHART|TC/TOC:)\s*\d{8}/[A-Z]\s+(\d{8}/[A-Z])",
                        # "ARR NAME": r"TC/TOC:\s*[^\s]+ [^\s]+\s+([^\d]+)",
                        # "DEP NAME": r"TC/TOC:\s*[^\s]+ [^\s]+\s+[^\d]+\s+([^\d]+)",
                        "Blocks In": r"Blocks In\s*(\d{2}:\d{2})",
                        "Position PLB/Step": r"Position PLB/Step\s*(\d{2}:\d{2})",
                        "Open Door": r"Open Door\s*(\d{2}:\d{2})",
                        "Passenger Deplane Start": r"Passenger Deplane\s*(\d{2}:\d{2})",
                        "Passenger Deplane Finish": r"Passenger Deplane\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "Cabin Cleaning Start": r"Cabin Cleaning\s*(\d{2}:\d{2})",
                        "Cabin Cleaning Finish": r"Cabin Cleaning\s*\d{2}:\d{2}.*\s*(\d{2}:\d{2})",
                        "Customs Clearance Start":r"Customs Clearance\s*(\d{2}:\d{2})",
                        "Customs Clearance Finish":r"Customs Clearance\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "Cabin Cleaning start":r"Cabin Cleaning\s*(\d{2}:\d{2})",
                        "Cabin Cleaning Finish":r"Cabin Cleaning\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "Galley Services Start": r"Galley Services\s*(\d{2}:\d{2})",
                        "Galley Services Finish": r"Galley Services\s*\d{2}:\d{2}.*\s*(\d{2}:\d{2})",
                        "Cabin Security Check Start": r"Cabin Security Check\s*(\d{2}:\d{2})",
                        "Cabin Security Check Finish": r"Cabin Security Check\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "Boarding Clearance": r"Boarding Clearance\s*(\d{2}:\d{2})",
                        "Passengers Enplane Start": r"Passengers Enplane\s*(\d{2}:\d{2})",
                        "Passengers Enplane Finish": r"Passengers Enplane\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "TOP Finalization Start":r"TOP Finalization\s*(\d{2}:\d{2})",
                        "TOP Finalization Finish":r"TOP Finalization\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "FWD Unloading Start":r"FWD Unloading\s*(\d{2}:\d{2})",
                        "FWD Unloading Finish":r"FWD Unloading\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "FWD Leading Start":r"FWD Loading\s*(\d{2}:\d{2})",
                         "FWD Leading Finish":r"FWD Loading\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "AFT Unloading Start":r"AFT Unloading\s*(\d{2}:\d{2})",
                        "AFT Unloading Finish":r"AFT Unloading\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "AFT Loading Start":r"AFT Loading\s*(\d{2}:\d{2})",
                        "AFT Loading Finish":r"AFT Loading\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                         "Bulk Unloading Start":r"Bulk Unloading\s*(\d{2}:\d{2})",
                        "Bulk Unloading Finish":r"Bulk Unloading\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "Bulk Loading Start":r"Bulk Loading\s*(\d{2}:\d{2})",
                        "Bulk Loading Finish":r"Bulk Loading\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "GPU/ACU Support Start":r"GPU/ACU Support\s*(\d{2}:\d{2})",
                         "GPU/ACU Support Finish":r"GPU/ACU Support\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "Refueling Start":r"Refueling\s*(\d{2}:\d{2})",
                        "Refueling Finish":r"Refueling\s*\d{2}:\d{2}\s*(\d{2}:\d{2})",
                        "Remove PLB/Step":r"Remove PLB/Step\s*(\d{2}:\d{2})",
                        "Close Door": r"Close Door\s*(\d{2}:\d{2})",
                        "Pushback/Block-out": r"Pushback/Block-out\s*(\d{2}:\d{2})",

                    }


                    # Extract data
                    extracted_data = extract_data_using_patterns(formatted_text, patterns)
                    img = Image.open(file_path)
                    crop_box = (1395, 130, img.width, 200)
                    crop_box2 = (1195, 130,1400 , 200)
                    crop_box3 = (250, 190,400 , 250)
                    crop_box4 = (250, 250,400 , 300)
                    crop_box5 = (650, 190,800 , 250)
                    crop_box6 = (650, 250,800 , 300)
                    crop_box7 = (650, 150,800 , 190)

                    cropped_img = img.crop(crop_box)
                    cropped_img.save("cropped_image.png")

                    cropped_img2 = img.crop(crop_box2)
                    cropped_img2.save("cropped_image2.png")

                    cropped_img3 = img.crop(crop_box3)
                    cropped_img3.save("cropped_image3.png")

                    cropped_img4 = img.crop(crop_box4)
                    cropped_img4.save("cropped_image4.png")

                    cropped_img5 = img.crop(crop_box5)
                    cropped_img5.save("cropped_image5.png")

                    cropped_img6 = img.crop(crop_box6)
                    cropped_img6.save("cropped_image6.png")

                    cropped_img7 = img.crop(crop_box7)
                    cropped_img7.save("cropped_image7.png")

                    dep=main2('cropped_image.png')
                    arr=main3('cropped_image2.png')
                    flight_arr=main6("Flight Arrival",'cropped_image3.png')
                    flight_dep=main6("Flight Departure",'cropped_image4.png')
                    from_=main6("From",'cropped_image5.png')
                    to_=main6("To",'cropped_image6.png')
                    ac_=main6("AC Type:",'cropped_image7.png')

                    extracted_data.update(arr)
                    extracted_data.update(dep)
                    extracted_data.update(flight_arr)
                    extracted_data.update(flight_dep)
                    extracted_data.update(from_)
                    extracted_data.update(to_)
                    extracted_data.update(ac_)
                
                    extracted_data['Station'] = "RUH"
                    print(extracted_data)
                    df = pd.read_csv('CRS - RUH copy.csv')
                    new_row_aligned = {col: extracted_data.get(col, None) for col in df.columns}
                    # Append the new row as a DataFrame
                    new_row_df = pd.DataFrame([new_row_aligned])

                    # Append to the original DataFrame
                    df = pd.concat([df, new_row_df], ignore_index=True)
                    # Save the updated DataFrame back to the CSV file

                    df['Flight Arrival'] = df['Flight Arrival'].apply(clean_flight_code)
                    df['Flight Departure'] = df['Flight Departure'].apply(clean_flight_code)
                    # Apply the function to From and To columns
                    df['From'] = df['From'].apply(check_3_letter_code)
                    df['To'] = df['To'].apply(check_3_letter_code)
                    df['AC Type:'] = df['AC Type:'].str.replace(r'(From:|To:|nan,4|TW|TIA|FALSE|TIZ|LAAN|RANI|TIC|aD)', '', regex=True).str.strip()
                    columns_to_clean = ['STA', 'ETA', 'ATA', 'STD', 'ETD', 'ATD','Blocks In',
                        'Position PLB/Step', 'Open Door', 'Passenger Deplane Start',
                        'Passenger Deplane Finish', 'Customs Clearance Start',
                        'Customs Clearance Finish', 'Cabin Cleaning start',
                        'Cabin Cleaning Finish', 'Galley Services Start',
                        'Galley Services Finish', 'Cabin Security Check Start',
                        'Cabin Security Check Finish', 'Boarding Clearance',
                        'Passengers Enplane Start', 'Passengers Enplane Finish',
                        'TOP Finalization Start', 'TOP Finalization Finish',
                        'FWD Unloading Start', 'FWD Unloading Finish', 'FWD Leading Start',
                        'FWD Leading Finish', 'AFT Unloading Start', 'AFT Unloading Finish',
                        'AFT Loading Start', 'AFT Loading Finish', 'Bulk Unloading Start',
                        'Bulk Unloading Finish', 'Bulk Loading Start', 'Bulk Loading Finish',
                        'GPU/ACU Support Start', 'GPU/ACU Support Finish', 'Refueling Start',
                        'Refueling Finish', 'Close Door', 'Remove PLB/Step',
                        'Pushback/Block-out']

                    # Apply the function to all the specified columns
                    for column in columns_to_clean:
                        df[column] = df[column].apply(clean_time_format)

                   
                    df = clean_date_column(df)
                    df.replace("Not found", "", inplace=True)
                    df = clean_name_columns(df, ['ARR NAME', 'DEP NAME'])
                    df = df.drop_duplicates()
                    # df = df.dropna(subset=['Flight Arrival', 'Flight Departure'], how='all')
                
                    df.to_csv('CRS - RUH copy.csv', index=False)

                   


def main2(image_path):
    """
    Main function to process the image and extract data points.
    :param image_path: Path to the image file.
    """
    # Extract text from the image
    raw_text = extract_text_from_image(image_path)
    if raw_text.startswith("Error:"):
        print(raw_text)
        return

    # Format the extracted text
    formatted_text = format_extracted_text(raw_text)
    formatted_text = formatted_text.replace("|", "")
    print(formatted_text)
    # Define regex patterns for data extraction
    patterns = {
    "DEP PRN": r"(\d{8})/",  # Match only the 8 digits before the slash
    "DEP NAME": r"\d{8}/[A-Z]\.\s*([A-Za-z][A-Za-z\-\.\s]+)"
}

            

    # Extract data
    extracted_data = extract_data_using_patterns(formatted_text, patterns)
    return extracted_data

def main3(image_path):
    """
    Main function to process the image and extract data points.
    :param image_path: Path to the image file.
    """
    # Extract text from the image
    raw_text = extract_text_from_image(image_path)
    if raw_text.startswith("Error:"):
        print(raw_text)
        return

    # Format the extracted text
    formatted_text = format_extracted_text(raw_text)
    formatted_text = formatted_text.replace("|", "")
    print(formatted_text)
    # Define regex patterns for data extraction
    patterns = {
    "ARR PRN": r"(\d{8})/",  # Match only the 8 digits before the slash
    "ARR NAME": r"\d{8}/[A-Z]\.\s*([A-Za-z][A-Za-z\-\.\s]+)",  # Match the name after ARR PRN
}

    
    # Extract data
    extracted_data = extract_data_using_patterns(formatted_text, patterns)
    return extracted_data




def main6(text, image_path):
    """
    Main function to process the image and extract data points.
    :param text: Key for the extracted data.
    :param image_path: Path to the image file.
    :return: Dictionary with extracted data or default value if an error occurs.
    """
    # Extract text from the image
    raw_text = extract_text_from_image(image_path)
    
    # Check for errors in text extraction
    if not raw_text or raw_text.startswith("Error:"):
        print(f"Error extracting text or no data found: {raw_text}")
        return {text: "Not found"}  # Return default value in case of error

    # Format the extracted text
    formatted_text = format_extracted_text(raw_text)
    if not formatted_text:  # Handle cases where formatting fails
        print("Warning: Text formatting failed. Returning default value.")
        return {text: "Not found"}

    # Clean up the formatted text
    formatted_text = formatted_text.replace("|", "")
    print(f"Formatted text: {formatted_text}")

    # Return extracted data
    extracted_data = {text: formatted_text}
    return extracted_data

# Run the main function
if __name__ == "__main__":    
    save_images('PDF')
    main('IMG')

    for filename in os.listdir('PDF'):
        shutil.move(f'PDF\\{filename}', 'PDFr')
    
    for filename in os.listdir('IMG'):
        shutil.move(f'IMG\\{filename}', 'IMGr')
